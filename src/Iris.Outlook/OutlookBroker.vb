Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' Implementação de <see cref="IOutlookBroker"/>.
    '''
    ''' Migrada do spike da Fase 0 componente a componente, não por cópia
    ''' integral. O que veio de lá e por quê:
    '''
    '''   • Thread STA dedicada com <c>Dispatcher.Run</c> — message pump de
    '''     verdade. Uma fila bloqueante processaria comandos e mataria a
    '''     entrega de eventos, que chegam como mensagens de janela.
    '''   • <c>IOleMessageFilter</c> com orçamento de retry, e o retry
    '''     DESLIGADO em toda mutação.
    '''   • <c>GuardResult</c> em tudo que sai da thread do broker.
    '''   • Liberação determinística com o pump ainda vivo no encerramento.
    '''
    ''' O que MUDOU em relação ao spike: <c>ReadAsync</c> e <c>MutateAsync</c>
    ''' são privados. A superfície pública é o contrato de operações
    ''' nomeadas, porque a assinatura genérica deixava um RCW escapar.
    ''' </summary>
    Public NotInheritable Class OutlookBroker
        Implements IOutlookBroker, IDisposable

        Private Const ProbeConnectedMs As Integer = 15000
        Private Const ProbeSearchingMs As Integer = 3000

        Private ReadOnly _log As ILog
        Private ReadOnly _ready As New ManualResetEventSlim(False)

        Private _thread As Thread
        Private _dispatcher As Dispatcher
        Private _filter As OutlookMessageFilter
        Private _watchdog As DispatcherTimer
        Private _startupError As Exception
        Private _disposed As Boolean

        ' Só podem ser tocados de dentro da thread do broker.
        Private _application As OL.Application
        Private _namespace As OL.NameSpace

        ' Inteiro, nao SessionState: e lido de outras threads, e
        ' Volatile.Read so opera sobre tipos que ele conhece. Passar o enum
        ' por CObj faria box e leria a COPIA, nao o campo.
        Private _state As Integer = CInt(SessionState.Disconnected)

        Public Event StateChanged As EventHandler(Of SessionState) Implements IOutlookBroker.StateChanged
        Public Event FolderInvalidated As EventHandler(Of FolderInvalidation) Implements IOutlookBroker.FolderInvalidated

        Public Sub New(log As ILog)
            _log = If(log, CType(New NullLog(), ILog))
        End Sub

        Public ReadOnly Property State As SessionState Implements IOutlookBroker.State
            Get
                Return CType(Volatile.Read(_state), SessionState)
            End Get
        End Property

        Public ReadOnly Property BrokerThreadId As Integer

        ' ===================================================================
        ' Ciclo de vida da thread
        ' ===================================================================

        Public Sub Start(Optional timeout As TimeSpan = Nothing)
            If timeout = Nothing Then timeout = TimeSpan.FromSeconds(10)

            _thread = New Thread(AddressOf ThreadBody) With {
                .Name = "IrisOutlookBroker",
                .IsBackground = True
            }
            _thread.SetApartmentState(ApartmentState.STA)
            _thread.Start()

            If Not _ready.Wait(timeout) Then
                Throw New TimeoutException($"Broker não subiu em {timeout.TotalSeconds:0}s.")
            End If
            If _startupError IsNot Nothing Then
                Throw New InvalidOperationException("Falha ao iniciar o broker.", _startupError)
            End If

            _log.Write(LogLevel.Info, "broker.start", $"thread={BrokerThreadId} STA")
        End Sub

        Private Sub ThreadBody()
            Try
                _dispatcher = Dispatcher.CurrentDispatcher
                _BrokerThreadId = Environment.CurrentManagedThreadId
                _filter = OutlookMessageFilter.Register()

                ' Vigia de reconexão: é o que faz "abrir o Outlook depois"
                ' conectar sozinho, sem o usuário reiniciar o Iris.
                _watchdog = New DispatcherTimer(DispatcherPriority.Background, _dispatcher) With {
                    .Interval = TimeSpan.FromMilliseconds(ProbeSearchingMs)
                }
                AddHandler _watchdog.Tick, AddressOf OnWatchdogTick
                _watchdog.Start()
            Catch ex As Exception
                _startupError = ex
                _ready.Set()
                Return
            End Try

            _ready.Set()

            ' O message pump.
            Dispatcher.Run()

            ' Rede de segurança: só roda se ninguém chamou Shutdown.
            Try
                ReleaseSessionCore()
                OutlookMessageFilter.Revoke()
            Catch
            End Try
        End Sub

        Private Sub OnWatchdogTick(sender As Object, e As EventArgs)
            Dim antes = State
            Dim agora = ProbeCore()
            _watchdog.Interval = TimeSpan.FromMilliseconds(
                If(agora = SessionState.Connected, ProbeConnectedMs, ProbeSearchingMs))
            If agora <> antes Then SetState(agora)
        End Sub

        Private Sub SetState(novo As SessionState)
            If Volatile.Read(_state) = CInt(novo) Then Return
            Volatile.Write(_state, CInt(novo))
            _log.Write(LogLevel.Info, "broker.state", novo.ToString())
            RaiseEvent StateChanged(Me, novo)
        End Sub

        ' ===================================================================
        ' Despacho — PRIVADO. Era público no spike, e por isso vazava RCW.
        ' ===================================================================

        Private Function InvokeAsync(Of T)(work As Func(Of T)) As Task(Of T)
            If _disposed Then Throw New ObjectDisposedException(NameOf(OutlookBroker))
            If _dispatcher Is Nothing Then Throw New InvalidOperationException("Broker não iniciado.")
            Return _dispatcher.InvokeAsync(Function() GuardResult(work())).Task
        End Function

        ''' <summary>Leitura idempotente: o message filter pode repetir.</summary>
        Private Async Function ReadAsync(Of T)(operation As String,
                                               work As Func(Of OL.Application, OL.NameSpace, OperationResult(Of T)),
                                               cancel As CancellationToken) As Task(Of OperationResult(Of T))
            Return Await RunAsync(operation, work, allowRetry:=True, cancel:=cancel)
        End Function

        ''' <summary>
        ''' Mutação: retry DESLIGADO. Criar, Save, Move, Delete e Send não
        ''' são idempotentes, e a Fase 0 pegou exatamente este erro — todas
        ''' as mutações do grupo D estavam rodando com retry ligado.
        ''' </summary>
        Private Async Function MutateAsync(Of T)(operation As String,
                                                 work As Func(Of OL.Application, OL.NameSpace, OperationResult(Of T)),
                                                 cancel As CancellationToken) As Task(Of OperationResult(Of T))
            Return Await RunAsync(operation, work, allowRetry:=False, cancel:=cancel)
        End Function

        Private Async Function RunAsync(Of T)(operation As String,
                                              work As Func(Of OL.Application, OL.NameSpace, OperationResult(Of T)),
                                              allowRetry As Boolean,
                                              cancel As CancellationToken) As Task(Of OperationResult(Of T))
            ' Cancelar antes de começar evita o trabalho. Depois que a
            ' chamada COM começou, cancelar só libera o CHAMADOR — a
            ' operação segue no broker.
            If cancel.IsCancellationRequested Then
                Return OperationResult(Of T).Fail(ErrorKind.Cancelled, "cancelado antes de iniciar")
            End If
            If State <> SessionState.Connected Then
                Return OperationResult(Of T).Fail(ErrorKind.NotConnected, State.ToString())
            End If

            Dim inicio = DateTime.UtcNow
            Try
                Dim resultado = Await InvokeAsync(
                    Function()
                        If _application Is Nothing OrElse _namespace Is Nothing Then
                            Return OperationResult(Of T).Fail(ErrorKind.NotConnected, "sem sessão")
                        End If
                        _filter.AllowRetry = allowRetry
                        Try
                            Return work(_application, _namespace)
                        Finally
                            _filter.AllowRetry = True
                        End Try
                    End Function)

                _log.Write(LogLevel.Debug, operation,
                           $"{resultado.Kind} em {(DateTime.UtcNow - inicio).TotalMilliseconds:0} ms")
                Return resultado

            Catch ex As COMException
                Dim kind = Classify(ex.HResult)
                _log.Write(LogLevel.Warn, operation, $"HRESULT 0x{ex.HResult:X8} -> {kind}")
                Return OperationResult(Of T).Fail(kind, $"0x{ex.HResult:X8}")
            Catch ex As Exception
                _log.Write(LogLevel.Error, operation, ex.GetType().Name)
                Return OperationResult(Of T).Fail(ErrorKind.Unexpected, ex.GetType().Name)
            End Try
        End Function

        ''' <summary>
        ''' Última barreira contra RCW atravessando a fronteira. A garantia
        ''' principal é estrutural — Model e Core nem miram Windows — e isto
        ''' é a rede de segurança.
        ''' </summary>
        Private Shared Function GuardResult(Of T)(value As T) As T
            If ComHelpers.ContainsComReference(value) Then
                Throw New InvalidOperationException(
                    "Um objeto COM tentou atravessar a fronteira do broker.")
            End If
            Return value
        End Function

        Private Shared Function Classify(hresult As Integer) As ErrorKind
            Select Case hresult
                Case &H80010001, &H8001010A : Return ErrorKind.Busy
                Case &H80010108, &H800706BA, &H800401FD : Return ErrorKind.NotConnected
                Case &H80070005 : Return ErrorKind.Denied
                Case Else : Return ErrorKind.Unexpected
            End Select
        End Function

        ' ===================================================================
        ' Sessão
        ' ===================================================================

        Public Async Function ConnectAsync(cancel As CancellationToken) As Task(Of SessionState) _
            Implements IOutlookBroker.ConnectAsync
            Dim resultado = Await InvokeAsync(Function() ConnectCore())
            SetState(resultado)
            Return resultado
        End Function

        Public Async Function ProbeAsync(cancel As CancellationToken) As Task(Of SessionState) _
            Implements IOutlookBroker.ProbeAsync
            Dim resultado = Await InvokeAsync(Function() ProbeCore())
            SetState(resultado)
            Return resultado
        End Function

        Private Function ConnectCore() As SessionState
            AssertOnBrokerThread()
            Try
                Dim attach = ComHelpers.GetRunningInstance("Outlook.Application")
                Select Case attach.Outcome
                    Case ComHelpers.AttachOutcome.Busy
                        Return SessionState.Busy
                    Case ComHelpers.AttachOutcome.NotRunning,
                         ComHelpers.AttachOutcome.NotRegistered,
                         ComHelpers.AttachOutcome.Failed
                        Return SessionState.Unavailable
                End Select

                Dim app = TryCast(attach.Instance, OL.Application)
                If app Is Nothing Then
                    ComHelpers.Release(attach.Instance)
                    Return SessionState.Unavailable
                End If

                ReleaseSessionCore()
                _application = app
                _namespace = app.GetNamespace("MAPI")
                Return SessionState.Connected

            Catch ex As COMException
                ReleaseSessionCore()
                Return SessionState.Busy
            End Try
        End Function

        Private Function ProbeCore() As SessionState
            AssertOnBrokerThread()
            If _application Is Nothing Then Return ConnectCore()

            Try
                Dim nome = _application.Name
                Return If(String.IsNullOrEmpty(nome), SessionState.Disconnected, SessionState.Connected)
            Catch ex As COMException
                ' Classificar pelo HRESULT: tratar toda COMException como
                ' morte derrubaria a sessão por uma recusa transitória.
                Select Case ex.HResult
                    Case &H80010108, &H800706BA, &H800401FD
                        ReleaseSessionCore()
                        Return ConnectCore()
                    Case &H80010001, &H8001010A
                        Return SessionState.Busy
                    Case Else
                        Return SessionState.Connected
                End Select
            End Try
        End Function

        ' ===================================================================
        ' Leitura
        ' ===================================================================

        Public Async Function GetStoresAsync(cancel As CancellationToken) _
            As Task(Of OperationResult(Of IReadOnlyList(Of StoreInfo))) Implements IOutlookBroker.GetStoresAsync

            Return Await ReadAsync(Of IReadOnlyList(Of StoreInfo))(
                "outlook.getStores",
                Function(app, ns)
                    Dim lista As New List(Of StoreInfo)()
                    Dim stores = ns.Stores
                    Try
                        For i = 1 To stores.Count
                            Dim bruto = stores.Item(i)
                            Dim store = TryCast(bruto, OL.Store)
                            If store Is Nothing Then
                                ComHelpers.Release(bruto)
                                Continue For
                            End If
                            Try
                                lista.Add(New StoreInfo With {
                                    .DisplayName = Safe(Function() store.DisplayName),
                                    .StoreId = Safe(Function() store.StoreID),
                                    .ExchangeStoreType = Safe(Function() store.ExchangeStoreType.ToString()),
                                    .IsCachedExchange = SafeBool(Function() store.IsCachedExchange),
                                    .RootFolder = RootFolderKey(store)
                                })
                            Finally
                                ComHelpers.Release(store)
                            End Try
                        Next
                    Finally
                        ComHelpers.Release(stores)
                    End Try
                    Return OperationResult(Of IReadOnlyList(Of StoreInfo)).Ok(lista)
                End Function, cancel)
        End Function

        Private Shared Function RootFolderKey(store As OL.Store) As FolderKey
            Dim raiz As OL.Folder = Nothing
            Try
                raiz = TryCast(store.GetRootFolder(), OL.Folder)
                If raiz Is Nothing Then Return New FolderKey("", "")
                Return New FolderKey(Safe(Function() raiz.EntryID), Safe(Function() raiz.StoreID))
            Catch
                Return New FolderKey("", "")
            Finally
                ComHelpers.Release(raiz)
            End Try
        End Function

        Public Async Function GetFolderChildrenAsync(parent As FolderKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of IReadOnlyList(Of FolderInfo))) Implements IOutlookBroker.GetFolderChildrenAsync

            Return Await ReadAsync(Of IReadOnlyList(Of FolderInfo))(
                "outlook.getFolderChildren",
                Function(app, ns)
                    Dim pai As OL.Folder = Nothing
                    Try
                        pai = TryCast(ns.GetFolderFromID(parent.EntryId, parent.StoreId), OL.Folder)
                    Catch ex As COMException
                        Return OperationResult(Of IReadOnlyList(Of FolderInfo)).Fail(
                            ErrorKind.NotFound, "pasta não encontrada")
                    End Try
                    If pai Is Nothing Then
                        Return OperationResult(Of IReadOnlyList(Of FolderInfo)).Fail(
                            ErrorKind.NotFound, "pasta não encontrada")
                    End If

                    Try
                        Dim lista As New List(Of FolderInfo)()
                        Dim filhos = pai.Folders
                        Try
                            For i = 1 To filhos.Count
                                Dim f = filhos.Item(i)
                                If f Is Nothing Then Continue For
                                Try
                                    Dim netos = 0
                                    Try : netos = f.Folders.Count : Catch : End Try
                                    lista.Add(New FolderInfo With {
                                        .Key = New FolderKey(Safe(Function() f.EntryID), Safe(Function() f.StoreID)),
                                        .Name = Safe(Function() f.Name),
                                        .DefaultItemType = Safe(Function() f.DefaultItemType.ToString()),
                                        .ItemCount = SafeInt(Function() f.Items.Count),
                                        .UnreadCount = SafeInt(Function() f.UnReadItemCount),
                                        .HasChildren = netos > 0
                                    })
                                Finally
                                    ComHelpers.Release(f)
                                End Try
                            Next
                        Finally
                            ComHelpers.Release(filhos)
                        End Try
                        Return OperationResult(Of IReadOnlyList(Of FolderInfo)).Ok(lista)
                    Finally
                        ComHelpers.Release(pai)
                    End Try
                End Function, cancel)
        End Function

        ' ===================================================================
        ' Ainda não implementadas — marcos 1.3 a 1.5
        ' ===================================================================

        Private Shared Function Pendente(Of T)(marco As String) As Task(Of OperationResult(Of T))
            Return Task.FromResult(OperationResult(Of T).Fail(
                ErrorKind.NotImplemented, $"previsto para o marco {marco}"))
        End Function

        Public Function GetMessagePageAsync(query As MessageQuery, offset As Integer, count As Integer,
                                            cancel As CancellationToken) As Task(Of OperationResult(Of MessagePage)) _
            Implements IOutlookBroker.GetMessagePageAsync
            Return Pendente(Of MessagePage)("1.3")
        End Function

        Public Function GetMessageDetailAsync(item As ItemKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of MessageDetail)) Implements IOutlookBroker.GetMessageDetailAsync
            Return Pendente(Of MessageDetail)("1.4")
        End Function

        Public Function SaveAttachmentAsync(attachment As AttachmentKey, destinationPath As String,
                                            overwrite As Boolean, cancel As CancellationToken) _
            As Task(Of OperationResult(Of String)) Implements IOutlookBroker.SaveAttachmentAsync
            Return Pendente(Of String)("1.4")
        End Function

        Public Function MarkReadAsync(item As ItemKey, isRead As Boolean, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.MarkReadAsync
            Return Pendente(Of Boolean)("1.4")
        End Function

        Public Function CreateDraftAsync(content As DraftContent, cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftKey)) Implements IOutlookBroker.CreateDraftAsync
            Return Pendente(Of DraftKey)("1.5")
        End Function

        Public Function CreateReplyDraftAsync(item As ItemKey, replyAll As Boolean, cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftKey)) Implements IOutlookBroker.CreateReplyDraftAsync
            Return Pendente(Of DraftKey)("1.5")
        End Function

        Public Function CreateForwardDraftAsync(item As ItemKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftKey)) Implements IOutlookBroker.CreateForwardDraftAsync
            Return Pendente(Of DraftKey)("1.5")
        End Function

        Public Function UpdateDraftAsync(draft As DraftKey, content As DraftContent, cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftKey)) Implements IOutlookBroker.UpdateDraftAsync
            Return Pendente(Of DraftKey)("1.5")
        End Function

        Public Function AddDraftAttachmentAsync(draft As DraftKey, filePath As String, cancel As CancellationToken) _
            As Task(Of OperationResult(Of AttachmentInfo)) Implements IOutlookBroker.AddDraftAttachmentAsync
            Return Pendente(Of AttachmentInfo)("1.5")
        End Function

        Public Function RemoveDraftAttachmentAsync(draft As DraftKey, attachment As AttachmentKey,
                                                   cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.RemoveDraftAttachmentAsync
            Return Pendente(Of Boolean)("1.5")
        End Function

        Public Function PrepareSendAsync(draft As DraftKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of SendPreview)) Implements IOutlookBroker.PrepareSendAsync
            Return Pendente(Of SendPreview)("1.5")
        End Function

        Public Function SendDraftAsync(draft As DraftKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.SendDraftAsync
            Return Pendente(Of Boolean)("1.5")
        End Function

        Public Function DeleteDraftAsync(draft As DraftKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.DeleteDraftAsync
            Return Pendente(Of Boolean)("1.5")
        End Function

        Public Function SubscribeFolderAsync(folder As FolderKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of SubscriptionToken)) Implements IOutlookBroker.SubscribeFolderAsync
            Return Pendente(Of SubscriptionToken)("1.3")
        End Function

        Public Function UnsubscribeFolderAsync(token As SubscriptionToken, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.UnsubscribeFolderAsync
            Return Pendente(Of Boolean)("1.3")
        End Function

        ' ===================================================================
        ' Encerramento
        ' ===================================================================

        Private Sub ReleaseSessionCore()
            ComHelpers.Release(_namespace)
            _namespace = Nothing
            ComHelpers.Release(_application)
            _application = Nothing
        End Sub

        Public Sub AssertOnBrokerThread()
            If Environment.CurrentManagedThreadId <> BrokerThreadId Then
                Throw New InvalidOperationException(
                    $"COM tocado fora da thread do broker (atual " &
                    $"{Environment.CurrentManagedThreadId}, broker {BrokerThreadId}).")
            End If
        End Sub

        ''' <summary>
        ''' A liberação acontece com o pump AINDA VIVO: desconectar sinks e
        ''' liberar proxies pode exigir processamento de mensagens.
        ''' </summary>
        Public Sub Shutdown(Optional timeout As TimeSpan = Nothing)
            If timeout = Nothing Then timeout = TimeSpan.FromSeconds(10)
            If _dispatcher Is Nothing Then Return

            Try
                _dispatcher.Invoke(
                    Sub()
                        If _watchdog IsNot Nothing Then
                            _watchdog.Stop()
                            RemoveHandler _watchdog.Tick, AddressOf OnWatchdogTick
                            _watchdog = Nothing
                        End If
                        ReleaseSessionCore()
                        OutlookMessageFilter.Revoke()
                    End Sub, DispatcherPriority.Send, Nothing, timeout)
            Catch
            End Try

            _dispatcher.InvokeShutdown()

            If _thread IsNot Nothing AndAlso Not _thread.Join(timeout) Then
                _log.Write(LogLevel.Error, "broker.shutdown", "thread não encerrou")
            End If
            _log.Write(LogLevel.Info, "broker.shutdown", "ok")
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            Try
                Shutdown()
            Catch
            End Try
            _ready.Dispose()
        End Sub

        ' ===================================================================
        ' Leitura defensiva de propriedades COM
        ' ===================================================================

        Private Shared Function Safe(getter As Func(Of String)) As String
            Try : Return If(getter(), "") : Catch : Return "" : End Try
        End Function

        Private Shared Function SafeInt(getter As Func(Of Integer)) As Integer
            Try : Return getter() : Catch : Return 0 : End Try
        End Function

        Private Shared Function SafeBool(getter As Func(Of Boolean)) As Boolean
            Try : Return getter() : Catch : Return False : End Try
        End Function

    End Class

End Namespace

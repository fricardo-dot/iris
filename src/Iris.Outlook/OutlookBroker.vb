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

        ''' <summary>
        ''' Sobe a cada aquisição de sessão COM. Ver SessionEpoch no contrato.
        ''' </summary>
        Private _epoca As Long = 0

        ''' <summary>
        ''' Probes seguidos que falharam sem classificação. Zera a cada probe
        ''' que dá certo. Existe para falha desconhecida não deixar o Iris
        ''' "ocupado" para sempre, preservando um RCW que talvez esteja morto.
        ''' </summary>
        Private _probesDesconhecidos As Integer = 0

        ' Só podem ser tocados de dentro da thread do broker.
        Private _application As OL.Application
        Private _namespace As OL.NameSpace

        ' Inteiro, nao SessionState: e lido de outras threads, e
        ' Volatile.Read so opera sobre tipos que ele conhece. Passar o enum
        ' por CObj faria box e leria a COPIA, nao o campo.
        Private _state As Integer = CInt(SessionState.Disconnected)

        ' 1 = o delegate da operacao ja comecou, logo pode ter surtido
        ' efeito. Serializado pela thread do broker.

        ''' <summary>
        ''' Assinaturas ativas, por id. Só podem ser tocadas na thread do
        ''' broker. Guardadas aqui porque a coleção Items precisa de
        ''' referência forte viva: se o GC a coletar, o event sink morre e os
        ''' eventos param sem erro nenhum (R7).
        ''' </summary>
        Private ReadOnly _subscriptions As New Dictionary(Of Integer, FolderSubscription)()

        Public Event StateChanged As EventHandler(Of SessionState) Implements IOutlookBroker.StateChanged
        Public Event FolderInvalidated As EventHandler(Of FolderInvalidation) Implements IOutlookBroker.FolderInvalidated

        Public Sub New(log As ILog)
            _log = If(log, CType(New NullLog(), ILog))
        End Sub

        Public Event SessionReplaced As EventHandler(Of Long) _
            Implements IOutlookBroker.SessionReplaced

        Public ReadOnly Property SessionEpoch As Long Implements IOutlookBroker.SessionEpoch
            Get
                Return Interlocked.Read(_epoca)
            End Get
        End Property

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

            ' Publica por ÚLTIMO. Ordem: adquirir, atualizar o estado, e só
            ' então anunciar.
            '
            ' Publicar antes do SetState era uma armadilha: o assinante
            ' enfileira o tratamento na UI, ele podia rodar antes de o
            ' estado virar Connected, encontrar a sessão como não conectada,
            ' limpar tudo e desistir de restaurar — e a mudança de estado
            ' seguinte fazia só a recarga comum, já sem a pasta que o
            ' usuário tinha aberto.
            PublicarSessaoSeHouver()
        End Sub

        ''' <summary>
        ''' Época adquirida e ainda não anunciada. Existe para separar
        ''' ADQUIRIR de PUBLICAR: ver o comentário em ConnectCore.
        ''' </summary>
        Private _epocaAPublicar As Long = 0

        ''' <summary>
        ''' Anuncia a sessão nova, se houver.
        '''
        ''' Cada assinante é chamado dentro do seu próprio Try. Um handler
        ''' que estoure não pode derrubar os outros nem fazer uma aquisição
        ''' que deu certo parecer que falhou — e assinante é código de fora,
        ''' então isso não é hipótese remota.
        '''
        ''' O que isto NÃO faz: não troca de thread. No caminho do watchdog
        ''' os handlers rodam na STA do broker, e um handler que chame o
        ''' broker de volta e ESPERE trava a STA. O contrato do evento diz
        ''' que handler devolve ao dispatcher dele e não bloqueia; isto aqui
        ''' protege contra exceção, não contra bloqueio.
        ''' </summary>
        Private Sub PublicarSessaoSeHouver()
            ' Interlocked: escrito na STA (dentro do ConnectCore) e lido
            ' aqui, que depois desta correção pode ser a thread do chamador.
            Dim epoca = Interlocked.Exchange(_epocaAPublicar, 0)
            If epoca = 0 Then Return

            Dim inscritos = SessionReplacedEvent
            If inscritos Is Nothing Then Return

            For Each alvo In inscritos.GetInvocationList()
                Try
                    CType(alvo, EventHandler(Of Long)).Invoke(Me, epoca)
                Catch ex As Exception
                    _log.Write(LogLevel.Error, "broker.session",
                               $"assinante falhou: {ex.GetType().Name}")
                End Try
            Next
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
            Return Await RunAsync(operation, work, allowRetry:=True, isMutation:=False, cancel:=cancel)
        End Function

        ''' <summary>
        ''' Mutação: retry DESLIGADO. Criar, Save, Move, Delete e Send não
        ''' são idempotentes, e a Fase 0 pegou exatamente este erro — todas
        ''' as mutações do grupo D estavam rodando com retry ligado.
        ''' </summary>
        ''' <summary>
        ''' Nem leitura nem mutação: opera sobre o item mas não deixa efeito
        ''' que sobreviva à chamada. Sem retry, porque repetir não é
        ''' obviamente seguro; sem Ambiguous, porque não há efeito que possa
        ''' ter acontecido pela metade.
        ''' </summary>
        Private Async Function SemRetryAsync(Of T)(operation As String,
                                                   work As Func(Of OL.Application, OL.NameSpace, OperationResult(Of T)),
                                                   cancel As CancellationToken) As Task(Of OperationResult(Of T))
            Return Await RunAsync(operation, work, allowRetry:=False, isMutation:=False, cancel:=cancel)
        End Function

        Private Async Function MutateAsync(Of T)(operation As String,
                                                 work As Func(Of OL.Application, OL.NameSpace, OperationResult(Of T)),
                                                 cancel As CancellationToken) As Task(Of OperationResult(Of T))
            Return Await RunAsync(operation, work, allowRetry:=False, isMutation:=True, cancel:=cancel)
        End Function

        Private Async Function RunAsync(Of T)(operation As String,
                                              work As Func(Of OL.Application, OL.NameSpace, OperationResult(Of T)),
                                              allowRetry As Boolean,
                                              isMutation As Boolean,
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

            ' A fase é LOCAL a esta invocação, e isso é a correção de F1-N.
            '
            ' Era um campo do broker: zerado na thread do CHAMADOR antes de
            ' postar, marcado na STA, e lido no Catch depois do Await. Uma
            ' operação concorrente — uma recarga de pasta disparada por
            ' evento, por exemplo — zerava o campo entre a falha de um Send e
            ' a classificação dela. O envio que talvez tivesse saído voltava
            ' como NotConnected, que é retentável.
            '
            ' A regra estava certa e lia estado errado.
            Dim fase As New FaseDaOperacao()

            Try
                Dim resultado = Await InvokeAsync(
                    Function()
                        If _application Is Nothing OrElse _namespace Is Nothing Then
                            Return OperationResult(Of T).Fail(ErrorKind.NotConnected, "sem sessão")
                        End If
                        ' AllowRetry continua no filtro, que é do broker: ele
                        ' é lido e escrito só aqui dentro, na STA, que roda
                        ' uma operação por vez.
                        _filter.AllowRetry = allowRetry
                        fase.Marcar()
                        Try
                            Return work(_application, _namespace)
                        Finally
                            _filter.AllowRetry = True
                        End Try
                    End Function)

                ' "ok", nao "None": ErrorKind.None significa sucesso, e ler
                ' "None" numa linha de log parece falha.
                _log.Write(LogLevel.Debug, operation,
                           $"{If(resultado.Succeeded, "ok", resultado.Kind.ToString())} " &
                           $"em {(DateTime.UtcNow - inicio).TotalMilliseconds:0} ms")
                Return resultado

            Catch ex As COMException
                Dim kind = ClassificarERegistrar(ex.HResult, isMutation, fase, operation)
                Return OperationResult(Of T).Fail(kind, $"0x{ex.HResult:X8}")
            Catch ex As Exception
                Dim kind = ClassificarERegistrar(Nothing, isMutation, fase, operation)
                _log.Write(LogLevel.Error, operation, ex.GetType().Name)
                Return OperationResult(Of T).Fail(kind, ex.GetType().Name)
            End Try
        End Function

        ''' <summary>
        ''' Fase de UMA invocação. Objeto, e não campo do broker, porque cada
        ''' chamada precisa da sua — ver o comentário em RunAsync.
        ''' </summary>
        Private NotInheritable Class FaseDaOperacao
            Private _iniciou As Integer

            ''' <summary>Chamado na STA; lido na thread do chamador.</summary>
            Public Sub Marcar()
                Volatile.Write(_iniciou, 1)
            End Sub

            Public ReadOnly Property Iniciou As Boolean
                Get
                    Return Volatile.Read(_iniciou) = 1
                End Get
            End Property
        End Class

        ''' <summary>
        ''' Decide pela política e registra. A REGRA mora em
        ''' <see cref="OutlookFailurePolicy"/>, no Core, onde teste alcança;
        ''' aqui fica o que é do broker: observar a fase e escrever o log.
        ''' </summary>
        Private Function ClassificarERegistrar(hresult As Integer?, isMutation As Boolean,
                                               fase As FaseDaOperacao,
                                               operation As String) As ErrorKind

            Dim kind = OutlookFailurePolicy.ClassifyFailure(hresult, isMutation, fase.Iniciou)

            If kind = ErrorKind.Ambiguous Then
                _log.Write(LogLevel.Warn, operation,
                           "AMBIGUO apos iniciar mutacao" &
                           If(hresult.HasValue, $" (0x{hresult.Value:X8})", ""))
            ElseIf hresult.HasValue Then
                _log.Write(LogLevel.Warn, operation, $"HRESULT 0x{hresult.Value:X8} -> {kind}")
            End If

            Return kind
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
            Return OutlookFailurePolicy.ClassifyFailure(hresult, isMutation:=False,
                                                        mutationAttemptStarted:=False)
        End Function

        ' ===================================================================
        ' Sessão
        ' ===================================================================

        Public Async Function ConnectAsync(cancel As CancellationToken) As Task(Of SessionState) _
            Implements IOutlookBroker.ConnectAsync
            Dim resultado = Await InvokeAsync(Function() ConnectCore())
            SetState(resultado)
            PublicarSessaoSeHouver()
            Return resultado
        End Function

        Public Async Function ProbeAsync(cancel As CancellationToken) As Task(Of SessionState) _
            Implements IOutlookBroker.ProbeAsync
            Dim resultado = Await InvokeAsync(Function() ProbeCore())
            SetState(resultado)
            PublicarSessaoSeHouver()
            Return resultado
        End Function

        Private Function ConnectCore() As SessionState
            AssertOnBrokerThread()
            Try
                ' Se ja ha sessao e ela responde, nao reconectar. Liberar a
                ' sessao antiga ANTES de adquirir a nova pode derrubar o
                ' proprio RCW recem-obtido, quando GetActiveObject devolve a
                ' mesma instancia.
                If _application IsNot Nothing Then
                    Try
                        Dim vivo = _application.Name
                        If Not String.IsNullOrEmpty(vivo) Then Return SessionState.Connected
                    Catch
                        ReleaseSessionCore()
                    End Try
                End If

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

                _application = app
                _namespace = app.GetNamespace("MAPI")

                ' Sessão NOVA. Tudo o que a anterior entregou — chaves,
                ' tokens de assinatura — deixou de valer neste instante, e as
                ' assinaturas em si já foram embora no ReleaseSessionCore.
                '
                ' ANOTA e não publica. Publicar daqui rodaria código de
                ' assinante dentro da rotina que acabou de instalar
                ' _application e _namespace: um assinante que bloqueasse
                ' esperando o broker, ou que usasse Dispatcher.Invoke,
                ' travaria a aquisição; e uma exceção dele faria uma conexão
                ' bem-sucedida parecer falha.
                Dim nova = Interlocked.Increment(_epoca)
                Interlocked.Exchange(_epocaAPublicar, nova)
                _probesDesconhecidos = 0
                _log.Write(LogLevel.Info, "broker.session", $"epoca {nova}")

                Return SessionState.Connected

            Catch ex As COMException
                ' Nem toda COMException e "ocupado": acesso negado e objeto
                ' desconectado sao coisas diferentes, e chamar tudo de Busy faz
                ' a UI prometer reconexao automatica que nao vai acontecer.
                ReleaseSessionCore()
                Select Case Classify(ex.HResult)
                    Case ErrorKind.Busy : Return SessionState.Busy
                    Case Else : Return SessionState.Unavailable
                End Select
            End Try
        End Function

        Private Function ProbeCore() As SessionState
            AssertOnBrokerThread()
            If _application Is Nothing Then Return ConnectCore()

            Try
                Dim nome = _application.Name
                _probesDesconhecidos = 0
                Return If(String.IsNullOrEmpty(nome), SessionState.Disconnected, SessionState.Connected)
            Catch ex As COMException
                ' Classificar pelo HRESULT: tratar toda COMException como
                ' morte derrubaria a sessão por uma recusa transitória.
                If OutlookFailurePolicy.IsSessionDead(ex.HResult) Then
                    ReleaseSessionCore()
                    Return ConnectCore()
                End If

                Select Case ex.HResult
                    Case &H80010001, &H8001010A
                        _probesDesconhecidos = 0
                        Return SessionState.Busy

                    Case Else
                        ' Preservar o RCW na primeira é certo: recusa
                        ' transitória não é morte. Insistir para sempre é que
                        ' não — se o RCW estiver de fato morto com um código
                        ' que ninguém previu, o Iris ficava "ocupado"
                        ' eternamente, sem erro, sem reconexão e sem sinal.
                        _probesDesconhecidos += 1
                        _log.Write(LogLevel.Warn, "broker.probe",
                                   $"HRESULT 0x{ex.HResult:X8} nao classificado " &
                                   $"({_probesDesconhecidos}x)")

                        If OutlookFailurePolicy.ShouldReattachAfterUnknown(_probesDesconhecidos) Then
                            _log.Write(LogLevel.Warn, "broker.probe",
                                       "desistindo do RCW e reanexando")
                            ReleaseSessionCore()
                            Return ConnectCore()
                        End If

                        Return SessionState.Busy
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
                                    ' f.Folders e f.Items sao objetos COM
                                    ' PROPRIOS. Escrever f.Folders.Count cria um
                                    ' RCW intermediario sem dono - o R7, que eu
                                    ' ja violei duas vezes antes desta. E aqui o
                                    ' vazamento se repetiria a cada pasta
                                    ' expandida na arvore.
                                    Dim netos = 0
                                    Dim subpastas As OL.Folders = Nothing
                                    Try
                                        subpastas = f.Folders
                                        netos = subpastas.Count
                                    Catch
                                    Finally
                                        ComHelpers.Release(subpastas)
                                    End Try

                                    Dim total = 0
                                    Dim itens As OL.Items = Nothing
                                    Try
                                        itens = f.Items
                                        total = itens.Count
                                    Catch
                                    Finally
                                        ComHelpers.Release(itens)
                                    End Try

                                    lista.Add(New FolderInfo With {
                                        .Key = New FolderKey(Safe(Function() f.EntryID), Safe(Function() f.StoreID)),
                                        .Name = Safe(Function() f.Name),
                                        .ContentKind = TipoDeConteudo(f),
                                        .ItemCount = total,
                                        .UnreadCount = SafeInt(Function() f.UnReadItemCount),
                                        .HasChildren = netos > 0,
                                        .IsHidden = PastaOculta(f)
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

        Public Async Function GetMessagePageAsync(query As MessageQuery, continuation As String,
                                                  targetCount As Integer,
                                                  cancel As CancellationToken) _
            As Task(Of OperationResult(Of MessagePage)) Implements IOutlookBroker.GetMessagePageAsync

            Return Await ReadAsync(Of MessagePage)(
                "outlook.getMessagePage",
                Function(app, ns) MessagePaging.ReadPage(ns, query, continuation, targetCount),
                cancel)
        End Function

        Public Async Function GetMessageDetailAsync(item As ItemKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of MessageDetail)) Implements IOutlookBroker.GetMessageDetailAsync

            Return Await ReadAsync(Of MessageDetail)(
                "outlook.getMessageDetail",
                Function(app, ns) MessageReading.ReadDetail(ns, item),
                cancel)
        End Function

        Public Async Function SaveAttachmentAsync(attachment As AttachmentKey, destinationPath As String,
                                                  overwrite As Boolean, cancel As CancellationToken) _
            As Task(Of OperationResult(Of String)) Implements IOutlookBroker.SaveAttachmentAsync

            ' MutateAsync: gravar em disco e efeito externo, nao leitura
            ' idempotente. Retry cego aqui poderia reescrever um arquivo.
            Return Await MutateAsync(Of String)(
                "outlook.saveAttachment",
                Function(app, ns) MessageReading.SaveAttachment(ns, attachment, destinationPath, overwrite),
                cancel)
        End Function

        Public Async Function MarkReadAsync(item As ItemKey, isRead As Boolean, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.MarkReadAsync

            Return Await MutateAsync(Of Boolean)(
                "outlook.markRead",
                Function(app, ns) MessageReading.SetReadState(ns, item, isRead),
                cancel)
        End Function

        ' Criar um rascunho GRAVA no store: vai por MutateAsync, com o
        ' retry do message filter desligado. Um retry aqui criaria dois
        ' rascunhos.

        Public Async Function CreateDraftAsync(content As DraftContent, cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateDraftAsync

            Return Await MutateAsync(Of DraftInfo)(
                "outlook.createDraft",
                Function(app, ns) DraftWriting.CreateNew(app, ns), cancel)
        End Function

        Public Async Function CreateReplyDraftAsync(item As ItemKey, replyAll As Boolean,
                                                    cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateReplyDraftAsync

            Return Await MutateAsync(Of DraftInfo)(
                "outlook.createReplyDraft",
                Function(app, ns) DraftWriting.CreateReply(ns, item, replyAll), cancel)
        End Function

        Public Async Function CreateForwardDraftAsync(item As ItemKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateForwardDraftAsync

            Return Await MutateAsync(Of DraftInfo)(
                "outlook.createForwardDraft",
                Function(app, ns) DraftWriting.CreateForward(ns, item), cancel)
        End Function

        Public Async Function UpdateDraftAsync(draft As DraftKey, content As DraftContent,
                                               cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.UpdateDraftAsync

            Return Await MutateAsync(Of DraftInfo)(
                "outlook.updateDraft",
                Function(app, ns) DraftWriting.Update(ns, draft, content), cancel)
        End Function

        Public Async Function AddDraftAttachmentAsync(draft As DraftKey, filePath As String,
                                                      cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.AddDraftAttachmentAsync

            Return Await MutateAsync(Of DraftInfo)(
                "outlook.addDraftAttachment",
                Function(app, ns) DraftWriting.AddAttachment(ns, draft, filePath), cancel)
        End Function

        Public Async Function RemoveDraftAttachmentAsync(draft As DraftKey, attachment As AttachmentKey,
                                                         cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.RemoveDraftAttachmentAsync

            Return Await MutateAsync(Of DraftInfo)(
                "outlook.removeDraftAttachment",
                Function(app, ns) DraftWriting.RemoveAttachment(ns, draft, attachment), cancel)
        End Function

        ''' <summary>
        ''' Descobre para quem a mensagem vai e por qual conta. NÃO envia.
        '''
        ''' Não é leitura pura — ResolveAll mexe nos Recipients — e por isso
        ''' não vai por ReadAsync: a regra "leitura tem retry, mutação não"
        ''' é absoluta, e abrir exceção para ela no código enquanto o
        ''' contrato a declara sem exceção é como não ter regra.
        '''
        ''' Vai por uma terceira classificação: SEM retry, porque não é
        ''' leitura; e SEM Ambiguous, porque falhar aqui não deixa efeito
        ''' nenhum no mundo — marcar como ambíguo travaria o rascunho de um
        ''' usuário que só queria revisar para quem ia mandar.
        ''' </summary>
        Public Async Function PrepareSendAsync(draft As DraftKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of SendPreview)) Implements IOutlookBroker.PrepareSendAsync

            Return Await SemRetryAsync(Of SendPreview)(
                "outlook.prepareSend",
                Function(app, ns) DraftWriting.PrepareSend(ns, draft), cancel)
        End Function

        ''' <summary>
        ''' Envia. Falha depois de o Send() começar vira Ambiguous pelo
        ''' classificador, e Ambiguous NUNCA é retentável — reenviar no
        ''' escuro é o único erro irreversível deste projeto.
        ''' </summary>
        Public Async Function SendDraftAsync(draft As DraftKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.SendDraftAsync

            Return Await MutateAsync(Of Boolean)(
                "outlook.sendDraft",
                Function(app, ns) DraftWriting.Send(ns, draft), cancel)
        End Function

        Public Async Function DeleteDraftAsync(draft As DraftKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.DeleteDraftAsync

            Return Await MutateAsync(Of Boolean)(
                "outlook.deleteDraft",
                Function(app, ns) DraftWriting.Delete(ns, draft), cancel)
        End Function

        Public Async Function SubscribeFolderAsync(folder As FolderKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of SubscriptionToken)) Implements IOutlookBroker.SubscribeFolderAsync

            Return Await ReadAsync(Of SubscriptionToken)(
                "outlook.subscribeFolder",
                Function(app, ns)
                    Dim pasta As OL.MAPIFolder = Nothing
                    Try
                        pasta = TryCast(ns.GetFolderFromID(folder.EntryId, folder.StoreId), OL.MAPIFolder)
                    Catch ex As Runtime.InteropServices.COMException When EhNaoEncontrado(ex.HResult)
                        ' Mesmo filtro do MessagePaging: ocupado, desconectado
                        ' e acesso negado nao sao "pasta nao existe", e levam
                        ' a UI a decisoes opostas. As demais sobem para o
                        ' classificador.
                        Return OperationResult(Of SubscriptionToken).Fail(ErrorKind.NotFound, "pasta")
                    End Try

                    If pasta Is Nothing Then
                        Return OperationResult(Of SubscriptionToken).Fail(ErrorKind.NotFound, "pasta")
                    End If

                    ' A assinatura vira dona da pasta: NAO liberar aqui.
                    Dim assinatura As New FolderSubscription(folder, pasta, AddressOf RaiseInvalidation)
                    _subscriptions(assinatura.Id) = assinatura

                    _log.Write(LogLevel.Info, "outlook.subscribeFolder", $"id={assinatura.Id}")
                    Return OperationResult(Of SubscriptionToken).Ok(
                        New SubscriptionToken(assinatura.Id, folder))
                End Function, cancel)
        End Function

        Public Async Function UnsubscribeFolderAsync(token As SubscriptionToken, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.UnsubscribeFolderAsync

            Return Await ReadAsync(Of Boolean)(
                "outlook.unsubscribeFolder",
                Function(app, ns)
                    Dim assinatura As FolderSubscription = Nothing
                    If Not _subscriptions.TryGetValue(token.Id, assinatura) Then
                        Return OperationResult(Of Boolean).Ok(False)
                    End If
                    _subscriptions.Remove(token.Id)
                    assinatura.Dispose()
                    Return OperationResult(Of Boolean).Ok(True)
                End Function, cancel)
        End Function

        ''' <summary>
        ''' Chamado da thread MTA de entrega. Só repassa o aviso — o
        ''' assinante decide o que fazer, e a resposta certa é RELER.
        ''' </summary>
        Private Sub RaiseInvalidation(invalidation As FolderInvalidation)
            RaiseEvent FolderInvalidated(Me, invalidation)
        End Sub

        ' ===================================================================
        ' Encerramento
        ' ===================================================================

        Private Sub ReleaseSessionCore()
            ' Assinaturas primeiro: Dispose desconecta os handlers antes de
            ' liberar os RCWs. Soltar a sessao com sinks conectados e o
            ' caminho para OUTLOOK.EXE orfao.
            For Each par In _subscriptions
                Try
                    par.Value.Dispose()
                Catch
                End Try
            Next
            _subscriptions.Clear()

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

            Dim limpezaOk = True
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
            Catch ex As Exception
                limpezaOk = False
                _log.Write(LogLevel.Error, "broker.shutdown", "limpeza falhou: " & ex.GetType().Name)
            End Try

            _dispatcher.InvokeShutdown()

            Dim encerrou = _thread Is Nothing OrElse _thread.Join(timeout)
            If encerrou Then
                _log.Write(LogLevel.Info, "broker.shutdown",
                           If(limpezaOk, "ok", "thread ok, limpeza falhou"))
            Else
                ' Registrar "ok" logo depois de o Join falhar era log
                ' mentiroso, e log mentiroso e pior que log ausente.
                _log.Write(LogLevel.Error, "broker.shutdown", "thread NAO encerrou no timeout")
            End If
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

        ''' <summary>
        ''' Traduz o tipo do COM para o vocabulario do Model. Sem isto, o
        ''' Core e a UI acabariam comparando a string "olMailItem" — nome de
        ''' membro de enum do interop vazando para as camadas de cima.
        ''' </summary>
        Private Shared Function TipoDeConteudo(f As OL.MAPIFolder) As FolderContentKind
            Try
                Select Case f.DefaultItemType
                    Case OL.OlItemType.olMailItem : Return FolderContentKind.Mail
                    Case OL.OlItemType.olAppointmentItem : Return FolderContentKind.Calendar
                    Case OL.OlItemType.olContactItem : Return FolderContentKind.Contacts
                    Case OL.OlItemType.olTaskItem : Return FolderContentKind.Tasks
                    Case OL.OlItemType.olNoteItem : Return FolderContentKind.Notes
                    Case OL.OlItemType.olJournalItem : Return FolderContentKind.Journal
                    Case Else : Return FolderContentKind.Unknown
                End Select
            Catch
                Return FolderContentKind.Unknown
            End Try
        End Function

        ''' <summary>
        ''' PR_ATTR_HIDDEN via PropertyAccessor. O PropertyAccessor e um
        ''' objeto COM proprio e precisa ser liberado — encadear
        ''' f.PropertyAccessor.GetProperty(...) seria o R7 mais uma vez.
        ''' </summary>
        Private Shared Function PastaOculta(f As OL.MAPIFolder) As Boolean
            Const TagOculta As String = "http://schemas.microsoft.com/mapi/proptag/0x10F4000B"
            Dim acessor As OL.PropertyAccessor = Nothing
            Try
                acessor = f.PropertyAccessor
                Dim valor = acessor.GetProperty(TagOculta)
                Return TypeOf valor Is Boolean AndAlso CBool(valor)
            Catch
                ' Pasta sem a propriedade: tratar como visivel.
                Return False
            Finally
                ComHelpers.Release(acessor)
            End Try
        End Function

        ''' <summary>
        ''' MAPI_E_NOT_FOUND e o "objeto nao existe" do MAPI; E_INVALIDARG
        ''' aparece quando o EntryID nao pertence mais ao store.
        ''' </summary>
        Private Shared Function EhNaoEncontrado(hresult As Integer) As Boolean
            Return hresult = &H8004010F OrElse hresult = &H80070057
        End Function

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

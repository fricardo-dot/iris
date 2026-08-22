Imports System.Threading
Imports System.Windows.Threading
Imports IrisSpike.Interop
Imports Outlook = Microsoft.Office.Interop.Outlook

Namespace Broker

    Public Enum SessionState
        Disconnected
        Connected
        Reconnecting
        ''' <summary>Outlook aberto, mas recusando chamadas (R13).</summary>
        Busy
        ''' <summary>Outlook não está em execução, ou nem instalado.</summary>
        Unavailable
    End Enum

    ''' <summary>
    ''' O broker do ESCOPO.md, seção 4. Única porta de entrada para o COM.
    '''
    ''' Três propriedades que a Fase 0 precisa provar, nesta ordem:
    '''
    '''   1. Thread STA dedicada. O OOM tem afinidade de thread; pegar um
    '''      objeto aqui e usar em outra thread produz falha errática.
    '''   2. MESSAGE PUMP de verdade. Uma fila bloqueante processaria
    '''      comandos e mataria a entrega de eventos do Outlook, que chegam
    '''      como mensagens de janela. Por isso Dispatcher.Run(), não um
    '''      While/Dequeue.
    '''   3. Só DTOs atravessam a fronteira. Nenhum RCW sobe para o chamador.
    '''
    ''' Tipagem forte em todo o acesso ao Outlook — o escopo descarta late
    ''' binding como estratégia principal (seção 10).
    ''' </summary>
    Public NotInheritable Class OutlookBroker
        Implements IDisposable

        Private ReadOnly _ready As New ManualResetEventSlim(False)
        Private _thread As Thread
        Private _dispatcher As Dispatcher
        Private _filter As OutlookMessageFilter
        Private _startupError As Exception
        Private _disposed As Boolean

        ' Sessão Outlook. Só pode ser tocada de dentro da thread do broker.
        Private _application As Outlook.Application
        Private _namespace As Outlook.NameSpace

        ''' <summary>
        ''' R7, tensão deliberada: as coleções Items com eventos assinados
        ''' precisam de referência FORTE viva. Se o GC coletar a coleção, o
        ''' event sink morre junto e os eventos param sem erro nenhum — é o
        ''' bug que se manifesta como "os eventos simplesmente pararam".
        ''' Soltar cedo demais aqui é tão ruim quanto nunca soltar.
        ''' </summary>
        ''' <remarks>
        ''' IDisposable, não Object: guardar só a referência mantinha a
        ''' coleção viva mas não permitia RemoveHandler, então cancelar a
        ''' assinatura era impossível. Ver FolderSubscription.
        ''' </remarks>
        Private ReadOnly _liveSinks As New List(Of IDisposable)()

        ' HRESULTs usados na classificação de ProbeCore.
        Private Const RPC_E_CALL_REJECTED As Integer = &H80010001
        Private Const RPC_E_SERVERCALL_RETRYLATER As Integer = &H8001010A
        Private Const RPC_E_DISCONNECTED As Integer = &H80010108
        Private Const RPC_S_SERVER_UNAVAILABLE As Integer = &H800706BA
        Private Const CO_E_OBJNOTCONNECTED As Integer = &H800401FD

        Public ReadOnly Property ThreadId As Integer
        Public ReadOnly Property Apartment As ApartmentState
        Public Property State As SessionState = SessionState.Disconnected

        ''' <summary>Diagnóstico da última tentativa de anexar.</summary>
        Public Property LastAttachOutcome As ComHelpers.AttachOutcome
        Public Property LastAttachHresult As Integer

        Public ReadOnly Property MessageFilter As OutlookMessageFilter
            Get
                Return _filter
            End Get
        End Property

        Public ReadOnly Property IsRunning As Boolean
            Get
                Return _dispatcher IsNot Nothing AndAlso Not _dispatcher.HasShutdownStarted
            End Get
        End Property

        ''' <summary>
        ''' Sobe a thread e espera o pump estar de pé. Síncrono de propósito:
        ''' sem broker não há nada a fazer.
        ''' </summary>
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
        End Sub

        Private Sub ThreadBody()
            Try
                _dispatcher = Dispatcher.CurrentDispatcher
                _ThreadId = Environment.CurrentManagedThreadId
                _Apartment = Thread.CurrentThread.GetApartmentState()

                ' Por thread — precisa ser aqui dentro, não no chamador.
                _filter = OutlookMessageFilter.Register()
            Catch ex As Exception
                _startupError = ex
                _ready.Set()
                Return
            End Try

            _ready.Set()

            ' ISTO é o message pump. Bloqueia até InvokeShutdown().
            Dispatcher.Run()

            ' Encerramento ordenado, ainda na thread dona dos objetos.
            Try
                ReleaseSessionCore()
                OutlookMessageFilter.Revoke()
            Catch
                ' Encerramento nunca deve lançar.
            End Try
        End Sub

        ''' <summary>
        ''' Executa na thread do broker e devolve um valor. Todo resultado
        ''' passa por <see cref="GuardResult"/> — é o ponto único por onde
        ''' qualquer coisa sai da thread do broker.
        ''' </summary>
        Public Function InvokeAsync(Of T)(work As Func(Of T)) As Task(Of T)
            EnsureUsable()
            Return _dispatcher.InvokeAsync(Function() GuardResult(work())).Task
        End Function

        ''' <summary>Executa na thread do broker sem retorno.</summary>
        Public Function InvokeAsync(work As Action) As Task
            EnsureUsable()
            Return _dispatcher.InvokeAsync(work).Task
        End Function

        ''' <summary>
        ''' Enfileira trabalho na thread do broker SEM bloquear o chamador.
        ''' Usado pelos callbacks de evento, que chegam numa thread MTA do
        ''' pool: bloquear ali para esperar a STA arriscaria deadlock quando
        ''' o evento e disparado de dentro de uma chamada que a propria STA
        ''' esta executando.
        ''' </summary>
        Public Sub PostToBrokerThread(work As Action)
            If _dispatcher Is Nothing OrElse _dispatcher.HasShutdownStarted Then Return
            _dispatcher.InvokeAsync(work)
        End Sub

        ''' <summary>
        ''' Acesso direto ao NameSpace para helpers que JÁ estão rodando na
        ''' thread do broker. Não passa por GuardResult de propósito: devolve
        ''' RCW por contrato, e quem chama é dono de liberar. Uso interno.
        ''' </summary>
        Public Function WithNamespace(Of T)(work As Func(Of Outlook.NameSpace, T)) As T
            RequireSession()
            Return work(_namespace)
        End Function

        ''' <summary>
        ''' Operação COM idempotente (leitura). O message filter pode repetir
        ''' a chamada se o Outlook estiver ocupado.
        ''' </summary>
        Public Function ReadAsync(Of T)(work As Func(Of Outlook.Application, Outlook.NameSpace, T)) As Task(Of T)
            Return InvokeAsync(Function()
                                   RequireSession()
                                   _filter.AllowRetry = True
                                   Return work(_application, _namespace)
                               End Function)
        End Function

        ''' <summary>
        ''' A assinatura genérica destes métodos permite escrever
        ''' <c>ReadAsync(Function(app, ns) ns.GetDefaultFolder(...))</c> e
        ''' devolver um RCW ao chamador — a fronteira da seção 4 é uma
        ''' convenção, não algo que o compilador imponha.
        '''
        ''' Aqui isso vira erro em tempo de execução. No produto, a correção
        ''' de verdade é outra: estes genéricos ficam privados e a API pública
        ''' expõe operações específicas (GetMailSummariesAsync,
        ''' CreateDraftAsync, SubscribeFolderAsync), sem lambda recebendo
        ''' Application ou NameSpace.
        ''' </summary>
        Private Shared Function GuardResult(Of T)(value As T) As T
            If ComHelpers.ContainsComReference(value) Then
                Throw New InvalidOperationException(
                    "Um objeto COM tentou atravessar a fronteira do broker. " &
                    "Só DTOs podem sair (ESCOPO.md, seção 4).")
            End If
            Return value
        End Function

        ''' <summary>
        ''' Operação com efeito colateral — Send() acima de todas. Desliga o
        ''' retry do message filter enquanto roda: uma chamada repetida aqui
        ''' envia o e-mail duas vezes (R13).
        ''' </summary>
        Public Function MutateAsync(Of T)(work As Func(Of Outlook.Application, Outlook.NameSpace, T)) As Task(Of T)
            Return InvokeAsync(Function()
                                   RequireSession()
                                   _filter.AllowRetry = False
                                   Try
                                       Return work(_application, _namespace)
                                   Finally
                                       _filter.AllowRetry = True
                                   End Try
                               End Function)
        End Function

        ''' <summary>
        ''' Anexa a um Outlook JÁ EM EXECUÇÃO. Não inicia o aplicativo.
        ''' Retorna o estado resultante, sem lançar: "Outlook não está
        ''' aberto" é um estado previsto do produto, não uma exceção.
        ''' </summary>
        Public Async Function ConnectAsync() As Task(Of SessionState)
            Dim result = Await InvokeAsync(Function() ConnectCore())
            State = result
            Return result
        End Function

        Private Function ConnectCore() As SessionState
            AssertOnBrokerThread()

            Try
                Dim attach = ComHelpers.GetRunningInstance("Outlook.Application")
                LastAttachOutcome = attach.Outcome
                LastAttachHresult = attach.Hresult

                Select Case attach.Outcome
                    Case ComHelpers.AttachOutcome.Busy
                        Return SessionState.Busy
                    Case ComHelpers.AttachOutcome.NotRunning,
                         ComHelpers.AttachOutcome.NotRegistered,
                         ComHelpers.AttachOutcome.Failed
                        Return SessionState.Unavailable
                End Select

                ' Referências curtas e nomeadas, nunca encadeadas (R7).
                Dim app = TryCast(attach.Instance, Outlook.Application)
                If app Is Nothing Then
                    ComHelpers.Release(attach.Instance)
                    Return SessionState.Unavailable
                End If

                ' Uma sessão anterior precisa sair antes de a nova entrar,
                ' senão reconectar vaza o Application e o NameSpace antigos.
                ReleaseSessionCore()

                _application = app
                _namespace = app.GetNamespace("MAPI")
                Return SessionState.Connected

            Catch ex As Runtime.InteropServices.COMException
                ' Anexou, mas GetNamespace foi recusado: Outlook ocupado.
                ReleaseSessionCore()
                LastAttachHresult = ex.HResult
                Return SessionState.Busy
            End Try
        End Function

        ''' <summary>
        ''' Detecta se a sessão morreu (Outlook fechado por baixo) e tenta
        ''' restabelecer. Critério E da Fase 0: isto decide se o broker tem
        ''' estado Reconnecting ou se o aplicativo inteiro precisa reiniciar.
        ''' </summary>
        Public Async Function ProbeAsync() As Task(Of SessionState)
            Dim result = Await InvokeAsync(Function() ProbeCore())
            State = result
            Return result
        End Function

        Private Function ProbeCore() As SessionState
            AssertOnBrokerThread()
            If _application Is Nothing Then Return ConnectCore()

            Try
                ' Toque barato só para ver se o RCW ainda responde.
                Dim name = _application.Name
                Return If(String.IsNullOrEmpty(name), SessionState.Disconnected, SessionState.Connected)

            Catch ex As Runtime.InteropServices.COMException
                LastAttachHresult = ex.HResult

                ' Tratar QUALQUER COMException como sessão morta derrubaria a
                ' conexão por uma recusa transitória — e reconectar é caro.
                ' Classificar pelo HRESULT é o que separa "morreu" de
                ' "está ocupado agora".
                Select Case ex.HResult
                    Case RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, CO_E_OBJNOTCONNECTED
                        State = SessionState.Reconnecting
                        ReleaseSessionCore()
                        Return ConnectCore()

                    Case RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER
                        ' Vivo, só indisponível neste instante. Manter a sessão.
                        Return SessionState.Busy

                    Case Else
                        ' Erro real e não classificado: não presumir morte.
                        Return SessionState.Connected
                End Select
            End Try
        End Function

        ''' <summary>
        ''' Mantém viva a coleção cujo evento foi assinado. Ver _liveSinks.
        ''' </summary>
        Public Sub TrackSink(sink As IDisposable)
            AssertOnBrokerThread()
            _liveSinks.Add(sink)
        End Sub

        ''' <summary>
        ''' Cancela e descarta uma assinatura específica, sem esperar o
        ''' encerramento do broker. Precisa rodar na thread do broker.
        ''' </summary>
        Public Sub ReleaseSink(sink As IDisposable)
            AssertOnBrokerThread()
            _liveSinks.Remove(sink)
            sink.Dispose()
        End Sub

        Public ReadOnly Property LiveSinkCount As Integer
            Get
                Return _liveSinks.Count
            End Get
        End Property

        Private Sub RequireSession()
            AssertOnBrokerThread()
            If _application Is Nothing OrElse _namespace Is Nothing Then
                Throw New InvalidOperationException(
                    "Sem sessão com o Outlook. Chame ConnectAsync() antes.")
            End If
        End Sub

        Private Sub ReleaseSessionCore()
            ' Dispose desconecta os handlers e depois libera o RCW. Só
            ' liberar, como antes, deixaria event sinks pendurados.
            For i = _liveSinks.Count - 1 To 0 Step -1
                Try
                    _liveSinks(i).Dispose()
                Catch
                    ' Encerramento nunca lança.
                End Try
            Next
            _liveSinks.Clear()

            ' Ordem inversa da aquisição (R7).
            ComHelpers.Release(_namespace)
            _namespace = Nothing
            ComHelpers.Release(_application)
            _application = Nothing
        End Sub

        Public Sub AssertOnBrokerThread()
            If Environment.CurrentManagedThreadId <> ThreadId Then
                Throw New InvalidOperationException(
                    $"Objeto COM tocado fora da thread do broker " &
                    $"(atual {Environment.CurrentManagedThreadId}, broker {ThreadId}). " &
                    "Ver R6 do ESCOPO.md.")
            End If
        End Sub

        Private Sub EnsureUsable()
            If _disposed Then Throw New ObjectDisposedException(NameOf(OutlookBroker))
            If _dispatcher Is Nothing Then Throw New InvalidOperationException("Broker não iniciado.")
        End Sub

        ''' <summary>
        ''' Encerramento ordenado. A liberação dos objetos COM acontece na
        ''' thread dona deles, depois que o pump para — ver ThreadBody.
        ''' </summary>
        Public Sub Shutdown(Optional timeout As TimeSpan = Nothing)
            If timeout = Nothing Then timeout = TimeSpan.FromSeconds(10)
            If _dispatcher Is Nothing Then Return

            ' A liberação precisa acontecer com o PUMP AINDA VIVO. Desconectar
            ' event sinks e liberar proxies COM pode exigir processamento de
            ' mensagens; fazer isso depois que Dispatcher.Run() retornou é
            ' pedir para travar ou deixar referência pendurada. A limpeza no
            ' fim de ThreadBody continua existindo, mas só como rede de
            ' segurança para o caminho em que ninguém chamou Shutdown().
            Try
                _dispatcher.Invoke(
                    Sub()
                        ReleaseSessionCore()
                        OutlookMessageFilter.Revoke()
                    End Sub, DispatcherPriority.Send, Nothing, timeout)
            Catch
                ' Se a limpeza ordenada falhar, ainda assim derruba a thread.
            End Try

            _dispatcher.InvokeShutdown()

            If _thread IsNot Nothing AndAlso Not _thread.Join(timeout) Then
                Throw New TimeoutException(
                    $"Thread do broker não encerrou em {timeout.TotalSeconds:0}s.")
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            Try
                Shutdown()
            Catch
                ' Dispose não lança.
            End Try
            _ready.Dispose()
        End Sub
    End Class

End Namespace

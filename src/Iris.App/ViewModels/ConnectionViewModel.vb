Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' Estado da conexão com o Outlook.
    '''
    ''' Depende de <see cref="IOutlookBroker"/>, que vive em Iris.Core.
    ''' NUNCA do assembly Iris.Outlook: a View e o ViewModel não podem
    ''' enxergar COM, e há teste arquitetural cobrando isso.
    '''
    ''' O evento do broker pode chegar em qualquer thread — na Fase 0
    ''' medimos callbacks do Outlook chegando em thread MTA do pool. Por
    ''' isso todo toque em propriedade observável é remarcado para o
    ''' Dispatcher da UI aqui, e não se confia na continuação do Await
    ''' voltar ao contexto certo.
    ''' </summary>
    Public NotInheritable Class ConnectionViewModel
        Inherits ObservableObject
        Implements IDisposable

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _ui As Threading.Dispatcher
        Private _disposed As Boolean

        Private _state As SessionState = SessionState.Disconnected
        Private _lastCheck As DateTime?
        Private _storeName As String = ""
        Private _detail As String = ""

        ' 0 = stores ainda não lidos nesta conexão. Interlocked porque duas
        ' origens de mudança de estado disputam a leitura.
        Private _storesLoaded As Integer = 0

        Public Sub New(broker As IOutlookBroker, uiDispatcher As Threading.Dispatcher)
            _broker = broker
            _ui = uiDispatcher

            RetryCommand = New AsyncRelayCommand(AddressOf RetryAsync)
            OpenOutlookCommand = New RelayCommand(AddressOf OpenOutlook)

            AddHandler _broker.StateChanged, AddressOf OnBrokerStateChanged
            _state = _broker.State
        End Sub

        Public ReadOnly Property RetryCommand As IAsyncRelayCommand
        Public ReadOnly Property OpenOutlookCommand As IRelayCommand

        Public Property State As SessionState
            Get
                Return _state
            End Get
            Private Set(value As SessionState)
                If SetProperty(_state, value) Then
                    OnPropertyChanged(NameOf(StatusLabel))
                    OnPropertyChanged(NameOf(Headline))
                    OnPropertyChanged(NameOf(Explanation))
                    OnPropertyChanged(NameOf(ShowOpenOutlook))
                    OnPropertyChanged(NameOf(IsWorking))
                End If
            End Set
        End Property

        Public Property LastCheck As DateTime?
            Get
                Return _lastCheck
            End Get
            Private Set(value As DateTime?)
                SetProperty(_lastCheck, value)
            End Set
        End Property

        Public Property StoreName As String
            Get
                Return _storeName
            End Get
            Private Set(value As String)
                If SetProperty(_storeName, value) Then
                    OnPropertyChanged(NameOf(HasStore))
                End If
            End Set
        End Property

        ''' <summary>
        ''' Booleano de propósito: BooleanToVisibilityConverter aplicado a
        ''' uma String não converte nada, e a linha apareceria vazia.
        ''' </summary>
        Public ReadOnly Property HasStore As Boolean
            Get
                Return Not String.IsNullOrEmpty(_storeName)
            End Get
        End Property

        Public Property Detail As String
            Get
                Return _detail
            End Get
            Private Set(value As String)
                SetProperty(_detail, value)
            End Set
        End Property

        ''' <summary>
        ''' Rótulo curto do indicador. Existe porque a versão anterior fazia
        ''' Binding direto no enum e a barra exibia "Connected" — nome de
        ''' identificador do código vazando para a interface, em inglês.
        ''' </summary>
        Public ReadOnly Property StatusLabel As String
            Get
                Select Case State
                    Case SessionState.Connected : Return "Conectado"
                    Case SessionState.Busy : Return "Outlook ocupado"
                    Case SessionState.Reconnecting : Return "Reconectando"
                    Case SessionState.Connecting : Return "Conectando"
                    Case SessionState.Unavailable : Return "Outlook fechado"
                    Case Else : Return "Desconectado"
                End Select
            End Get
        End Property

        ''' <summary>
        ''' A frase principal. Em "fechado" ela é uma INSTRUÇÃO, porque a
        ''' ação corretiva é do usuário — dizer só "desconectado" deixaria
        ''' a pessoa sem saber o que fazer.
        ''' </summary>
        Public ReadOnly Property Headline As String
            Get
                Select Case State
                    Case SessionState.Connected : Return "Conectado ao Outlook"
                    Case SessionState.Busy : Return "O Outlook está ocupado"
                    Case SessionState.Reconnecting : Return "Reconectando…"
                    Case SessionState.Connecting : Return "Conectando…"
                    Case Else : Return "Abra o Outlook clássico para continuar"
                End Select
            End Get
        End Property

        Public ReadOnly Property Explanation As String
            Get
                Select Case State
                    Case SessionState.Connected
                        Return "A sessão está pronta. O Iris usa a conta que já está autenticada no Outlook."
                    Case SessionState.Busy
                        Return "Ele pode estar iniciando, sincronizando ou esperando uma janela aberta. " &
                               "O Iris tenta de novo sozinho."
                    Case SessionState.Reconnecting, SessionState.Connecting
                        Return "Restabelecendo a sessão com o Outlook."
                    Case Else
                        Return "O Iris usa a sessão já autenticada do Outlook, sem pedir senha nem " &
                               "configurar conta. Assim que ele terminar de abrir, a conexão é feita " &
                               "automaticamente."
                End Select
            End Get
        End Property

        Public ReadOnly Property ShowOpenOutlook As Boolean
            Get
                Return State = SessionState.Unavailable OrElse State = SessionState.Disconnected
            End Get
        End Property

        Public ReadOnly Property IsWorking As Boolean
            Get
                Return State = SessionState.Connecting OrElse State = SessionState.Reconnecting
            End Get
        End Property

        ''' <summary>
        ''' Primeira conexão. Depois disso, o watchdog do broker mantém o
        ''' estado atualizado sozinho.
        ''' </summary>
        Public Async Function InitializeAsync() As Task
            Await RetryAsync()
        End Function

        Private Async Function RetryAsync() As Task
            Dim resultado = Await _broker.ProbeAsync(CancellationToken.None)
            Await ApplyAsync(resultado)
        End Function

        Private Async Function ApplyAsync(novo As SessionState) As Task
            Dim nomeStore = _storeName

            If novo <> SessionState.Connected Then
                ' Saiu de Connected: libera para reler quando voltar.
                Interlocked.Exchange(_storesLoaded, 0)
                nomeStore = ""
            End If

            ' A guarda precisa ser ATÔMICA, não uma leitura de _state.
            ' A conexão inicial e o watchdog chegam aqui em paralelo, e
            ' _state só muda depois, no BeginInvoke — então ambos liam o
            ' valor antigo e ambos buscavam os stores. Chamada de COM
            ' redundante ocupa a fila única do broker à toa (F1-F).
            Dim precisaLerStores = novo = SessionState.Connected AndAlso
                                   Interlocked.CompareExchange(_storesLoaded, 1, 0) = 0

            If precisaLerStores Then
                Dim stores = Await _broker.GetStoresAsync(CancellationToken.None)
                If stores.Succeeded AndAlso stores.Value.Count > 0 Then
                    nomeStore = stores.Value(0).DisplayName
                    If stores.Value.Count > 1 Then
                        nomeStore &= $" (+{stores.Value.Count - 1})"
                    End If
                Else
                    ' Falhou: libera a flag, senão a caixa ficaria sem nome
                    ' até a próxima desconexão.
                    Interlocked.Exchange(_storesLoaded, 0)
                End If
            End If

            OnUi(Sub()
                     State = novo
                     StoreName = nomeStore
                     LastCheck = DateTime.Now
                 End Sub)
        End Function

        Private Sub OnBrokerStateChanged(sender As Object, novo As SessionState)
            ' Pode chegar em qualquer thread. Nunca tocar propriedade
            ' observável daqui direto.
            Dim ignorado = ApplyAsync(novo)
        End Sub

        Private Sub OnUi(action As Action)
            If _ui.CheckAccess() Then
                action()
            Else
                _ui.BeginInvoke(action)
            End If
        End Sub

        ''' <summary>
        ''' Abre o Outlook de forma INTERATIVA, pelo shell. É diferente de
        ''' criar uma instância COM: aquela viria sem perfil interativo. O
        ''' watchdog do broker continua esperando a sessão real aparecer.
        ''' </summary>
        Private Sub OpenOutlook()
            Try
                Dim psi As New Global.System.Diagnostics.ProcessStartInfo("outlook.exe") With {
                    .UseShellExecute = True
                }
                Global.System.Diagnostics.Process.Start(psi)
                Detail = "Aguardando o Outlook abrir…"
            Catch
                Detail = "Não foi possível iniciar o Outlook daqui. Abra-o pelo menu Iniciar."
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            RemoveHandler _broker.StateChanged, AddressOf OnBrokerStateChanged
        End Sub

    End Class

End Namespace

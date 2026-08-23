Imports System.Collections.ObjectModel
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' O que um nó precisa para viver, num objeto só.
    '''
    ''' Cada nó guardava broker, dispatcher e callbacks separadamente.
    ''' Com o contexto, acrescentar uma dependência não exige mudar a
    ''' assinatura de todos os construtores, e fica explícito que os nós
    ''' COMPARTILHAM os mesmos serviços em vez de cada um ter os seus.
    ''' </summary>
    Public NotInheritable Class FolderTreeContext
        Public ReadOnly Property Broker As IOutlookBroker
        Public ReadOnly Property Ui As Global.System.Windows.Threading.Dispatcher
        Public ReadOnly Property Policy As FolderVisibilityPolicy

        Private ReadOnly _onSelected As Action(Of FolderNodeViewModel)
        Private ReadOnly _onError As Action(Of ErrorKind)
        Private ReadOnly _observe As Action(Of Task, String)

        Public Sub New(broker As IOutlookBroker,
                       ui As Global.System.Windows.Threading.Dispatcher,
                       policy As FolderVisibilityPolicy,
                       onSelected As Action(Of FolderNodeViewModel),
                       onError As Action(Of ErrorKind),
                       observe As Action(Of Task, String))
            Me.Broker = broker
            Me.Ui = ui
            Me.Policy = policy
            _onSelected = onSelected
            _onError = onError
            _observe = observe
        End Sub

        Public Sub NotifySelected(node As FolderNodeViewModel)
            _onSelected(node)
        End Sub

        Public Sub ReportError(kind As ErrorKind)
            _onError(kind)
        End Sub

        Public Sub Observe(work As Task, operation As String)
            _observe(work, operation)
        End Sub
    End Class

    ''' <summary>
    ''' A árvore de pastas.
    '''
    ''' Regras decididas explicitamente, porque "reflete a árvore do Outlook"
    ''' não é critério verificável:
    '''
    '''   • Entram todos os stores que o perfil expõe.
    '''   • O primeiro nível carrega junto com a conexão — abrir o Iris e ver
    '''     a árvore vazia até clicar seria pior que esperar meio segundo.
    '''   • Níveis seguintes carregam sob demanda, um por expansão.
    '''   • A ordem é a que o Outlook devolve. Reordenar alfabeticamente
    '''     brigaria com a ordem que o usuário já conhece do Outlook.
    '''   • O que aparece é decidido pela <see cref="FolderVisibilityPolicy"/>,
    '''     que vive no Core: é política de aplicação, não detalhe visual.
    ''' </summary>
    Public NotInheritable Class FolderTreeViewModel
        Inherits ObservableObject

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _ui As Global.System.Windows.Threading.Dispatcher
        Private ReadOnly _observe As Action(Of Task, String)
        Private ReadOnly _context As FolderTreeContext

        ''' <summary>
        ''' Geração da árvore. Sem ela, um reload iniciado antes de a sessão
        ''' cair terminaria depois do Clear e repovoaria a árvore com pastas
        ''' de uma sessão que não existe mais.
        ''' </summary>
        Private _generation As Integer = 0

        Private _isLoading As Boolean
        Private _selected As FolderNodeViewModel
        Private _errorMessage As String = ""

        Public Sub New(broker As IOutlookBroker,
                       ui As Global.System.Windows.Threading.Dispatcher,
                       observe As Action(Of Task, String))
            _broker = broker
            _ui = ui
            _observe = observe
            Policy = New FolderVisibilityPolicy()

            _context = New FolderTreeContext(broker, ui, Policy,
                                             AddressOf OnNodeSelected,
                                             AddressOf OnNodeError,
                                             observe)

            ReloadCommand = New AsyncRelayCommand(AddressOf ReloadAsync)
        End Sub

        Public ReadOnly Property Roots As New ObservableCollection(Of FolderNodeViewModel)()
        Public ReadOnly Property ReloadCommand As IAsyncRelayCommand
        Public ReadOnly Property Policy As FolderVisibilityPolicy

        Public Property IsLoading As Boolean
            Get
                Return _isLoading
            End Get
            Private Set(value As Boolean)
                SetProperty(_isLoading, value)
            End Set
        End Property

        Public Property Selected As FolderNodeViewModel
            Get
                Return _selected
            End Get
            Private Set(value As FolderNodeViewModel)
                SetProperty(_selected, value)
            End Set
        End Property

        ''' <summary>
        ''' Mensagem traduzida a partir do <see cref="ErrorKind"/>. O
        ''' <c>Detail</c> do resultado é diagnóstico e não vai para a tela.
        ''' </summary>
        Public Property ErrorMessage As String
            Get
                Return _errorMessage
            End Get
            Private Set(value As String)
                If SetProperty(_errorMessage, value) Then
                    OnPropertyChanged(NameOf(HasError))
                End If
            End Set
        End Property

        Public ReadOnly Property HasError As Boolean
            Get
                Return Not String.IsNullOrEmpty(_errorMessage)
            End Get
        End Property

        Public Async Function ReloadAsync() As Task
            Dim geracao = Interlocked.Increment(_generation)

            IsLoading = True
            ErrorMessage = ""
            Try
                Dim stores = Await _broker.GetStoresAsync(CancellationToken.None)
                If Not Atual(geracao) Then Return

                If Not stores.Succeeded Then
                    ErrorMessage = Traduzir(stores.Kind)
                    Return
                End If

                Dim visiveis As New List(Of FolderInfo)()
                For Each store In stores.Value
                    Dim filhos = Await _broker.GetFolderChildrenAsync(
                        store.RootFolder, CancellationToken.None)
                    If Not Atual(geracao) Then Return
                    If Not filhos.Succeeded Then Continue For

                    visiveis.AddRange(filhos.Value.Where(Function(f) Policy.IsVisible(f)))
                Next

                Await _ui.InvokeAsync(
                    Sub()
                        If Not Atual(geracao) Then Return
                        Roots.Clear()
                        For Each f In visiveis
                            Roots.Add(New FolderNodeViewModel(f, _context))
                        Next
                    End Sub).Task
            Finally
                IsLoading = False
            End Try
        End Function

        ''' <summary>
        ''' Esvazia a árvore e INVALIDA qualquer carga em voo: incrementar a
        ''' geração é o que impede uma resposta em trânsito de repovoar as
        ''' pastas depois de a sessão cair.
        ''' </summary>
        Public Sub Clear()
            Interlocked.Increment(_generation)
            Roots.Clear()
            Selected = Nothing
            ErrorMessage = ""
        End Sub

        Private Function Atual(geracao As Integer) As Boolean
            Return Volatile.Read(_generation) = geracao
        End Function

        Private Sub OnNodeSelected(node As FolderNodeViewModel)
            Selected = node
        End Sub

        Private Sub OnNodeError(kind As ErrorKind)
            _observe(_ui.InvokeAsync(Sub() ErrorMessage = Traduzir(kind)).Task, "folders.error")
        End Sub

        ''' <summary>
        ''' A UI traduz o <see cref="ErrorKind"/>; ela nunca exibe o
        ''' <c>Detail</c>, que é diagnóstico e pode conter dado da caixa.
        ''' </summary>
        Private Shared Function Traduzir(kind As ErrorKind) As String
            Select Case kind
                Case ErrorKind.NotConnected : Return "Sem conexão com o Outlook."
                Case ErrorKind.Busy : Return "O Outlook está ocupado. Tentando de novo em instantes."
                Case ErrorKind.NotFound : Return "A pasta não existe mais."
                Case ErrorKind.Denied : Return "Acesso negado pela política."
                Case ErrorKind.NotImplemented : Return "Ainda não implementado."
                Case Else : Return "Não foi possível ler as pastas."
            End Select
        End Function

    End Class

End Namespace

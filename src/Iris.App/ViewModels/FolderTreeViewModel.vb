Imports System.Collections.ObjectModel
Imports System.Threading
Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' A árvore de pastas.
    '''
    ''' Regras decididas explicitamente, porque "reflete a árvore do Outlook"
    ''' não é critério verificável:
    '''
    '''   • Entram todos os stores que o perfil expõe. Cada um vira uma raiz.
    '''   • A raiz de cada store é a pasta raiz dele, e o primeiro nível é
    '''     carregado junto — abrir o Iris e ver a árvore vazia até clicar
    '''     seria pior que esperar meio segundo.
    '''   • Níveis seguintes carregam sob demanda, um por expansão.
    '''   • A ordem é a que o Outlook devolve. Reordenar alfabeticamente
    '''     brigaria com a ordem que o usuário já conhece do Outlook.
    '''   • Pastas OCULTAS (PR_ATTR_HIDDEN) não aparecem. São internas do
    '''     Outlook — "Conversation Action Settings", "Quick Step Settings",
    '''     "Raiz do Yammer" — e o próprio Outlook não as mostra.
    '''   • Só pastas de E-MAIL aparecem na Fase 1. Calendário, Contatos,
    '''     Tarefas e Observações existem no store, mas exibi-las agora
    '''     seria oferecer uma porta que não abre: elas voltam nas Fases 5
    '''     a 7, quando os módulos existirem.
    ''' </summary>
    Public NotInheritable Class FolderTreeViewModel
        Inherits ObservableObject

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _ui As Global.System.Windows.Threading.Dispatcher

        Private _isLoading As Boolean
        Private _selected As FolderNodeViewModel
        Private _errorMessage As String = ""

        Public Sub New(broker As IOutlookBroker, ui As Global.System.Windows.Threading.Dispatcher)
            _broker = broker
            _ui = ui
            ReloadCommand = New AsyncRelayCommand(AddressOf ReloadAsync)
        End Sub

        Public ReadOnly Property Roots As New ObservableCollection(Of FolderNodeViewModel)()
        Public ReadOnly Property ReloadCommand As IAsyncRelayCommand

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
            Set(value As FolderNodeViewModel)
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

        Public ReadOnly Property IsEmpty As Boolean
            Get
                Return Roots.Count = 0
            End Get
        End Property

        Public Async Function ReloadAsync() As Task
            IsLoading = True
            ErrorMessage = ""
            Try
                Dim stores = Await _broker.GetStoresAsync(CancellationToken.None)
                If Not stores.Succeeded Then
                    ErrorMessage = Traduzir(stores.Kind)
                    Return
                End If

                Dim raizes As New List(Of FolderNodeViewModel)()

                For Each store In stores.Value
                    Dim filhos = Await _broker.GetFolderChildrenAsync(
                        store.RootFolder, CancellationToken.None)
                    If Not filhos.Succeeded Then Continue For

                    For Each f In filhos.Value
                        If Not Exibir(f) Then Continue For
                        raizes.Add(New FolderNodeViewModel(f, _broker, _ui, AddressOf OnNodeError))
                    Next
                Next

                Await _ui.InvokeAsync(
                    Sub()
                        Roots.Clear()
                        For Each r In raizes
                            Roots.Add(r)
                        Next
                        OnPropertyChanged(NameOf(IsEmpty))
                    End Sub).Task
            Finally
                IsLoading = False
            End Try
        End Function

        ''' <summary>
        ''' A política de exibição vive AQUI, não no broker: o broker reporta
        ''' o que existe, a apresentação decide o que mostrar.
        ''' </summary>
        Friend Shared Function Exibir(f As FolderInfo) As Boolean
            If f.IsHidden Then Return False
            Return String.Equals(f.DefaultItemType, "olMailItem", StringComparison.Ordinal)
        End Function

        Public Sub Clear()
            Roots.Clear()
            Selected = Nothing
            ErrorMessage = ""
            OnPropertyChanged(NameOf(IsEmpty))
        End Sub

        Private Sub OnNodeError(kind As ErrorKind, detail As String)
            Dim ignorado = _ui.InvokeAsync(Sub() ErrorMessage = Traduzir(kind))
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

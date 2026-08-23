Imports System.Collections.ObjectModel
Imports System.Diagnostics
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' A lista de mensagens de uma pasta.
    '''
    ''' PAGINAÇÃO VOLÁTIL, assumida e não escondida (FASE1.md seção 5).
    ''' Offset numa pasta viva não é estável: se uma mensagem chega no topo
    ''' entre duas páginas, um item duplica e outro é pulado. A Fase 1 aceita
    ''' isso; o que ela NÃO aceita é pretender o contrário. Por isso:
    '''
    '''   • Toda consulta carrega uma geração.
    '''   • Resposta de geração vencida é DESCARTADA, nunca anexada.
    '''   • Mudança na pasta recarrega do topo em vez de remendar a lista.
    '''
    ''' Se isso não ficar utilizável na prática, a decisão é antecipar a
    ''' Fase 2 — este marco é ponto de decisão, não só entrega.
    ''' </summary>
    Public NotInheritable Class MessageListViewModel
        Inherits ObservableObject

        ''' <summary>
        ''' 30 itens. A Fase 0 mediu ~16 ms por item, então uma página custa
        ''' cerca de meio segundo. Páginas maiores estouram o orçamento de
        ''' 1 s do critério de aceite.
        ''' </summary>
        Public Const PageSize As Integer = 30

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _ui As Global.System.Windows.Threading.Dispatcher
        Private ReadOnly _observe As Action(Of Task, String)

        Private _generation As Long = 0
        Private _folder As FolderKey
        Private _folderName As String = ""
        Private _sort As MessageSort = MessageSort.ReceivedDesc

        Private _loading As Integer = 0
        Private _isLoading As Boolean
        Private _selected As MailSummary
        Private _total As Integer
        Private _hasMore As Boolean
        Private _errorMessage As String = ""
        Private _lastPageMs As Double

        Public Sub New(broker As IOutlookBroker,
                       ui As Global.System.Windows.Threading.Dispatcher,
                       observe As Action(Of Task, String))
            _broker = broker
            _ui = ui
            _observe = observe
            LoadMoreCommand = New AsyncRelayCommand(AddressOf LoadMoreAsync, AddressOf PodeCarregarMais)
            ReloadCommand = New AsyncRelayCommand(Function() ReloadAsync(preservarSelecao:=True))
        End Sub

        Public ReadOnly Property Messages As New ObservableCollection(Of MailSummary)()
        Public ReadOnly Property LoadMoreCommand As IAsyncRelayCommand
        Public ReadOnly Property ReloadCommand As IAsyncRelayCommand

        Public Property FolderName As String
            Get
                Return _folderName
            End Get
            Private Set(value As String)
                SetProperty(_folderName, value)
            End Set
        End Property

        Public Property IsLoading As Boolean
            Get
                Return _isLoading
            End Get
            Private Set(value As Boolean)
                SetProperty(_isLoading, value)
            End Set
        End Property

        ''' <summary>
        ''' ListBox.SelectedItem aceita TwoWay, diferente do TreeView. Por
        ''' isso aqui a seleção é uma propriedade simples.
        ''' </summary>
        Public Property Selected As MailSummary
            Get
                Return _selected
            End Get
            Set(value As MailSummary)
                SetProperty(_selected, value)
            End Set
        End Property

        Public Property Total As Integer
            Get
                Return _total
            End Get
            Private Set(value As Integer)
                If SetProperty(_total, value) Then
                    OnPropertyChanged(NameOf(StatusLine))
                End If
            End Set
        End Property

        Public Property HasMore As Boolean
            Get
                Return _hasMore
            End Get
            Private Set(value As Boolean)
                If SetProperty(_hasMore, value) Then
                    LoadMoreCommand.NotifyCanExecuteChanged()
                End If
            End Set
        End Property

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

        ''' <summary>
        ''' Tempo da última página. Existe para o critério de aceite do 1.3
        ''' ser verificável olhando a tela, e não por impressão.
        ''' </summary>
        Public Property LastPageMs As Double
            Get
                Return _lastPageMs
            End Get
            Private Set(value As Double)
                If SetProperty(_lastPageMs, value) Then
                    OnPropertyChanged(NameOf(StatusLine))
                End If
            End Set
        End Property

        Public ReadOnly Property StatusLine As String
            Get
                If _total = 0 Then Return ""
                Return $"{Messages.Count} de {_total} · última página {_lastPageMs:0} ms"
            End Get
        End Property

        Public ReadOnly Property IsEmpty As Boolean
            Get
                Return Messages.Count = 0
            End Get
        End Property

        ' ===================================================================

        Public Async Function ShowFolderAsync(folder As FolderKey, nome As String) As Task
            _folder = folder
            FolderName = nome
            Await ReloadAsync(preservarSelecao:=False)
        End Function

        Public Sub Clear()
            Interlocked.Increment(_generation)
            _folder = Nothing
            FolderName = ""
            Messages.Clear()
            Selected = Nothing
            Total = 0
            HasMore = False
            ErrorMessage = ""
            OnPropertyChanged(NameOf(IsEmpty))
        End Sub

        ''' <summary>
        ''' Recarrega do topo.
        '''
        ''' A seleção é preservada por <see cref="ItemKey"/>, não pelo objeto:
        ''' o DTO é recriado a cada leitura, então guardar a referência
        ''' perderia a seleção em toda recarga.
        ''' </summary>
        Public Async Function ReloadAsync(preservarSelecao As Boolean) As Task
            Dim chaveSelecionada = If(preservarSelecao AndAlso Selected IsNot Nothing,
                                      Selected.Key, Nothing)

            Dim geracao = Interlocked.Increment(_generation)

            Await OnUiAsync(
                Sub()
                    Messages.Clear()
                    ErrorMessage = ""
                    OnPropertyChanged(NameOf(IsEmpty))
                End Sub)

            Await CarregarPaginaAsync(geracao, offset:=0, chaveParaSelecionar:=chaveSelecionada)
        End Function

        Private Function PodeCarregarMais() As Boolean
            Return _hasMore AndAlso Volatile.Read(_loading) = 0
        End Function

        Public Async Function LoadMoreAsync() As Task
            If Not PodeCarregarMais() Then Return
            Await CarregarPaginaAsync(Volatile.Read(_generation), Messages.Count, Nothing)
        End Function

        ''' <summary>
        ''' Uma requisição por vez. Sem esta trava, rolagem rápida enfileira
        ''' páginas na fila única da STA e trava a interface — o F1-F.
        ''' </summary>
        Private Async Function CarregarPaginaAsync(geracao As Long, offset As Integer,
                                                   chaveParaSelecionar As ItemKey) As Task
            If _folder Is Nothing Then Return
            If Interlocked.CompareExchange(_loading, 1, 0) <> 0 Then Return

            IsLoading = True
            LoadMoreCommand.NotifyCanExecuteChanged()

            Try
                Dim consulta = New MessageQuery(_folder, _sort, geracao)
                Dim cronometro = Stopwatch.StartNew()
                Dim resultado = Await _broker.GetMessagePageAsync(
                    consulta, offset, PageSize, CancellationToken.None)
                cronometro.Stop()

                ' Geração vencida: a resposta é de uma pasta ou de uma
                ' sessão que o usuário já deixou para trás.
                If Volatile.Read(_generation) <> geracao Then Return

                If Not resultado.Succeeded Then
                    Await OnUiAsync(Sub() ErrorMessage = Traduzir(resultado.Kind))
                    Return
                End If

                Dim pagina = resultado.Value

                Await OnUiAsync(
                    Sub()
                        If Volatile.Read(_generation) <> geracao Then Return

                        For Each m In pagina.Items
                            Messages.Add(m)
                        Next

                        Total = pagina.TotalAtRead
                        HasMore = pagina.HasMore
                        LastPageMs = cronometro.Elapsed.TotalMilliseconds
                        OnPropertyChanged(NameOf(IsEmpty))
                        OnPropertyChanged(NameOf(StatusLine))

                        If chaveParaSelecionar IsNot Nothing Then
                            Selected = Messages.FirstOrDefault(
                                Function(m) chaveParaSelecionar.Equals(m.Key))
                        End If
                    End Sub)
            Finally
                Interlocked.Exchange(_loading, 0)
                IsLoading = False
                LoadMoreCommand.NotifyCanExecuteChanged()
            End Try
        End Function

        ''' <summary>
        ''' A pasta mudou. Recarrega do topo em vez de aplicar a transição do
        ''' evento: a Fase 0 mostrou que o estado lido depois pode não ser o
        ''' que causou o evento, e que a ordem dos eventos não é confiável.
        ''' </summary>
        Public Sub OnFolderInvalidated(folder As FolderKey)
            If _folder Is Nothing OrElse Not _folder.Equals(folder) Then Return
            _observe(ReloadAsync(preservarSelecao:=True), "messages.invalidated")
        End Sub

        Private Async Function OnUiAsync(action As Action) As Task
            If _ui.CheckAccess() Then
                action()
                Return
            End If
            Await _ui.InvokeAsync(action).Task
        End Function

        Private Shared Function Traduzir(kind As ErrorKind) As String
            Select Case kind
                Case ErrorKind.NotConnected : Return "Sem conexão com o Outlook."
                Case ErrorKind.Busy : Return "O Outlook está ocupado."
                Case ErrorKind.NotFound : Return "A pasta não existe mais."
                Case ErrorKind.Denied : Return "Acesso negado pela política."
                Case Else : Return "Não foi possível ler as mensagens."
            End Select
        End Function

    End Class

End Namespace

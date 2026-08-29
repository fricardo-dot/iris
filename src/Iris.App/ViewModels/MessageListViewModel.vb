Imports System.Collections.ObjectModel
Imports System.Collections.Generic
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
    ''' Um pedido de página, com tudo que ele precisa capturado junto.
    '''
    ''' Pasta, geração, cursor e seleção viajam como UMA unidade. Ler
    ''' <c>_folder</c> mutável depois de adquirir a trava permitia que uma
    ''' operação de uma geração acabasse consultando a pasta de outra.
    ''' </summary>
    Friend NotInheritable Class PageRequest
        Public ReadOnly Property Folder As FolderKey
        Public ReadOnly Property Sort As MessageSort
        Public ReadOnly Property Generation As Long
        ''' <summary>Cursor opaco. Nothing pede a primeira pagina.</summary>
        Public ReadOnly Property Cursor As String
        Public ReadOnly Property SelectionKey As ItemKey
        Public ReadOnly Property IsReload As Boolean

        Public Sub New(folder As FolderKey, sort As MessageSort, generation As Long,
                       cursor As String, selectionKey As ItemKey, isReload As Boolean)
            Me.Folder = folder
            Me.Sort = sort
            Me.Generation = generation
            Me.Cursor = cursor
            Me.SelectionKey = selectionKey
            Me.IsReload = isReload
        End Sub
    End Class

    ''' <summary>
    ''' A lista de mensagens de uma pasta.
    '''
    ''' PAGINAÇÃO VOLÁTIL, assumida e não escondida (FASE1.md seção 5).
    ''' Offset numa pasta viva não é estável: se uma mensagem chega no topo
    ''' entre duas páginas, um item duplica e outro é pulado. A Fase 1 aceita
    ''' isso; o que ela NÃO aceita é pretender o contrário.
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
        Private ReadOnly _gate As New Object()

        Private _generation As Long = 0
        Private _folder As FolderKey
        Private _folderName As String = ""
        Private _sort As MessageSort = MessageSort.ReceivedDesc
        ''' <summary>
        ''' Cursor da proxima pagina. Nothing significa 'primeira pagina',
        ''' e por isso PRECISA ser zerado em toda troca de pasta, de
        ''' ordenacao e em recarga: cursor sobrevivente pediria a
        ''' continuacao de uma lista que nao esta mais na tela.
        ''' </summary>
        Private _nextCursor As String = Nothing

        ''' <summary>
        ''' Um pedido em execução e, no máximo, UM pendente — o último vence.
        '''
        ''' A versão anterior simplesmente desistia quando havia carga em
        ''' curso. O efeito era grave: trocar de pasta durante um
        ''' carregamento deixava a lista VAZIA até o usuário clicar de novo,
        ''' porque o pedido novo era descartado e o antigo invalidado pela
        ''' geração.
        ''' </summary>
        Private _running As Boolean
        Private _pending As PageRequest

        Private _isLoading As Boolean
        Private _isRestoringSelection As Boolean
        Private _selected As MessageRowViewModel
        Private _total As Integer
        Private _hasMore As Boolean
        Private _hasFolder As Boolean
        Private _errorMessage As String = ""
        Private _lastPageMs As Double
        Private _skipped As Integer

        ''' <summary>
        ''' Quantas células ausentes viraram valor nas páginas desta pasta.
        '''
        ''' ------------------------------------------------------------------
        ''' Existe porque o contador nasceu, em 28/08/2026, <b>sem consumidor</b>
        ''' — o número era calculado no <c>MessagePage</c> e morria ali. A
        ''' revisão externa foi direta: <i>"o silêncio não saiu; o número morre
        ''' no MessagePage"</i>.
        '''
        ''' É literalmente o erro que este projeto conta seis vezes: proteção que
        ''' existe e não está ligada a nada. Um contador de fabricação que
        ''' ninguém lê é pior que nenhum, porque dá a impressão de que alguém
        ''' está olhando.
        ''' </summary>
        Private _fabricadas As Integer

        Public Sub New(broker As IOutlookBroker,
                       ui As Global.System.Windows.Threading.Dispatcher,
                       observe As Action(Of Task, String))
            _broker = broker
            _ui = ui
            _observe = observe
            LoadMoreCommand = New AsyncRelayCommand(AddressOf LoadMoreAsync, AddressOf PodeCarregarMais)
            ReloadCommand = New AsyncRelayCommand(Function() ReloadAsync(preservarSelecao:=True))
        End Sub

        ''' <summary>
        ''' MessageRowViewModel, nao MailSummary: o DTO nao notifica mudanca,
        ''' entao marcar como lida nao repintaria a linha.
        ''' </summary>
        Public ReadOnly Property Messages As New ObservableCollection(Of MessageRowViewModel)()
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
                If SetProperty(_isLoading, value) Then AtualizarEstados()
            End Set
        End Property

        ''' <summary>ListBox.SelectedItem aceita TwoWay, diferente do TreeView.</summary>
        Public Property Selected As MessageRowViewModel
            Get
                Return _selected
            End Get
            Set(value As MessageRowViewModel)
                SetProperty(_selected, value)
            End Set
        End Property

        Public Property Total As Integer
            Get
                Return _total
            End Get
            Private Set(value As Integer)
                If SetProperty(_total, value) Then OnPropertyChanged(NameOf(StatusLine))
            End Set
        End Property

        Public Property HasMore As Boolean
            Get
                Return _hasMore
            End Get
            Private Set(value As Boolean)
                If SetProperty(_hasMore, value) Then LoadMoreCommand.NotifyCanExecuteChanged()
            End Set
        End Property

        ''' <summary>
        ''' A lista esta trocando o conteudo e vai restaurar a selecao pela
        ''' chave. Existe para distinguir "a colecao foi esvaziada por dentro"
        ''' de "o usuario desmarcou".
        '''
        ''' Antes isso era inferido de IsLoading, o que era amplo demais:
        ''' IsLoading tambem e verdadeiro durante "Carregar mais", quando
        ''' nenhuma selecao esta sendo restaurada, e nesse intervalo uma
        ''' desmarcacao legitima era ignorada.
        ''' </summary>
        Public ReadOnly Property IsRestoringSelection As Boolean
            Get
                Return _isRestoringSelection
            End Get
        End Property

        ''' <summary>Uma pasta foi escolhida — ainda que esteja vazia.</summary>
        Public Property HasFolder As Boolean
            Get
                Return _hasFolder
            End Get
            Private Set(value As Boolean)
                If SetProperty(_hasFolder, value) Then AtualizarEstados()
            End Set
        End Property

        Public Property ErrorMessage As String
            Get
                Return _errorMessage
            End Get
            Private Set(value As String)
                If SetProperty(_errorMessage, value) Then
                    OnPropertyChanged(NameOf(HasError))
                    AtualizarEstados()
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
                If SetProperty(_lastPageMs, value) Then OnPropertyChanged(NameOf(StatusLine))
            End Set
        End Property

        Public ReadOnly Property StatusLine As String
            Get
                If Not _hasFolder Then Return ""
                Dim linha = $"{Messages.Count} de {_total} · última página {_lastPageMs:0} ms"
                ' "28 de 30" sem explicação vira mistério.
                If _skipped > 0 Then linha &= $" · {_skipped} item(ns) ignorado(s)"
                ' AUSENCIA QUE VIROU VALOR. Tamanho 0, "nao lida" e "sem anexo"
                ' que na verdade sao "o Outlook nao respondeu". A medicao de
                ' 28/08 achou ZERO destes em 1.109 linhas -- entao esta frase
                ' nao deve aparecer, e se aparecer e informacao de verdade.
                If _fabricadas > 0 Then
                    linha &= $" · {_fabricadas} campo(s) que o Outlook nao entregou"
                End If
                Return linha
            End Get
        End Property

        ''' <summary>
        ''' Três estados que antes eram um só. "IsEmpty" dizia ao mesmo tempo
        ''' "escolha uma pasta", "esta pasta está vazia", "está carregando" e
        ''' "deu erro" — e a tela mostrava "Selecione uma pasta" para uma
        ''' caixa legitimamente vazia.
        ''' </summary>
        Public ReadOnly Property ShowSelectPrompt As Boolean
            Get
                Return Not _hasFolder AndAlso Not HasError
            End Get
        End Property

        Public ReadOnly Property ShowEmptyFolder As Boolean
            Get
                Return _hasFolder AndAlso Messages.Count = 0 AndAlso
                       Not _isLoading AndAlso Not HasError
            End Get
        End Property

        ''' <summary>
        ''' <b>Lista vazia na tela não é pasta vazia.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>A TELA DIZIA AS DUAS COISAS AO MESMO TEMPO</b>
        '''
        ''' O texto era fixo — "Esta pasta está vazia" —, e saía sempre que a
        ''' lista convertida ficava sem linha. Com uma página de
        ''' <c>TotalAtStart = 1</c> e <c>SkippedCount = 1</c>, a mesma tela
        ''' mostrava <i>"Esta pasta está vazia"</i> no meio e
        ''' <i>"0 de 1 · 1 item ignorado"</i> no rodapé. Uma das duas frases
        ''' estava mentindo, e era a de cima — a que ocupa a tela inteira.
        '''
        ''' É a mesma família do <c>0 compromisso(s)</c> da agenda e do
        ''' <i>"calendário vazio"</i> do roteiro: <b>tratar "não observei" como
        ''' "observei e não há"</b>. Três instâncias, três superfícies, o mesmo
        ''' erro — e este é o único que estava na tela principal.
        '''
        ''' Os três casos, em ordem de força do que se pode afirmar:
        ''' <list type="bullet">
        ''' <item>a leitura perdeu item ⇒ não se afirma nada sobre a pasta;</item>
        ''' <item>a pasta declara N e a leitura trouxe zero ⇒ diz-se o N;</item>
        ''' <item>a pasta declara zero e nada foi perdido ⇒ aí sim, vazia.</item>
        ''' </list>
        ''' </summary>
        Public ReadOnly Property EmptyMessage As String
            Get
                If _skipped > 0 Then
                    Return $"Nenhuma mensagem para mostrar — {_skipped} item(ns) desta " &
                           "pasta a leitura não conseguiu trazer."
                End If
                If _fabricadas > 0 Then
                    Return "Nenhuma mensagem para mostrar, e a leitura teve campos que o " &
                           "Outlook não entregou."
                End If
                If _total > 0 Then
                    Return $"Nenhuma mensagem para mostrar, e a pasta declara {_total} " &
                           "item(ns). A leitura não trouxe nenhum."
                End If
                Return "Esta pasta está vazia."
            End Get
        End Property

        Private Sub AtualizarEstados()
            OnPropertyChanged(NameOf(ShowSelectPrompt))
            OnPropertyChanged(NameOf(ShowEmptyFolder))
            ' A MENSAGEM MUDA COM _skipped E _total, que mudam pagina a pagina.
            ' Sem esta linha o texto ficaria correto por acaso -- o da primeira
            ' pagina, mostrado para sempre.
            OnPropertyChanged(NameOf(EmptyMessage))
            OnPropertyChanged(NameOf(StatusLine))
        End Sub

        ' ===================================================================

        Public Async Function ShowFolderAsync(folder As FolderKey, nome As String) As Task
            _folder = folder
            FolderName = nome
            HasFolder = True
            Await ReloadAsync(preservarSelecao:=False)
        End Function

        Public Sub Clear()
            Interlocked.Increment(_generation)
            SyncLock _gate
                _pending = Nothing
            End SyncLock

            _folder = Nothing
            _nextCursor = Nothing
            FolderName = ""
            HasFolder = False
            Messages.Clear()
            Selected = Nothing
            Total = 0
            HasMore = False
            ErrorMessage = ""
            _skipped = 0
            _fabricadas = 0
            AtualizarEstados()
        End Sub

        Public Async Function ReloadAsync(preservarSelecao As Boolean) As Task
            If _folder Is Nothing Then Return

            Dim chave = If(preservarSelecao AndAlso Selected IsNot Nothing, Selected.Key, Nothing)
            Dim geracao = Interlocked.Increment(_generation)

            _nextCursor = Nothing
            Await Despachar(New PageRequest(_folder, _sort, geracao, Nothing, chave, isReload:=True))
        End Function

        Private Function PodeCarregarMais() As Boolean
            Return _hasMore AndAlso Not _isLoading
        End Function

        Public Async Function LoadMoreAsync() As Task
            If Not _hasMore OrElse _folder Is Nothing Then Return

            ' NextOffset, não Messages.Count: o broker examina posições e
            ' pode devolver menos DTOs do que examinou.
            Await Despachar(New PageRequest(_folder, _sort, Volatile.Read(_generation),
                                            _nextCursor, Nothing, isReload:=False))
        End Function

        ''' <summary>
        ''' Uma operação por vez, com o último pedido sempre preservado.
        ''' </summary>
        Private Async Function Despachar(pedido As PageRequest) As Task
            SyncLock _gate
                If _running Then
                    ' O último vence, e nada se perde: quem chegou agora
                    ' substitui o pendente e será executado ao fim da
                    ' operação corrente.
                    _pending = pedido
                    Return
                End If
                _running = True
            End SyncLock

            Try
                Dim atual = pedido
                Do
                    Await ExecutarAsync(atual)

                    SyncLock _gate
                        atual = _pending
                        _pending = Nothing

                        ' Observar a fila vazia e desligar _running precisa
                        ' ser ATOMICO. Com o desligamento no Finally havia
                        ' uma janela: o worker saia do lock vendo vazio, um
                        ' pedido novo chegava, via _running=True, guardava-se
                        ' como pendente — e entao o Finally desligava
                        ' _running sem que ninguem o executasse. O pedido
                        ' ficava abandonado ate outra acao do usuario.
                        If atual Is Nothing Then
                            _running = False
                            Return
                        End If
                    End SyncLock
                Loop
            Catch
                ' Rede de seguranca so para excecao: no caminho normal o
                ' desligamento ja aconteceu dentro do lock.
                SyncLock _gate
                    _running = False
                End SyncLock
                Throw
            End Try
        End Function

        Private Async Function ExecutarAsync(pedido As PageRequest) As Task
            ' Revalida ANTES de ir ao broker: o pedido pode ter envelhecido
            ' enquanto esperava na fila.
            If Volatile.Read(_generation) <> pedido.Generation Then Return

            IsLoading = True
            LoadMoreCommand.NotifyCanExecuteChanged()

            Try
                If pedido.IsReload Then
                    Await OnUiAsync(
                        Sub()
                            If Volatile.Read(_generation) <> pedido.Generation Then Return
                            ' Marcado ANTES do Clear: limpar a colecao anula a
                            ' selecao do ListBox, e quem escuta precisa saber
                            ' que aquilo e interno, nao o usuario.
                            _isRestoringSelection = pedido.SelectionKey IsNot Nothing
                            Messages.Clear()
                            ErrorMessage = ""
                            _skipped = 0
                            _fabricadas = 0
                            AtualizarEstados()
                        End Sub)
                End If

                Dim consulta = New MessageQuery(pedido.Folder, pedido.Sort, pedido.Generation)
                Dim cronometro = Stopwatch.StartNew()
                Dim resultado = Await _broker.GetMessagePageAsync(
                    consulta, pedido.Cursor, PageSize, CancellationToken.None)
                cronometro.Stop()

                If Volatile.Read(_generation) <> pedido.Generation Then Return

                If Not resultado.Succeeded Then
                    Await OnUiAsync(Sub() ErrorMessage = Traduzir(resultado.Kind))
                    Return
                End If

                Dim pagina = resultado.Value
                ' A geracao da PAGINA e conferida aqui, antes de qualquer
                ' coisa tocar na colecao. Antes isso era feito depois da
                ' deduplicacao e do Add: se o broker devolvesse geracao
                ' vencida, os itens ja tinham entrado na tela.
                If pagina.Generation <> pedido.Generation Then Return

                Await OnUiAsync(
                    Sub()
                        If Volatile.Read(_generation) <> pedido.Generation Then Return

                        ' Deduplicação por chave: a paginação é volátil, e
                        ' uma mensagem que chegou no topo entre páginas
                        ' apareceria duas vezes na tela.
                        Dim existentes = New HashSet(Of ItemKey)(Messages.Select(Function(m) m.Key))
                        For Each m In pagina.Items
                            If existentes.Add(m.Key) Then Messages.Add(New MessageRowViewModel(m))
                        Next

                        _nextCursor = pagina.NextCursor
                        _skipped += pagina.SkippedCount
                        _fabricadas += pagina.FabricatedCells
                        ' TotalAtStart so vem na primeira pagina; nas demais
                        ' o valor anterior e mantido.
                        If pagina.TotalAtStart.HasValue Then Total = pagina.TotalAtStart.Value
                        HasMore = pagina.HasMore
                        LastPageMs = cronometro.Elapsed.TotalMilliseconds
                        AtualizarEstados()

                        If pedido.SelectionKey IsNot Nothing Then
                            ' Limite conhecido: só reencontra a seleção se
                            ' ela estiver na parte recarregada.
                            Selected = Messages.FirstOrDefault(
                                Function(m) pedido.SelectionKey.Equals(m.Key))
                            _isRestoringSelection = False
                        End If
                    End Sub)
            Finally
                ' Sempre desmarcado: se a pagina falhou ou a geracao venceu,
                ' a restauracao nao vai acontecer e deixar a flag ligada
                ' faria a UI ignorar desmarcacoes de verdade para sempre.
                _isRestoringSelection = False
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

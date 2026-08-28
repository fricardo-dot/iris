Imports System.Collections.ObjectModel
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    Public Enum NodeLoadState
        Unloaded
        Loading
        Loaded
    End Enum

    ''' <summary>
    ''' Um nó da árvore de pastas, carregado sob demanda.
    '''
    ''' Carregar a árvore inteira na abertura seria simples e errado: a Fase
    ''' 0 mediu ~16 ms por item só de montar DTO, e uma caixa com muitas
    ''' pastas pagaria isso a cada abertura. Cada nível é buscado quando
    ''' expande, uma vez só.
    ''' </summary>
    Public NotInheritable Class FolderNodeViewModel
        Inherits ObservableObject

        Private ReadOnly _context As FolderTreeContext

        Private _state As NodeLoadState = NodeLoadState.Unloaded
        Private _loadGeneration As Integer = 0
        Private ReadOnly _gate As New Object()

        Private _isExpanded As Boolean
        Private _isSelected As Boolean
        Private _isLoading As Boolean
        Private _unreadCount As Integer
        Private _itemCount As Integer
        Private _hasUnrealizedChildren As Boolean

        Public Sub New(info As FolderInfo, context As FolderTreeContext,
                       Optional parent As FolderNodeViewModel = Nothing)
            _context = context
            Me.Parent = parent

            Key = info.Key
            Name = info.Name
            ContentKind = info.ContentKind
            _unreadCount = info.UnreadCount
            _itemCount = info.ItemCount
            _hasUnrealizedChildren = info.HasChildren
        End Sub

        ''' <summary>
        ''' O nó de cima, ou Nothing na raiz.
        '''
        ''' Existe para reconstruir o CAMINHO de uma pasta. Depois de uma
        ''' troca de sessão a árvore é refeita do zero, e sem o caminho só dá
        ''' para reencontrar pasta de topo — subpasta é materializada ao
        ''' expandir, então nem existe como nó para procurar.
        ''' </summary>
        Public ReadOnly Property Parent As FolderNodeViewModel

        Public ReadOnly Property Key As FolderKey
        Public ReadOnly Property Name As String

        ''' <summary>
        ''' <b>O que a pasta guarda</b> — correio, calendário, contatos, tarefas.
        '''
        ''' O broker já media isto e o nó jogava fora, exatamente como o
        ''' <c>MessageClass</c> era medido e substituído por constante na
        ''' varredura. Dado que se mede e se descarta é dado que alguém vai
        ''' acabar reinventando pior.
        '''
        ''' Quem precisou dele foi a agenda: apontá-la para a Caixa de Entrada
        ''' faria a tela mostrar "0 compromissos" sobre uma pasta que não tem
        ''' compromisso por definição — e zero por engano é indistinguível de
        ''' zero por medida.
        ''' </summary>
        Public ReadOnly Property ContentKind As FolderContentKind
        Public ReadOnly Property Children As New ObservableCollection(Of FolderNodeViewModel)()

        ''' <summary>
        ''' Existem filhos ainda não buscados.
        '''
        ''' Substitui o nó-marcador da versão anterior. Um filho falso na
        ''' coleção é um item de verdade para o WPF: ganha container, entra
        ''' na navegação por teclado e aparece para a automação, sem nome e
        ''' sem sentido. O template usa esta propriedade para desenhar a seta
        ''' sem que nada falso exista na árvore.
        ''' </summary>
        Public Property HasUnrealizedChildren As Boolean
            Get
                Return _hasUnrealizedChildren
            End Get
            Private Set(value As Boolean)
                If SetProperty(_hasUnrealizedChildren, value) Then
                    OnPropertyChanged(NameOf(CanExpand))
                End If
            End Set
        End Property

        Public ReadOnly Property CanExpand As Boolean
            Get
                Return _hasUnrealizedChildren OrElse Children.Count > 0
            End Get
        End Property

        ''' <summary>
        ''' Eventualmente consistente por desenho: o Outlook atualiza a
        ''' contagem de forma assíncrona, e o Iris não vai bloquear a árvore
        ''' para conferir um número que muda sozinho.
        ''' </summary>
        Public Property UnreadCount As Integer
            Get
                Return _unreadCount
            End Get
            Set(value As Integer)
                If SetProperty(_unreadCount, value) Then
                    OnPropertyChanged(NameOf(HasUnread))
                End If
            End Set
        End Property

        Public Property ItemCount As Integer
            Get
                Return _itemCount
            End Get
            Set(value As Integer)
                SetProperty(_itemCount, value)
            End Set
        End Property

        Public ReadOnly Property HasUnread As Boolean
            Get
                Return _unreadCount > 0
            End Get
        End Property

        Public Property IsExpanded As Boolean
            Get
                Return _isExpanded
            End Get
            Set(value As Boolean)
                If Not SetProperty(_isExpanded, value) Then Return
                If value Then BeginLoadChildren()
            End Set
        End Property

        ''' <summary>
        ''' TwoWay com o TreeViewItem. <c>TreeView.SelectedItem</c> é somente
        ''' leitura, então ligar a seleção ao ViewModel passa por aqui — sem
        ''' isto, clicar numa pasta não informava nada a ninguém.
        ''' </summary>
        Public Property IsSelected As Boolean
            Get
                Return _isSelected
            End Get
            Set(value As Boolean)
                If Not SetProperty(_isSelected, value) Then Return
                If value Then _context.NotifySelected(Me)
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

        ' ===================================================================
        ' Carregamento
        ' ===================================================================

        ''' <summary>
        ''' Sinal da carga em andamento. Quem chega no meio espera ESTE, em
        ''' vez de desistir.
        '''
        ''' Sem ele, encontrar o nó em Loading fazia EnsureChildrenAsync
        ''' voltar na hora — e quem estava restaurando um caminho descia para
        ''' um nível ainda vazio e concluía "a pasta não existe". Confundir
        ''' "ainda não terminou" com "não existe" deixa o usuário sem seleção
        ''' e sem assinatura, em silêncio.
        ''' </summary>
        Private _sinalDeCarga As TaskCompletionSource(Of Boolean)

        ''' <summary>
        ''' Garante os filhos carregados e DÁ PARA ESPERAR.
        '''
        ''' O caminho normal — expandir com o mouse — dispara e esquece,
        ''' porque a UI não pode travar. Restaurar um caminho precisa do
        ''' contrário: só dá para procurar no nível seguinte depois que o
        ''' atual terminou.
        ''' </summary>
        Public Async Function EnsureChildrenAsync() As Task
            If Not CanExpand Then Return

            Dim geracao As Integer
            Dim esperar As TaskCompletionSource(Of Boolean) = Nothing
            Dim meu As TaskCompletionSource(Of Boolean) = Nothing

            SyncLock _gate
                ' Unloaded / Loading / Loaded em vez de um booleano: antes,
                ' "já comecei" e "já terminei" eram o mesmo valor, e uma
                ' invalidação no meio do caminho deixava o nó inconsistente.
                If _state = NodeLoadState.Loaded Then Return

                If _state = NodeLoadState.Loading Then
                    esperar = _sinalDeCarga
                Else
                    _state = NodeLoadState.Loading
                    _loadGeneration += 1
                    geracao = _loadGeneration
                    meu = New TaskCompletionSource(Of Boolean)()
                    _sinalDeCarga = meu
                End If
            End SyncLock

            If esperar IsNot Nothing Then
                Await esperar.Task
                Return
            End If

            ' Estado Loading sem sinal: não deveria acontecer, mas voltar
            ' calado aqui reintroduziria o defeito.
            If meu Is Nothing Then Return

            Try
                Await LoadChildrenAsync(geracao)
            Catch
                ' Quem espera decide pelo estado do nó, não pela exceção.
            Finally
                meu.TrySetResult(True)
            End Try
        End Function

        ''' <summary>
        ''' Expandir com o mouse: dispara e esquece, porque a UI não pode
        ''' travar esperando o Outlook. Passa pelo MESMO caminho, para não
        ''' existirem duas máquinas de estado de carga.
        ''' </summary>
        Private Sub BeginLoadChildren()
            _context.Observe(EnsureChildrenAsync(), "folders.loadChildren")
        End Sub

        ''' <summary>
        ''' A geração impede uma resposta antiga de repovoar o nó depois de
        ''' ele ter sido invalidado ou de a sessão ter caído.
        ''' </summary>
        Private Async Function LoadChildrenAsync(geracao As Integer) As Task
            IsLoading = True
            Try
                Dim resultado = Await _context.Broker.GetFolderChildrenAsync(Key, CancellationToken.None)

                If Not Atual(geracao) Then Return

                If Not resultado.Succeeded Then
                    ' Libera para tentar de novo: falhar por Outlook ocupado
                    ' não pode condenar o nó a ficar vazio para sempre.
                    Recolher(geracao)
                    _context.ReportError(resultado.Kind)
                    Return
                End If

                Dim visiveis = resultado.Value.Where(Function(f) _context.Policy.IsVisible(f)).ToList()

                Await _context.Ui.InvokeAsync(
                    Sub()
                        If Not Atual(geracao) Then Return
                        Children.Clear()
                        For Each f In visiveis
                            Children.Add(New FolderNodeViewModel(f, _context, Me))
                        Next
                        HasUnrealizedChildren = False
                        OnPropertyChanged(NameOf(CanExpand))
                        SyncLock _gate
                            _state = NodeLoadState.Loaded
                        End SyncLock
                    End Sub).Task

            Catch ex As Exception
                ' Sem este Catch, uma exceção real deixava o nó preso em
                ' Loading para sempre, e a Task era descartada em silêncio.
                Recolher(geracao)
                _context.ReportError(ErrorKind.Unexpected)
                Throw
            Finally
                IsLoading = False
            End Try
        End Function

        Private Function Atual(geracao As Integer) As Boolean
            SyncLock _gate
                Return _loadGeneration = geracao
            End SyncLock
        End Function

        Private Sub Recolher(geracao As Integer)
            SyncLock _gate
                If _loadGeneration <> geracao Then Return
                _state = NodeLoadState.Unloaded
            End SyncLock
        End Sub

        ''' <summary>
        ''' Descarta os filhos e permite buscar de novo.
        '''
        ''' PRESERVA a expansão de propósito: o usuário abriu aquele ramo, e
        ''' recolhê-lo porque uma mensagem chegou seria puni-lo por usar o
        ''' aplicativo. Precisa rodar na thread de UI, porque mexe em
        ''' ObservableCollection.
        ''' </summary>
        Public Sub Invalidate()
            SyncLock _gate
                _loadGeneration += 1
                _state = NodeLoadState.Unloaded
            End SyncLock

            Children.Clear()
            HasUnrealizedChildren = True
            OnPropertyChanged(NameOf(CanExpand))

            If _isExpanded Then BeginLoadChildren()
        End Sub

        ''' <summary>Este nó e a subárvore já materializada.</summary>
        Public Iterator Function Descendants() As IEnumerable(Of FolderNodeViewModel)
            Yield Me
            For Each c In Children
                For Each d In c.Descendants()
                    Yield d
                Next
            Next
        End Function

    End Class

End Namespace

Imports System.Collections.ObjectModel
Imports System.Threading
Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' Um nó da árvore de pastas, carregado sob demanda.
    '''
    ''' Carregar a árvore inteira na abertura seria simples e errado: uma
    ''' caixa com muitas pastas custaria segundos, e a Fase 0 mediu ~16 ms
    ''' por item só de montar DTO. Cada nível é buscado quando expande, uma
    ''' vez só.
    ''' </summary>
    Public NotInheritable Class FolderNodeViewModel
        Inherits ObservableObject

        ''' <summary>
        ''' Filho falso, para o TreeView desenhar a seta de expansão sem que
        ''' os filhos reais existam ainda. É substituído no primeiro
        ''' expandir.
        ''' </summary>
        Private Shared ReadOnly Marcador As FolderNodeViewModel = New FolderNodeViewModel()

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _ui As Global.System.Windows.Threading.Dispatcher
        Private ReadOnly _onError As Action(Of ErrorKind, String)

        Private _isExpanded As Boolean
        Private _isLoading As Boolean
        Private _loaded As Integer = 0

        Private Sub New()
            ' Só para o marcador.
            Name = ""
        End Sub

        Public Sub New(info As FolderInfo, broker As IOutlookBroker,
                       ui As Global.System.Windows.Threading.Dispatcher, onError As Action(Of ErrorKind, String))
            _broker = broker
            _ui = ui
            _onError = onError

            Key = info.Key
            Name = info.Name
            UnreadCount = info.UnreadCount
            ItemCount = info.ItemCount

            If info.HasChildren Then Children.Add(Marcador)
        End Sub

        Public ReadOnly Property Key As FolderKey
        Public ReadOnly Property Name As String
        Public ReadOnly Property Children As New ObservableCollection(Of FolderNodeViewModel)()

        ''' <summary>
        ''' Eventualmente consistente por desenho: o Outlook atualiza a
        ''' contagem de forma assíncrona, e o Iris não vai bloquear a árvore
        ''' para conferir um número que muda sozinho.
        ''' </summary>
        Public Property UnreadCount As Integer
        Public Property ItemCount As Integer

        Public ReadOnly Property HasUnread As Boolean
            Get
                Return UnreadCount > 0
            End Get
        End Property

        Public Property IsExpanded As Boolean
            Get
                Return _isExpanded
            End Get
            Set(value As Boolean)
                If Not SetProperty(_isExpanded, value) Then Return
                If value Then LoadChildrenOnce()
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
        ''' Uma vez só, mesmo que o usuário recolha e expanda de novo.
        ''' Interlocked porque expandir rápido duas vezes dispararia duas
        ''' buscas concorrentes e duplicaria os filhos.
        ''' </summary>
        Private Sub LoadChildrenOnce()
            If Interlocked.CompareExchange(_loaded, 1, 0) <> 0 Then Return
            Dim ignorado = LoadChildrenAsync()
        End Sub

        Private Async Function LoadChildrenAsync() As Task
            IsLoading = True
            Try
                Dim resultado = Await _broker.GetFolderChildrenAsync(Key, CancellationToken.None)

                If Not resultado.Succeeded Then
                    ' Libera para tentar de novo: falhar por Outlook ocupado
                    ' não pode condenar o nó a ficar vazio para sempre.
                    Interlocked.Exchange(_loaded, 0)
                    _onError(resultado.Kind, resultado.Detail)
                    Return
                End If

                Dim filhos = resultado.Value
                Await _ui.InvokeAsync(
                    Sub()
                        Children.Clear()
                        For Each f In filhos
                            If Not FolderTreeViewModel.Exibir(f) Then Continue For
                            Children.Add(New FolderNodeViewModel(f, _broker, _ui, _onError))
                        Next
                    End Sub).Task
            Finally
                IsLoading = False
            End Try
        End Function

        ''' <summary>
        ''' Descarta os filhos e permite buscar de novo. Usado quando a pasta
        ''' é invalidada — reler é a resposta certa, nunca remendar (R6).
        ''' </summary>
        Public Sub Invalidate()
            Interlocked.Exchange(_loaded, 0)
            Children.Clear()
            Children.Add(Marcador)
            _isExpanded = False
            OnPropertyChanged(NameOf(IsExpanded))
        End Sub

    End Class

End Namespace

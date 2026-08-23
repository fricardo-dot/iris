Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' Junta o estado da conexão com o conteúdo da janela.
    '''
    ''' A árvore só existe quando há sessão. Ao cair a conexão, ela é
    ''' esvaziada — mostrar pastas de uma sessão morta seria mentir sobre
    ''' dado que não pode mais ser lido.
    ''' </summary>
    Public NotInheritable Class MainViewModel
        Inherits ObservableObject
        Implements IDisposable

        Private ReadOnly _watcher As FolderWatcher
        Private _disposed As Boolean
        Private _wasConnected As Boolean

        Public Sub New(broker As IOutlookBroker, ui As Global.System.Windows.Threading.Dispatcher)
            Connection = New ConnectionViewModel(broker, ui)
            Folders = New FolderTreeViewModel(broker, ui, AddressOf Connection.Observe)
            Messages = New MessageListViewModel(broker, ui, AddressOf Connection.Observe)
            _watcher = New FolderWatcher(broker, ui, AddressOf Connection.Observe,
                                         AddressOf Messages.OnFolderInvalidated)

            AddHandler Folders.PropertyChanged, AddressOf OnFoldersChanged

            AddHandler Connection.PropertyChanged, AddressOf OnConnectionChanged
        End Sub

        Public ReadOnly Property Connection As ConnectionViewModel
        Public ReadOnly Property Folders As FolderTreeViewModel
        Public ReadOnly Property Messages As MessageListViewModel

        ''' <summary>
        ''' Enquanto não há sessão, o card de conexão ocupa a janela. Quando
        ''' há, dá lugar ao conteúdo — e o indicador da barra continua
        ''' contando o estado, sem trocar a tela inteira a cada oscilação.
        ''' </summary>
        Public ReadOnly Property ShowContent As Boolean
            Get
                Return Connection.State = SessionState.Connected
            End Get
        End Property

        Public Async Function InitializeAsync() As Task
            Await Connection.InitializeAsync()
            SyncContentWithSession()
        End Function

        Private Sub OnConnectionChanged(sender As Object, e As ComponentModel.PropertyChangedEventArgs)
            If e.PropertyName <> NameOf(ConnectionViewModel.State) Then Return
            SyncContentWithSession()
        End Sub

        ''' <summary>
        ''' Sincroniza o conteúdo com o estado atual da sessão.
        '''
        ''' Chamado também DEPOIS de InitializeAsync, e não só por
        ''' PropertyChanged: se o broker já estivesse conectado quando o
        ''' ViewModel nasceu, o estado inicial seria Connected, o probe
        ''' devolveria Connected de novo, SetProperty não dispararia evento
        ''' nenhum — e a árvore nunca carregaria.
        ''' </summary>
        Public Sub SyncContentWithSession()
            OnPropertyChanged(NameOf(ShowContent))

            Dim conectado = Connection.State = SessionState.Connected
            If conectado = _wasConnected Then Return
            _wasConnected = conectado

            If conectado Then
                Connection.Observe(Folders.ReloadAsync(), "folders.reload")
            Else
                Folders.Clear()
                Messages.Clear()
                Connection.Observe(_watcher.UnwatchAsync(), "watcher.unwatch")
            End If
        End Sub

        ''' <summary>
        ''' Selecionar uma pasta é o que dispara a lista. A pasta manda; a
        ''' lista obedece — e nunca o contrário.
        ''' </summary>
        Private Sub OnFoldersChanged(sender As Object, e As ComponentModel.PropertyChangedEventArgs)
            If e.PropertyName <> NameOf(FolderTreeViewModel.Selected) Then Return

            Dim pasta = Folders.Selected
            If pasta Is Nothing Then
                Messages.Clear()
                Return
            End If

            Connection.Observe(Messages.ShowFolderAsync(pasta.Key, pasta.Name), "messages.showFolder")
            ' Observar a pasta exibida e o que faz a lista se atualizar
            ' sozinha quando chega mensagem nova.
            Connection.Observe(_watcher.WatchAsync(pasta.Key), "watcher.watch")
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            RemoveHandler Connection.PropertyChanged, AddressOf OnConnectionChanged
            RemoveHandler Folders.PropertyChanged, AddressOf OnFoldersChanged
            _watcher.Dispose()
            Connection.Dispose()
        End Sub

    End Class

End Namespace

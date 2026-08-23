Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
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
        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _ui As Global.System.Windows.Threading.Dispatcher

        ''' <summary>
        ''' Época de sessão que este ViewModel já tratou. Duas coisas podem
        ''' disparar a recarga — o evento de substituição e a transição de
        ''' estado — e a primeira que chegar reivindica a época; a segunda
        ''' vira no-op. Sem isso, abrir o Iris recarregaria a árvore duas
        ''' vezes.
        ''' </summary>
        Private _epocaVista As Long
        Private _disposed As Boolean
        Private _wasConnected As Boolean

        Public Sub New(broker As IOutlookBroker, ui As Global.System.Windows.Threading.Dispatcher,
                       saveFile As ISaveFileService, pickFile As IPickFileService)
            _broker = broker
            _ui = ui
            _epocaVista = broker.SessionEpoch

            Connection = New ConnectionViewModel(broker, ui)
            Folders = New FolderTreeViewModel(broker, ui, AddressOf Connection.Observe)
            Messages = New MessageListViewModel(broker, ui, AddressOf Connection.Observe)
            Detail = New MessageDetailViewModel(broker, ui, AddressOf Connection.Observe, saveFile)
            Composer = New ComposerViewModel(broker, ui, AddressOf Connection.Observe, pickFile)
            _watcher = New FolderWatcher(broker, ui, AddressOf Connection.Observe,
                                         AddressOf Messages.OnFolderInvalidated)

            NewMessageCommand = New AsyncRelayCommand(Function() Composer.NewMessageAsync(),
                                                      Function() PodeCompor)
            ReplyCommand = New AsyncRelayCommand(Function() ResponderAsync(replyAll:=False),
                                                 Function() PodeResponder)
            ReplyAllCommand = New AsyncRelayCommand(Function() ResponderAsync(replyAll:=True),
                                                    Function() PodeResponder)
            ForwardCommand = New AsyncRelayCommand(AddressOf EncaminharAsync,
                                                   Function() PodeEncaminhar)

            AddHandler Composer.PropertyChanged, AddressOf OnComposerChanged

            ' O detalhe chega DEPOIS da seleção, e é ele que diz se a leitura
            ' veio completa. Sem escutar isto, responder ficaria habilitado
            ' com base no estado da mensagem anterior.
            AddHandler Detail.PropertyChanged, AddressOf OnDetailChanged

            AddHandler Messages.PropertyChanged, AddressOf OnMessagesChanged

            AddHandler Folders.PropertyChanged, AddressOf OnFoldersChanged

            AddHandler Connection.PropertyChanged, AddressOf OnConnectionChanged

            AddHandler broker.SessionReplaced, AddressOf OnSessionReplaced
        End Sub

        Public ReadOnly Property Connection As ConnectionViewModel
        Public ReadOnly Property Folders As FolderTreeViewModel
        Public ReadOnly Property Messages As MessageListViewModel
        Public ReadOnly Property Detail As MessageDetailViewModel
        Public ReadOnly Property Composer As ComposerViewModel

        Public ReadOnly Property NewMessageCommand As IAsyncRelayCommand
        Public ReadOnly Property ReplyCommand As IAsyncRelayCommand
        Public ReadOnly Property ReplyAllCommand As IAsyncRelayCommand
        Public ReadOnly Property ForwardCommand As IAsyncRelayCommand

        ''' <summary>
        ''' Um compositor por vez. Dois rascunhos abertos na mesma janela
        ''' precisariam de dois autosaves concorrentes na mesma fila da STA,
        ''' e o segundo só serviria para o usuário perder de vista qual dos
        ''' dois ele está prestes a enviar.
        ''' </summary>
        Public ReadOnly Property PodeCompor As Boolean
            Get
                Return ShowContent AndAlso Not Composer.IsOpen
            End Get
        End Property

        ''' <summary>
        ''' Responder depende dos destinatários LIDOS. Se a leitura veio
        ''' incompleta, responder a todos responderia a menos gente do que a
        ''' mensagem tem, sem nada indicar isso.
        ''' </summary>
        Public ReadOnly Property PodeResponder As Boolean
            Get
                Return PodeCompor AndAlso Messages.Selected IsNot Nothing AndAlso Detail.CanReply
            End Get
        End Property

        ''' <summary>
        ''' Encaminhar leva os anexos. Lista de anexos incompleta significa
        ''' mandar para fora sem conseguir conferir o que vai junto.
        ''' </summary>
        Public ReadOnly Property PodeEncaminhar As Boolean
            Get
                ' SÓ os anexos. Exigir CanReply aqui bloquearia encaminhar por
                ' causa de uma lista de destinatários incompleta — e
                ' encaminhar não usa os destinatários da mensagem original.
                ' Era regra mais restritiva que a documentada, o que é a
                ' forma silenciosa de a regra escrita deixar de valer.
                Return PodeCompor AndAlso Messages.Selected IsNot Nothing AndAlso
                       Detail.CanForward
            End Get
        End Property

        Private Function ResponderAsync(replyAll As Boolean) As Task
            Dim linha = Messages.Selected
            If linha Is Nothing Then Return Task.CompletedTask
            Return Composer.ReplyAsync(linha.Key, replyAll)
        End Function

        Private Function EncaminharAsync() As Task
            Dim linha = Messages.Selected
            If linha Is Nothing Then Return Task.CompletedTask
            Return Composer.ForwardAsync(linha.Key)
        End Function

        Private Sub OnDetailChanged(sender As Object, e As ComponentModel.PropertyChangedEventArgs)
            If e.PropertyName <> NameOf(MessageDetailViewModel.CanReply) AndAlso
               e.PropertyName <> NameOf(MessageDetailViewModel.CanForward) Then Return
            AtualizarComandosDeComposicao()
        End Sub

        Private Sub OnComposerChanged(sender As Object, e As ComponentModel.PropertyChangedEventArgs)
            If e.PropertyName <> NameOf(ComposerViewModel.IsOpen) Then Return
            AtualizarComandosDeComposicao()
        End Sub

        Private Sub AtualizarComandosDeComposicao()
            OnPropertyChanged(NameOf(PodeCompor))
            OnPropertyChanged(NameOf(PodeResponder))
            OnPropertyChanged(NameOf(PodeEncaminhar))
            NewMessageCommand.NotifyCanExecuteChanged()
            ReplyCommand.NotifyCanExecuteChanged()
            ReplyAllCommand.NotifyCanExecuteChanged()
            ForwardCommand.NotifyCanExecuteChanged()
        End Sub

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

        ''' <summary>
        ''' A janela pode fechar agora?
        '''
        ''' Quem decide e o compositor, nao a janela: a janela so sabe de
        ''' chrome e Win32, e o que esta em jogo aqui e texto do usuario.
        ''' </summary>
        Public Function CanCloseWindow() As Boolean
            Return Composer.RequestCloseFromWindow()
        End Function

        ''' <summary>
        ''' Chega em THREAD DO BROKER, como todo evento dele. Nada de tocar
        ''' em ViewModel aqui: devolve ao dispatcher da UI primeiro.
        ''' </summary>
        Private Sub OnSessionReplaced(sender As Object, novaEpoca As Long)
            _ui.BeginInvoke(New Action(Sub() AplicarSessaoNova(novaEpoca)))
        End Sub

        ''' <summary>
        ''' A sessão do Outlook foi substituída — ele morreu e voltou.
        '''
        ''' Tudo o que a sessão anterior entregou deixou de valer: chaves de
        ''' pasta, de item e de assinatura. Continuar mostrando a árvore
        ''' antiga seria mostrar dado que não pode mais ser lido, e foi
        ''' exatamente isso que acontecia antes — em silêncio, e para sempre,
        ''' porque nenhum evento era emitido quando o estado não mudava.
        ''' </summary>
        Private Sub AplicarSessaoNova(novaEpoca As Long)
            If novaEpoca = _epocaVista Then Return
            _epocaVista = novaEpoca

            ' O compositor é avisado, não limpo: o texto é do usuário.
            Composer.OnSessionReplaced(novaEpoca)

            _watcher.OnSessionReplaced()

            ' Guarda o CAMINHO antes do Clear, que zera a seleção. A
            ' chave sozinha só reencontraria pasta de topo.
            Dim anterior = FolderTreeViewModel.CaminhoDe(Folders.Selected)

            Folders.Clear()
            Messages.Clear()
            Detail.Clear()

            If Connection.State <> SessionState.Connected Then Return

            Connection.Observe(RecarregarERestaurarAsync(anterior), "folders.reload")
        End Sub

        ''' <summary>
        ''' Recarrega a árvore e devolve o usuário à pasta em que ele estava.
        '''
        ''' Sem isto, reconectar deixava a árvore certa e a seleção vazia: a
        ''' pasta aberta parava de ser observada e só voltava a atualizar
        ''' quando o usuário clicasse nela de novo, sem nada indicar que
        ''' precisava. Não é a falha silenciosa permanente de antes, mas
        ''' continua sendo atualização que some sem avisar.
        '''
        ''' Reselecionar dispara o fluxo normal — mostrar a pasta e assiná-la
        ''' — em vez de duplicar essa lógica aqui.
        ''' </summary>
        Private Async Function RecarregarERestaurarAsync(anterior As List(Of FolderKey)) As Task
            Await Folders.ReloadAsync()

            If anterior Is Nothing OrElse anterior.Count = 0 Then Return
            If Await Folders.TrySelectAsync(anterior) Then Return

            ' Não achou: era subpasta, ou a pasta não existe mais nesta
            ' sessão. Ficar sem seleção é o comportamento honesto — melhor
            ' que escolher outra pasta por conta própria.
        End Function

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
            AtualizarComandosDeComposicao()

            Dim conectado = Connection.State = SessionState.Connected
            If conectado = _wasConnected Then Return
            _wasConnected = conectado

            If conectado Then
                _epocaVista = _broker.SessionEpoch
                Connection.Observe(Folders.ReloadAsync(), "folders.reload")
            Else
                Folders.Clear()
                Messages.Clear()
                Detail.Clear()

                ' O compositor NAO e limpado aqui. Arvore, lista e leitor
                ' mostram dado do Outlook e sem sessao viram mentira; o
                ' texto do compositor e trabalho do usuario. Fecha-lo na
                ' queda apagaria o que ele escreveu por um motivo que nao e
                ' dele. O rascunho ja esta gravado, e as operacoes falham
                ' com NotConnected ate a sessao voltar.

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
                Detail.Clear()
                Return
            End If

            ' Trocar de pasta esvazia o leitor: manter a mensagem anterior
            ' aberta enquanto a lista mostra outra pasta seria mentir sobre
            ' onde o usuario esta.
            Detail.Clear()

            Connection.Observe(Messages.ShowFolderAsync(pasta.Key, pasta.Name), "messages.showFolder")
            ' Observar a pasta exibida e o que faz a lista se atualizar
            ' sozinha quando chega mensagem nova.
            Connection.Observe(_watcher.WatchAsync(pasta.Key), "watcher.watch")
        End Sub

        ''' <summary>
        ''' Selecionar uma mensagem alimenta o leitor. O leitor decide
        ''' sozinho quando pedir o conteudo — ele tem debounce proprio.
        ''' </summary>
        Private Sub OnMessagesChanged(sender As Object, e As ComponentModel.PropertyChangedEventArgs)
            If e.PropertyName <> NameOf(MessageListViewModel.Selected) Then Return

            ' Uma recarga chama Messages.Clear(), e limpar a
            ' ObservableCollection zera a seleção do ListBox por um instante
            ' antes de ela ser restaurada pela chave. Isso NÃO é o usuário
            ' desmarcando nada — e tratar como se fosse limpava o leitor e
            ' forçava uma segunda leitura do corpo.
            '
            ' A condição é o estado EXPLÍCITO de restauração, não IsLoading:
            ' IsLoading também vale durante "Carregar mais", e ali uma
            ' desmarcação de verdade seria ignorada.
            If Messages.Selected Is Nothing AndAlso Messages.IsRestoringSelection Then Return

            Detail.Show(Messages.Selected)
            AtualizarComandosDeComposicao()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            RemoveHandler Connection.PropertyChanged, AddressOf OnConnectionChanged
            RemoveHandler Folders.PropertyChanged, AddressOf OnFoldersChanged
            RemoveHandler Messages.PropertyChanged, AddressOf OnMessagesChanged
            RemoveHandler Composer.PropertyChanged, AddressOf OnComposerChanged
            RemoveHandler Detail.PropertyChanged, AddressOf OnDetailChanged
            RemoveHandler _broker.SessionReplaced, AddressOf OnSessionReplaced
            _watcher.Dispose()
            Composer.Dispose()
            Detail.Dispose()
            Connection.Dispose()
        End Sub

    End Class

End Namespace

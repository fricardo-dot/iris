Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Core
Imports Iris.Assist
Imports Iris.Integration
Imports Iris.Integration.Assist.Http
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

        ''' <summary>
        ''' A pasta cujo acervo a tela mostra.
        '''
        ''' Fixa em 1 por enquanto, e isso e limitacao declarada: mapear pasta
        ''' do Outlook para chave do cache e trabalho da fase seguinte, junto
        ''' com a varredura disparada pelo app. Hoje o cache so tem o que uma
        ''' importacao manual colocou nele.
        ''' </summary>
        Private Const folderKeyDoAcervo As Long = 1

        Public ReadOnly Property Acervo As AcervoViewModel
        Public ReadOnly Property AcervoIndisponivel As String

        ''' <summary>
        ''' A IA — que hoje serve para dizer que <b>não está habilitada</b>.
        '''
        ''' A composição usa <c>ActivationRecord.DaProducao</c>, que é
        ''' <c>Nothing</c>, e <see cref="AssistenteIndisponivel"/> como provedor.
        ''' As duas coisas são a §28.2 em forma de código: o mecanismo está
        ''' inteiro e o que falta é decisão do usuário.
        ''' </summary>
        Public ReadOnly Property Assistente As AssistenteViewModel

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

            ' O ACERVO — o que a varredura ja guardou no cache.
            '
            ' Nao substitui a lista: ela continua lendo AO VIVO do Outlook, e e
            ' isso que o usuario opera. O acervo e outra coisa, e a §23 explica
            ' por que — em modo cached ele e um arquivo historico conservador,
            ' nao o estado corrente da caixa.
            '
            ' Se o cache nao abrir, o motivo fica VISIVEL. Cache que falha em
            ' silencio vira tela vazia, e tela vazia e indistinguivel de "nao
            ' ha nada guardado".
            Dim motivo As String = Nothing
            Acervo = AcervoViewModel.Abrir(ui, folderKeyDoAcervo, motivo)
            AcervoIndisponivel = motivo

            Assistente = MontarAssistente(ui)

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

            ' A IA perde o contexto: sessão nova é outra ligação com o Outlook,
            ' e um resumo pedido na anterior descreve mensagens que talvez nem
            ' sejam mais as mesmas.
            Assistente?.Trocou()

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

            ' Trocar de PASTA tambem: o contexto da IA e a mensagem aberta, e
            ' ela muda junto.
            Assistente?.Trocou()

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

            ' TROCOU DE MENSAGEM. Sem isto, um resumo pedido para a anterior
            ' voltaria e apareceria embaixo desta — um resumo errado com cara de
            ' certo, que e pior que resumo nenhum porque ninguem desconfia.
            '
            ' A protecao existia no ViewModel e nao estava LIGADA a nada: os
            ' testes chamavam Trocou() direto, e na aplicacao ninguem chamava.
            Assistente?.Trocou()
            AtualizarComandosDeComposicao()
        End Sub

        ''' <summary>
        ''' <b>A composição da IA — e ela é fechada.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>A RECONCILIAÇÃO RODA AQUI, E NÃO NA TELA</b>
        '''
        ''' O que ficou <i>em voo</i> numa execução que morreu vira ambíguo na
        ''' abertura seguinte. Isso é recuperação de segurança, não um número
        ''' para mostrar: roda na <b>composição</b>, antes de a IA ficar apta a
        ''' transmitir, e se falhar o egress fica fechado.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O QUE A PRODUÇÃO MONTA</b>
        '''
        ''' <c>ActivationRecord.DaProducao</c> é <c>Nothing</c>, e o provedor é
        ''' <see cref="AssistenteIndisponivel"/>. Não é lacuna: é a §28.2 — a
        ''' política corporativa aplicável não é inferível desta máquina, e a
        ''' escolha do provedor e da credencial é do usuário.
        '''
        ''' Sem cache aberto não há diário, e sem diário a IA fica desligada:
        ''' transmitir sem poder registrar seria pior que não transmitir.
        ''' </summary>
        ''' <summary>
        ''' O que a IA olharia: a pasta aberta e a mensagem selecionada.
        '''
        ''' Uma mensagem, e não a thread inteira. Reconstruir conversa exige
        ''' correlacionar por <c>ConversationID</c> e por assunto, e o ESCOPO diz
        ''' que reconstrução perfeita de conversa está fora do escopo — juntar
        ''' mensagens que <i>parecem</i> da mesma conversa para mandá-las juntas
        ''' seria decidir divulgação por semelhança.
        ''' </summary>
        Private Function SelecaoParaIa() As (Pasta As FolderKey,
                                             Itens As IReadOnlyList(Of ItemKey))
            Dim vazio = CType(Array.Empty(Of ItemKey)(), IReadOnlyList(Of ItemKey))
            Dim pasta = Folders.Selected?.Key
            Dim selecionada = Messages.Selected
            If selecionada Is Nothing Then Return (pasta, vazio)
            Return (pasta, {selecionada.Key})
        End Function

        ''' <summary>
        ''' <b>O adaptador do provedor que a ativação declara — e só ele.</b>
        '''
        ''' Era um <c>If</c> que instanciava o OpenRouter para qualquer ativação
        ''' carregada. Uma ativação que declarasse outro provedor, com outro
        ''' endereço, receberia o protocolo do OpenRouter e a credencial
        ''' guardada sob o nome dele — conteúdo da caixa indo para um endereço
        ''' arbitrário escrito num arquivo.
        '''
        ''' A lista é <b>fechada</b>: provedor não reconhecido não vira
        ''' adaptador genérico, vira <see cref="AssistenteIndisponivel"/>. Não há
        ''' caso <c>Else</c> que "tenta assim mesmo".
        ''' </summary>
        Friend Shared Function ProvedorPara(ativacao As ActivationLoadResult) As IAssistantProvider
            If Not ativacao.Carregou Then Return New AssistenteIndisponivel()

            If OpenRouterAssistantProvider.Atende(ativacao.Record) Then
                Return New OpenRouterAssistantProvider(ativacao.Record,
                                                       CredencialDoWindows.Leitor())
            End If

            Return New AssistenteIndisponivel()
        End Function

        ''' <summary>
        ''' O que a cerimônia tem a dizer, em português — e o que fica dito
        ''' mesmo quando ela dá certo.
        '''
        ''' Carregou <b>e</b> a política foi verificada: nada. Nos outros casos,
        ''' uma frase que diz onde olhar. O nome do campo com problema entra; o
        ''' <b>valor</b> nunca — o arquivo tem endpoint e identificador de
        ''' pasta, e esta frase vai para a tela.
        ''' </summary>
        Private Shared Function AtivacaoEmPortugues(r As ActivationLoadResult) As String
            If r.Carregou Then
                If r.Record.PoliticaCorporativaVerificada Then Return ""
                Return "A política corporativa aplicável NÃO foi verificada. " &
                       "A IA está ligada por decisão sua, e a autorização vence em " &
                       r.Record.Ate.Value.ToLocalTime().ToString("dd/MM/yyyy") & "."
            End If

            Dim onde = If(r.Campo.Length > 0, $" (campo ""{r.Campo}"")", "")

            Select Case r.Falha
                Case ActivationLoadFailure.Ausente
                    ' O caso NORMAL: nao ha arquivo, e a IA nasce desligada.
                    ' Nao e erro, e dizer "erro" mandaria o usuario procurar
                    ' defeito onde ha so ausencia de decisao.
                    Return ""
                Case ActivationLoadFailure.CampoDesconhecido
                    Return "A ativação não foi aceita: há um campo que o Iris não " &
                           "conhece" & onde & ". Pode ser erro de digitação num campo que importa."
                Case ActivationLoadFailure.CampoDuplicado
                    Return "A ativação não foi aceita: um campo aparece duas vezes" & onde &
                           ", e não dá para saber qual delas vale."
                Case ActivationLoadFailure.CampoFaltando
                    Return "A ativação não foi aceita: falta um campo obrigatório" & onde & "."
                Case ActivationLoadFailure.TipoErrado
                    Return "A ativação não foi aceita: um campo está com o tipo errado" &
                           onde & "."
                Case ActivationLoadFailure.ValorInvalido
                    Return "A ativação não foi aceita: há um valor que não dá para " &
                           "interpretar" & onde & ". Datas precisam de fuso, e nomes de " &
                           "operação e de rótulo vão por extenso, nunca por número."
                Case ActivationLoadFailure.JsonInvalido
                    Return "A ativação não foi aceita: o arquivo não é JSON válido. " &
                           "Comentário e vírgula sobrando não são aceitos."
                Case ActivationLoadFailure.Incompleta
                    Return "A ativação não foi aceita: falta preencher campo obrigatório, " &
                           "ou não há pasta nenhuma autorizada."
                Case ActivationLoadFailure.Incoerente
                    Return "A ativação não foi aceita: ela declara algo que não pode ser " &
                           "declarado — prazo invertido, ou desfecho de leitura que não é " &
                           "prova de nada."
                Case ActivationLoadFailure.EndpointInseguro
                    Return "A ativação não foi aceita: o endereço não é HTTPS."
                Case ActivationLoadFailure.PrazoLongoDemais
                    Return "A ativação não foi aceita: o prazo passa de 90 dias."
                Case ActivationLoadFailure.NaoEhArquivoComum
                    Return "A ativação não foi aceita: o caminho não é um arquivo comum."
                Case ActivationLoadFailure.PermissaoRuim
                    Return "A ativação não foi aceita: mais alguém pode escrever no " &
                           "arquivo, e quem pode escrever nele escolhe para onde o seu " &
                           "e-mail vai. Rode tools\conferir-permissao.ps1 para ver quem é " &
                           "e como tirar."
                Case ActivationLoadFailure.GrandeDemais
                    Return "A ativação não foi aceita: o arquivo é grande demais."
                Case Else
                    Return "A ativação não foi aceita, e não foi possível ler o arquivo."
            End Select
        End Function

        Private Function MontarAssistente(ui As Global.System.Windows.Threading.Dispatcher) _
                                          As AssistenteViewModel
            Dim relogio As Func(Of DateTimeOffset) = Function() DateTimeOffset.Now

            ' A CERIMONIA, lida do arquivo. Sem arquivo — o caso normal — a
            ' ativacao e Nothing e a politica nega tudo, exatamente como antes.
            ' Com arquivo invalido, tambem nega, e o motivo aparece na faixa.
            ' A conferencia de permissao vem DAQUI, e nao de dentro do
            ' carregador: ela e do Windows, e o Iris.Integration nao tem
            ' dependencia de plataforma. Aqui e onde ela pode existir.
            Dim ativacao = ActivationLoader.Carregar(
                relogio(), AddressOf PermissaoDoArquivo.SoMinha)
            Dim politica As New DisclosurePolicy(ativacao.Record)

            Dim diario As IDisclosureJournal = Acervo?.Diario
            Dim reconciliacao = If(diario Is Nothing,
                                   ReconciliationResult.NaoRodou(),
                                   ReconciliationResult.Rodar(diario, relogio()))

            ' O PROVEDOR — UM SO, e o mesmo para o transmissor e o contexto.
            '
            ' Havia dois `New AssistenteIndisponivel()` separados, e o destino
            ' que o contexto declarava vinha de um objeto diferente do que
            ' transmitia. Enquanto os dois eram indisponiveis dava na mesma;
            ' com provedor de verdade, seriam duas verdades sobre para onde o
            ' conteudo vai.
            '
            ' Sem ativacao valida continua sendo o AssistenteIndisponivel, que
            ' nao tem destino: o portao recusa antes de qualquer leitura.
            Dim provedor = ProvedorPara(ativacao)

            Dim transmissor As New AssistTransmitter(
                politica, New CapabilityLedger(),
                If(diario, CType(New DiarioAusente(), IDisclosureJournal)),
                provedor, relogio)

            ' O CONTEXTO DE VERDADE. Ler a mensagem, classificar e montar o
            ' envelope sao requisitos do Iris, e independem de qual API vai
            ' receber os bytes.
            Dim contexto As New ContextoDoOutlook(
                _broker, provedor.Destino, AddressOf SelecaoParaIa)

            Dim vm As New AssistenteViewModel(ui, transmissor, politica, relogio,
                                              reconciliacao, contexto,
                                              New RascunhoDoCompositor(Composer),
                                              AtivacaoEmPortugues(ativacao))

            ' O aviso da abertura entra na tela mesmo sem ninguem pedir nada:
            ' "pode ter saido conteudo e ninguem sabe" nao espera interacao.
            ' O ESTADO DO COMPOSITOR MUDA O QUE A IA PODE FAZER.
            '
            ' `PodeRedigir` exige rascunho editavel, e sem isto o botao ficaria
            ' com o estado do momento em que a janela abriu: habilitado com o
            ' compositor fechado, ou habilitado durante a confirmacao de envio,
            ' quando os campos estao travados de proposito.
            AddHandler Composer.PropertyChanged,
                Sub(remetente As Object, arg As ComponentModel.PropertyChangedEventArgs)
                    If arg.PropertyName = NameOf(ComposerViewModel.PodeEditar) OrElse
                       arg.PropertyName = NameOf(ComposerViewModel.IsOpen) OrElse
                       arg.PropertyName = NameOf(ComposerViewModel.State) Then
                        vm.Avaliar()
                    End If
                End Sub

            vm.Avaliar()
            Return vm
        End Function

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
            Acervo?.Dispose()
        End Sub

    End Class

End Namespace

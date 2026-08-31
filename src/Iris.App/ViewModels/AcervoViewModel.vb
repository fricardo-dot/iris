Imports System.IO
Imports System.Windows.Threading
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Cache
Imports Iris.Core
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.Assist
Imports Iris.Integration
Imports Iris.Integration.Outlook
Imports Iris.Model
Imports Iris.Sync

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' O acervo — o que o Iris guardou do que já viu — e a ressalva que ele
    ''' obriga a mostrar junto.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTE VIEWMODEL NÃO É</b>
    '''
    ''' Não é a lista de mensagens. A lista continua lendo <b>ao vivo</b> do
    ''' Outlook, pelo broker, e é isso que o usuário opera.
    '''
    ''' Este aqui mostra o <b>acervo</b>: o que a varredura publicou no cache.
    ''' São coisas diferentes e a §23 explica por quê — em modo cached o cache é
    ''' um arquivo histórico conservador, não o estado corrente da caixa. Pode
    ''' faltar mensagem que existe no servidor, e pode conter mensagem que o
    ''' usuário já apagou.
    '''
    ''' Trocar a lista para ler daqui é a fase seguinte, e exige que o cache
    ''' ganhe busca, ordenação e reconciliação com o que está ao vivo (§26.3).
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O DRENO</b>
    '''
    ''' A §26.2 exige o consumidor ligado ao <see cref="PublicationDrain"/> na
    ''' inicialização <b>e</b> durante a execução, e proíbe a UI de contornar o
    ''' dreno lendo o manifesto direto como substituto da dívida registrada.
    '''
    ''' Por isso este ViewModel <b>nunca</b> chama <c>ManifestReader</c>: ele
    ''' observa o <see cref="AcervoService"/>, e quem atualiza o serviço é o
    ''' dreno entregando a geração. A leitura direta acontece uma vez só, na
    ''' construção do serviço, para a tela não abrir vazia.
    ''' </summary>
    Public NotInheritable Class AcervoViewModel
        Inherits ObservableObject
        Implements IDisposable

        Private ReadOnly _ui As Dispatcher
        Private ReadOnly _db As CacheDatabase
        Private ReadOnly _servico As AcervoService
        Private ReadOnly _dreno As PublicationDrain

        ''' <summary>
        ''' O acervo de todas as pastas, para a busca. Alimentado pelo mesmo
        ''' dreno que alimenta o painel — ver o construtor.
        ''' </summary>
        Private ReadOnly _todasAsPastas As AcervoDeTodasAsPastas

        ''' <summary>Os dois consumidores, como um só, para o dreno.</summary>
        Private ReadOnly _consumidores As IPublicationConsumer
        Private ReadOnly _broker As IOutlookBroker
        ''' <summary>O arquivo do cache — a varredura abre o dela por aqui.</summary>
        Private ReadOnly _caminho As String

        ''' <summary>
        ''' <b>Quem executa a varredura.</b> <c>Nothing</c> usa o caminho real.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>ISTO EXISTE PARA UMA GUARDA PODER SER TESTADA</b>
        '''
        ''' A guarda que importa aqui é o <c>If _disposed Then Return</c> logo
        ''' depois do <c>Await</c>: fechar a janela no meio de uma varredura
        ''' deixava a continuação escrever num <c>_db</c> já descartado. Ela
        ''' nasceu de um defeito, e o defeito nasceu de <i>outra</i> correção.
        '''
        ''' Sem este ponto de injeção, prová-la exigia segurar uma varredura de
        ''' verdade no meio — e nesta base já se escreveu um teste de
        ''' concorrência com <c>Barrier</c> que <b>passava com o defeito
        ''' presente</b>. Ele foi apagado, e a lacuna ficou declarada.
        '''
        ''' Com o executor injetável o teste vira determinístico: ele sinaliza
        ''' que entrou, segura, deixa o <c>Dispose</c> acontecer, libera, e
        ''' confere que nada foi tocado depois.
        '''
        ''' <c>Friend</c>, e não público: é ponto de teste, não configuração.
        ''' </summary>
        Friend Property ExecutorDaVarredura _
            As Func(Of FolderKey, String, StoreInfo, CancellationToken, ResultadoDaVarredura)
        ''' <summary>
        ''' O voo corrente da varredura, para o descarte poder cancelá-lo.
        ''' </summary>
        Private _cancelamentoDaVarredura As CancellationTokenSource
        Private _alvoDoOutlook As FolderKey
        Private _nomeDoAlvo As String = ""
        Private _storeDoAlvo As StoreInfo
        Private ReadOnly _relogio As DispatcherTimer
        Private _disposed As Boolean

        ''' <summary>
        ''' O diário do egress, sobre o mesmo banco.
        '''
        ''' Mora aqui porque é aqui que o cache está aberto — e não porque o
        ''' acervo tenha algo a ver com a IA. Sem cache não há diário, e sem
        ''' diário a IA fica desligada: transmitir sem poder registrar seria
        ''' pior que não transmitir.
        ''' </summary>
        Public ReadOnly Property Diario As IDisclosureJournal

        ''' <summary>
        ''' Onde o cache mora. Em <c>%LOCALAPPDATA%</c> e não ao lado do
        ''' executável: o executável pode estar em Program Files, onde escrever
        ''' exige elevação, e um cache que só funciona com privilégio não é um
        ''' cache.
        ''' </summary>
        Public Shared Function CaminhoPadrao() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "cache.db")
        End Function

        ''' <summary>
        ''' Abre o acervo, ou devolve <c>Nothing</c> com o motivo em
        ''' <paramref name="motivoDaFalha"/>.
        '''
        ''' FALHA FECHADO e VISÍVEL: cache que não abre não pode virar tela
        ''' vazia silenciosa, porque tela vazia é indistinguível de "não há
        ''' nada guardado".
        ''' </summary>
        ''' <param name="caminho">
        ''' Onde o cache mora. <c>Nothing</c> usa <see cref="CaminhoPadrao"/>.
        '''
        ''' Existe para o teste poder apontar para um arquivo descartável: sem
        ''' isto, qualquer teste deste ViewModel abriria o cache <b>de verdade
        ''' do usuário</b> — e a revisão apontou, com razão, que os caminhos
        ''' concorrentes daqui não tinham cobertura nenhuma justamente porque
        ''' ninguém conseguia instanciá-lo sem tocar na máquina.
        ''' </param>
        Public Shared Function Abrir(ui As Dispatcher, folderKey As Long,
                                     ByRef motivoDaFalha As String,
                                     Optional broker As IOutlookBroker = Nothing,
                                     Optional caminho As String = Nothing) As AcervoViewModel
            motivoDaFalha = Nothing
            Try
                ' GetFullPath: o parametro e String e nada impedia um caminho
                ' RELATIVO. "cache.db" faria GetDirectoryName devolver vazio, e
                ' CreateDirectory("") lanca -- o cache "nao abriria" por um
                ' motivo que nao tem nada a ver com o cache.
                caminho = IO.Path.GetFullPath(
                    If(String.IsNullOrWhiteSpace(caminho), CaminhoPadrao(), caminho))
                Directory.CreateDirectory(Path.GetDirectoryName(caminho))

                Dim falha As OpenFailure = Nothing
                Dim db = CacheDatabase.Open(caminho, CacheSchema.Intended(), falha)
                If db Is Nothing Then
                    motivoDaFalha = $"o cache não abriu ({falha})"
                    Return Nothing
                End If
                Return New AcervoViewModel(ui, db, folderKey, broker, caminho)
            Catch ex As Exception
                motivoDaFalha = $"o cache não abriu ({ex.GetType().Name}: {ex.Message})"
                Return Nothing
            End Try
        End Function

        Private Sub New(ui As Dispatcher, db As CacheDatabase, folderKey As Long,
                        broker As IOutlookBroker, caminho As String)
            _ui = ui
            _db = db
            _caminho = caminho
            ' Sem broker nao ha varredura, e o botao fica desabilitado. E o
            ' caso dos testes que so olham o lado de leitura.
            _broker = broker
            _servico = New AcervoService(db, folderKey)
            _dreno = New PublicationDrain(db)

            ' O SEGUNDO CONSUMIDOR, e o dreno alimenta OS DOIS.
            '
            ' A busca precisa de todas as pastas; o painel mostra uma. Dois
            ' drenos seriam pior: cada um marcaria a geracao como entregue por
            ' conta propria, e o segundo nunca veria o que o primeiro drenou.
            ' Um dreno, um consumidor composto, e o fan-out la dentro.
            _todasAsPastas = New AcervoDeTodasAsPastas(db)
            _consumidores = New ConsumidorComposto(_servico, _todasAsPastas)
            Diario = New SqliteDisclosureJournal(db)

            AddHandler _servico.Mudou, AddressOf AoMudar

            DrenarCommand = New RelayCommand(AddressOf Drenar)
            VarrerCommand = New AsyncRelayCommand(AddressOf VarrerAsync, Function() PodeVarrer)

            ' Na INICIALIZACAO.
            Drenar()

            ' E DURANTE A EXECUCAO. O intervalo e folgado de proposito: nada no
            ' app publica ainda, entao o dreno so tem trabalho quando uma
            ' varredura rodar. Bater no banco a cada segundo para nao achar
            ' nada seria custo sem informacao.
            _relogio = New DispatcherTimer(DispatcherPriority.Background, ui) With {
                .Interval = TimeSpan.FromSeconds(30)}
            AddHandler _relogio.Tick, Sub() Drenar()
            _relogio.Start()

            Atualizar()
        End Sub

        Public ReadOnly Property DrenarCommand As RelayCommand

        ''' <summary>
        ''' <b>Varre a pasta selecionada.</b> Botão, e não automático.
        '''
        ''' ------------------------------------------------------------------
        ''' Varrer é caro — a §D5 mediu o custo por lote — e <b>escreve no
        ''' cache</b>. Disparar a cada clique numa pasta gastaria COM sem
        ''' ninguém pedir, e numa caixa grande travaria a troca de pasta.
        ''' O momento e o custo ficam na mão de quem opera.
        ''' </summary>
        Public ReadOnly Property VarrerCommand As AsyncRelayCommand

        Private _varrendo As Boolean
        Public Property Varrendo As Boolean
            Get
                Return _varrendo
            End Get
            Private Set(valor As Boolean)
                SetProperty(_varrendo, valor)
                OnPropertyChanged(NameOf(PodeVarrer))
                VarrerCommand.NotifyCanExecuteChanged()
            End Set
        End Property

        ''' <summary>
        ''' Dá para varrer? Exige pasta escolhida, broker, e não estar varrendo.
        '''
        ''' <b>Não</b> exige ambiente autorizado: quem recusa por isso é a
        ''' <see cref="VarreduraDaPasta"/>, e a recusa dela <b>explica</b>. Um
        ''' botão desabilitado sem motivo na tela é o defeito que a faixa da IA
        ''' já teve — o usuário clica e nada acontece.
        ''' </summary>
        Public ReadOnly Property PodeVarrer As Boolean
            Get
                Return _broker IsNot Nothing AndAlso
                       _alvoDoOutlook IsNot Nothing AndAlso Not Varrendo
            End Get
        End Property

        Private _itens As Integer
        Public Property Itens As Integer
            Get
                Return _itens
            End Get
            Private Set(value As Integer)
                SetProperty(_itens, value)
            End Set
        End Property

        Private _ressalva As String
        ''' <summary>
        ''' O que a UI é obrigada a mostrar junto do acervo. <c>Nothing</c>
        ''' quando não há ressalva — que hoje nunca acontece em modo cached.
        ''' </summary>
        Public Property Ressalva As String
            Get
                Return _ressalva
            End Get
            Private Set(value As String)
                SetProperty(_ressalva, value)
            End Set
        End Property

        Private _temAlgoADizer As Boolean
        Public Property TemAlgoADizer As Boolean
            Get
                Return _temAlgoADizer
            End Get
            Private Set(value As Boolean)
                SetProperty(_temAlgoADizer, value)
            End Set
        End Property

        Private _travado As String
        ''' <summary>
        ''' Preenchido quando a fila de publicação travou na cabeça. A §26.2
        ''' exige que falha persistente do consumidor apareça — não fique
        ''' bloqueando em silêncio.
        ''' </summary>
        Public Property Travado As String
            Get
                Return _travado
            End Get
            Private Set(value As String)
                SetProperty(_travado, value)
            End Set
        End Property

        ' ==============================================================

        Private Sub Drenar()
            If _disposed Then Return
            Try
                _dreno.Drenar(_consumidores)

                ' E SO DEPOIS DE DRENAR O ACERVO DE TODAS AS PASTAS CARREGA.
                '
                ' Ele nasce vazio de proposito -- ler no construtor seria ler
                ' na frente do dreno quando ha publicacao pendente de uma
                ' queda anterior. Aqui a leitura acontece DEPOIS de a fila ter
                ' sido entregue, entao o retrato nunca esta a frente dela.
                '
                ' Chamar mesmo quando nada foi drenado e o ponto: sem entrega
                ' pendente nenhum Receber dispara, e sem isto a busca ficaria
                ' vazia para sempre numa abertura normal.
                If _todasAsPastas.Recarregado = 0 Then _todasAsPastas.Recarregar()
                Travado = Nothing
            Catch ex As Exception
                ' Consumidor que falha trava a cabeca da fila DE PROPOSITO —
                ' marcar como drenada uma geracao nao recebida seria perder em
                ' silencio. O que nao pode e o bloqueio ser invisivel.
                Dim g = _dreno.TravadoEm()
                Travado = If(g.HasValue,
                    $"A atualização do acervo parou na geração {g.Value}: {_dreno.UltimoErro(g.Value)}",
                    $"A atualização do acervo falhou: {ex.Message}")
            End Try
            Atualizar()
        End Sub

        Private Sub AoMudar(sender As Object, e As EventArgs)
            If _ui.CheckAccess() Then Atualizar() Else _ui.BeginInvoke(CType(AddressOf Atualizar, Action))
        End Sub

        ''' <summary>
        ''' <b>Passa a mostrar a pasta que o usuário selecionou.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' Era a constante 1 — a pasta que uma importação manual de teste tinha
        ''' criado. A lista ao lado mostrava uma pasta e o acervo mostrava
        ''' outra, sem nada dizendo isso.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>NAVEGAR NÃO CRIA PASTA</b>
        '''
        ''' Isto usava <c>ResolvedorDoAcervo.Pasta</c>, que <b>cria</b> — então
        ''' cada clique inseria <c>store</c> e <c>folder</c> no cache, antes de
        ''' qualquer cerimônia de ambiente. Não gravava conteúdo, então o gate
        ''' não era teatro; mas o contrato "nenhuma pasta antes da autorização"
        ''' estava quebrado, e o teste que afirmava <c>folder = 0</c> só
        ''' exercitava a varredura direta — nunca este caminho.
        '''
        ''' Pasta nunca vista resolve para <b>nada</b>, e o acervo diz "ainda não
        ''' foi varrida", que é a verdade. Quem cria é a varredura, depois do
        ''' gate.
        ''' </summary>
        Public Sub Apontar(pasta As FolderKey, nome As String, store As StoreInfo)
            _alvoDoOutlook = pasta
            _nomeDoAlvo = If(nome, "")
            _storeDoAlvo = store

            If pasta Is Nothing OrElse String.IsNullOrWhiteSpace(pasta.EntryId) Then
                ' ESVAZIA. Sem isto o acervo continuava mostrando itens e
                ' ressalva da ULTIMA pasta depois de a selecao sumir -- numeros
                ' sem dono, descrevendo uma pasta que ninguem esta olhando.
                _servico.Apontar(0L)
                OnPropertyChanged(NameOf(PodeVarrer))
                VarrerCommand.NotifyCanExecuteChanged()
                Atualizar()
                Return
            End If

            Try
                ' SemPasta = 0: nao existe folder_key 0, entao o manifesto sai
                ' vazio e a faixa diz "nao tem acervo publicado".
                _servico.Apontar(If(New ResolvedorDoAcervo(_db).PastaExistente(
                    pasta.StoreId, pasta.EntryId), 0L))
            Catch ex As Exception
                ' Leitura tambem falha: banco travado, arquivo sumindo. Nao
                ' pode derrubar a troca de pasta -- a lista ao lado continua
                ' funcionando, e o acervo e o painel secundario.
                Travado = "Nao consegui apontar o acervo para esta pasta."
            End Try

            OnPropertyChanged(NameOf(PodeVarrer))
            VarrerCommand.NotifyCanExecuteChanged()
            Atualizar()
        End Sub

        ''' <summary>
        ''' <b>Varre, numa conexão só dela.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>CONEXÃO PRÓPRIA, E ISSO NÃO É ZELO EXCESSIVO</b>
        '''
        ''' A varredura roda em <c>Task.Run</c>, e enquanto ela corre a interface
        ''' segue viva: o usuário troca de pasta (que lê o cache), o
        ''' temporizador do dreno bate a cada 30 s (que lê e escreve). Enquanto
        ''' tudo isso usava <c>_db.Connection</c>, três threads compartilhavam
        ''' <b>um objeto de conexão</b> — e <c>SqliteConnection</c> não oferece
        ''' isso como contrato.
        '''
        ''' Os <c>BEGIN IMMEDIATE</c> do resolvedor protegem concorrência
        ''' <i>entre conexões</i>, no arquivo. Eles não tornam um objeto de
        ''' conexão seguro para uso simultâneo, e eu tinha confundido as duas
        ''' coisas.
        '''
        ''' Abrir a segunda conexão custa uma migração-e-introspecção por
        ''' varredura. Varredura é rara e disparada por clique; o preço é
        ''' pequeno perto de corrupção intermitente que ninguém reproduz.
        ''' </summary>
        Private Async Function VarrerAsync() As Task
            If Not PodeVarrer Then Return

            Dim pasta = _alvoDoOutlook
            Dim nome = _nomeDoAlvo
            Dim store = _storeDoAlvo
            ' A pasta a que ESTE voo pertence. Depois do Await o usuario pode
            ' ter trocado, e o desfecho de A nao pode aparecer na faixa de B.
            Dim doVoo = pasta

            Dim cts As New CancellationTokenSource()
            _cancelamentoDaVarredura = cts

            Varrendo = True
            Travado = Nothing
            Try
                Dim executor = If(ExecutorDaVarredura, AddressOf VarrerDeVerdade)
                Dim r = Await Task.Run(
                    Function() executor(pasta, nome, store, cts.Token))

                ' A JANELA FECHOU NO MEIO.
                '
                ' Defeito que nasceu na correcao anterior: o Dispose cancela e
                ' NAO espera, e o SweepRunner captura o cancelamento e devolve
                ' um SweepResult -- entao o Await nao lanca, e a execucao
                ' seguia por Recarregar() e Atualizar() sobre um _db ja
                ' descartado. O Catch convertia a corrida em "A varredura
                ' falhou", numa tela que ja nao existe.
                If _disposed Then Return

                ' TROCOU DE PASTA: o voo terminou e publicou, e isso vale. O
                ' que nao vale e escrever o desfecho dele na faixa de outra
                ' pasta -- seria uma recusa da pasta A aparecendo sob o nome
                ' da B.
                '
                ' Equals, e nao ReferenceEquals: FolderKey e imutavel e tem
                ' igualdade de VALOR por (EntryId, StoreId). Comparar por
                ' referencia e mais estreito que a identidade de dominio --
                ' recarregar a arvore reconstroi o objeto da MESMA pasta, e o
                ' resultado seria descartado a toa.
                If Not doVoo.Equals(_alvoDoOutlook) Then Return

                If r Is Nothing Then
                    Travado = "O cache nao abriu para a varredura."
                Else
                    Travado = EmPortugues(r)
                End If
                _servico.Recarregar()
                Atualizar()
            Catch ex As OperationCanceledException
                ' A janela fechou no meio. Nao ha tela para avisar.
            Catch ex As Exception
                If Not _disposed Then Travado = "A varredura falhou."
            Finally
                If ReferenceEquals(_cancelamentoDaVarredura, cts) Then
                    _cancelamentoDaVarredura = Nothing
                End If
                cts.Dispose()
                ' Varrendo dispara OnPropertyChanged e NotifyCanExecuteChanged.
                ' Num objeto descartado isso e mexer em tela que ja saiu.
                If Not _disposed Then Varrendo = False
            End Try
        End Function

        ''' <summary>
        ''' A varredura de verdade, na conexão dela. Ver o doc de
        ''' <see cref="ExecutorDaVarredura"/> para por que isto é substituível.
        ''' </summary>
        Private Function VarrerDeVerdade(pasta As FolderKey, nome As String,
                                         store As StoreInfo,
                                         ct As CancellationToken) As ResultadoDaVarredura
            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(_caminho, CacheSchema.Intended(), falha)
                If db Is Nothing Then Return Nothing
                Return New VarreduraDaPasta(_broker, db).Executar(pasta, nome, store, ct)
            End Using
        End Function

        ''' <summary>
        ''' O desfecho da varredura em português.
        '''
        ''' A recusa por ambiente não autorizado <b>diz o que fazer</b>: sem
        ''' isso o usuário clica, nada acontece, e a única pista mora num banco
        ''' SQLite. Foi exatamente o que a faixa da IA já fez com o botão
        ''' desabilitado.
        ''' </summary>
        Private Shared Function EmPortugues(r As ResultadoDaVarredura) As String
            Select Case r.Recusa
                Case RecusaDeVarredura.AmbienteNaoAutorizado
                    Return "Este ambiente ainda não foi autorizado para varredura " &
                           $"({r.Ambiente}). Autorize com:  " &
                           $"dotnet run --project tools\Iris.CrashHarness -- ambiente --autorizar {r.ChaveDoAmbiente}"
                Case RecusaDeVarredura.SemPasta
                    Return "Escolha uma pasta antes de varrer."
                Case RecusaDeVarredura.StoreDesconhecido
                    Return "Não sei a qual conta esta pasta pertence."
                Case RecusaDeVarredura.Falhou
                    Return "A varredura não terminou, e a tentativa foi descartada."
            End Select

            If r.Varredura Is Nothing Then Return "A varredura não produziu desfecho."
            If r.Varredura.Conclusion = SweepConclusion.Publicada Then Return Nothing

            ' Sem o Motivo: ele e texto, e pode carregar assunto de mensagem.
            Return $"A varredura terminou como {r.Varredura.Conclusion} e nada foi publicado."
        End Function

        Private Sub Atualizar()
            Dim m = _servico.Atual
            Itens = m.Items.Count
            Ressalva = m.Ressalva
            TemAlgoADizer = (m.Ressalva IsNot Nothing) OrElse (Travado IsNot Nothing)
        End Sub

        ''' <summary>
        ''' <b>A porta da busca.</b>
        '''
        ''' Existe aqui, e não no <c>BuscaViewModel</c>, por causa da §26.2 e da
        ''' <c>ArchitectureTests</c>: a camada de apresentação não instancia
        ''' leitor de cache. Quem já tem o banco aberto é este ViewModel, e
        ''' quem o fecha no <c>Dispose</c> também.
        '''
        ''' Devolver o resultado direto — em vez de mantê-lo — é deliberado:
        ''' busca não tem estado que sobreviva à pergunta, e guardar o último
        ''' resultado aqui criaria uma segunda cópia do acervo esperando ficar
        ''' velha.
        '''
        ''' <b>E ela lê o acervo DRENADO, não o banco.</b> Até 28/08/2026 à
        ''' tarde a busca abria o <c>ManifestReader</c> por conta própria, que é
        ''' o contorno da §26.2 — ela podia mostrar uma geração que o painel ao
        ''' lado ainda não tinha recebido, dois lugares da mesma janela
        ''' discordando.
        '''
        ''' Agora as duas leem o mesmo retrato <b>quando as duas entregas
        ''' concluem</b>. Elas são sequenciais: se a primeira conclui e a
        ''' segunda falha, o painel fica à frente da busca, e assim permanece
        ''' enquanto a falha se repetir. Nada se perde — a pendência fica no
        ''' banco —, mas não é entrega atômica, e o
        ''' <c>ConsumidorComposto</c> diz isso.
        ''' </summary>
        Public Function Procurar(termo As String) As Iris.Integration.ResultadoDaBusca
            If _disposed Then
                ' Janela fechando com a caixa de busca ainda em foco. Lançar
                ' aqui viraria exceção numa tela que já saiu; devolver nulo
                ' faria o chamador achar que não achou nada.
                Throw New ObjectDisposedException(NameOf(AcervoViewModel))
            End If
            Return New Iris.Integration.BuscaNoAcervo(_todasAsPastas, _dreno).Procurar(termo)
        End Function

        ''' <summary>
        ''' <b>A fila de respostas, pela mesma porta da busca.</b>
        '''
        ''' Quem tem o banco é o acervo, e é a §26.2: o <c>MainViewModel</c> não
        ''' instancia leitor de cache. A fila é mais um leitor, e entra do mesmo
        ''' jeito.
        ''' </summary>
        Public Function MontarAFila(eu As Iris.Model.MinhasIdentidades,
                                    agora As DateTimeOffset,
                                    fuso As TimeZoneInfo,
                                    dispensadas As Collections.Generic.IEnumerable(Of String),
                                    ignorados As Iris.Model.MinhasIdentidades) _
                                    As Iris.Model.ResultadoDaFila
            If _disposed Then Throw New ObjectDisposedException(NameOf(AcervoViewModel))

            Dim leitor As New Iris.Integration.FilaDoAcervo(_todasAsPastas)
            Return leitor.Montar(eu, agora, fuso, leitor.AcharOsEnviados(),
                                 dispensadas, ignorados)
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True

            ' CANCELA A VARREDURA ANTES DE FECHAR O BANCO.
            '
            ' Sem isto, fechar a janela no meio de uma varredura deixava a
            ' tarefa seguindo com o broker e com um CacheDatabase que este
            ' Dispose ia fechar debaixo dela. O token passado era None, entao
            ' nao havia nem como pedir para parar.
            '
            ' Nao ESPERA a tarefa: bloquear o dispatcher no fechamento
            ' travaria a janela. A varredura tem conexao propria, entao o que
            ' ela ainda escrever nao passa por _db.
            Try
                _cancelamentoDaVarredura?.Cancel()
            Catch
            End Try

            _relogio?.Stop()
            RemoveHandler _servico.Mudou, AddressOf AoMudar
            _db?.Dispose()
        End Sub

    End Class

End Namespace

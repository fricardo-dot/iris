Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports Iris.App.ViewModels
Imports Iris.Model
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>GUARDAS DE DESCARTE DA JANELA PRINCIPAL.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE NAO EXISTIA PRIMEIRO TESTE</b>
'''
''' O relatorio da Fase 2 disse que estas guardas se sustentavam por leitura
''' de codigo, e deu a receita — broker falso e fonte bloqueavel. A receita
''' estava certa e incompleta: faltava dizer que <c>MainViewModel</c> <b>nao
''' podia ser construido</b> num teste.
'''
''' O construtor chamava <c>AcervoViewModel.Abrir</c> sem caminho, e sem
''' caminho o acervo abre o cache <b>do usuario</b>. Instanciar a janela
''' principal numa suite mexeria no banco de producao dele. Nao era teimosia
''' de teste; era dependencia dura.
'''
''' O caminho virou parametro opcional. Producao continua sem escolher.
'''
''' ------------------------------------------------------------------
''' <b>O QUE CADA TESTE COBRE</b>
'''
''' Sao quatro caminhos, e eles nao sao igualmente testaveis. Os que sao
''' estao aqui com controle negativo medido; o que nao e esta dito no fim do
''' arquivo, com o motivo, em vez de virar linha verde que nao prova nada.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class JanelaPrincipalTests

    Private Shared ReadOnly Raiz As New FolderKey("raiz", "store-1")
    Private Shared ReadOnly Entrada As New FolderKey("entrada", "store-1")

    Private _pasta As String
    Private _caminho As String

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-janela-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_pasta)
        _caminho = Path.Combine(_pasta, "cache.db")
    End Sub

    <TestCleanup>
    Public Sub Limpar()
        SqliteConnection.ClearAllPools()
        Try
            If Directory.Exists(_pasta) Then Directory.Delete(_pasta, True)
        Catch
        End Try
    End Sub

    Private Shared Sub NoDispatcherAsync(corpo As Func(Of Dispatcher, Task))
        Dim erro As Exception = Nothing
        Dim t As New Thread(
            Sub()
                Dim d = Dispatcher.CurrentDispatcher
                d.BeginInvoke(
                    Async Sub()
                        Try
                            Await corpo(d)
                        Catch ex As Exception
                            erro = ex
                        Finally
                            d.InvokeShutdown()
                        End Try
                    End Sub)
                Dispatcher.Run()
            End Sub)
        t.SetApartmentState(ApartmentState.STA)
        t.Start()
        Assert.IsTrue(t.Join(TimeSpan.FromSeconds(30)), "a thread STA nao terminou")
        If erro IsNot Nothing Then Throw erro
    End Sub

    Private Shared Function Broker() As FakeBroker
        Dim b As New FakeBroker()
        b.EstadoDaSessao = SessionState.Connected
        b.ComStore("Caixa do teste", "store-1")
        b.ComPasta(Raiz, "Caixa de Entrada", "entrada", temFilhas:=True)
        Return b
    End Function

    Private Function Janela(b As FakeBroker, d As Dispatcher) As MainViewModel
        Return New MainViewModel(b, d, New FakeSaveFile(), New FakePickFile(), _caminho)
    End Function

    Private Shared Function PediuStores(b As FakeBroker) As Integer
        Return b.Chamadas.FindAll(Function(x) x = "GetStores").Count
    End Function

    ''' <summary>
    ''' <b>InitializeAsync NAO espera a arvore.</b>
    '''
    ''' Ela dispara <c>Folders.ReloadAsync()</c> por <c>Connection.Observe</c>,
    ''' que e dispara-e-esquece de proposito: a UI nao pode travar esperando o
    ''' Outlook. Quem escreve teste aqui precisa saber disso, ou mede a arvore
    ''' antes de ela existir — foi o que aconteceu na primeira versao deste
    ''' arquivo, e o sintoma foi um indice fora de faixa.
    '''
    ''' O <c>Await</c> dentro do laco devolve o controle ao dispatcher, que e
    ''' quem povoa as colecoes.
    ''' </summary>
    Private Shared Async Function Assentar(condicao As Func(Of Boolean)) As Task
        For i = 1 To 200
            If condicao() Then Return
            Await Task.Delay(5)
        Next
    End Function

    ''' <summary>Deixa o que estiver em voo terminar, sem esperar nada em especial.</summary>
    Private Shared Async Function Assentar() As Task
        For i = 1 To 20
            Await Task.Delay(5)
        Next
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>Controle: sem fechar nada, abrir a janela carrega a arvore.</b>
    '''
    ''' Primeiro de todos, e de proposito. Uma janela que simplesmente nunca
    ''' carrega passaria em todos os testes de "nao carregou depois de
    ''' fechar" — a armadilha exata que o CLAUDE.md descreve.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_abrir_a_janela_carrega_a_arvore()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                Using vm = Janela(b, d)
                    Await vm.InitializeAsync()
                    Await Assentar(Function() vm.Folders.Roots.Count > 0)
                    Assert.IsTrue(PediuStores(b) > 0, "controle: a abertura tinha de ler os stores")
                    Assert.AreEqual(1, vm.Folders.Roots.Count, "controle: a arvore tinha de carregar")
                    Assert.AreEqual("Caixa de Entrada", vm.Folders.Roots(0).Name)
                    ' O par do teste de fechamento: aqui as filhas SAO pedidas.
                    Assert.IsTrue(b.PedidosDeFilhas.Count > 0,
                        "controle: a recarga normal pede as filhas da raiz")
                End Using
            End Function)
    End Sub

    ''' <summary>
    ''' <b>Fechar durante a abertura NAO dispara recarga nenhuma.</b>
    '''
    ''' <c>InitializeAsync</c> espera o probe e so entao chama
    ''' <c>SyncContentWithSession</c>, que dispara a recarga da arvore E a dos
    ''' stores. Sem o <c>If _disposed Then Return</c> no meio, fechar a janela
    ''' durante a abertura poe as duas para rodar num ViewModel ja descartado
    ''' — e a dos stores termina chamando <c>Acervo.Apontar</c>, que fala com
    ''' um banco fechado.
    '''
    ''' <b>Controle negativo confirmado:</b> removendo essa linha, o teste
    ''' falha com 1 leitura de stores.
    ''' </summary>
    <TestMethod>
    Public Sub Fechar_durante_a_abertura_NAO_dispara_recarga()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                b.TravaDoProbe = New TaskCompletionSource(Of Boolean)()
                Dim vm = Janela(b, d)

                Dim voo = vm.InitializeAsync()
                Assert.IsTrue(b.Chamadas.Contains("Probe"),
                    "controle: a abertura tinha de estar parada no probe")

                vm.Dispose()
                b.TravaDoProbe.SetResult(True)
                Await voo

                Assert.AreEqual(0, PediuStores(b),
                    "a janela fechada disparou a recarga de stores")
                Assert.AreEqual(0, vm.Folders.Roots.Count,
                    "a janela fechada carregou a arvore")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>Fechar durante a recarga da arvore NAO repovoa as colecoes.</b>
    '''
    ''' O <c>Dispose</c> chama <c>Folders.Clear()</c> e <c>Messages.Clear()</c>
    ''' justamente para subir a geracao das duas: nenhuma tem ciclo de vida
    ''' proprio, e sem isso uma recarga iniciada antes do fechamento escreve
    ''' nas colecoes de uma janela que ja saiu.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DUAS VERSOES ERRADAS ANTES DESTA, E AS DUAS PASSAVAM</b>
    '''
    ''' A primeira esperava 100 ms fixos: media relogio, nao guarda. A segunda
    ''' instalava a trava <b>antes</b> do <c>InitializeAsync</c> — e a primeira
    ''' leitura de stores interceptada nem era a da arvore, era a do
    ''' <c>ConnectionViewModel</c>. Com ela presa ali, o <c>SyncContentWithSession</c>
    ''' nunca acontecia, <c>Folders.ReloadAsync</c> nunca comecava, e o
    ''' <c>Dispose</c> caia antes de existir recarga. O teste era outra
    ''' variacao de "fechar durante a abertura" com nome errado.
    '''
    ''' Esta versao abre a janela <b>inteira</b> primeiro, so entao instala a
    ''' trava, e dispara a recarga da arvore <b>explicitamente</b>, guardando a
    ''' tarefa para poder esperar por ela.
    '''
    ''' <b>Controle negativo confirmado:</b> removendo o <c>Folders.Clear()</c>
    ''' do <c>Dispose</c>, este teste falha.
    ''' </summary>
    <TestMethod>
    Public Sub Fechar_durante_a_recarga_da_arvore_NAO_repovoa()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                Dim vm = Janela(b, d)

                ' A janela abre INTEIRA primeiro.
                Await vm.InitializeAsync()
                Await Assentar(Function() vm.Folders.Roots.Count > 0)
                Assert.AreEqual(1, vm.Folders.Roots.Count, "controle: a arvore carregou")

                ' So agora a fonte trava, e a recarga e disparada por nome.
                b.TravaDosStores = New TaskCompletionSource(Of Boolean)()
                Dim antes = b.PedidosDeFilhas.Count
                Dim recarga = vm.Folders.ReloadAsync()
                Await Assentar(Function() b.Chamadas.FindAll(Function(x) x = "GetStores").Count > 1)
                Assert.IsFalse(recarga.IsCompleted,
                    "controle: a recarga tinha de estar presa no broker")

                vm.Dispose()
                b.TravaDosStores.SetResult(True)
                Await recarga

                Assert.AreEqual(antes, b.PedidosDeFilhas.Count,
                    "a recarga da janela fechada seguiu para as filhas")
                Assert.AreEqual(0, vm.Folders.Roots.Count,
                    "a recarga da janela fechada repovoou a arvore")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>Restauracao de sessao VENCIDA nao escolhe pasta na arvore nova.</b>
    '''
    ''' Este e o caminho mais sutil dos quatro, e o comentario da producao
    ''' conta a historia: as geracoes internas da arvore impedem uma carga
    ''' velha de repovoar; elas <b>nao</b> impedem o CHAMADOR velho de
    ''' escolher na arvore nova.
    '''
    ''' A sessao E2 comeca a reconstruir com o caminho guardado em E1. A E3
    ''' chega antes de ela terminar. Sem a conferencia de epoca, a restauracao
    ''' da E2 ainda chamaria <c>TrySelectAsync</c> — instalando selecao por
    ''' iniciativa de uma sessao que ja acabou.
    '''
    ''' <b>Controle negativo confirmado:</b> removendo as duas conferencias de
    ''' <c>minhaEpoca</c> do <c>RecarregarERestaurarAsync</c>, o teste falha
    ''' com a pasta selecionada.
    ''' </summary>
    <TestMethod>
    Public Sub Restauracao_de_sessao_vencida_NAO_seleciona_pasta()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                Dim vm = Janela(b, d)
                Await vm.InitializeAsync()
                Await Assentar(Function() vm.Folders.Roots.Count > 0)

                ' O usuario estava numa pasta quando a sessao caiu.
                vm.Folders.Roots(0).IsSelected = True
                Assert.IsNotNull(vm.Folders.Selected, "controle: a pasta tinha de estar selecionada")

                ' E2 comeca a restaurar, e fica presa no broker.
                b.TravaDosStores = New TaskCompletionSource(Of Boolean)()
                b.SubstituirSessao()
                ' O evento e marshalado com BeginInvoke: a troca de sessao so
                ' acontece quando o dispatcher rodar.
                Await Assentar(Function() vm.Folders.Selected Is Nothing)
                Assert.IsNull(vm.Folders.Selected, "controle: a queda limpa a selecao")
                Assert.IsTrue(PediuStores(b) > 1,
                    "controle: a restauracao da E2 tinha de estar parada no broker")

                ' E3 chega antes de a E2 terminar.
                b.SubstituirSessao()
                Await Assentar()

                b.TravaDosStores.SetResult(True)

                ' MARCO, E NAO RELOGIO: a restauracao da E3 recarrega a arvore,
                ' e so depois de ela estar de pe e que faz sentido perguntar
                ' se alguem selecionou pasta nela. Esperar tempo fixo aqui
                ' poderia afirmar "ninguem selecionou" antes de a restauracao
                ' defeituosa sequer chegar ao TrySelectAsync.
                Await Assentar(Function() vm.Folders.Roots.Count > 0)
                Assert.AreEqual(1, vm.Folders.Roots.Count,
                    "controle: a arvore da sessao nova tinha de ter carregado")
                Await Assentar()

                Assert.IsNull(vm.Folders.Selected,
                    "a restauracao da sessao vencida escolheu pasta na arvore nova")
                vm.Dispose()
            End Function)
    End Sub


    ''' <summary>
    ''' <b>O ACERVO E A AGENDA NÃO APARECEM JUNTOS — medido no ViewModel, e não
    ''' no XAML.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' O <c>BindingsDaJanelaTests</c> confere que as duas faixas declaram a
    ''' mesma linha da grade e que o acervo depende de <c>MostrarAcervo</c>. Isso
    ''' prova a ligação; não prova a <b>fórmula</b>. A revisão externa de 28/08
    ''' foi explícita: aquele teste só olhava se a propriedade existia.
    '''
    ''' Aqui a transição é exercitada de verdade — selecionar uma pasta de
    ''' correio, depois uma de calendário, depois nenhuma — porque é a transição
    ''' que produz a sobreposição, e não a declaração estática.
    '''
    ''' E o motivo de isto importar tanto está escrito na história do projeto: a
    ''' faixa do acervo ficou invisível por dias porque duas bordas dividiam a
    ''' <c>Grid.Row="2"</c> e ninguém tinha feito a pergunta ENTRE as faixas.
    ''' </summary>
    <TestMethod>
    Public Sub Acervo_e_agenda_nunca_aparecem_juntos()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                b.ComPasta(Raiz, "Calendário", "cal", temFilhas:=False)
                b.MarcarComoCalendario("cal")

                Using vm = Janela(b, d)
                    Await vm.InitializeAsync()
                    Await Assentar(Function() vm.Folders.Roots.Count > 1)

                    Dim correio = vm.Folders.Roots.First(Function(f) f.Name = "Caixa de Entrada")
                    Dim calendario = vm.Folders.Roots.First(Function(f) f.Name = "Calendário")

                    ''' Nada selecionado: nem um nem outro reivindica a linha.
                    Assert.IsFalse(vm.Agenda.TemPasta, "controle: agenda comeca vazia")

                    ''' Pasta de CORREIO: o acervo aparece, a agenda nao.
                    correio.IsSelected = True
                    Await Assentar(Function() vm.MostrarAcervo)
                    Assert.IsTrue(vm.MostrarAcervo, "em pasta de correio o acervo aparece")
                    Assert.IsFalse(vm.Agenda.TemPasta, "em pasta de correio a agenda NAO aparece")

                    ''' Pasta de CALENDARIO: troca de dono, e nunca os dois.
                    calendario.IsSelected = True
                    Await Assentar(Function() vm.Agenda.TemPasta)
                    Assert.IsTrue(vm.Agenda.TemPasta, "em pasta de calendario a agenda aparece")
                    Assert.IsFalse(vm.MostrarAcervo,
                        "as duas faixas apareceram juntas na mesma linha da grade")
                End Using
            End Function)
    End Sub

    ''' <summary>
    ''' <b>LISTA VAZIA NA TELA NÃO É PASTA VAZIA.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A TELA DIZIA AS DUAS COISAS AO MESMO TEMPO</b>
    '''
    ''' O texto do meio era fixo — <i>"Esta pasta está vazia"</i> — e aparecia
    ''' sempre que a lista convertida ficava sem linha. Com uma página de
    ''' <c>TotalAtStart = 1</c> e <c>SkippedCount = 1</c>, a mesma tela mostrava
    ''' <i>"Esta pasta está vazia"</i> no meio e <i>"0 de 1 · 1 item
    ''' ignorado"</i> no rodapé. Uma das duas estava mentindo, e era a que ocupa
    ''' a tela inteira.
    '''
    ''' É a terceira instância da mesma família em três passadas — o
    ''' <c>0 compromisso(s)</c> da agenda, o <i>"calendário vazio"</i> do
    ''' roteiro, e esta. <b>Tratar "não observei" como "observei e não há".</b>
    ''' Esta é a única que estava na tela principal.
    '''
    ''' <b>Controle negativo:</b> fazendo o <c>EmptyMessage</c> ignorar o
    ''' <c>_skipped</c>, a asserção do meio cai.
    '''
    ''' <b>E o que ele NÃO controla:</b> este teste não lê o XAML, então devolver
    ''' o texto fixo à tela o deixaria verde. Eu tinha escrito aqui que devolver
    ''' o literal derrubaria a asserção, e não derruba — a revisão externa pegou.
    ''' Quem fecha esse metro é
    ''' <c>BindingsDaJanelaTests.O_texto_da_pasta_vazia_vem_do_ViewModel</c>.
    ''' </summary>
    <TestMethod>
    Public Sub Lista_vazia_com_item_ignorado_NAO_diz_que_a_pasta_esta_vazia()
        NoDispatcherAsync(
            Async Function(d) As Task
                Dim b As New FakeBroker()
                b.EstadoDaSessao = SessionState.Connected
                ' A GERACAO IMPORTA: o ExecutarAsync descarta a pagina cuja
                ' Generation nao bate com a do pedido, e o primeiro
                ' ShowFolderAsync incrementa para 1. Sem isso o teste ve
                ' "0 de 0" e a pagina nunca chega -- foi o que aconteceu na
                ' primeira versao deste teste, e o sintoma nao aponta para ca.
                b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
                    New MessagePage With {
                        .Generation = 1,
                        .Items = New List(Of MailSummary)(),
                        .TotalAtStart = 1,
                        .SkippedCount = 1})

                Dim painel As New MessageListViewModel(b, d, Sub(t, nome)
                                                             End Sub)
                Await painel.ShowFolderAsync(New FolderKey("entry-1", "store-1"), "Caixa de Entrada")

                Assert.AreEqual(0, painel.Messages.Count, "controle: a pagina veio sem linha")
                Assert.IsTrue(painel.ShowEmptyFolder, "controle: a faixa do vazio tinha de aparecer")
                StringAssert.Contains(painel.StatusLine, "1 item(ns) ignorado(s)",
                    "controle: o rodape ja dizia que perdeu um item")

                ' O CONSERTO.
                Assert.IsFalse(painel.EmptyMessage.Contains("Esta pasta está vazia"),
                    "a tela afirma que a pasta esta vazia enquanto o rodape diz " &
                    "que a leitura perdeu um item")
                StringAssert.Contains(painel.EmptyMessage, "não conseguiu trazer")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>CONTAGEM QUE FALHOU NÃO É PASTA VAZIA.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O ZERO QUE ERA DOIS ESTADOS</b>
    '''
    ''' <c>MessagePage.TotalAtStart</c> é <c>Integer?</c> porque
    ''' <c>ContarItens</c> devolve <c>Nothing</c> quando <c>Items.Count</c>
    ''' lança. O <c>_total</c> guardava só o número, então "a pasta declara
    ''' zero" e "não consegui contar" viravam o mesmo zero — e a tela dizia
    ''' <i>"Esta pasta está vazia"</i> nos dois.
    '''
    ''' <b>Este teste nasceu de um controle negativo que PASSOU.</b> Eu tinha
    ''' acabado de escrever o <c>_totalConhecido</c>, apaguei o ramo dele para
    ''' conferir, e a suíte inteira continuou verde: a correção não tinha
    ''' nenhum teste. É o bloqueio sem controle negativo que o CLAUDE.md
    ''' descreve, cometido no mesmo dia em que eu o citei.
    ''' </summary>
    <TestMethod>
    Public Sub Contagem_que_FALHOU_nao_vira_pasta_vazia()
        NoDispatcherAsync(
            Async Function(d) As Task
                Dim b As New FakeBroker()
                b.EstadoDaSessao = SessionState.Connected
                b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
                    New MessagePage With {
                        .Generation = 1,
                        .Items = New List(Of MailSummary)(),
                        .TotalAtStart = Nothing,
                        .SkippedCount = 0})

                Dim painel As New MessageListViewModel(b, d, Sub(t, nome)
                                                             End Sub)
                Await painel.ShowFolderAsync(New FolderKey("entry-1", "store-1"), "Caixa de Entrada")

                Assert.IsTrue(painel.ShowEmptyFolder, "controle: a faixa do vazio aparece")
                StringAssert.Contains(painel.StatusLine, "0 de ?",
                    "o rodape afirma um total que ninguem contou")

                Assert.IsFalse(painel.EmptyMessage.Contains("Esta pasta está vazia"),
                    "a contagem falhou e a tela afirmou que a pasta esta vazia")
                StringAssert.Contains(painel.EmptyMessage, "não consegui saber quantos")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>A PASTA NOVA NÃO HERDA O TOTAL DA ANTERIOR.</b>
    '''
    ''' No reload, <c>Messages</c>, <c>_skipped</c> e <c>_fabricadas</c> zeravam
    ''' e o total <b>não</b>. Então uma pasta cuja contagem falha "declarava" o
    ''' total da pasta que o usuário estava olhando antes — número com dono
    ''' errado, que é o mesmo defeito que o <c>Perder_a_selecao_ESVAZIA_o_acervo</c>
    ''' pegou no acervo.
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_nova_NAO_herda_o_total_da_anterior()
        NoDispatcherAsync(
            Async Function(d) As Task
                Dim b As New FakeBroker()
                b.EstadoDaSessao = SessionState.Connected

                ' PASTA A: contagem boa, cinco itens declarados.
                b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
                    New MessagePage With {
                        .Generation = 1,
                        .Items = New List(Of MailSummary)() From {
                            New MailSummary With {
                                .Key = New ItemKey("E-1", "store-1"),
                                .Subject = "uma", .SenderName = "quem"}},
                        .TotalAtStart = 5,
                        .SkippedCount = 0})

                Dim painel As New MessageListViewModel(b, d, Sub(t, nome)
                                                             End Sub)
                Await painel.ShowFolderAsync(New FolderKey("entry-1", "store-1"), "Caixa de Entrada")
                Assert.AreEqual(1, painel.Messages.Count, "controle: a pasta A carregou")
                StringAssert.Contains(painel.StatusLine, "1 de 5", "controle: o total de A")

                ' PASTA B: a contagem falha, e nada dela pode vir de A.
                b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
                    New MessagePage With {
                        .Generation = 2,
                        .Items = New List(Of MailSummary)(),
                        .TotalAtStart = Nothing,
                        .SkippedCount = 0})
                Await painel.ShowFolderAsync(New FolderKey("entry-2", "store-1"), "Outra")

                Assert.IsTrue(painel.ShowEmptyFolder)
                Assert.IsFalse(painel.EmptyMessage.Contains("5"),
                    "a pasta nova declarou o total da anterior")

                ' EXIGIR O ESTADO CERTO, e nao so a ausencia do errado: sem
                ' esta linha o teste tambem passaria se B virasse "zero
                ' CONHECIDO", que e outra afirmacao falsa.
                StringAssert.Contains(painel.StatusLine, "0 de ?",
                    "o total de B nao e zero: e desconhecido")
                StringAssert.Contains(painel.EmptyMessage, "não consegui saber quantos")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>TROCAR DE PASTA DURANTE UMA CARGA NÃO DEIXA OS NÚMEROS DA ANTERIOR.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O NOME MUDAVA E OS NÚMEROS FICAVAM</b>
    '''
    ''' O <c>Despachar</c> tem fila de um: com uma operação em voo, o pedido novo
    ''' vira <c>_pending</c> e volta na hora. O <c>ShowFolderAsync</c> terminava
    ''' com o nome de B na tela e as mensagens, o total e o descarte de A — e o
    ''' <c>_totalConhecido = False</c> só chegava quando o pendente começasse. Se
    ''' a operação de A travasse, durava para sempre.
    '''
    ''' O teste sequencial não alcança isso, e a revisão externa apontou: ele
    ''' prova o reload EXECUTADO, e o defeito está no instante lógico da troca.
    ''' Aqui a página de A fica <b>presa</b> na trava do duplo, que é o que torna
    ''' a fila de um observável.
    '''
    ''' <b>Controle negativo:</b> tirando o <c>LimparConteudo()</c> do
    ''' <c>ShowFolderAsync</c>, as asserções do meio caem.
    ''' </summary>
    <TestMethod>
    Public Sub Trocar_de_pasta_DURANTE_a_carga_nao_deixa_os_numeros_da_anterior()
        NoDispatcherAsync(
            Async Function(d) As Task
                Dim b As New FakeBroker()
                b.EstadoDaSessao = SessionState.Connected
                b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
                    New MessagePage With {
                        .Generation = 1,
                        .Items = New List(Of MailSummary)() From {
                            New MailSummary With {
                                .Key = New ItemKey("E-1", "store-1"),
                                .Subject = "uma", .SenderName = "quem"}},
                        .TotalAtStart = 5,
                        .SkippedCount = 0,
                        .NextCursor = "c1"})

                Dim painel As New MessageListViewModel(b, d, Sub(t, nome)
                                                             End Sub)
                ' A PAGINA DE A DEMORA DE PROPOSITO. Sem isso a duracao dela e
                ' zero, e a assercao de que a duracao NAO vaza para B passaria
                ' com a correcao desfeita -- foi o que aconteceu aqui.
                Dim primeira As New TaskCompletionSource(Of Boolean)()
                b.TravaDaPagina = primeira
                Dim carga = painel.ShowFolderAsync(New FolderKey("entry-1", "store-1"), "Caixa de Entrada")
                Await Task.Delay(40)
                primeira.SetResult(True)
                Await carga
                b.TravaDaPagina = Nothing

                Assert.AreEqual(1, painel.Messages.Count, "controle: a pasta A carregou")
                StringAssert.Contains(painel.StatusLine, "1 de 5", "controle: o total de A")
                Assert.IsTrue(painel.LastPageMs > 0,
                    $"controle: a pagina de A tinha de custar tempo, custou {painel.LastPageMs}")

                ' UM LOAD MORE DE A FICA EM VOO, preso na trava.
                '
                ' TEM DE SER LOAD MORE, e nao reload: o reload limpa a tela
                ' logo no comeco dele, entao o estado ja estaria limpo por
                ' outro motivo e o teste passaria com a correcao desfeita.
                ' Foi o que aconteceu na primeira versao deste teste.
                b.TravaDaPagina = New TaskCompletionSource(Of Boolean)()
                Dim emVoo = painel.LoadMoreAsync()

                ' E O USUARIO TROCA PARA B. O ShowFolderAsync volta na hora,
                ' porque o Despachar so enfileira.
                Await painel.ShowFolderAsync(New FolderKey("entry-2", "store-1"), "Outra")

                Assert.AreEqual("Outra", painel.FolderName, "controle: o nome ja e o de B")

                ' O CONSERTO: os numeros de A nao podem estar na tela de B.
                Assert.AreEqual(0, painel.Messages.Count,
                    "a tela de B continuou mostrando as mensagens de A")
                Assert.IsFalse(painel.StatusLine.Contains("de 5"),
                    "o rodape de B continuou contando pela pasta A")
                StringAssert.Contains(painel.StatusLine, "0 de ?")

                ' A DURACAO TAMBEM E CONTEUDO, e tambem vazava: B mostrava a
                ' "ultima pagina" de A, inclusive sem nunca ter tido pagina.
                Assert.AreEqual(0.0, painel.LastPageMs,
                    "o rodape de B mostrou a duracao da pagina da pasta A")

                ' Solta a trava para nao deixar operacao pendurada.
                b.TravaDaPagina.SetResult(True)
                Await emVoo
            End Function)
    End Sub

    ''' <summary>
    ''' <b>CONTROLE POSITIVO: pasta que declara zero e não perdeu nada É vazia.</b>
    '''
    ''' Sem ele, um <c>EmptyMessage</c> que nunca dissesse "vazia" passaria no
    ''' teste de cima — que é o bloqueio sem controle negativo que o CLAUDE.md
    ''' descreve.
    '''
    ''' <b>E ele exige que a página TENHA SIDO APLICADA.</b> A primeira versão
    ''' não exigia: <c>_total</c>, <c>_skipped</c> e <c>_fabricadas</c> já nascem
    ''' zero, então o teste passaria com a página ignorada — a revisão externa
    ''' apontou. O <c>_totalConhecido</c> resolve isso por construção: sem a
    ''' página, o total é <b>desconhecido</b> e a frase é outra. A asserção do
    ''' <c>StatusLine</c> deixa isso explícito.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_pasta_que_declara_zero_e_dita_vazia()
        NoDispatcherAsync(
            Async Function(d) As Task
                Dim b As New FakeBroker()
                b.EstadoDaSessao = SessionState.Connected
                ' A GERACAO IMPORTA: o ExecutarAsync descarta a pagina cuja
                ' Generation nao bate com a do pedido, e o primeiro
                ' ShowFolderAsync incrementa para 1. Sem isso o teste ve
                ' "0 de 0" e a pagina nunca chega -- foi o que aconteceu na
                ' primeira versao deste teste, e o sintoma nao aponta para ca.
                b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
                    New MessagePage With {
                        .Generation = 1,
                        .Items = New List(Of MailSummary)(),
                        .TotalAtStart = 0,
                        .SkippedCount = 0})

                Dim painel As New MessageListViewModel(b, d, Sub(t, nome)
                                                             End Sub)
                Await painel.ShowFolderAsync(New FolderKey("entry-1", "store-1"), "Caixa de Entrada")

                Assert.IsTrue(painel.ShowEmptyFolder)

                ' A PAGINA CHEGOU: sem ela o total seria desconhecido, e o
                ' rodape diria "0 de ?" em vez de "0 de 0".
                StringAssert.Contains(painel.StatusLine, "0 de 0",
                    "controle: a pagina nao foi aplicada, e o teste passaria " &
                    "pelo estado inicial")
                StringAssert.Contains(painel.EmptyMessage, "Esta pasta está vazia")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>O DUPLO QUEBRA NA HORA quando chamam a página fora da alçada.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>UMA TRAVA NOVA QUASE CUSTOU ESSA PROPRIEDADE</b>
    '''
    ''' Ao acrescentar a <c>TravaDaPagina</c>, eu embrulhei todo o
    ''' <c>GetMessagePageAsync</c> num método <c>Async</c> — e com isso a
    ''' chamada não configurada deixou de lançar e passou a devolver uma
    ''' <c>Task</c> com falha. Teste que esquecesse o <c>Await</c> passaria em
    ''' silêncio, que é exatamente o contrário do que este duplo existe para
    ''' fazer. A revisão externa pegou e pediu regressão; é esta.
    '''
    ''' <b>Controle negativo:</b> voltando o embrulho <c>Async</c>, a exceção
    ''' vira <c>Task</c> com falha e este teste cai.
    ''' </summary>
    <TestMethod>
    Public Sub O_duplo_lanca_NA_HORA_para_pagina_fora_da_alcada()
        Dim b As New FakeBroker()

        ' Sem RespostaDaPagina configurada: e chamada fora da alcada.
        Dim explodiu = False
        Try
            ' De proposito SEM Await: o que se cobra e a excecao SINCRONA.
            Dim ignorado = b.GetMessagePageAsync(
                New MessageQuery(New FolderKey("entry-1", "store-1"),
                                 MessageSort.ReceivedDesc, 1),
                Nothing, 50, CancellationToken.None)
        Catch ex As NotSupportedException
            explodiu = True
        End Try

        Assert.IsTrue(explodiu,
            "a chamada fora da alcada devolveu Task com falha em vez de lancar: " &
            "um teste que esquecesse o Await passaria em silencio")
    End Sub

End Class

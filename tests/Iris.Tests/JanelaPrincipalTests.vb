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
    ''' <b>O CONTROLE NEGATIVO DESTE NAO DISTINGUE, E ISSO FICA ESCRITO</b>
    '''
    ''' Removendo o <c>Folders.Clear()</c> do <c>Dispose</c>, este teste
    ''' <b>continua verde</b>. Medi o estado do broker nos dois casos e ele e
    ''' identico:
    '''
    ''' <code>
    '''   chamadas = Probe | GetStores | GetStores-fim
    '''   filhas   = (vazio)
    '''   raizes   = 0
    ''' </code>
    '''
    ''' Ou seja: a recarga resume depois da trava e para antes de pedir filha
    ''' nenhuma, <b>com ou sem</b> a guarda do <c>Dispose</c>. Quem a segura e
    ''' outra coisa dentro do proprio <c>FolderTreeViewModel</c> — a cadeia de
    ''' <c>Atual(geracao)</c> que <c>ArvoreDePastasTests</c> ja cobre.
    '''
    ''' Entao este teste <b>nao</b> prova a guarda do <c>Dispose</c>. Ele prova
    ''' a propriedade composta — <i>fechar durante a recarga nao repovoa</i> —
    ''' e fica como cerca de regressao, dito com esse nome. Chamar de cobertura
    ''' seria repetir o erro que a Fase 2 documentou: linha verde que passa com
    ''' o defeito presente gasta a confianca do numero que ela mesma produz.
    '''
    ''' A guarda do <c>Dispose</c> isolada continua sem teste, e esta no
    ''' relatorio.
    ''' </summary>
    <TestMethod>
    Public Sub Fechar_durante_a_recarga_da_arvore_NAO_repovoa_CERCA()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                b.TravaDosStores = New TaskCompletionSource(Of Boolean)()
                Dim vm = Janela(b, d)

                Dim voo = vm.InitializeAsync()
                Assert.IsTrue(PediuStores(b) > 0,
                    "controle: a recarga tinha de estar parada no broker")

                vm.Dispose()
                b.TravaDosStores.SetResult(True)
                Await voo

                ' O SINAL E O PEDIDO DAS FILHAS, e ele tem de NAO acontecer.
                '
                ' A primeira versao esperava 100 ms fixos e passava ate com o
                ' Folders.Clear() removido do Dispose: media relogio, nao
                ' guarda. A segunda tentou esperar o pedido das filhas como
                ' marco -- e ele nunca chega, justamente porque a guarda
                ' funciona: o ReloadAsync ve a geracao vencida logo depois dos
                ' stores e volta antes de pedir filha nenhuma.
                '
                ' Entao o pedido das filhas e o SINTOMA da guarda ausente. A
                ' espera sai cedo se ele acontecer, e vai ate o fim se nao
                ' acontecer. Quem sustenta que isto nao e vacuo e o controle
                ' positivo la em cima, que cobra o pedido das filhas na
                ' recarga normal.
                Await Assentar(Function() b.PedidosDeFilhas.Count > 0)

                Assert.AreEqual(0, b.PedidosDeFilhas.Count,
                    $"a recarga da janela fechada seguiu para as filhas. chamadas={String.Join(",", b.Chamadas)} filhas={String.Join(",", b.PedidosDeFilhas)} raizes={vm.Folders.Roots.Count}")
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
                Await Assentar()

                Assert.IsNull(vm.Folders.Selected,
                    "a restauracao da sessao vencida escolheu pasta na arvore nova")
                vm.Dispose()
            End Function)
    End Sub

End Class

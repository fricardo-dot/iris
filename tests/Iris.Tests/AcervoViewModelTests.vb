Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports Iris.App.ViewModels
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Model
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O acervo pelo caminho da interface — que era o que faltava.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ESTE ARQUIVO EXISTE</b>
'''
''' A revisão apontou o buraco com precisão: os testes do
''' <c>ResolvedorDoAcervo</c> provam que <c>PastaExistente</c> não cria, e
''' continuariam <b>verdes</b> se <c>AcervoViewModel.Apontar</c> voltasse a
''' chamar <c>Pasta</c>, que cria. O contrato "navegar não é autorizar" vivia
''' na camada de baixo, e quem o quebrava era a de cima.
'''
''' Instanciar este ViewModel exigia um <c>Dispatcher</c> de verdade e abria
''' o cache <b>do usuário</b> — por isso ninguém o testava. O caminho do cache
''' passou a ser injetável só por causa disto.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class AcervoViewModelTests

    Private _pasta As String
    Private _caminho As String

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-acervovm-" & Guid.NewGuid().ToString("N"))
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

    ''' <summary>
    ''' Roda o corpo numa STA com <c>Dispatcher</c> de verdade: o ViewModel
    ''' cria um <c>DispatcherTimer</c> no construtor, e sem dispatcher ele nem
    ''' nasce.
    ''' </summary>
    Private Shared Sub NoDispatcher(corpo As Action(Of Dispatcher))
        Dim erro As Exception = Nothing
        Dim t As New Thread(
            Sub()
                Dim d = Dispatcher.CurrentDispatcher
                d.BeginInvoke(
                    Sub()
                        Try
                            corpo(d)
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

    ''' <summary>
    ''' Como <see cref="NoDispatcher"/>, mas o corpo pode <c>Await</c>.
    '''
    ''' Necessário para segurar a varredura no meio: esperar o sinal
    ''' bloqueando o dispatcher travaria a própria continuação que se quer
    ''' observar.
    ''' </summary>
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

    Private Function AbrirVm(d As Dispatcher) As AcervoViewModel
        Dim motivo As String = Nothing
        Dim vm = AcervoViewModel.Abrir(d, 0, motivo, New FakeBroker(), _caminho)
        Assert.IsNotNull(vm, motivo)
        Return vm
    End Function

    Private Function Contar(tabela As String) As Integer
        Using conn As New SqliteConnection($"Data Source={_caminho}")
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = $"SELECT COUNT(*) FROM {tabela}"
                Return Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>Clicar numa pasta NÃO cria linha no cache.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' É o teste que faltava. <c>Apontar</c> chamava o resolvedor que cria, e
    ''' cada clique na árvore inseria <c>store</c> e <c>folder</c> <b>antes de
    ''' qualquer cerimônia de ambiente</b>. Não gravava conteúdo — o gate não
    ''' era teatro — mas o contrato "nenhuma pasta antes da autorização"
    ''' estava quebrado, e nenhum teste passava por aqui.
    ''' </summary>
    <TestMethod>
    Public Sub Clicar_numa_pasta_NAO_cria_linha_no_cache()
        NoDispatcher(
            Sub(d)
                Using vm = AbrirVm(d)
                    vm.Apontar(New FolderKey("entry-1", "store-1"), "Caixa de Entrada",
                               New StoreInfo() With {.StoreId = "store-1"})

                    Assert.AreEqual(0, Contar("store"), "navegar criou store")
                    Assert.AreEqual(0, Contar("folder"), "navegar criou pasta")
                End Using
            End Sub)
    End Sub

    ''' <summary>
    ''' <b>E a faixa diz a verdade sobre a pasta nunca varrida.</b>
    '''
    ''' O controle negativo do teste acima: sem ele, um <c>Apontar</c> que
    ''' simplesmente não fizesse nada passaria — e o acervo mostraria os
    ''' números da pasta anterior.
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_nunca_varrida_diz_que_nao_foi_varrida()
        NoDispatcher(
            Sub(d)
                Using vm = AbrirVm(d)
                    vm.Apontar(New FolderKey("entry-1", "store-1"), "Caixa",
                               New StoreInfo() With {.StoreId = "store-1"})

                    Assert.AreEqual(0, vm.Itens)
                    StringAssert.Contains(vm.Ressalva, "ainda não foi varrida")
                End Using
            End Sub)
    End Sub

    ''' <summary>
    ''' <b>Perder a seleção esvazia o acervo.</b>
    '''
    ''' Sem isto, <c>Itens</c> e <c>Ressalva</c> continuavam descrevendo a
    ''' última pasta depois de a seleção sumir — números sem dono, falando de
    ''' uma pasta que ninguém está olhando.
    ''' </summary>
    <TestMethod>
    Public Sub Perder_a_selecao_ESVAZIA_o_acervo()
        NoDispatcher(
            Sub(d)
                Using vm = AbrirVm(d)
                    Semear()
                    vm.Apontar(New FolderKey("entry-1", "store-1"), "Caixa",
                               New StoreInfo() With {.StoreId = "store-1"})
                    ' CONTROLE POSITIVO: a pasta semeada tem itens de verdade.
                    '
                    ' A primeira versao deste teste so olhava a Ressalva, porque
                    ' semear association parecia caro. Provava que o servico
                    ' passou a apontar para a sentinela zero -- e NAO provava o
                    ' "esvazia" do nome, porque Itens ja era 0 antes.
                    Assert.AreEqual(2, vm.Itens, "controle: a pasta semeada tem itens")
                    StringAssert.Contains(vm.Ressalva, "Acervo parcial",
                                          "controle: a pasta semeada foi varrida")

                    vm.Apontar(Nothing, Nothing, Nothing)

                    Assert.AreEqual(0, vm.Itens,
                        "o acervo continuou contando os itens da pasta anterior")
                    StringAssert.Contains(vm.Ressalva, "ainda não foi varrida",
                        "e continuou descrevendo ela")
                    Assert.IsFalse(vm.PodeVarrer)
                End Using
            End Sub)
    End Sub

    ''' <summary>
    ''' <b>Sem pasta escolhida não dá para varrer, e com pasta dá.</b>
    '''
    ''' <c>PodeVarrer</c> <b>não</b> olha autorização de ambiente de propósito:
    ''' quem recusa por isso é a varredura, e a recusa dela explica. Botão
    ''' desabilitado sem motivo na tela é o defeito que a faixa da IA já teve.
    ''' </summary>
    <TestMethod>
    Public Sub PodeVarrer_segue_a_selecao_e_nao_a_autorizacao()
        NoDispatcher(
            Sub(d)
                Using vm = AbrirVm(d)
                    Assert.IsFalse(vm.PodeVarrer, "sem pasta nao ha o que varrer")

                    vm.Apontar(New FolderKey("entry-1", "store-1"), "Caixa",
                               New StoreInfo() With {.StoreId = "store-1"})

                    Assert.IsTrue(vm.PodeVarrer,
                        "o ambiente nao esta autorizado, e mesmo assim o botao " &
                        "tem de estar de pe: quem explica a recusa e a varredura")
                End Using
            End Sub)
    End Sub

    ' ==================================================================

    ''' <summary>Uma pasta com duas mensagens publicadas, direto no banco.</summary>
    Private Sub Semear()
        Using conn As New SqliteConnection($"Data Source={_caminho}")
            conn.Open()
            ' A pasta nasce SEM geracao publicada: ela referencia generation,
            ' que so existe depois. A cabeca e apontada no UPDATE do fim, que e
            ' a mesma ordem que o CacheWriter usa.
            '
            ' (Comentario AQUI: em VB a continuacao implicita de { } nao aceita
            ' linha so de comentario, e o erro sai na linha anterior.)
            For Each sql In {
                "INSERT INTO environment_profile (environment_key, fingerprint, provider, " &
                "  cached_mode, policy_version, allowed) VALUES (1,'f','ExchangeCached',1,1,1)",
                "INSERT INTO store (store_key, provider_store_id) VALUES (1,'store-1')",
                "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
                "  reconcile_epoch, stability) VALUES (1,1,'entry-1',0,'estavel')",
                "INSERT INTO scan_attempt (attempt_key, folder_key, environment_key, " &
                "  universe_fingerprint, algorithm_version, reconcile_epoch, attempt_number, " &
                "  stage, rows_read, started_at) " &
                "VALUES (1,1,1,'u',1,0,1,'publicada',2,'2026-08-27T00:00:00.0000000+00:00')",
                "INSERT INTO coverage_observation (coverage_key, folder_key, environment_key, " &
                "  universe_fingerprint, coverage, source, observed_at) " &
                "VALUES (1,1,1,'u','parcial','varredura','2026-08-27T00:00:00.0000000+00:00')",
                "INSERT INTO generation (generation_key, folder_key, attempt_key, coverage_kind, " &
                "  coverage_key, universe_fingerprint, rows_read, count_before, count_after, " &
                "  discarded, distinct_keys, reconcile_epoch, published_at) " &
                "VALUES (1,1,1,'completa',1,'u',2,2,2,0,2,0,'2026-08-27T00:00:00.0000000+00:00')",
                "INSERT INTO item (item_key, created_at) VALUES (1,'2026-08-27T00:00:00.0000000+00:00')",
                "INSERT INTO item (item_key, created_at) VALUES (2,'2026-08-27T00:00:00.0000000+00:00')",
                "INSERT INTO incarnation (incarnation_key, item_key, folder_key, " &
                "  provider_entry_id, first_seen_generation, last_seen_generation) " &
                "VALUES (1,1,1,'E-1',1,1)",
                "INSERT INTO incarnation (incarnation_key, item_key, folder_key, " &
                "  provider_entry_id, first_seen_generation, last_seen_generation) " &
                "VALUES (2,2,1,'E-2',1,1)",
                "INSERT INTO association (association_key, item_key, folder_key, presence, " &
                "  observability, version, generation_key) " &
                "VALUES (1,1,1,'presente','observavel',1,1)",
                "INSERT INTO association (association_key, item_key, folder_key, presence, " &
                "  observability, version, generation_key) " &
                "VALUES (2,2,1,'presente','observavel',1,1)",
                "UPDATE folder SET published_generation_key = 1 WHERE folder_key = 1"} _

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = sql
                    cmd.ExecuteNonQuery()
                End Using
            Next
        End Using
    End Sub

    ''' <summary>
    ''' <b>Fechar a janela com a varredura EM VOO não toca o que já morreu.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O TESTE QUE FALTAVA, E POR QUE ELE DEMOROU</b>
    '''
    ''' Este é o defeito que nasceu de uma correção: o <c>Dispose</c> cancela e
    ''' <b>não espera</b> — para não travar o dispatcher no fechamento — e o
    ''' <c>SweepRunner</c> <b>captura</b> o cancelamento e devolve um
    ''' <c>SweepResult</c>. Então o <c>Await</c> não lança, e a continuação
    ''' seguia por <c>Recarregar()</c> e <c>Atualizar()</c> sobre um <c>_db</c>
    ''' já descartado.
    '''
    ''' A primeira tentativa de provar isso foi um teste com <c>Barrier</c> e
    ''' duas <c>Task</c>, que <b>passava com o defeito presente</b> — a janela
    ''' era curta demais para as duas se cruzarem. Foi apagado, e a lacuna
    ''' ficou escrita. O executor injetável é o que a torna determinística: a
    ''' varredura só devolve quando este teste mandar.
    '''
    ''' O sinal observado é o <c>Travado</c>. O executor devolve
    ''' <c>Nothing</c>, que no caminho normal viraria "O cache nao abriu para a
    ''' varredura" — então, se a guarda falhar, a frase aparece num objeto
    ''' descartado, e é exatamente isso que se cobra não acontecer.
    ''' </summary>
    <TestMethod>
    Public Sub Fechar_com_a_varredura_EM_VOO_nao_toca_o_que_morreu()
        NoDispatcherAsync(
            Async Function(d) As Task
                Dim entrou As New ManualResetEventSlim(False)
                Dim liberar As New ManualResetEventSlim(False)
                Dim vm = AbrirVm(d)

                vm.ExecutorDaVarredura =
                    Function(pasta, nome, store, ct)
                        entrou.Set()
                        liberar.Wait(TimeSpan.FromSeconds(10))
                        Return Nothing
                    End Function

                vm.Apontar(New FolderKey("entry-1", "store-1"), "Caixa",
                           New StoreInfo() With {.StoreId = "store-1"})
                Assert.IsTrue(vm.PodeVarrer, "controle: da para varrer")

                Dim voo = vm.VarrerCommand.ExecuteAsync(Nothing)

                ' Esperar FORA do dispatcher: bloquear aqui travaria a
                ' continuacao que se quer observar.
                '
                ' E O RESULTADO DA ESPERA E COBRADO. Descartar o Boolean de
                ' Wait era o defeito deste teste: no timeout ele seguia, via
                ' Varrendo = True e descartava -- exercitando "cancelado ANTES
                ' de o executor entrar", que e outro caso. A sincronizacao
                ' existia justamente para tirar temporizacao daqui, e o
                ' descarte a colocava de volta.
                Dim entrouDeVerdade = Await Task.Run(
                    Function() entrou.Wait(TimeSpan.FromSeconds(10)))
                Assert.IsTrue(entrouDeVerdade, "o executor nao chegou a entrar")
                Assert.IsTrue(vm.Varrendo, "controle: o voo esta em andamento")

                vm.Dispose()
                liberar.Set()
                Await voo

                Assert.IsNull(vm.Travado,
                    "a continuacao escreveu na tela depois do Dispose")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>Controle negativo: sem o Dispose, a mesma varredura escreve.</b>
    '''
    ''' Sem ele, um <c>Travado</c> que nunca fosse escrito — ou um comando que
    ''' não rodasse — faria o teste acima passar sem provar nada.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_sem_fechar_a_varredura_ESCREVE_na_tela()
        NoDispatcherAsync(
            Async Function(d) As Task
                Using vm = AbrirVm(d)
                    vm.ExecutorDaVarredura = Function(pasta, nome, store, ct) Nothing

                    vm.Apontar(New FolderKey("entry-1", "store-1"), "Caixa",
                               New StoreInfo() With {.StoreId = "store-1"})

                    Await vm.VarrerCommand.ExecuteAsync(Nothing)

                    Assert.IsNotNull(vm.Travado,
                        "sem Dispose no meio, o desfecho TEM de chegar a tela")
                    StringAssert.Contains(vm.Travado, "cache")
                End Using
            End Function)
    End Sub

End Class

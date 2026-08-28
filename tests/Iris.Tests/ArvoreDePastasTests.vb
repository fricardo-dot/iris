Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports Iris.App.ViewModels
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>TROCA DE SESSAO DURANTE A EXPANSAO DA ARVORE.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ESTE ARQUIVO NAO EXISTIA</b>
'''
''' O relatorio da Fase 2 listou esta categoria como "sustentada por leitura
''' de codigo", com a receita escrita: <i>pede um broker que segure o
''' carregamento de filhos</i>. O <c>FakeBroker</c> respondia "fora da
''' alcada" para stores e filhas, entao nao havia como dizer "a sessao caiu
''' ENQUANTO isto carregava" — e a frase seguinte do relatorio era a que
''' doia: <i>remover a guarda hoje deixa os 805 verdes</i>.
'''
''' A fonte bloqueavel entrou no <c>FakeBroker</c>, e estes testes existem
''' porque ela existe.
'''
''' ------------------------------------------------------------------
''' <b>AS DUAS GERACOES SAO DIFERENTES, E ISSO IMPORTA</b>
'''
''' A arvore tem uma geracao (<c>FolderTreeViewModel._generation</c>, que o
''' <c>Clear</c> incrementa) e cada NO tem a sua
''' (<c>FolderNodeViewModel._loadGeneration</c>, que o <c>Invalidate</c>
''' incrementa). Sao mecanismos distintos, protegendo coisas distintas: uma
''' impede repovoar a RAIZ depois da sessao cair, a outra impede repovoar um
''' RAMO depois de ele ser invalidado. Cada uma tem o seu teste, e o controle
''' negativo de uma deixa a outra verde.
''' </summary>
<TestClass>
Public Class ArvoreDePastasTests

    Private Shared ReadOnly Raiz As New FolderKey("raiz", "store-1")
    Private Shared ReadOnly Entrada As New FolderKey("entrada", "store-1")

    ' Roda o corpo numa STA com Dispatcher de verdade: a arvore usa
    ' Ui.InvokeAsync para povoar as ObservableCollection.
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
        b.ComStore("Caixa do teste", "store-1")
        b.ComPasta(Raiz, "Caixa de Entrada", "entrada", temFilhas:=True)
        Return b
    End Function

    Private Shared Function Arvore(b As FakeBroker, d As Dispatcher) As FolderTreeViewModel
        Return New FolderTreeViewModel(b, d, Sub(t As Task, nome As String)
                                             End Sub)
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>Controle: sem interferencia, a arvore carrega.</b>
    '''
    ''' Existe para que os dois testes seguintes signifiquem alguma coisa. Sem
    ''' ele, uma arvore que simplesmente nunca carrega passaria nos dois — e
    ''' e exatamente a armadilha que o CLAUDE.md descreve: um compositor que
    ''' nunca envia passa em todos os testes de nao enviar errado.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_interferencia_a_arvore_carrega()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                Dim arv = Arvore(b, d)
                Await arv.ReloadAsync()
                Assert.AreEqual(1, arv.Roots.Count, "controle: a arvore tinha de ter carregado")
                Assert.AreEqual("Caixa de Entrada", arv.Roots(0).Name)
            End Function)
    End Sub

    ''' <summary>
    ''' <b>A sessao cai no meio do reload, e a arvore NAO se repovoa.</b>
    '''
    ''' O <c>Clear</c> incrementa a geracao justamente para que uma resposta em
    ''' transito nao devolva pastas de uma sessao que nao existe mais. Mostrar
    ''' a arvore da sessao morta seria pior que mostrar arvore nenhuma: o
    ''' usuario clicaria numa pasta que o broker ja nao consegue abrir.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTE TESTE PROVA, MEDIDO E NAO SUPOSTO</b>
    '''
    ''' Eu ia escrever que ele cobre o <c>Atual(geracao)</c> de dentro do
    ''' <c>InvokeAsync</c>. Fui medir, e nao cobre:
    '''
    '''   • so a 1a guarda removida .... <b>passa</b> (a 2a segura)
    '''   • so a 3a guarda removida .... <b>passa</b> (a 1a segura)
    '''   • as tres removidas .......... <b>falha</b>
    '''
    ''' Os tres <c>Atual(geracao)</c> do <c>ReloadAsync</c> sao uma
    ''' <b>corrente</b>: parar o broker em qualquer ponto faz a proxima
    ''' conferencia pegar, e cada uma e independentemente suficiente para o
    ''' ponto em que esta. Entao o que este teste prova e a propriedade da
    ''' CADEIA — <i>o reload nao repovoa a arvore depois do Clear</i> — e nao
    ''' uma guarda especifica.
    '''
    ''' A terceira, dentro do <c>InvokeAsync</c>, so seria alcancavel se a
    ''' geracao mudasse entre a segunda conferencia e o despacho para a UI.
    ''' Escrever isso de um teste que roda NA propria UI nao da; fica como
    ''' defesa em profundidade, e dito assim.
    ''' </summary>
    <TestMethod>
    Public Sub Clear_durante_o_reload_NAO_repovoa_a_arvore()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                b.TravaDosStores = New TaskCompletionSource(Of Boolean)()
                Dim arv = Arvore(b, d)

                Dim voo = arv.ReloadAsync()
                Assert.IsTrue(b.Chamadas.Contains("GetStores"),
                    "controle: o reload tinha de estar parado dentro do broker")

                arv.Clear()
                b.TravaDosStores.SetResult(True)
                Await voo

                Assert.AreEqual(0, arv.Roots.Count,
                    "o reload da sessao antiga repovoou a arvore depois do Clear")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>O no e invalidado no meio da expansao, e os filhos velhos NAO entram.</b>
    '''
    ''' Aqui o no NAO esta expandido de proposito. <c>Invalidate</c> so
    ''' redispara a carga quando <c>_isExpanded</c> e verdadeiro, e um segundo
    ''' voo bem-sucedido povoaria <c>Children</c> por conta propria — o teste
    ''' passaria sem que a guarda existisse, medindo o voo novo em vez do
    ''' velho.
    '''
    ''' Com o no recolhido, o unico candidato a povoar <c>Children</c> e o voo
    ''' que ficou para tras. Se ele conseguir, a guarda nao esta la.
    '''
    ''' <b>Medido, mesmo padrao do teste da arvore:</b> removendo so a
    ''' conferencia de depois do broker, passa; removendo so a de dentro do
    ''' <c>InvokeAsync</c>, passa; removendo <b>as duas</b>, falha com 1
    ''' filho. Sao independentemente suficientes, e o que fica provado e a
    ''' propriedade do par, nao de cada uma.
    ''' </summary>
    <TestMethod>
    Public Sub Invalidate_durante_a_expansao_NAO_repovoa_o_no()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                b.ComPasta(Entrada, "Uma subpasta", "sub")
                Dim arv = Arvore(b, d)
                Await arv.ReloadAsync()

                Dim no = arv.Roots(0)
                Assert.IsTrue(no.CanExpand, "controle: o no tinha de poder expandir")
                Assert.IsFalse(no.IsExpanded, "controle: o no NAO pode estar expandido")

                b.TravaDasFilhas = New TaskCompletionSource(Of Boolean)()
                Dim voo = no.EnsureChildrenAsync()
                Assert.IsTrue(b.PedidosDeFilhas.Contains("store-1|entrada"),
                    "controle: a expansao tinha de estar parada dentro do broker")

                no.Invalidate()
                b.TravaDasFilhas.SetResult(True)
                Await voo

                Assert.AreEqual(0, no.Children.Count,
                    "o voo da geracao vencida povoou o no depois do Invalidate")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>E quem chega no meio da carga ESPERA, em vez de concluir que a pasta nao existe.</b>
    '''
    ''' A outra metade do <c>_sinalDeCarga</c>, e a que ja custou um defeito
    ''' real: encontrar o no em <c>Loading</c> fazia <c>EnsureChildrenAsync</c>
    ''' voltar na hora, e quem restaurava um caminho descia para um nivel ainda
    ''' vazio e concluia "a pasta nao existe" — sem selecao e sem assinatura,
    ''' em silencio.
    '''
    ''' Duas chamadas concorrentes tem de produzir <b>um</b> pedido ao broker,
    ''' e as duas so voltam com os filhos ja no lugar.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A ASSERCAO QUE FALTAVA, E SEM ELA O TESTE ERA DECORATIVO</b>
    '''
    ''' A primeira versao criava a segunda chamada e so olhava para ela depois
    ''' de liberar a trava. Com o defeito antigo presente — o ramo
    ''' <c>Loading</c> voltando na hora — o teste continuava <b>verde</b>:
    ''' havia um pedido so, a primeira chamada acabava povoando os filhos, e o
    ''' <c>WhenAll</c> terminava.
    '''
    ''' O que separa "esperou" de "desistiu" e o estado da segunda tarefa
    ''' <b>enquanto o broker ainda esta travado</b>. Se ela ja terminou ali,
    ''' ela desistiu.
    ''' </summary>
    <TestMethod>
    Public Sub Quem_chega_no_meio_da_carga_ESPERA_o_mesmo_voo()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                b.ComPasta(Entrada, "Uma subpasta", "sub")
                Dim arv = Arvore(b, d)
                Await arv.ReloadAsync()

                Dim no = arv.Roots(0)
                b.TravaDasFilhas = New TaskCompletionSource(Of Boolean)()

                Dim primeiro = no.EnsureChildrenAsync()
                Dim segundo = no.EnsureChildrenAsync()

                ' AQUI, com a trava ainda fechada. Quem espera nao pode ter
                ' terminado; quem desiste ja terminou.
                Assert.IsFalse(segundo.IsCompleted,
                    "o segundo pedido voltou na hora em vez de esperar o primeiro")
                Assert.IsFalse(primeiro.IsCompleted,
                    "controle: o primeiro tinha de estar preso no broker")

                b.TravaDasFilhas.SetResult(True)
                Await Task.WhenAll(primeiro, segundo)

                Dim pedidos = b.PedidosDeFilhas.FindAll(Function(x) x = "store-1|entrada").Count
                Assert.AreEqual(1, pedidos,
                    "o segundo pedido nao esperou o primeiro: dois voos para a mesma pasta")
                Assert.AreEqual(1, no.Children.Count,
                    "quem esperou tinha de voltar com os filhos no lugar")
            End Function)
    End Sub

End Class

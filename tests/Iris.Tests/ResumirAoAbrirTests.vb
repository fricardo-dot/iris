Imports System.IO
Imports System.Threading.Tasks
Imports Iris.App.ViewModels
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>RESUMIR AO ABRIR — a Fase 1.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTE ARQUIVO PRENDE</b>
'''
''' Abrir uma mensagem passa a mandar conteúdo para fora sem clique nenhum.
''' Isso é uma mudança de <b>categoria</b>, e não de conforto, e é por isso que
''' aqui há mais teste de <i>quando NÃO vai</i> do que de quando vai:
'''
''' <list type="number">
''' <item>Desligado por padrão, e a decisão mora em disco.</item>
''' <item>Mensagem já resumida não é resumida de novo — ir e voltar na lista
''' não pode cobrar duas vezes pelo mesmo texto.</item>
''' <item>Descer a lista com a seta não dispara um pedido por linha: a espera
''' existe justamente para a troca seguinte cancelá-la <b>antes</b> de haver
''' pedido. Cancelar não desfaz requisição que já saiu.</item>
''' <item>Sem chave — troca de pasta, desmarcação — não resume nada.</item>
''' </list>
'''
''' ------------------------------------------------------------------
''' <b>O MARCADOR É REAL, E ISSO É DE PROPÓSITO</b>
'''
''' O interruptor grava em <c>%LOCALAPPDATA%\Iris</c>, e estes testes mexem no
''' arquivo de verdade. Um duplo em memória provaria que a propriedade guarda um
''' booleano, que não é a pergunta — a pergunta é se a decisão <b>sobrevive ao
''' fechamento do programa</b>. O estado anterior é salvo e devolvido no fim.
''' </summary>
' NAO PARALELIZAR: o marcador em disco e um so, e Montar mexe em estado Shared.
<TestClass>
<DoNotParallelize>
Public Class ResumirAoAbrirTests

    Private _estadoAnterior As Boolean

    <TestInitialize>
    Public Sub Guardar()
        _estadoAnterior = File.Exists(AssistenteViewModel.CaminhoDoResumoAutomatico())
        Apagar()
    End Sub

    <TestCleanup>
    Public Sub Devolver()
        If _estadoAnterior Then
            Dim caminho = AssistenteViewModel.CaminhoDoResumoAutomatico()
            Directory.CreateDirectory(Path.GetDirectoryName(caminho))
            File.WriteAllText(caminho, "devolvido pelo teste")
        Else
            Apagar()
        End If
    End Sub

    Private Shared Sub Apagar()
        Dim caminho = AssistenteViewModel.CaminhoDoResumoAutomatico()
        If File.Exists(caminho) Then File.Delete(caminho)
    End Sub

    Private Shared Function Chave(n As Integer) As ItemKey
        Return New ItemKey($"E-{n}", "store-1")
    End Function

    ''' <summary>Um assistente com a espera zerada — o teste não pode esperar.</summary>
    Private Shared Function Montar(p As AssistenteViewModelTests.ProvedorControlado) _
                                   As AssistenteViewModel
        Dim vm = AssistenteViewModelTests.Montar(AssistenteViewModelTests.Ativacao(), p,
                                                 AssistenteViewModelTests.Pronta())
        vm.EsperaAntesDeResumir = TimeSpan.Zero
        Return vm
    End Function

    <TestMethod>
    Public Sub Nasce_DESLIGADO()
        Dim vm = Montar(New AssistenteViewModelTests.ProvedorControlado())

        Assert.IsFalse(vm.ResumirAoAbrir,
            "mandar conteudo para fora sem clique nao pode ser o padrao")
    End Sub

    ''' <summary>
    ''' A decisão sobrevive ao programa. Um interruptor que volta a zero a cada
    ''' abertura não é um interruptor — é um botão que finge.
    ''' </summary>
    <TestMethod>
    Public Sub A_decisao_fica_em_DISCO()
        Dim p As New AssistenteViewModelTests.ProvedorControlado()
        Montar(p).ResumirAoAbrir = True

        Assert.IsTrue(Montar(p).ResumirAoAbrir,
            "outro assistente, no mesmo perfil, tem de ver a mesma decisao")

        Montar(p).ResumirAoAbrir = False
        Assert.IsFalse(Montar(p).ResumirAoAbrir, "e desligar tambem tem de ficar")
    End Sub

    <TestMethod>
    Public Async Function Ligado_abrir_a_mensagem_RESUME() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo"}
        Dim vm = Montar(p)
        vm.ResumirAoAbrir = True

        vm.Trocou(Chave(1))
        Await vm.EsperarOResumoAutomatico()

        Assert.AreEqual(1, p.Chamadas, "com o interruptor ligado, abrir tem de resumir")
        Assert.AreEqual("o resumo", vm.Resultado)
    End Function

    ''' <summary>
    ''' <b>O controle negativo do interruptor.</b> Sem ele, um assistente que
    ''' resumisse sempre passaria em todos os outros testes deste arquivo.
    ''' </summary>
    <TestMethod>
    Public Async Function DESLIGADO_abrir_a_mensagem_nao_manda_nada() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo"}
        Dim vm = Montar(p)

        vm.Trocou(Chave(1))
        Await vm.EsperarOResumoAutomatico()

        Assert.AreEqual(0, p.Chamadas,
            "desligado, abrir uma mensagem nao pode mandar nada para fora")
        Assert.AreEqual("", vm.Resultado)
    End Function

    ''' <summary>
    ''' Ir e voltar não cobra duas vezes. É a memória por mensagem fazendo o
    ''' trabalho — e sem esta conferência o automático transformaria cada
    ''' vaivém na lista numa chamada nova.
    ''' </summary>
    <TestMethod>
    Public Async Function Voltar_para_a_mensagem_ja_resumida_NAO_pede_de_novo() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo"}
        Dim vm = Montar(p)
        vm.ResumirAoAbrir = True

        vm.Trocou(Chave(1))
        Await vm.EsperarOResumoAutomatico()
        vm.Trocou(Chave(2))
        Await vm.EsperarOResumoAutomatico()
        vm.Trocou(Chave(1))
        Await vm.EsperarOResumoAutomatico()

        Assert.AreEqual(2, p.Chamadas,
            "a terceira troca voltou para uma mensagem ja resumida: tem de vir " &
            "da memoria, e nao do provedor")
        Assert.AreEqual("o resumo", vm.Resultado, "e o resumo tem de estar na tela")
    End Function

    ''' <summary>
    ''' <b>A ESPERA DEIXA PASSAR — o controle positivo do caminho lento.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ESTE TESTE FALTAVA, E A FALTA DELE ESCONDEU O RECURSO INTEIRO</b>
    '''
    ''' Todos os testes deste arquivo zeravam <c>EsperaAntesDeResumir</c>, e
    ''' espera zero <b>pula o <c>Delay</c></b>. O único que usava espera real
    ''' cobrava <i>zero chamadas</i> — e zero era exatamente o que o defeito
    ''' produzia: a espera lia um <c>CancellationTokenSource</c> que era
    ''' <c>Nothing</c> na maior parte do tempo, estourava dentro de uma task
    ''' que ninguém aguardava, e o resumo automático nunca acontecia.
    '''
    ''' Em produção a espera é 800 ms, então <b>só</b> o caminho quebrado
    ''' rodava. A suíte inteira verde, e o recurso nunca funcionou. Achado por
    ''' revisão externa.
    '''
    ''' A espera aqui é curta mas <b>não é zero</b>, que é o ponto: o
    ''' <c>Delay</c> tem de ser percorrido de verdade.
    ''' </summary>
    <TestMethod>
    Public Async Function Depois_da_ESPERA_o_resumo_acontece() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo"}
        Dim vm = Montar(p)
        vm.EsperaAntesDeResumir = TimeSpan.FromMilliseconds(20)
        vm.ResumirAoAbrir = True

        vm.Trocou(Chave(1))
        Await vm.EsperarOResumoAutomatico()

        Assert.AreEqual(1, p.Chamadas,
            "a espera terminou e ninguem pediu nada: o caminho lento esta quebrado")
        Assert.AreEqual("o resumo", vm.Resultado)
    End Function

    ''' <summary>
    ''' <b>Descer a lista com a seta não dispara um pedido por linha.</b>
    '''
    ''' A espera é real — 5 segundos — e nenhuma das oito trocas chega a pedir,
    ''' porque a seguinte cancela a espera da anterior <b>antes</b> de haver
    ''' pedido. Cancelar depois não serviria: requisição que saiu não volta, e
    ''' o duplo do provedor desta base existe para lembrar disso.
    '''
    ''' O controle positivo do fim usa espera <b>curta e não nula</b>, e não
    ''' zero: com zero ele passaria pelo atalho que pula o <c>Delay</c>, e
    ''' provaria menos do que promete — foi assim que a versão anterior deste
    ''' teste passou com a espera inteiramente quebrada.
    ''' </summary>
    <TestMethod>
    Public Async Function Descer_a_lista_depressa_NAO_dispara_um_pedido_por_linha() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo"}
        Dim vm = Montar(p)
        vm.EsperaAntesDeResumir = TimeSpan.FromSeconds(5)
        vm.ResumirAoAbrir = True

        For n = 1 To 8
            vm.Trocou(Chave(n))
        Next

        Assert.AreEqual(0, p.Chamadas,
            "oito linhas atravessadas viraram pedido: a espera nao esta " &
            "segurando nada")

        vm.EsperaAntesDeResumir = TimeSpan.FromMilliseconds(20)
        vm.Trocou(Chave(9))
        Await vm.EsperarOResumoAutomatico()
        Assert.AreEqual(1, p.Chamadas,
            "a linha em que se PARA tem de ser resumida -- e pelo caminho da " &
            "espera, senao este teste passaria num assistente que nunca resume")
    End Function

    ''' <summary>
    ''' Sem chave não há mensagem: trocar de pasta ou desmarcar não é abrir
    ''' nada, e não pode gastar chamada.
    ''' </summary>
    <TestMethod>
    Public Async Function Sem_chave_nao_resume() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo"}
        Dim vm = Montar(p)
        vm.ResumirAoAbrir = True

        vm.Trocou()
        Await vm.EsperarOResumoAutomatico()

        Assert.AreEqual(0, p.Chamadas, "trocar de pasta nao e abrir uma mensagem")
    End Function

    ' ==================================================================
    ' O RESUMO DE UMA LINHA

    ''' <summary>
    ''' A frase e o resto saem do <b>mesmo</b> texto. Dois pedidos seriam a
    ''' mesma leitura cobrada duas vezes.
    ''' </summary>
    <TestMethod>
    Public Async Function A_primeira_linha_e_o_resto_saem_do_mesmo_texto() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {
            .Texto = "Caroline mandou os códigos novos." & vbCrLf & vbCrLf &
                     "• 2348 virou 3387" & vbCrLf & "• 2280 virou 3389"}
        Dim vm = Montar(p)

        Await vm.Resumir()

        Assert.AreEqual("Caroline mandou os códigos novos.", vm.ResumoDeUmaLinha)
        StringAssert.Contains(vm.ResumoDetalhado, "2348 virou 3387")
        Assert.IsTrue(vm.TemResumoDetalhado)
        Assert.AreEqual(1, p.Chamadas, "uma leitura, e nao duas")
    End Function

    ''' <summary>
    ''' <b>O modelo ignorou o formato.</b> Um parágrafo só, sem linha em branco.
    ''' A frase vira o parágrafo inteiro e o detalhe some — pior que o pedido,
    ''' melhor que vazio, e <b>nada inventado</b>.
    ''' </summary>
    <TestMethod>
    Public Async Function Sem_o_formato_pedido_nada_e_inventado() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {
            .Texto = "um parágrafo único que o modelo devolveu do jeito dele"}
        Dim vm = Montar(p)

        Await vm.Resumir()

        Assert.AreEqual("um parágrafo único que o modelo devolveu do jeito dele",
                        vm.ResumoDeUmaLinha)
        Assert.AreEqual("", vm.ResumoDetalhado)
        Assert.IsFalse(vm.TemResumoDetalhado, "caixa vazia nao pode ocupar espaco")
    End Function

End Class

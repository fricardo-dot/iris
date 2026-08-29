Imports Iris.Model
Imports System.Linq
Imports Iris.Cache
Imports Iris.Integration
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O SEGUNDO PASSE DA BUSCA — e o número que o encomendou.</b>
'''
''' ------------------------------------------------------------------
''' <b>A MEDIÇÃO QUE FECHOU A FASE 4</b>
'''
''' O ESCOPO deixou a busca semântica parada com oito decisões abertas, e
''' nomeou a condição para reavaliar: <i>evidência de que a busca textual
''' normalizada não resolve</i>. Ninguém tinha produzido a evidência.
'''
''' Em 29/08/2026 ela foi medida (<c>tools/medir-busca.py</c>): 300 mensagens
''' do acervo real, consultas derivadas do <b>próprio assunto</b> de cada uma —
''' sem oráculo, porque a mensagem é a resposta certa por construção.
'''
''' <code>
''' exato                  100,0%      erro_de_digitacao     0,4%
''' sem_acento             100,0%      flexao_de_numero      0,0%
''' caixa_alta             100,0%
''' fora_de_ordem          100,0%
''' prefixo_da_palavra     100,0%
''' assunto_mais_remetente 100,0%
''' </code>
'''
''' <b>Tudo o que a busca prometia valia. As duas falhas eram mecânicas — e
''' nenhuma delas precisa de significado para ser consertada.</b> Numa caixa em
''' português, "reuniões" não acha "reunião" porque o singular não é subcadeia
''' do plural; e uma letra trocada zera a busca.
'''
''' Isso é o oposto de "a busca textual não resolve, precisamos de embeddings".
''' É "a busca textual resolve o que promete, e o que falta custa um radical
''' pobre e uma distância de edição".
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTA SUÍTE PRENDE</b>
'''
''' Duas coisas, e a segunda é a que importa mais: que os casos novos passem, e
''' que os <b>seis de 100% continuem em 100%</b>. Um conserto que trocasse
''' defeito por defeito passaria em metade destes testes.
''' </summary>
<TestClass>
Public Class BuscaToleranteTests

    Private Shared Function Item(assunto As String, remetente As String) As ManifestItem
        Return New ManifestItem("id", assunto, remetente, "2026-08-29T00:00:00Z",
                                False, Iris.Sync.PresenceState.Presente)
    End Function

    ''' <summary>
    ''' Monta a busca de verdade contra itens de mentira, e roda a tela em
    ''' cima do resultado. O ResultadoDaBusca tem construtor Friend e a suite
    ''' esta no assembly de amigos, entao nao ha banco no caminho.
    ''' </summary>
    Private Shared Function Procurar(termo As String,
                                     itens As ManifestItem()) As Iris.App.ViewModels.BuscaViewModel
        Dim t As New TermoDeBusca(termo)
        Dim achados = itens.
                      Select(Function(i) New AchadoDaBusca(1, "Caixa de Entrada", i, t.Grau(i))).
                      Where(Function(x) x.Grau <> GrauDoAchado.Nenhum).
                      ToList()
        Dim pasta As New PastaConsultada(1, "Caixa de Entrada", 1,
                                         Iris.Sync.FolderCoverage.Parcial,
                                         "2026-08-29T00:00:00Z", "Acervo parcial.", 9)
        Dim r As New ResultadoDaBusca(t, achados, {pasta},
                                      Array.Empty(Of PastaConsultada)())

        Dim vm As New Iris.App.ViewModels.BuscaViewModel(Function(x) r)
        vm.Termo = termo
        vm.ProcurarCommand.Execute(Nothing)
        Return vm
    End Function

    Private Shared Function Grau(termo As String, assunto As String,
                                 Optional remetente As String = "Quem Manda") As GrauDoAchado
        Return New TermoDeBusca(termo).Grau(Item(assunto, remetente))
    End Function

    ' ==================================================================
    ' A LINHA DE BASE. Estes seis eram 100% antes, e continuam sendo.

    ''' <summary>
    ''' <b>Os seis casos medidos em 100% continuam EXATOS.</b>
    '''
    ''' Não "continuam achando" — continuam achando <b>no primeiro passe</b>. Um
    ''' conserto que rebaixasse um destes a "aproximado" estaria degradando a
    ''' certeza para consertar o palpite, que é o negócio ao contrário.
    '''
    ''' <b>Controle negativo:</b> fazendo o <c>Grau</c> devolver
    ''' <c>Aproximado</c> antes de testar a subcadeia, todos os seis caem.
    ''' </summary>
    <DataTestMethod>
    <DataRow("contrato aditivo", "Contrato aditivo de fornecimento", "exato")>
    <DataRow("regulatorio", "Assunto Regulatório da semana", "sem acento")>
    <DataRow("CONTRATO", "Contrato aditivo", "caixa alta")>
    <DataRow("aditivo contrato", "Contrato aditivo", "fora de ordem")>
    <DataRow("regulat", "Regulatório", "pedaço da palavra")>
    <DataRow("contrato manda", "Contrato aditivo", "assunto + remetente")>
    Public Sub Os_seis_casos_de_cem_por_cento_continuam_EXATOS(termo As String,
                                                               assunto As String,
                                                               caso As String)
        Assert.AreEqual(GrauDoAchado.Exato, Grau(termo, assunto),
            $"o caso '{caso}' era 100% exato na medição de 29/08 e deixou de ser")
    End Sub

    ' ==================================================================
    ' O QUE A MEDIÇÃO ACHOU QUEBRADO

    ''' <summary>
    ''' <b>Flexão de número — era 0% de 64 consultas.</b>
    '''
    ''' Só as terminações em que o singular <b>não é subcadeia</b> do plural.
    ''' Onde ele é ("contratos" contém "contrato"), o primeiro passe já achava.
    ''' </summary>
    <DataTestMethod>
    <DataRow("reunioes", "Reunião de diretoria")>
    <DataRow("contratuais", "Cláusula contratual")>
    <DataRow("armazens", "Armazém central")>
    <DataRow("reuniao", "Pauta das reuniões")>
    Public Sub Flexao_de_numero_acha_por_APROXIMACAO(termo As String, assunto As String)
        Assert.AreEqual(GrauDoAchado.Aproximado, Grau(termo, assunto),
            $"'{termo}' não achou '{assunto}': flexão era 0% de 64 na medição")
    End Sub

    ''' <summary>
    ''' <b>Erro de digitação — era 0,4% de 242 consultas.</b>
    '''
    ''' Uma letra: trocada, faltando, ou sobrando.
    ''' </summary>
    <DataTestMethod>
    <DataRow("contrado", "Contrato aditivo", "letra trocada")>
    <DataRow("contrto", "Contrato aditivo", "letra faltando")>
    <DataRow("contratto", "Contrato aditivo", "letra sobrando")>
    Public Sub Erro_de_digitacao_acha_por_APROXIMACAO(termo As String,
                                                      assunto As String, caso As String)
        Assert.AreEqual(GrauDoAchado.Aproximado, Grau(termo, assunto),
            $"'{caso}' não achou: erro de digitação era 0,4% de 242 na medição")
    End Sub

    ' ==================================================================
    ' E O QUE A TOLERÂNCIA NÃO PODE FAZER

    ''' <summary>
    ''' <b>DUAS LETRAS ERRADAS NÃO É ERRO DE DIGITAÇÃO, É OUTRA PALAVRA.</b>
    '''
    ''' Este é o controle que impede a tolerância de virar "acha tudo". Um
    ''' aproximado que casa com meio dicionário não é aproximado — é a lista
    ''' inteira com outro nome, e a medição já mostrou que esta busca tem
    ''' problema de ruído (67,7% das consultas por pedaço de palavra casam com
    ''' mais de dez mensagens).
    ''' </summary>
    ' "fatura" NAO entra aqui, e a primeira versao deste teste errou nisso:
    ' "fatura" e subcadeia de "Faturamento", entao o PRIMEIRO passe casa --
    ' e casar por pedaco de palavra e o caso que a medicao deu 100% e que
    ' eu disse que ficaria. O teste acusava a tolerancia por um acerto que
    ' nao e dela.
    <DataTestMethod>
    <DataRow("contrado", "Contrabando apreendido")>
    <DataRow("relatorio", "Repositorio central", "")>
    <DataRow("casa", "caso", "")>
    Public Sub Longe_demais_NAO_acha(termo As String, assunto As String,
                                     Optional ignorado As String = "")
        Assert.AreEqual(GrauDoAchado.Nenhum, Grau(termo, assunto, remetente:=""),
            $"'{termo}' achou '{assunto}': a tolerância virou 'acha tudo'")
    End Sub

    ''' <summary>
    ''' <b>Palavra curta não ganha tolerância a erro.</b>
    '''
    ''' Distância 1 sobre quatro letras casa metade do dicionário. O corte é em
    ''' cinco, e é arbitrário no valor e não no motivo.
    ''' </summary>
    <TestMethod>
    Public Sub Palavra_curta_nao_tolera_erro()
        ' "nota" vs "nata": UMA letra de diferenca, quatro letras. O primeiro
        ' corte deste teste usava "nota" vs "nada", que sao DUAS letras -- nao
        ' casava com piso nem sem piso, e o controle negativo passou quando
        ' devia falhar. Um teste que nao alcanca a guarda nao prova a guarda.
        Assert.AreEqual(GrauDoAchado.Nenhum, Grau("nota", "Nata fresca", remetente:=""),
            "tolerou erro numa palavra de quatro letras: distancia 1 sobre " &
            "palavra curta casa metade do dicionario")

        ' CONTROLE: acima do piso, a mesma diferenca de uma letra casa.
        Assert.AreEqual(GrauDoAchado.Aproximado, Grau("notas", "Natas frescas", remetente:=""),
            "controle: cinco letras tinham de tolerar")
    End Sub

    ''' <summary>
    ''' <b>E o radical não inventa palavra curta.</b>
    '''
    ''' A regra que conserta "contratuais" → "contratual" transformaria "mais"
    ''' em "mal" se não houvesse piso de tamanho. Um radical que fabrica palavra
    ''' de três letras gera ruído em cima de ruído.
    ''' </summary>
    <TestMethod>
    Public Sub O_radical_tem_piso_de_tamanho()
        Assert.AreEqual("mais", TermoDeBusca.Radical("mais"),
                        "o radical inventou 'mal' a partir de 'mais'")
        Assert.AreEqual("reuniao", TermoDeBusca.Radical("reunioes"),
                        "controle: acima do piso o radical funciona")
    End Sub

    ''' <summary>
    ''' <b>Conjunção continua valendo no segundo passe.</b>
    '''
    ''' Duas palavras exigem as duas, aproximadas ou não. Um segundo passe que
    ''' virasse disjunção devolveria tudo o que tem uma delas — ruído com cara
    ''' de resultado, que é o que o comentário do <c>TermoDeBusca</c> já
    ''' proibia para o primeiro passe.
    ''' </summary>
    <TestMethod>
    Public Sub A_conjuncao_vale_no_segundo_passe()
        ' "contrado" casa por aproximação; "jacaré" não casa de jeito nenhum.
        Assert.AreEqual(GrauDoAchado.Nenhum,
                        Grau("contrado jacare", "Contrato aditivo", remetente:=""),
            "o segundo passe virou disjunção: achou com uma palavra só")
    End Sub

    ' ==================================================================

    ''' <summary>
    ''' <b>A distância de edição, sozinha.</b>
    '''
    ''' Testada direto porque ela para no segundo erro em vez de calcular a
    ''' matriz — uma otimização que é fácil de escrever quase certa.
    ''' </summary>
    <DataTestMethod>
    <DataRow("abc", "abc", True)>
    <DataRow("abc", "abd", True)>
    <DataRow("abc", "abcd", True)>
    <DataRow("abcd", "abc", True)>
    <DataRow("abc", "xbd", False)>
    <DataRow("abc", "abcde", False)>
    <DataRow("", "a", True)>
    <DataRow("abc", "cba", False)>
    Public Sub A_distancia_para_no_segundo_erro(a As String, b As String, esperado As Boolean)
        Assert.AreEqual(esperado, TermoDeBusca.DistanciaAte1(a, b),
                        $"'{a}' vs '{b}'")
    End Sub

    ''' <summary>
    ''' <b>Zero achados continua sendo zero achados.</b>
    '''
    ''' A tolerância não pode transformar "não achei" em "achei alguma coisa" —
    ''' seria a §23 pelo avesso: fabricar presença em vez de ausência.
    ''' </summary>
    <TestMethod>
    Public Sub Termo_sem_nada_a_ver_continua_sem_achar()
        Assert.AreEqual(GrauDoAchado.Nenhum,
                        Grau("jacare bicicleta", "Contrato aditivo", remetente:=""))
    End Sub

    <TestMethod>
    Public Sub Termo_vazio_nao_acha_nada()
        Assert.AreEqual(GrauDoAchado.Nenhum, Grau("", "Contrato aditivo"))
        Assert.AreEqual(GrauDoAchado.Nenhum, Grau("   ", "Contrato aditivo"))
    End Sub

    ''' <summary>
    ''' <b>A ORDEM: EXATOS PRIMEIRO.</b>
    '''
    ''' Um achado aproximado é um palpite bom. Palpite no topo da lista empurra
    ''' a certeza para baixo da dobra — quem olha os três primeiros resultados
    ''' veria três palpites, e a lista teria mentido pela ordem sem mentir por
    ''' nenhuma linha.
    '''
    ''' <b>Controle negativo:</b> tirando o <c>OrderBy</c> do
    ''' <c>BuscaViewModel</c>, este teste cai.
    ''' </summary>
    <TestMethod>
    Public Sub Os_exatos_vem_antes_dos_aproximados()
        Dim vm = Procurar("contrado",
                          {Item("Contrato aditivo", ""),
                           Item("Contrado assinado", ""),
                           Item("Contrado revisado", "")})

        Assert.AreEqual(3, vm.Achados.Count, "controle: os tres tinham de casar")
        Assert.IsFalse(vm.Achados(0).Aproximado,
            "um palpite ficou no topo, empurrando a certeza para baixo")
        Assert.IsTrue(vm.Achados(vm.Achados.Count - 1).Aproximado,
            "controle: havia aproximado para vir depois")
    End Sub

    ''' <summary>
    ''' <b>A contagem de aproximados é o que a tela mostra.</b>
    '''
    ''' E ela é <b>zero</b> quando tudo casou exato — senão a frase de ressalva
    ''' apareceria sobre uma lista que não tem palpite nenhum, e ressalva que
    ''' aparece sempre é ressalva que ninguém lê.
    ''' </summary>
    <TestMethod>
    Public Sub A_contagem_de_aproximados_e_a_da_tela()
        Dim so_exatos = Procurar("contrato", {Item("Contrato aditivo", "")})
        Assert.AreEqual(0, so_exatos.Aproximados)
        Assert.IsFalse(so_exatos.TemAproximados, "ressalvou uma lista sem palpite")
        Assert.AreEqual("", so_exatos.FraseDosAproximados)

        Dim com_palpite = Procurar("contrado", {Item("Contrato aditivo", "")})
        Assert.AreEqual(1, com_palpite.Aproximados)
        Assert.IsTrue(com_palpite.TemAproximados)
        StringAssert.Contains(com_palpite.FraseDosAproximados, "aproximação",
            "a frase nao diz que sao aproximados")
        StringAssert.Contains(com_palpite.FraseDosAproximados, "letra",
            "a frase nao diz aproximados COMO -- ressalva decorativa")
    End Sub

End Class

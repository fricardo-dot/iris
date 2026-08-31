Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Assist
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A SUPERFÍCIE DO CLASSIFICADOR — Fase 4.</b>
'''
''' ------------------------------------------------------------------
''' <b>ONDE MORA A BARREIRA CONTRA O E-MAIL QUE DÁ ORDENS</b>
'''
''' O corpo de cada mensagem veio de fora, e um e-mail pode dizer <i>"ignore as
''' instruções acima"</i>. Pedir ao modelo, em português, que trate o corpo como
''' dado é persuasão — e persuasão não é barreira.
'''
''' A barreira tem duas partes, e as duas ficam presas aqui:
'''
''' <list type="number">
''' <item><b>A forma da resposta.</b> Sai <c>{item_key, label, confidence}</c>
''' com o rótulo restrito a um enum — e mais nada.</item>
''' <item><b>A ficha opaca.</b> O identificador que vai no fio é cunhado por
''' lote e não aparece em corpo de e-mail nenhum, então <b>o conteúdo não tem
''' como nomear a mensagem do vizinho</b>. Sem isso, um e-mail hostil pedia que
''' a classificação dele fosse escrita na chave de outro, e as contagens
''' fechavam.</item>
''' </list>
'''
''' ------------------------------------------------------------------
''' <b>A REGRA DURA, E O CONTRAPONTO</b>
'''
''' Desencontro de identidade derruba o lote <b>inteiro</b>: se uma veio
''' trocada, não há razão para crer nas outras. Rótulo inventado derruba
''' <b>só o item</b>, e sem esse contraponto a regra viraria "qualquer coisa
''' estranha derruba tudo" — ver
''' <see cref="Rotulo_fora_do_enum_invalida_SO_O_ITEM"/>.
''' </summary>
<TestClass>
Public Class LoteDeClassificacaoTests

    Private Shared Function Chave(id As String,
                                  Optional store As String = "store-1") As ItemKey
        Return New ItemKey(id, store)
    End Function

    ''' <summary>
    ''' <b>Chama-se <c>Montar</c>, e não <c>Lote</c>.</b> Com o nome <c>Lote</c>,
    ''' o local <c>lote</c> de cada teste eclipsava a função — VB é
    ''' insensível a maiúsculas — e o compilador reclamava de inferência de
    ''' tipo, longe da causa. É a armadilha número um do CLAUDE.md.
    ''' </summary>
    Private Shared Function Montar(ParamArray chaves As ItemKey()) As LoteDeClassificacao
        Return LoteDeClassificacao.Preparar(chaves.ToList())
    End Function

    ' ==================================================================
    ' A FICHA

    ''' <summary>
    ''' <b>O EntryID não sai da máquina.</b> A ficha é cunhada aqui, vale só
    ''' neste lote, e é ela que vai no envelope — o identificador de uma mensagem
    ''' real fica deste lado, pela mesma regra que o mantém fora do log.
    ''' </summary>
    <TestMethod>
    Public Sub A_ficha_nao_carrega_o_EntryID()
        Dim lote = Montar(Chave("E-SECRETO-1"), Chave("E-SECRETO-2"))

        For Each ficha In {lote.FichaDe(Chave("E-SECRETO-1")),
                           lote.FichaDe(Chave("E-SECRETO-2"))}
            Assert.IsTrue(ficha.Length > 0, "a mensagem do lote nao tem ficha")
            Assert.IsFalse(ficha.Contains("SECRETO"),
                "o EntryID foi para o fio dentro da ficha")
        Next

        Assert.AreNotEqual(lote.FichaDe(Chave("E-SECRETO-1")),
                           lote.FichaDe(Chave("E-SECRETO-2")),
                           "duas mensagens receberam a mesma ficha")
    End Sub

    ''' <summary>
    ''' <b>O CONTROLE DA FICHA OPACA.</b>
    '''
    ''' Um e-mail hostil que conheça o <c>EntryID</c> do vizinho — porque o viu
    ''' num cabeçalho, num encaminhamento, onde for — não consegue usá-lo: o fio
    ''' não fala <c>EntryID</c>. Uma resposta que traga o identificador de verdade
    ''' é recusada como item que não foi enviado.
    ''' </summary>
    <TestMethod>
    Public Sub Resposta_com_o_ENTRYID_de_verdade_e_recusada()
        Dim lote = Montar(Chave("E-1"), Chave("E-2"))

        Dim r = lote.Conferir(
            "[{""item_key"":""E-1"",""label"":""fyi""}," &
            " {""item_key"":""E-2"",""label"":""fyi""}]")

        Assert.IsFalse(r.IdentidadesConferem,
            "a resposta nomeou as mensagens pelo EntryID e foi aceita: o fio " &
            "voltou a falar a lingua que o conteudo do e-mail conhece")
    End Sub

    ''' <summary>
    ''' <b>Duas caixas podem repetir o EntryID</b>, e a identidade é
    ''' <c>EntryID + StoreID</c>. Com a chave reduzida ao EntryID, o lote de dois
    ''' virava um e a resposta com um item só era dada por completa.
    ''' </summary>
    <TestMethod>
    Public Sub O_mesmo_EntryID_em_DUAS_caixas_sao_dois_itens()
        Dim lote = Montar(Chave("E-1", "store-A"), Chave("E-1", "store-B"))

        Assert.AreEqual(2, lote.Quantos, "as duas caixas viraram um item so")

        Dim so1 = lote.Conferir(
            $"[{{""item_key"":""{lote.FichaDe(Chave("E-1", "store-A"))}"",""label"":""fyi""}}]")
        Assert.IsFalse(so1.IdentidadesConferem,
            "um item so respondeu por duas mensagens diferentes")
    End Sub

    ''' <summary>
    ''' A mesma mensagem duas vezes é erro de quem monta o lote, e não
    ''' normalização: colapsá-las faria um lote de três com uma duplicata ser
    ''' respondido com dois itens e dado por completo.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_REPETIDA_nao_monta_lote()
        Assert.IsNull(Montar(Chave("E-1"), Chave("E-2"), Chave("E-1")))
    End Sub

    <TestMethod>
    Public Sub Lote_vazio_nulo_ou_grande_demais_nao_monta()
        Assert.IsNull(LoteDeClassificacao.Preparar(Nothing))
        Assert.IsNull(LoteDeClassificacao.Preparar(Array.Empty(Of ItemKey)()))
        Assert.IsNull(Montar(Chave("E-1"), Nothing))
        Assert.IsNull(Montar(Chave("")), "chave vazia nao identifica mensagem nenhuma")

        Dim demais = Enumerable.Range(1, LoteDeClassificacao.MaximoDeItens + 1).
                     Select(Function(i) Chave($"E-{i}")).ToList()
        Assert.IsNull(LoteDeClassificacao.Preparar(demais))
    End Sub

    ' ==================================================================
    ' O CAMINHO BOM

    <TestMethod>
    Public Sub Um_lote_bem_formado_volta_casado_por_ficha()
        Dim lote = Montar(Chave("E-1"), Chave("E-2"))
        Dim f1 = lote.FichaDe(Chave("E-1"))
        Dim f2 = lote.FichaDe(Chave("E-2"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{f1}"",""label"":""precisa_de_mim"",""confidence"":0.9}}," &
            $" {{""item_key"":""{f2}"",""label"":""newsletter"",""confidence"":0.5}}]")

        Assert.IsTrue(r.IdentidadesConferem, r.Motivo)
        Assert.AreEqual(Rotulo.PrecisaDeMim, r.Rotulos(Chave("E-1")))
        Assert.AreEqual(Rotulo.Newsletter, r.Rotulos(Chave("E-2")))
        Assert.AreEqual(0.9, r.Confiancas(Chave("E-1")), 0.0001)
    End Sub

    ''' <summary>
    ''' A ordem da resposta não importa — o casamento é por ficha. Se importasse,
    ''' um modelo que reordenasse a lista trocaria todos os rótulos de lugar sem
    ''' nada falhar.
    ''' </summary>
    <TestMethod>
    Public Sub A_ORDEM_da_resposta_nao_importa()
        Dim lote = Montar(Chave("E-1"), Chave("E-2"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{lote.FichaDe(Chave("E-2"))}"",""label"":""fyi""}}," &
            $" {{""item_key"":""{lote.FichaDe(Chave("E-1"))}"",""label"":""promocao""}}]")

        Assert.IsTrue(r.IdentidadesConferem, r.Motivo)
        Assert.AreEqual(Rotulo.Promocao, r.Rotulos(Chave("E-1")))
        Assert.AreEqual(Rotulo.Fyi, r.Rotulos(Chave("E-2")))
    End Sub

    ' ==================================================================
    ' A IDENTIDADE

    <TestMethod>
    Public Sub Ficha_que_nao_e_do_lote_INVALIDA_tudo()
        Dim lote = Montar(Chave("E-1"), Chave("E-2"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{lote.FichaDe(Chave("E-1"))}"",""label"":""fyi""}}," &
            " {""item_key"":""i99"",""label"":""fyi""}]")

        Assert.IsFalse(r.IdentidadesConferem)
        Assert.AreEqual(0, r.Rotulos.Count, "nao pode aproveitar a parte que casou")
    End Sub

    <TestMethod>
    Public Sub Ficha_repetida_INVALIDA_tudo()
        Dim lote = Montar(Chave("E-1"), Chave("E-2"))
        Dim f1 = lote.FichaDe(Chave("E-1"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{f1}"",""label"":""fyi""}}," &
            $" {{""item_key"":""{f1}"",""label"":""promocao""}}]")

        Assert.IsFalse(r.IdentidadesConferem)
        StringAssert.Contains(r.Motivo, "duas vezes")
    End Sub

    ''' <summary>
    ''' Item enviado que não voltou. Silêncio não é "sem rótulo": é uma resposta
    ''' que não corresponde ao pedido, e aceitar o pedaço gravaria uma
    ''' classificação parcial que ninguém sabe que é parcial.
    ''' </summary>
    <TestMethod>
    Public Sub Item_que_nao_voltou_INVALIDA_tudo()
        Dim lote = Montar(Chave("E-1"), Chave("E-2"), Chave("E-3"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{lote.FichaDe(Chave("E-1"))}"",""label"":""fyi""}}]")

        Assert.IsFalse(r.IdentidadesConferem)
        StringAssert.Contains(r.Motivo, "3")
    End Sub

    ' ==================================================================
    ' A SUPERFICIE

    ''' <summary>
    ''' <b>O contraponto da regra dura.</b> Rótulo que não existe é a única
    ''' inconsistência que não sugere troca de identidade: o modelo escreveu uma
    ''' palavra inventada, e a mensagem fica <b>sem</b> rótulo em vez de com um
    ''' rótulo inventado. Sem isto, a regra viraria "qualquer coisa estranha
    ''' derruba tudo" e a varredura nunca terminaria.
    ''' </summary>
    <TestMethod>
    Public Sub Rotulo_fora_do_enum_invalida_SO_O_ITEM()
        Dim lote = Montar(Chave("E-1"), Chave("E-2"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{lote.FichaDe(Chave("E-1"))}"",""label"":""urgentissimo""}}," &
            $" {{""item_key"":""{lote.FichaDe(Chave("E-2"))}"",""label"":""fyi""}}]")

        Assert.IsTrue(r.IdentidadesConferem, r.Motivo)
        Assert.IsFalse(r.Rotulos.ContainsKey(Chave("E-1")), "aceitou um rotulo inventado")
        Assert.AreEqual(Rotulo.Fyi, r.Rotulos(Chave("E-2")), "o item bom foi junto")
        CollectionAssert.AreEqual({Chave("E-1")}, r.SemRotulo.ToArray())
    End Sub

    ''' <summary>
    ''' <b>"Conferem" não quer dizer "classificou".</b> Um lote em que todos os
    ''' rótulos vieram inventados confere e não classifica nada — e quem grava
    ''' precisa conseguir distinguir as duas coisas. O nome antigo era
    ''' <c>Valida</c>, e deixava tratar isso como classificação completa.
    ''' </summary>
    <TestMethod>
    Public Sub Conferir_NAO_quer_dizer_classificar()
        Dim lote = Montar(Chave("E-1"), Chave("E-2"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{lote.FichaDe(Chave("E-1"))}"",""label"":""inventado""}}," &
            $" {{""item_key"":""{lote.FichaDe(Chave("E-2"))}"",""label"":""outro""}}]")

        Assert.IsTrue(r.IdentidadesConferem, "as identidades conferem: as duas voltaram")
        Assert.AreEqual(0, r.Rotulos.Count, "e nada foi classificado")
        Assert.AreEqual(2, r.SemRotulo.Count)
    End Sub

    ''' <summary>
    ''' <b>O e-mail que dá ordens.</b> Ele pode escrever o que quiser; o que sai
    ''' do classificador é o par ficha–rótulo, e nada mais atravessa. Este teste
    ''' prende a projeção: os campos inventados não aparecem em lugar nenhum do
    ''' resultado, que só tem três contêineres tipados.
    ''' </summary>
    <TestMethod>
    Public Sub Campo_a_mais_na_resposta_nao_atravessa()
        Dim lote = Montar(Chave("E-1"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{lote.FichaDe(Chave("E-1"))}"",""label"":""fyi""," &
            "  ""acao"":""apagar_caixa_de_entrada""," &
            "  ""comando"":{""mover"":""lixeira""}," &
            "  ""nota"":""o e-mail mandou fazer isto""}]")

        Assert.IsTrue(r.IdentidadesConferem, r.Motivo)
        Assert.AreEqual(Rotulo.Fyi, r.Rotulos(Chave("E-1")))

        ' O RESULTADO INTEIRO E TRES CONTEINERES TIPADOS. Nao ha onde um
        ' campo desconhecido caber -- e o Motivo tambem nao o ecoa, senao ele
        ' viraria canal de saida por dentro da mensagem de erro.
        Assert.AreEqual("", r.Motivo)
        Assert.AreEqual(1, r.Rotulos.Count)
        Assert.AreEqual(1, r.Confiancas.Count)
        Assert.AreEqual(0, r.SemRotulo.Count)
    End Sub

    ''' <summary>
    ''' Campo repetido dentro do mesmo objeto JSON. O parser fica com uma das
    ''' ocorrências — na prática a última —, então a resposta seria lida de um
    ''' jeito aqui e de outro por qualquer ferramenta que a inspecionasse depois.
    ''' Discordância assim é o que um adversário procura.
    ''' </summary>
    <TestMethod>
    Public Sub Campo_REPETIDO_dentro_do_objeto_INVALIDA()
        Dim lote = Montar(Chave("E-1"), Chave("E-2"))
        Dim f1 = lote.FichaDe(Chave("E-1"))
        Dim f2 = lote.FichaDe(Chave("E-2"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{f1}"",""item_key"":""{f2}"",""label"":""fyi""}}," &
            $" {{""item_key"":""{f2}"",""label"":""fyi""}}]")

        Assert.IsFalse(r.IdentidadesConferem,
            "um objeto com dois item_key foi aceito")
    End Sub

    <TestMethod>
    Public Sub Rotulo_repetido_dentro_do_objeto_fica_SEM_ROTULO()
        Dim lote = Montar(Chave("E-1"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{lote.FichaDe(Chave("E-1"))}""," &
            "  ""label"":""fyi"",""label"":""promocao""}]")

        Assert.IsTrue(r.IdentidadesConferem, r.Motivo)
        Assert.AreEqual(0, r.Rotulos.Count, "escolheu um dos dois labels")
        Assert.AreEqual(1, r.SemRotulo.Count)
    End Sub

    <TestMethod>
    Public Sub Resposta_que_nao_e_lista_INVALIDA()
        Dim lote = Montar(Chave("E-1"))
        Dim f = lote.FichaDe(Chave("E-1"))

        For Each lixo In {"não sou JSON",
                          $"{{""item_key"":""{f}"",""label"":""fyi""}}",
                          "",
                          $"[""{f}""]"}
            Assert.IsFalse(lote.Conferir(lixo).IdentidadesConferem,
                $"aceitou uma resposta que nao e a lista: {lixo}")
        Next
    End Sub

    ''' <summary>
    ''' <b>Resposta gigante é recusada antes do parser.</b> Um vetor enorme ou
    ''' uma cadeia enorme não estouram o limite de profundidade, e o
    ''' <c>JsonDocument</c> materializaria tudo antes de a conferência ver o
    ''' primeiro item. A barreira de tipo não é barreira de disponibilidade.
    ''' </summary>
    <TestMethod>
    Public Sub Resposta_GIGANTE_e_recusada_sem_ser_lida()
        Dim lote = Montar(Chave("E-1"))
        Dim enorme = "[""" & New String("a"c, LoteDeClassificacao.MaximoDaResposta) & """]"

        Dim r = lote.Conferir(enorme)

        Assert.IsFalse(r.IdentidadesConferem)
        StringAssert.Contains(r.Motivo, "grande demais")
    End Sub

    ' ==================================================================
    ' A CONFIANCA

    ''' <summary>
    ''' Confiança ausente, fora da faixa ou ilegível vira <b>zero</b>, e não um
    ''' palpite: número que não veio não pode virar certeza.
    ''' </summary>
    <TestMethod>
    Public Sub Confianca_ausente_ou_absurda_vira_ZERO()
        Dim chaves = {Chave("E-1"), Chave("E-2"), Chave("E-3"), Chave("E-4")}
        Dim lote = LoteDeClassificacao.Preparar(chaves.ToList())

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{lote.FichaDe(chaves(0))}"",""label"":""fyi""}}," &
            $" {{""item_key"":""{lote.FichaDe(chaves(1))}"",""label"":""fyi"",""confidence"":7}}," &
            $" {{""item_key"":""{lote.FichaDe(chaves(2))}"",""label"":""fyi"",""confidence"":-1}}," &
            $" {{""item_key"":""{lote.FichaDe(chaves(3))}"",""label"":""fyi"",""confidence"":""muita""}}]")

        Assert.IsTrue(r.IdentidadesConferem, r.Motivo)
        ' 'aChave', e nao 'chave': o local eclipsaria a funcao Chave(), e o
        ' compilador reclamaria de argumento faltando -- longe da causa.
        For Each aChave In chaves
            Assert.AreEqual(0.0, r.Confiancas(aChave), 0.0001,
                "numero que nao veio virou certeza")
        Next
    End Sub

    <TestMethod>
    Public Sub A_caixa_das_letras_do_rotulo_nao_importa()
        Dim lote = Montar(Chave("E-1"), Chave("E-2"))

        Dim r = lote.Conferir(
            $"[{{""item_key"":""{lote.FichaDe(Chave("E-1"))}"",""label"":""FYI""}}," &
            $" {{""item_key"":""{lote.FichaDe(Chave("E-2"))}"",""label"":""Precisa_De_Mim""}}]")

        Assert.IsTrue(r.IdentidadesConferem, r.Motivo)
        Assert.AreEqual(Rotulo.Fyi, r.Rotulos(Chave("E-1")))
        Assert.AreEqual(Rotulo.PrecisaDeMim, r.Rotulos(Chave("E-2")))
    End Sub

    ''' <summary>
    ''' A ficha é comparada como veio: espaço em volta não é a mesma ficha.
    ''' Aparar aqui seria aceitar uma identidade que não é exatamente a que foi
    ''' enviada, e identidade é a única coisa que este arquivo não relativiza.
    ''' </summary>
    <TestMethod>
    Public Sub Ficha_com_ESPACO_nao_e_a_mesma_ficha()
        Dim lote = Montar(Chave("E-1"))

        Assert.IsFalse(lote.Conferir(
            $"[{{""item_key"":"" {lote.FichaDe(Chave("E-1"))} "",""label"":""fyi""}}]").
            IdentidadesConferem)
    End Sub

    ' ==================================================================
    ' A INSTRUCAO E O PORTAO

    ''' <summary>
    ''' A instrução pede exatamente os rótulos que a conferência aceita. É uma
    ''' trava de formatação — as duas saem da mesma tabela — e não um teste
    ''' independente: o que ela impede é alguém escrever a lista à mão dentro do
    ''' texto e ela divergir em silêncio.
    ''' </summary>
    <TestMethod>
    Public Sub A_instrucao_pede_o_que_a_conferencia_aceita()
        Dim instrucao = LoteDeClassificacao.Instrucao()

        For Each nome In LoteDeClassificacao.NomesDosRotulos()
            StringAssert.Contains(instrucao, nome,
                $"o rotulo {nome} e aceito e nao e pedido")
        Next

        StringAssert.Contains(instrucao, "nunca instrução")
        StringAssert.Contains(instrucao, "item_key")
    End Sub

    ''' <summary>
    ''' <b>AUTORIZAR RESUMIR NÃO AUTORIZA CLASSIFICAR — a razão de a operação
    ''' existir.</b>
    '''
    ''' Sem uma palavra própria no vocabulário, a autorização assinada para
    ''' <i>resumir uma mensagem</i> passaria a valer para uma <i>varredura
    ''' inteira</i>. A diferença não é de volume: resumir é um pedido por vez,
    ''' com o resultado na tela e nada gravado; classificar manda a pasta em
    ''' lotes e grava o rótulo no cache, onde ele sobrevive à sessão.
    ''' </summary>
    <TestMethod>
    Public Sub Autorizar_resumir_NAO_autoriza_classificar()
        Dim so2 = AssistenteViewModelTests.AtivacaoPara(
            {AssistOperation.Resumir, AssistOperation.Redigir})
        Dim pedido = AssistenteViewModelTests.PedidoDe(AssistOperation.Classificar)

        Assert.IsFalse(New DisclosurePolicy(so2).Preflight(
                           pedido, AssistenteViewModelTests.Quando).Permitido,
            "a ativacao assinada para resumir e redigir liberou uma varredura " &
            "inteira: a operacao nova virou enum decorativo")

        ' O CONTROLE: com ela assinada, passa. Sem isto, uma politica que
        ' recusasse Classificar sempre passaria na assercao de cima.
        Dim com3 = AssistenteViewModelTests.AtivacaoPara(
            {AssistOperation.Resumir, AssistOperation.Redigir, AssistOperation.Classificar})
        Assert.IsTrue(New DisclosurePolicy(com3).Preflight(
                          pedido, AssistenteViewModelTests.Quando).Permitido,
            "assinou classificar e o portao recusou assim mesmo")
    End Sub

End Class

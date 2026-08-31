Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A CAIXA DIVIDIDA — Fase 7.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTE ARQUIVO PRENDE</b>
'''
''' Uma coisa, e ela é a forma de tudo aqui: <b>uma mensagem aparece em uma
''' gaveta só</b>. Uma caixa dividida em que o mesmo e-mail aparece três vezes
''' não está dividida, está triplicada — e o dono passa a ter de lembrar se já
''' tratou aquilo em outra gaveta.
'''
''' ------------------------------------------------------------------
''' <b>O CONTROLE NEGATIVO</b>
'''
''' <see cref="A_gaveta_das_nao_classificadas_existe_MESMO_VAZIA"/> e
''' <see cref="Gaveta_de_rotulo_sem_mensagem_NAO_some"/>. Sem eles, uma
''' implementação que só criasse gaveta com conteúdo passaria em todo o resto
''' daqui — e seria exatamente a que faz a caixa parecer completa quando a
''' varredura classificou quarenta de novecentas.
''' </summary>
<TestClass>
Public Class CaixasSeparadasTests

    Private Shared Function Mensagem(id As String,
                                     Optional assunto As String = "assunto",
                                     Optional dia As Integer = 1) As MensagemNaFila
        Return New MensagemNaFila(
            New ItemKey(id, "store-1"), "conversa-" & id, assunto,
            "Alguém", "alguem@exemplo.com",
            New DateTimeOffset(2026, 8, dia, 12, 0, 0, TimeSpan.Zero))
    End Function

    Private Shared Function Rotulos(ParamArray pares As String()) _
                                    As IReadOnlyDictionary(Of ItemKey, String)
        Dim mapa As New Dictionary(Of ItemKey, String)()
        For i = 0 To pares.Length - 1 Step 2
            mapa(New ItemKey(pares(i), "store-1")) = pares(i + 1)
        Next
        Return mapa
    End Function

    Private Shared Function Casadas(id As String, ParamArray regras As String()) _
                                    As IReadOnlyDictionary(Of ItemKey, IReadOnlyList(Of String))
        Return New Dictionary(Of ItemKey, IReadOnlyList(Of String)) From {
            {New ItemKey(id, "store-1"), regras.ToList()}}
    End Function

    Private Shared Function Nenhuma() _
                            As IReadOnlyDictionary(Of ItemKey, IReadOnlyList(Of String))
        Return New Dictionary(Of ItemKey, IReadOnlyList(Of String))()
    End Function

    Private Shared Function Achar(gavetas As IReadOnlyList(Of Gaveta),
                                  nome As String) As Gaveta
        Return gavetas.Single(Function(g) g.Nome = nome)
    End Function

    ' ==================================================================
    ' A GAVETA ÚNICA

    ''' <summary>
    ''' <b>A decisão central.</b> A mensagem satisfaz uma regra do dono e tem
    ''' rótulo; ela vai para a regra, e <b>não</b> aparece na gaveta do rótulo.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_com_regra_E_rotulo_aparece_SO_na_regra()
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("a")},
            Rotulos("a", "precisa_de_mim"),
            Casadas("a", "clientes reclamando"),
            {"clientes reclamando"})

        Assert.AreEqual(1, Achar(gavetas, "clientes reclamando").Quantas)
        Assert.AreEqual(0, Achar(gavetas, "Esperam você").Quantas)

        ' E o total das gavetas e o total das mensagens: nada duplicou.
        Assert.AreEqual(1, gavetas.Sum(Function(g) g.Quantas))
    End Sub

    ''' <summary>
    ''' Duas regras satisfeitas: vale a <b>primeira do arquivo dele</b>. A ordem
    ''' em que o modelo devolveu as regras casadas não quer dizer nada; a ordem
    ''' do arquivo ele escreveu.
    ''' </summary>
    <TestMethod>
    Public Sub Duas_regras_satisfeitas_vale_a_PRIMEIRA_do_arquivo()
        ' O modelo devolveu as regras casadas na ordem INVERSA, de proposito.
        ' (Comentario aqui, e nao dentro da lista de argumentos: em VB a
        ' continuacao implicita nao aceita uma linha so de comentario, e o erro
        ' sai na linha seguinte. Armadilha do CLAUDE.md.)
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("a")},
            Rotulos(),
            Casadas("a", "sobre boleto", "clientes reclamando"),
            {"clientes reclamando", "sobre boleto"})

        Assert.AreEqual(1, Achar(gavetas, "clientes reclamando").Quantas)
        Assert.AreEqual(0, Achar(gavetas, "sobre boleto").Quantas)
    End Sub

    <TestMethod>
    Public Sub Mensagem_sem_regra_vai_para_a_gaveta_do_rotulo()
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("a")}, Rotulos("a", "newsletter"), Nenhuma(), {"uma regra"})

        Assert.AreEqual(1, Achar(gavetas, "Newsletters").Quantas)
        Assert.AreEqual(0, Achar(gavetas, "uma regra").Quantas)
    End Sub

    <TestMethod>
    Public Sub Mensagem_sem_rotulo_nenhum_vai_para_as_nao_classificadas()
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("a")}, Rotulos(), Nenhuma(), Array.Empty(Of String)())

        Assert.AreEqual(1, Achar(gavetas, CaixasSeparadas.NomeDasNaoClassificadas).Quantas)
    End Sub

    ''' <summary>
    ''' <b>A regra que o dono apagou não ressuscita.</b> A classificação ficou
    ''' gravada com o texto da regra de ontem; se ela virasse gaveta, o programa
    ''' estaria insistindo numa pergunta que ele parou de fazer.
    ''' </summary>
    <TestMethod>
    Public Sub Regra_casada_que_saiu_do_arquivo_e_IGNORADA()
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("a")},
            Rotulos("a", "fyi"),
            Casadas("a", "uma regra que ele apagou"),
            {"a regra de hoje"})

        Assert.IsFalse(gavetas.Any(Function(g) g.Nome = "uma regra que ele apagou"))
        Assert.AreEqual(1, Achar(gavetas, "Só para saber").Quantas)
    End Sub

    ''' <summary>
    ''' Rótulo que este programa não conhece cai nas não classificadas, e não
    ''' numa gaveta própria: ele só pode ter vindo de um banco gravado por outra
    ''' versão, e inventar uma gaveta mostraria ao dono uma categoria que
    ''' ninguém aqui sabe explicar.
    ''' </summary>
    <TestMethod>
    Public Sub Rotulo_desconhecido_nao_inventa_gaveta()
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("a")}, Rotulos("a", "urgentissimo"), Nenhuma(),
            Array.Empty(Of String)())

        Assert.IsFalse(gavetas.Any(Function(g) g.Nome = "urgentissimo"))
        Assert.AreEqual(1, Achar(gavetas, CaixasSeparadas.NomeDasNaoClassificadas).Quantas)
    End Sub

    ' ==================================================================
    ' AS GAVETAS QUE NÃO SOMEM

    ''' <summary>
    ''' <b>O controle negativo.</b> Se a varredura classificou quarenta de
    ''' novecentas, as oitocentas e sessenta restantes não são "sem
    ''' importância": são desconhecidas. Esconder a gaveta faria a caixa
    ''' dividida parecer completa.
    ''' </summary>
    <TestMethod>
    Public Sub A_gaveta_das_nao_classificadas_existe_MESMO_VAZIA()
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("a")}, Rotulos("a", "fyi"), Nenhuma(), Array.Empty(Of String)())

        Dim resto = Achar(gavetas, CaixasSeparadas.NomeDasNaoClassificadas)
        Assert.IsTrue(resto.Vazia)
    End Sub

    ''' <summary>
    ''' <b>O outro controle negativo.</b> "Nenhuma mensagem espera você" e "esta
    ''' divisão não existe" são coisas diferentes, e uma tela que só mostra
    ''' gaveta com conteúdo muda de forma a cada varredura — ninguém acha nada
    ''' duas vezes.
    ''' </summary>
    <TestMethod>
    Public Sub Gaveta_de_rotulo_sem_mensagem_NAO_some()
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("a")}, Rotulos("a", "fyi"), Nenhuma(), Array.Empty(Of String)())

        Assert.IsTrue(Achar(gavetas, "Esperam você").Vazia)
        Assert.IsTrue(Achar(gavetas, "Promoções").Vazia)
    End Sub

    <TestMethod>
    Public Sub Caixa_sem_mensagem_nenhuma_ainda_tem_todas_as_gavetas()
        Dim gavetas = CaixasSeparadas.Dividir(
            Array.Empty(Of MensagemNaFila)(), Nothing, Nothing, {"uma regra"})

        ' Seis rotulos + a regra do dono + as nao classificadas.
        Assert.AreEqual(8, gavetas.Count)
        Assert.IsTrue(gavetas.All(Function(g) g.Vazia))
    End Sub

    ' ==================================================================
    ' A ORDEM

    ''' <summary>
    ''' As gavetas do dono vêm primeiro, na ordem do arquivo dele; depois os
    ''' rótulos, na ordem de quem cobra; as não classificadas por último.
    '''
    ''' Ordenar por tamanho faria a caixa mudar de forma a cada varredura.
    ''' </summary>
    <TestMethod>
    Public Sub A_ordem_das_gavetas_e_FIXA()
        Dim nomes = CaixasSeparadas.Dividir(
            Array.Empty(Of MensagemNaFila)(), Nothing, Nothing,
            {"primeira regra", "segunda regra"}).
            Select(Function(g) g.Nome).ToList()

        CollectionAssert.AreEqual(
            {"primeira regra", "segunda regra",
             "Esperam você", "Você já respondeu", "Só para saber",
             "Avisos automáticos", "Promoções", "Newsletters",
             CaixasSeparadas.NomeDasNaoClassificadas}, nomes.ToArray())
    End Sub

    <TestMethod>
    Public Sub Dentro_da_gaveta_a_mais_recente_vem_primeiro()
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("velha", dia:=1), Mensagem("nova", dia:=20)},
            Rotulos("velha", "fyi", "nova", "fyi"), Nenhuma(),
            Array.Empty(Of String)())

        Dim gaveta = Achar(gavetas, "Só para saber")
        Assert.AreEqual("nova", gaveta.Mensagens.First().Chave.EntryId)
    End Sub

    <TestMethod>
    Public Sub Regra_repetida_no_arquivo_nao_vira_duas_gavetas()
        Dim gavetas = CaixasSeparadas.Dividir(
            Array.Empty(Of MensagemNaFila)(), Nothing, Nothing,
            {"a mesma", "a mesma"})

        ' .Where().Count() e nao .Count(predicado): a PROPRIEDADE Count da
        ' colecao eclipsa a extensao Count(Of T) do LINQ, e o compilador
        ' reclama de indexar um Integer -- longe da causa, como sempre.
        Assert.AreEqual(1, gavetas.Where(Function(g) g.Nome = "a mesma").Count())
    End Sub

    <TestMethod>
    Public Sub Regra_em_branco_nao_vira_gaveta()
        Dim gavetas = CaixasSeparadas.Dividir(
            Array.Empty(Of MensagemNaFila)(), Nothing, Nothing, {"   ", ""})

        Assert.AreEqual(7, gavetas.Count)
    End Sub

    ''' <summary>
    ''' A gaveta do dono se identifica como tal. A tela precisa: uma regra que
    ''' ele escreveu e uma categoria que o programa inventou não merecem o
    ''' mesmo tratamento — a primeira ele pode corrigir.
    ''' </summary>
    <TestMethod>
    Public Sub A_gaveta_do_dono_diz_que_e_dele()
        Dim gavetas = CaixasSeparadas.Dividir(
            Array.Empty(Of MensagemNaFila)(), Nothing, Nothing, {"uma regra"})

        Assert.IsTrue(Achar(gavetas, "uma regra").DoDono)
        Assert.AreEqual("", Achar(gavetas, "uma regra").Rotulo)
        Assert.IsFalse(Achar(gavetas, "Promoções").DoDono)
        Assert.AreEqual("promocao", Achar(gavetas, "Promoções").Rotulo)
    End Sub


    ' ==================================================================
    ' NOME NÃO É IDENTIDADE

    ''' <summary>
    ''' <b>Uma regra chamada como um rótulo não duplica mensagem.</b>
    '''
    ''' A identidade da gaveta era o texto do dono, e aí uma regra chamada
    ''' <i>fyi</i> colidia com a chave reservada: as duas gavetas passavam a
    ''' compartilhar a mesma lista e a mensagem aparecia nas duas — quebrando
    ''' justamente o contrato central deste arquivo.
    '''
    ''' Duas gavetas mostrando o mesmo nome é confuso e verdadeiro; a mesma
    ''' mensagem nas duas é falso.
    ''' </summary>
    <TestMethod>
    Public Sub Regra_com_nome_de_rotulo_nao_duplica_a_mensagem()
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("a")}, Rotulos("a", "fyi"), Casadas("a", "fyi"), {"fyi"})

        Assert.AreEqual(1, gavetas.Sum(Function(g) g.Quantas))
        Assert.AreEqual(8, gavetas.Count)
    End Sub

    ''' <summary>
    ''' O mesmo com o nome da gaveta residual, que é a que mais dói: ela é a
    ''' única que diz a verdade sobre a cobertura, e uma cópia dela marcada como
    ''' "do dono" faria a conta das não classificadas aparecer duas vezes.
    ''' </summary>
    <TestMethod>
    Public Sub Regra_com_o_nome_da_gaveta_residual_nao_duplica_nada()
        Dim gavetas = CaixasSeparadas.Dividir(
            {Mensagem("a")}, Rotulos(), Nenhuma(),
            {CaixasSeparadas.NomeDasNaoClassificadas})

        Assert.AreEqual(1, gavetas.Sum(Function(g) g.Quantas))

        ' A do dono existe e esta vazia; a de verdade tem a mensagem.
        Dim doDono = gavetas.Single(Function(g) g.DoDono)
        Assert.IsTrue(doDono.Vazia)
        Assert.AreEqual(1, gavetas.Last().Quantas)
    End Sub
End Class

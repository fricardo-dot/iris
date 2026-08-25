Imports System.Linq
Imports Iris.Assist
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A fronteira entre o que o Outlook entrega e o que pode virar bytes.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ISTO PRECISOU EXISTIR</b>
'''
''' O <c>MessagePart</c> <i>afirmava</i> que o corpo já era texto seguro, e
''' nada cobrava. Qualquer chamador podia passar HTML, um <c>cid:</c> ou um
''' data URI, e o envelope carregaria tudo para dentro do provedor — que
''' então buscaria a referência, por um caminho que o portão nunca viu.
'''
''' Escaping de JSON não faz nada disso: ele impede quebrar a estrutura do
''' documento, e não impede o conteúdo ser o que não devia.
''' </summary>
<TestClass>
Public Class ContentPipelineTests

    Private Shared Function Chave() As ItemKey
        Return New ItemKey("E-1", "store-1")
    End Function

    Private Shared Function Preparar(corpo As String,
                                     Optional html As Boolean = False,
                                     Optional completo As Boolean = True,
                                     Optional assunto As String = "assunto",
                                     Optional de As String = "fulano@exemplo.invalido",
                                     Optional para As String() = Nothing) As ContentResult
        Return ContentPipeline.Preparar(Chave(), "CK-1", assunto, de,
                                        If(para, {"beltrano@exemplo.invalido"}),
                                        corpo, html, completo)
    End Function

    ' ==================================================================
    ' Controle positivo

    <TestMethod>
    Public Sub Texto_simples_passa_intacto()
        Dim r = Preparar("olá, tudo bem?")

        Assert.IsTrue(r.Ok, $"recusou por {r.Recusa}")
        Assert.AreEqual("olá, tudo bem?", r.Parte.Corpo)
    End Sub

    ' ==================================================================
    ' HTML

    ''' <summary>
    ''' Tag some, estrutura vira quebra de linha, entidade é decodificada.
    ''' </summary>
    <TestMethod>
    Public Sub HTML_vira_texto()
        Dim r = Preparar("<p>primeiro</p><p>segundo &amp; terceiro</p>", html:=True)

        Assert.IsTrue(r.Ok, $"recusou por {r.Recusa}")
        StringAssert.Contains(r.Parte.Corpo, "primeiro")
        StringAssert.Contains(r.Parte.Corpo, "segundo & terceiro")
        Assert.IsFalse(r.Parte.Corpo.Contains("<p>"), "tag nao pode sobrar")
    End Sub

    ''' <summary>
    ''' <b>Comentário e <c>script</c> saem ANTES das tags.</b>
    '''
    ''' Se saíssem depois, o conteúdo deles viraria texto visível — e é
    ''' justamente ali que mora texto que o usuário nunca viu na tela. Um
    ''' resumo que "leu" um comentário HTML resumiria algo que ninguém
    ''' escreveu para ser lido.
    ''' </summary>
    <TestMethod>
    Public Sub Comentario_e_script_NAO_viram_texto()
        Dim r = Preparar("<p>visivel</p><!-- ESCONDIDO-1 -->" &
                         "<script>var x = 'ESCONDIDO-2';</script>" &
                         "<style>.a{content:'ESCONDIDO-3'}</style>", html:=True)

        Assert.IsTrue(r.Ok, $"recusou por {r.Recusa}")
        StringAssert.Contains(r.Parte.Corpo, "visivel")
        For Each escondido In {"ESCONDIDO-1", "ESCONDIDO-2", "ESCONDIDO-3"}
            Assert.IsFalse(r.Parte.Corpo.Contains(escondido),
                           $"{escondido} virou texto: " & r.Parte.Corpo)
        Next
    End Sub

    ''' <summary>Atributo não vira texto — inclusive o que carrega instrução.</summary>
    <TestMethod>
    Public Sub Atributo_nao_vira_texto()
        Dim r = Preparar("<div title=""ESCONDIDO"" data-x=""TAMBEM"">visivel</div>", html:=True)

        Assert.IsTrue(r.Ok)
        Assert.IsFalse(r.Parte.Corpo.Contains("ESCONDIDO"))
        Assert.IsFalse(r.Parte.Corpo.Contains("TAMBEM"))
    End Sub

    ' ==================================================================
    ' Referência embutida

    ''' <summary>
    ''' <b><c>cid:</c> e data URI RECUSAM.</b>
    '''
    ''' Não é higiene: uma referência dessas dentro do envelope faz o provedor
    ''' buscá-la, e aí o conteúdo sai por um caminho que o portão nunca viu.
    ''' </summary>
    <DataTestMethod>
    <DataRow("veja a imagem cid:abc123@empresa")>
    <DataRow("data:image/png;base64,AAAA")>
    <DataRow("<img src=""cid:logo"">texto</img>")>
    Public Sub Referencia_embutida_RECUSA(corpo As String)
        Dim r = Preparar(corpo, html:=corpo.StartsWith("<"))

        Assert.IsFalse(r.Ok, "passou: " & corpo)
        Assert.AreEqual(ContentRefusal.ReferenciaEmbutida, r.Recusa)
    End Sub

    ''' <summary>
    ''' E em <b>qualquer</b> campo, não só no corpo. Um <c>cid:</c> no assunto
    ''' é tão capaz de virar busca remota quanto um no corpo.
    ''' </summary>
    <TestMethod>
    Public Sub Referencia_embutida_no_ASSUNTO_tambem_recusa()
        Assert.AreEqual(ContentRefusal.ReferenciaEmbutida,
                        Preparar("corpo normal", assunto:="cid:coisa").Recusa)
        Assert.AreEqual(ContentRefusal.ReferenciaEmbutida,
                        Preparar("corpo normal", de:="data:text/x").Recusa)
        Assert.AreEqual(ContentRefusal.ReferenciaEmbutida,
                        Preparar("corpo normal", para:={"cid:x@y"}).Recusa)
    End Sub

    ''' <summary>
    ''' O contraponto: URL comum <b>não</b> recusa.
    '''
    ''' Sem ele, um pipeline que recusasse tudo passaria em todos os testes de
    ''' recusa acima — e nenhuma mensagem real, que quase sempre tem um link,
    ''' chegaria a ser resumida.
    ''' </summary>
    <TestMethod>
    Public Sub URL_comum_NAO_recusa()
        Dim r = Preparar("veja em https://exemplo.invalido/pagina")

        Assert.IsTrue(r.Ok, $"recusou por {r.Recusa}")
        StringAssert.Contains(r.Parte.Corpo, "https://exemplo.invalido/pagina")
    End Sub

    ' ==================================================================
    ' Corpo incompleto

    ''' <summary>
    ''' <b>Corpo pela metade RECUSA.</b>
    '''
    ''' Um resumo feito sobre meio corpo e apresentado como resumo é pior que
    ''' nenhum resumo. É a regra da §29.1 — "membro não comprovadamente
    ''' permitido nega" — aplicada ao conteúdo.
    ''' </summary>
    <TestMethod>
    Public Sub Corpo_INCOMPLETO_recusa()
        Dim r = Preparar("comeco do corpo e", completo:=False)

        Assert.IsFalse(r.Ok)
        Assert.AreEqual(ContentRefusal.CorpoIncompleto, r.Recusa)
    End Sub

    ' ==================================================================
    ' Controle e Unicode

    ''' <summary>
    ''' Caractere de controle e marcador de direção saem.
    '''
    ''' Um marcador de direção no meio de uma frase muda o que se <b>lê</b> sem
    ''' mudar o que está escrito — e um resumo feito sobre a leitura invertida
    ''' descreveria outra coisa.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_e_marcador_de_direcao_saem()
        Dim r = Preparar("antes" & ChrW(&H202E) & "depois" & ChrW(0) & "fim")

        Assert.IsTrue(r.Ok, $"recusou por {r.Recusa}")
        Assert.AreEqual("antesdepoisfim", r.Parte.Corpo)
    End Sub

    ''' <summary>
    ''' <b>Unicode de verdade é preservado.</b>
    '''
    ''' Emoji, acento e combinante ficam como o usuário escreveu. Normalizar
    ''' mudaria o texto dele, e o que se ganharia — estabilidade de bytes — o
    ''' envelope já garante de outro jeito.
    ''' </summary>
    <TestMethod>
    Public Sub Emoji_acento_e_combinante_ficam()
        Const original = "ação 🙂 a" & ChrW(&H301) & " ok"
        Dim r = Preparar(original)

        Assert.IsTrue(r.Ok)
        Assert.AreEqual(original, r.Parte.Corpo)
    End Sub

    ''' <summary>Quebra de linha vira <c>LF</c>, e sequência longa é encurtada.</summary>
    <TestMethod>
    Public Sub Quebra_de_linha_e_normalizada()
        Dim r = Preparar("a" & vbCrLf & "b" & vbCr & "c" & vbLf & vbLf & vbLf & vbLf & "d")

        Assert.IsTrue(r.Ok)
        Assert.IsFalse(r.Parte.Corpo.Contains(vbCr), "CR nao pode sobrar")
        Assert.AreEqual("a" & vbLf & "b" & vbLf & "c" & vbLf & vbLf & "d", r.Parte.Corpo)
    End Sub

    ' ==================================================================
    ' Tetos

    <TestMethod>
    Public Sub Assunto_longo_demais_RECUSA()
        Assert.AreEqual(ContentRefusal.CampoLongoDemais,
                        Preparar("corpo", assunto:=New String("a"c,
                                 ContentPipeline.MaxAssunto + 1)).Recusa)
    End Sub

    <TestMethod>
    Public Sub Corpo_longo_demais_RECUSA()
        Assert.AreEqual(ContentRefusal.CampoLongoDemais,
                        Preparar(New String("a"c, ContentPipeline.MaxCorpo + 1)).Recusa)
    End Sub

    <TestMethod>
    Public Sub Destinatarios_demais_RECUSA()
        Dim muitos = Enumerable.Range(1, ContentPipeline.MaxDestinatarios + 1).
                     Select(Function(i) $"p{i}@exemplo.invalido").ToArray()

        Assert.AreEqual(ContentRefusal.CampoLongoDemais, Preparar("corpo", para:=muitos).Recusa)
    End Sub

    ''' <summary>Depois de limpar, corpo vazio recusa — não vira envelope oco.</summary>
    <TestMethod>
    Public Sub Corpo_que_some_na_limpeza_RECUSA()
        Assert.AreEqual(ContentRefusal.SemTexto, Preparar("<style>.a{}</style>", html:=True).Recusa)
        Assert.AreEqual(ContentRefusal.SemTexto, Preparar("   " & vbTab & "  ").Recusa)
    End Sub


    ' ==================================================================
    ' Referência codificada

    ''' <summary>
    ''' <b>Referência escrita com entidade HTML também recusa.</b>
    '''
    ''' <c>&lt;img src="cid&amp;#58;logo"&gt;</c> não contém <c>cid:</c> no cru, e
    ''' a remoção da tag apagaria a evidência — mas o navegador do provedor lê a
    ''' entidade do mesmo jeito. Procurar só no cru era uma barreira que
    ''' qualquer remetente atravessava escrevendo dois pontos de outro jeito.
    ''' </summary>
    <DataTestMethod>
    <DataRow("<img src=""cid&#58;logo"">texto</img>")>
    <DataRow("<img src=""data&#58;image/png;base64,AAAA"">texto</img>")>
    <DataRow("&#99;&#105;&#100;&#58;abc")>
    Public Sub Referencia_com_ENTIDADE_tambem_recusa(corpo As String)
        Dim r = Preparar(corpo, html:=True)

        Assert.IsFalse(r.Ok, "passou: " & corpo)
        Assert.AreEqual(ContentRefusal.ReferenciaEmbutida, r.Recusa)
    End Sub

    ' ==================================================================
    ' A versão

    ''' <summary>
    ''' <b>Sem <c>PR_CHANGE_KEY</c>, o pipeline recusa.</b>
    '''
    ''' É ela que prende o corpo à leitura que o classificou. O 3.0 mediu a
    ''' propriedade vindo em 20 de 20 itens; faltar aqui é sinal de que alguém
    ''' montou por outro caminho.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_ChangeKey_o_pipeline_RECUSA()
        Dim r = ContentPipeline.Preparar(Chave(), "", "assunto", "de@x.invalido",
                                         {"para@x.invalido"}, "corpo", False, True)

        Assert.IsFalse(r.Ok)
        Assert.AreEqual(ContentRefusal.SemVersao, r.Recusa)
    End Sub

    ''' <summary>E ela sai na parte, para o envelope prender.</summary>
    <TestMethod>
    Public Sub A_ChangeKey_sai_na_parte()
        Dim r = ContentPipeline.Preparar(Chave(), "CK-7", "assunto", "de@x.invalido",
                                         {"para@x.invalido"}, "corpo", False, True)

        Assert.IsTrue(r.Ok)
        Assert.AreEqual("CK-7", r.Parte.ChangeKey)
    End Sub

End Class

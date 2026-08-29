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
        ' Pelo caminho PUBLICO — o snapshot. Sem isto, nenhum teste
        ' atravessaria a fronteira nova, e ela existiria sem prova.
        Return ContentPipeline.Preparar(
            New MessageSnapshot(Chave(), "CK-1", assunto, de,
                                If(para, {"beltrano@exemplo.invalido"}),
                                corpo, html, completo, temAnexo:=False))
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
        Dim r = ContentPipeline.Preparar(
            New MessageSnapshot(Chave(), "", "assunto", "de@x.invalido",
                                {"para@x.invalido"}, "corpo", False, True, temAnexo:=False))

        Assert.IsFalse(r.Ok)
        Assert.AreEqual(ContentRefusal.SemVersao, r.Recusa)
    End Sub

    ''' <summary>E ela sai na parte, para o envelope prender.</summary>
    <TestMethod>
    Public Sub A_ChangeKey_sai_na_parte()
        Dim r = ContentPipeline.Preparar(
            New MessageSnapshot(Chave(), "CK-7", "assunto", "de@x.invalido",
                                {"para@x.invalido"}, "corpo", False, True, temAnexo:=False))

        Assert.IsTrue(r.Ok)
        Assert.AreEqual("CK-7", r.Parte.ChangeKey)
    End Sub


    ' ==================================================================
    ' HTML mal formado

    ''' <summary>
    ''' <b>FECHAMENTO COM ESPAÇO É FECHAMENTO — e o segredo saía por essa fresta.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DUAS REGRAS QUE NÃO CONCORDAVAM</b>
    '''
    ''' O <c>HtmlInterpretavel</c> conta <c>"&lt;/script"</c> — <b>sem</b> o
    ''' <c>&gt;</c> —, então <c>&lt;/script &gt;</c> conta como fechamento e o
    ''' HTML passa como interpretável. Mas o padrão que <i>remove o bloco</i>
    ''' exigia <c>&lt;/script&gt;</c> exato, e não removia nada.
    '''
    ''' Resultado: o bloco não saía, a limpeza genérica de tags comia só as
    ''' tags, e <c>segredo()</c> ia para o provedor <b>como se fosse a
    ''' mensagem</b> — que é exatamente o dano que o teste "bloco sem fechar"
    ''' logo abaixo existe para impedir, chegando pelo outro lado.
    '''
    ''' Hoje o fluxo do Outlook monta o snapshot com <c>ehHtml = False</c>,
    ''' então isto era latente. <b>Latente é o estado em que todos os outros
    ''' desta família estavam</b> quando chegaram à tela.
    '''
    ''' <b>E as três últimas linhas vieram da revisão seguinte:</b> eu tinha
    ''' trocado <c>&lt;/&gt;</c> por <c>&lt;/\s*&gt;</c>, que fecha o
    ''' espaço e <b>não</b> fecha a família. A contagem aceita <i>qualquer
    ''' coisa</i> que comece com <c>&lt;/script</c>, e o parser HTML também
    ''' trata <c>&lt;/script x&gt;</c> e <c>&lt;/script/&gt;</c> como
    ''' fechamento. Consertar o caso que o revisor citou e deixar os irmãos é
    ''' o erro que este projeto já cometeu quatro vezes; agora os dois lados
    ''' usam o mesmo critério.
    '''
    ''' <b>Controle negativo:</b> devolvendo <c>&lt;/&gt;</c> — ou o
    ''' <c>&lt;/\s*&gt;</c> intermediário — ao padrão de bloco, este teste
    ''' cai.
    ''' </summary>
    <DataTestMethod>
    <DataRow("<p>visivel</p><script>SEGREDO</script >")>
    <DataRow("<p>visivel</p><script>SEGREDO</script  >")>
    <DataRow("<p>visivel</p><style>SEGREDO</style >")>
    <DataRow("<p>visivel</p><SCRIPT>SEGREDO</SCRIPT >")>
    <DataRow("<p>visivel</p><script>SEGREDO</script x>")>
    <DataRow("<p>visivel</p><script>SEGREDO</script/>")>
    <DataRow("<p>visivel</p><style>SEGREDO</style tipo=1>")>
    Public Sub Fechamento_com_ESPACO_remove_o_bloco(corpo As String)
        Dim r = Preparar(corpo, html:=True)

        Assert.IsTrue(r.Ok, "recusou HTML que e interpretavel: " & corpo)
        Assert.IsFalse(r.Parte.Corpo.Contains("SEGREDO"),
            "o conteudo do bloco saiu como texto para o provedor: " & r.Parte.Corpo)
        StringAssert.Contains(r.Parte.Corpo, "visivel",
            "controle: o texto legitimo tinha de sobreviver")
    End Sub

    ''' <summary>
    ''' <b>Bloco sem fechar RECUSA — porque a alternativa é ele virar texto.</b>
    '''
    ''' A remoção de bloco é por expressão regular, e expressão regular não é
    ''' parser: <c>&lt;script&gt;segredo</c>, sem fechar, não casa com o padrão
    ''' de bloco. O <c>&lt;script&gt;</c> então some junto com as outras tags e
    ''' <c>segredo</c> vira texto legítimo — que é justamente o texto que o
    ''' usuário <b>não</b> viu na tela, indo para o provedor como se fosse a
    ''' mensagem.
    '''
    ''' Um parser de verdade resolveria melhor. Enquanto não há, o mínimo
    ''' honesto é recusar o que não dá para interpretar.
    ''' </summary>
    <DataTestMethod>
    <DataRow("<p>visivel</p><script>SEGREDO")>
    <DataRow("<p>visivel</p><style>SEGREDO")>
    <DataRow("<p>visivel</p><!-- SEGREDO")>
    <DataRow("<script>a</script><script>SEGREDO")>
    Public Sub Bloco_sem_fechar_RECUSA(corpo As String)
        Dim r = Preparar(corpo, html:=True)

        Assert.IsFalse(r.Ok, "passou, e o segredo teria virado texto: " & corpo)
        Assert.AreEqual(ContentRefusal.HtmlIlegivel, r.Recusa)
    End Sub

    ''' <summary>
    ''' E o controle: HTML bem formado com bloco <b>continua</b> passando.
    '''
    ''' Sem ele, um pipeline que recusasse todo HTML com <c>script</c> passaria
    ''' nos testes acima — e nenhuma mensagem real de newsletter chegaria a ser
    ''' resumida.
    ''' </summary>
    <TestMethod>
    Public Sub HTML_bem_formado_com_bloco_PASSA()
        Dim r = Preparar("<p>visivel</p><script>SEGREDO</script>" &
                         "<style>.a{}</style><!-- nota -->", html:=True)

        Assert.IsTrue(r.Ok, $"recusou por {r.Recusa}")
        StringAssert.Contains(r.Parte.Corpo, "visivel")
        Assert.IsFalse(r.Parte.Corpo.Contains("SEGREDO"))
    End Sub

    ' ==================================================================
    ' O snapshot

    ''' <summary>
    ''' <b>Os amigos do <c>Iris.Model</c> são exatamente estes três.</b>
    '''
    ''' <c>Iris.Outlook</c> é a borda que lê do provider — a única camada de
    ''' <b>produção</b> capaz de montar um <c>MessageSnapshot</c>. Os outros
    ''' dois são aparato de teste: <c>Iris.Tests</c> e o harness de crash, que
    ''' precisa produzir uma capability de verdade em vez de fabricar uma.
    '''
    ''' O teste existe para que acrescentar um quarto seja uma <b>decisão</b> e
    ''' não um descuido: cada amigo é mais uma camada capaz de montar o
    ''' snapshot, e o tipo existe justamente para limitar quem o faz.
    ''' </summary>
    <TestMethod>
    Public Sub Os_amigos_do_Model_sao_exatamente_estes()
        Dim amigos = GetType(MessageSnapshot).Assembly.
            GetCustomAttributes(GetType(Runtime.CompilerServices.InternalsVisibleToAttribute), False).
            Cast(Of Runtime.CompilerServices.InternalsVisibleToAttribute)().
            Select(Function(a) a.AssemblyName).OrderBy(Function(n) n).ToArray()

        CollectionAssert.AreEqual({"Iris.CrashHarness", "Iris.Outlook", "Iris.Tests"}, amigos,
            "mais um amigo e mais uma camada capaz de montar o snapshot")
    End Sub

    ''' <summary>
    ''' <b>O caminho público é o snapshot, e ele vem de uma leitura só.</b>
    '''
    ''' A sobrecarga com item, versão, assunto e corpo separados preservava o
    ''' par (item, versão) e não provava nada sobre ele: qualquer chamador
    ''' passava o item aprovado, a versão aprovada, e um corpo qualquer. Agora
    ''' ela é <c>Friend</c>, e o público recebe o objeto que a borda do provider
    ''' montou.
    ''' </summary>
    <TestMethod>
    Public Sub A_sobrecarga_publica_recebe_SNAPSHOT()
        Dim publicas = GetType(ContentPipeline).GetMethods(
            Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static).
            Where(Function(m) m.Name = "Preparar").ToList()

        Assert.AreEqual(1, publicas.Count, "so um caminho publico")
        Assert.AreEqual(GetType(MessageSnapshot), publicas(0).GetParameters()(0).ParameterType)
    End Sub


    ' ==================================================================
    ' Anexo: a barreira que fecha a corrida do portão

    ''' <summary>
    ''' <b>Snapshot com anexo não vira bytes.</b>
    '''
    ''' O portão já nega mensagem com anexo, e isso não basta: ele classifica
    ''' numa visita ao COM, e o corpo é lido em outra. Um anexo acrescentado
    ''' entre as duas passaria pela classificação e entraria no envelope.
    '''
    ''' Aqui a verificação está presa ao <b>mesmo snapshot</b> que vira bytes.
    ''' </summary>
    <TestMethod>
    Public Sub Snapshot_com_anexo_e_recusado()
        Dim r = ContentPipeline.Preparar(Snapshot(temAnexo:=True))

        Assert.IsFalse(r.Ok)
        Assert.AreEqual(ContentRefusal.Anexo, r.Recusa)
    End Sub

    ''' <summary>
    ''' <b>"Não deu para contar" também não vira bytes.</b>
    '''
    ''' <c>Nothing</c> é o que a leitura devolve quando a contagem falhou —
    ''' guarda do Object Model, classe inesperada, erro de COM. Deixar passar
    ''' seria transformar ignorância em prova de ausência, que é a forma exata
    ''' da falha aberta que o 3.0 já custou uma vez.
    ''' </summary>
    <TestMethod>
    Public Sub Snapshot_com_anexo_DESCONHECIDO_e_recusado()
        Dim r = ContentPipeline.Preparar(Snapshot(temAnexo:=Nothing))

        Assert.IsFalse(r.Ok)
        Assert.AreEqual(ContentRefusal.Anexo, r.Recusa)
    End Sub

    ''' <summary>
    ''' Controle negativo: sem anexo, o mesmo snapshot <b>passa</b>.
    '''
    ''' Sem ele, um pipeline que recusasse tudo passaria nos dois testes de
    ''' cima — e nenhuma mensagem chegaria à IA, pelo motivo errado.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_sem_anexo_o_mesmo_snapshot_PASSA()
        Dim r = ContentPipeline.Preparar(Snapshot(temAnexo:=False))

        Assert.IsTrue(r.Ok, "so o anexo podia ser a diferenca")
        Assert.AreEqual(ContentRefusal.Nenhuma, r.Recusa)
    End Sub

    ''' <summary>O mesmo snapshot, mudando só o anexo.</summary>
    Private Shared Function Snapshot(temAnexo As Boolean?) As MessageSnapshot
        Return New MessageSnapshot(New ItemKey("E-1", "store-1"), "CK-1",
                                   "assunto", "de@x.invalido", {"para@x.invalido"},
                                   "um corpo qualquer", False, True, temAnexo)
    End Function

End Class

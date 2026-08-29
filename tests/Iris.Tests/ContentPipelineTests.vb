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
    '''
    ''' ------------------------------------------------------------------
    ''' <b>AS TRÊS ÚLTIMAS LINHAS: FECHAMENTO SEM O <c>&gt;</c></b>
    '''
    ''' A contagem procurava a substring <c>"&lt;/script"</c>, então
    ''' <c>&lt;/script</c> <i>sem o terminador</i> contava como fechamento: o
    ''' balanço fechava, o HTML passava, e o padrão de bloco — que precisa do
    ''' <c>&gt;</c> — não removia nada. Sobrava <c>SEGREDO&lt;/script</c> no
    ''' texto que vai para o provedor.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>AS TRÊS ÚLTIMAS: FECHAMENTO FALSO, E O CASO QUE EU MESMO ABRI</b>
    '''
    ''' Enquanto o teste era <i>contar</i> aberturas e fechamentos, ele aceitava
    ''' fechamento falso vindo de comentário ou de atributo. E a minha correção
    ''' anterior — contar só fechamento terminado — <b>abriu um caso novo</b>:
    ''' <c>&lt;!-- &lt;/script&gt; --&gt;&lt;script&gt;SEGREDO&lt;/script</c>
    ''' passou a ser aceito, porque o fechamento de dentro do comentário
    ''' equilibrava a abertura real. A contagem antiga recusava esse.
    '''
    ''' <b>Três versões consertando contagem com contagem.</b> A pergunta certa
    ''' não é "está balanceado", é <i>"sobrou alguma coisa que eu não soube
    ''' remover"</i> — e é isso que o código faz agora: tira comentário, tira
    ''' bloco, e recusa se ainda restar <c>&lt;script</c> em qualquer forma.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>AS DUAS ÚLTIMAS: O DANO PELO OUTRO LADO</b>
    '''
    ''' A penúltima é o comentário pela mesma doença: <c>--&gt;</c> vem
    ''' <i>antes</i> de <c>&lt;!--</c>, a contagem fecha, o removedor não acha
    ''' par nenhum, e o texto que o navegador trata como comentário aberto sai
    ''' como mensagem. Eu tinha trocado a regra de <c>script</c> e deixado o
    ''' comentário na contagem.
    '''
    ''' A última é o dano <b>invertido</b>, e foi o revisor que apontou: com
    ''' abertura <i>e</i> fechamento falsos em atributos, o removedor come o
    ''' <c>VISIVEL</c> que está entre os dois. Não sobra nada, a verificação
    ''' aceita, e o que vai para o provedor perdeu texto que o usuário vê.
    ''' Consertar o vazamento e criar um sumiço é o mesmo erro de sinal
    ''' trocado.
    '''
    ''' <b>Controle negativo:</b> qualquer uma das versões de contagem deixa
    ''' passar pelo menos uma destas linhas; e sem o
    ''' <c>MarcadorForaDeLugar</c>, as duas de atributo passam.
    ''' </summary>
    <DataTestMethod>
    <DataRow("<p>visivel</p><script>SEGREDO")>
    <DataRow("<p>visivel</p><style>SEGREDO")>
    <DataRow("<p>visivel</p><!-- SEGREDO")>
    <DataRow("<script>a</script><script>SEGREDO")>
    <DataRow("<p>visivel</p><script>SEGREDO</script")>
    <DataRow("<p>visivel</p><script>SEGREDO</script x")>
    <DataRow("<p>visivel</p><style>SEGREDO</style")>
    <DataRow("<!-- </script> --><script>SEGREDO")>
    <DataRow("<!-- </script> --><script>SEGREDO</script")>
    <DataRow("<p title=""</script>"">visivel</p><script>SEGREDO")>
    <DataRow("<p>VISIVEL</p>--><!-- marcador >SEGREDO")>
    <DataRow("<p title=""<script>"">VISIVEL</p><p title=""</script>"">FIM</p>")>
    <DataRow("<p>ANTES</p><p title=<script>>VISIVEL</p><p title=</script>><p>DEPOIS</p>")>
    <DataRow("<p>a < b e c > d</p>")>
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

    ''' <summary>
    ''' <b>O AUTÔMATO DAS ASPAS, caso a caso.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ELE RECUSA, ENTÃO ERRAR PARA O LADO ERRADO CUSTA CARO</b>
    '''
    ''' <c>MarcadorForaDeLugar</c> é o que decide se um HTML é recusado
    ''' por conter <c>&lt;</c> dentro de um valor de atributo. Um falso
    ''' positivo recusa mensagem legítima; um falso negativo deixa o removedor
    ''' comer texto visível. Os dois lados doem, então cada estado tem caso.
    '''
    ''' Estes casos são os que a revisão externa ia pedir e eu escrevi antes:
    ''' aspas simples dentro de duplas, <c>&gt;</c> dentro de aspas, tag não
    ''' fechada no fim, <c>&lt;</c> solto no texto, e aspas fora de tag.
    ''' </summary>
    <TestMethod>
    Public Sub O_automato_das_aspas_caso_a_caso()
        ' ACHA: < dentro de valor de atributo, nas duas aspas.
        Assert.IsTrue(ContentPipeline.MarcadorForaDeLugar("<p title=""<script>"">x</p>"))
        Assert.IsTrue(ContentPipeline.MarcadorForaDeLugar("<p title='<script>'>x</p>"))

        ' ASPA SIMPLES DENTRO DE DUPLA nao fecha a dupla: o < depois dela
        ' continua sendo dentro do atributo.
        Assert.IsTrue(ContentPipeline.MarcadorForaDeLugar("<p t=""ele's <b"">x</p>"))

        ' ATRIBUTO SEM ASPAS -- o irmao direto, e ele passava.
        Assert.IsTrue(ContentPipeline.MarcadorForaDeLugar(
            "<p title=<script>>VISIVEL</p><p title=</script>>"))

        ' "<" SOLTO NO TEXTO -- este eu tinha marcado como "nao acha", e a
        ' revisao mostrou que a limpeza generica come "< b e c >" e entrega
        ' "a d". Texto visivel some, que e o dano invertido do vazamento.
        Assert.IsTrue(ContentPipeline.MarcadorForaDeLugar("<p>a < b e c > d</p>"))

        ' NAO ACHA -- e cada um destes seria um falso positivo caro:
        '
        ' aspas fora de tag sao texto comum.
        Assert.IsFalse(ContentPipeline.MarcadorForaDeLugar("ele disse ""oi"" e foi <p>x</p>"))
        ' > dentro de aspas nao termina a tag, mas tambem nao e marcador.
        Assert.IsFalse(ContentPipeline.MarcadorForaDeLugar("<p title=""a > b"">x</p>"))
        ' HTML comum, com atributo depois de atributo.
        Assert.IsFalse(ContentPipeline.MarcadorForaDeLugar(
            "<a href=""http://x.invalido/?a=1&b=2"" title='dois'>ok</a>"))

        ' TEXTO CRU: dentro de script/style nao ha marcacao, e este era o
        ' falso positivo que a revisao achou -- o "<" da string JS abria uma
        ' tag ficticia e o </script> derrubava HTML legitimo.
        Assert.IsFalse(ContentPipeline.MarcadorForaDeLugar(
            "<p>VISIVEL</p><script>const lt = ""<"";</script>"))
        Assert.IsFalse(ContentPipeline.MarcadorForaDeLugar(
            "<style>a[href^=""<""] { color: red }</style>"))

        ' Comentario condicional da Microsoft, onipresente em newsletter.
        Assert.IsFalse(ContentPipeline.MarcadorForaDeLugar(
            "<!--[if mso]><table><tr><td>x</td></tr></table><![endif]--><p>oi</p>"))

        ' Tag que nao fecha ate o fim do texto: quem recusa isso, se for o
        ' caso, e a regra da sobra -- aqui nao ha o que declarar.
        Assert.IsFalse(ContentPipeline.MarcadorForaDeLugar("<p title=""aberto"))
        ' Vazio e nulo nao explodem nem acusam.
        Assert.IsFalse(ContentPipeline.MarcadorForaDeLugar(""))
        Assert.IsFalse(ContentPipeline.MarcadorForaDeLugar(Nothing))
    End Sub

    ''' <summary>
    ''' <b>CONTROLE POSITIVO DA REGRA NOVA: HTML comum continua passando.</b>
    '''
    ''' Sem isto, um pipeline que recusasse todo HTML passaria em todos os
    ''' testes de recusa acima — que é o bloqueio sem controle negativo que o
    ''' CLAUDE.md descreve, e que eu já cometi quatro vezes nesta série.
    '''
    ''' Inclui o comentário, porque a regra do comentário também mudou nesta
    ''' passada e é a que tinha mais risco de recusar demais.
    ''' </summary>
    <DataTestMethod>
    <DataRow("<p>oi</p><a href=""http://x.invalido"" title='t'>link</a>")>
    <DataRow("<!-- comentario normal --><p>oi</p>")>
    <DataRow("<p>oi</p><script>var s = 'a > b';</script><p>tchau</p>")>
    <DataRow("<div><span>a</span> &gt; <span>b</span></div>")>
    <DataRow("<p title=""a > b"">com maior dentro do atributo</p>")>
    <DataRow("<!--[if mso]><table><tr><td>x</td></tr></table><![endif]--><p>oi</p>")>
    <DataRow("<p>oi</p><script>const lt = ""<"";</script><p>tchau</p>")>
    <DataRow("<p>Use a seta --> para continuar.</p>")>
    Public Sub Controle_HTML_comum_continua_passando(corpo As String)
        Dim r = Preparar(corpo, html:=True)
        Assert.IsTrue(r.Ok, $"recusou HTML comum por {r.Recusa}: " & corpo)
    End Sub

End Class

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports Iris.App.ViewModels
Imports Iris.Assist
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Model
Imports Iris.Sync
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A borda em lote, do cache ao provedor e de volta ao cache.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ESTE ARQUIVO EXISTE</b>
'''
''' Durante dez etapas, <c>ClassificarUmaPasta</c> foi testado contra delegates
''' de mentira: um que devolvia partes prontas e outro que devolvia JSON pronto.
''' Isso prova o <i>miolo</i> — os lotes, o controle, a conferência — e não prova
''' nada sobre o que estava faltando, que era justamente <b>ligar os dois fios</b>
''' ao Outlook e ao provedor.
'''
''' Aqui os delegates são os de produção. O caminho é: cache semeado →
''' <c>ClassificarUmaPasta</c> → <c>BordaEmLote</c> → <c>ContextoDoOutlook</c> →
''' <c>DisclosurePolicy</c> → <c>CapabilityLedger</c> → <c>AssistTransmitter</c> →
''' provedor, e a resposta volta pelo mesmo caminho até virar linha no cache.
'''
''' <b>O controle negativo é o teste mais importante do arquivo</b>: sem ativação
''' assinada, o provedor não é chamado nenhuma vez. Sem ele, todo o resto aqui
''' provaria apenas que <i>alguma coisa</i> acontece.
''' </summary>
' TOCA SQLITE: nao roda em paralelo com as outras. Foi assim que a falha
' rara de 25/08/2026 apareceu, e ha um meta-teste que cobra isto -- ele
' pegou esta classe no mesmo dia em que ela nasceu.
<TestClass>
<DoNotParallelize>
Public Class BordaEmLoteTests

    Private Shared ReadOnly Quando As New DateTimeOffset(2026, 9, 1, 12, 0, 0, TimeSpan.Zero)
    Private Const Endereco As String = "https://provedor.invalido/v1"
    Private Const EntradaDaPasta As String = "f-1"

    ''' <summary>A pasta como o portão a identifica. Ver o <c>Semear</c>.</summary>
    Private Shared ReadOnly Pasta As New FolderKey(EntradaDaPasta, "store-1")

    ' ==================================================================
    ' O CAMINHO INTEIRO

    ''' <summary>
    ''' <b>Três mensagens saem do cache, passam pelo provedor e voltam rotuladas.</b>
    '''
    ''' O que cada asserção cobre, e por que não bastava a anterior:
    '''
    ''' <list type="number">
    ''' <item>o provedor <b>foi chamado</b> — sem isto, um caminho que recusa
    ''' tudo em silêncio passaria;</item>
    ''' <item>os <b>corpos</b> chegaram — sem isto, um envelope vazio passaria;</item>
    ''' <item>os rótulos <b>estão no cache</b> — sem isto, a resposta teria sido
    ''' recebida e jogada fora.</item>
    ''' </list>
    ''' </summary>
    <TestMethod>
    Public Sub A_pasta_inteira_atravessa_a_borda_e_volta_rotulada()
        Comigo(Sub(db)
                   Dim pasta = Semear(db, {"a", "b", "c"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim r = Classificar(db, pasta, provedor, ComAtivacao())

                   Assert.AreEqual(MotivoDaClassificacao.Passou, r.Motivo)
                   Assert.AreEqual(3, r.Pedidos)
                   Assert.AreEqual(3, r.Classificados,
                       "a passagem inteira rodou e nenhuma mensagem foi rotulada")

                   Assert.AreEqual(1, provedor.Chamadas,
                       "três mensagens cabem num lote só")

                   ' O CONJUNTO INTEIRO, e nao "algum corpo contem 'a'".
                   '
                   ' A versao anterior sobreviveria a uma sabotagem que mandasse
                   ' "a" certo e esvaziasse ou duplicasse "b" e "c": o provedor
                   ' responde por ficha, entao a contagem de tres rotulos ainda
                   ' fecharia. Achado por revisao externa em 01/09/2026.
                   Dim corpos = provedor.PorFicha().Values.
                                Where(Function(c) c.StartsWith("corpo de ")).
                                OrderBy(Function(c) c, StringComparer.Ordinal).ToList()
                   CollectionAssert.AreEqual(
                       {"corpo de a", "corpo de b", "corpo de c"}, corpos.ToArray(),
                       "os tres corpos tinham de chegar, cada um uma vez")

                   Dim guardados = New RotulosNoCache(db).Publicados(pasta)
                   Assert.AreEqual(3, guardados.Count,
                       "a resposta voltou e o cache continuou vazio")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo, e ele vale por todos os outros.</b>
    '''
    ''' Sem ativação assinada o portão nega cada lote <i>antes</i> de qualquer
    ''' leitura de corpo. O provedor não é chamado nenhuma vez.
    '''
    ''' Sem este teste, uma borda que simplesmente nunca enviasse passaria em
    ''' todos os testes de "não envia errado" — é a regra do CLAUDE.md, e este é
    ''' o lugar dela.
    ''' </summary>
    <TestMethod>
    Public Sub SEM_ativacao_o_provedor_nao_e_chamado_nenhuma_vez()
        Comigo(Sub(db)
                   Dim pasta = Semear(db, {"a", "b", "c"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim r = Classificar(db, pasta, provedor, ativacao:=Nothing)

                   Assert.AreEqual(0, provedor.Chamadas,
                       "CONTEUDO SAIU DA MAQUINA SEM AUTORIZACAO")
                   Assert.AreEqual(0, r.Classificados)
                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(pasta).Count)

                   ' E a passagem PERCORREU: recusa de lote, e não desistência
                   ' antes de começar. As duas dão zero rótulos e são estados
                   ' diferentes -- uma diz "o portão negou", a outra diria que
                   ' nem havia o que classificar.
                   Assert.AreEqual(MotivoDaClassificacao.Passou, r.Motivo)
                   Assert.AreEqual(1, r.LotesRecusados)
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>A ficha de cada mensagem viaja com o corpo daquela mensagem.</b>
    '''
    ''' É a asserção que o alinhamento posicional da leitura em lote existe para
    ''' garantir. Trocar duas fichas não quebra nada visível: o modelo responde,
    ''' a conferência passa, os rótulos entram — <b>nas mensagens erradas</b>.
    '''
    ''' O corpo de cada mensagem carrega o próprio sufixo, então o par
    ''' ficha↔corpo é conferível do lado de fora.
    ''' </summary>
    <TestMethod>
    Public Sub Cada_ficha_viaja_com_o_corpo_da_PROPRIA_mensagem()
        Comigo(Sub(db)
                   Dim pasta = Semear(db, {"a", "b", "c"})
                   Dim provedor As New ProvedorQueClassifica()

                   ' A FONTE INDEPENDENTE: o mapa ficha -> chave que a
                   ' ClassificarUmaPasta montou, capturado ANTES de a borda ler
                   ' qualquer corpo. Sem ele, este teste comparava dois
                   ' derivados do mesmo envelope e passava com as fichas
                   ' invertidas -- provado invertendo-as de proposito.
                   Dim pedidos As New List(Of PedidoDeParte)()
                   Classificar(db, pasta, provedor, ComAtivacao(), pedidos)

                   Dim vistos = provedor.PorFicha()
                   ' Quatro partes: as três mensagens e o controle do lote, que
                   ' não corresponde a mensagem nenhuma.
                   Assert.AreEqual(4, vistos.Count)
                   Assert.AreEqual(3, pedidos.Count, "controle: três pedidos")
                   Assert.AreEqual(1, provedor.Chamadas, "controle: o lote saiu")

                   For Each p In pedidos
                       Dim esperado = "corpo de " & SufixoDe(p.Chave)
                       Dim chegou As String = Nothing
                       Assert.IsTrue(vistos.TryGetValue(p.Ficha, chegou),
                           $"a ficha de {p.Chave.EntryId} não chegou ao provedor")
                       Assert.AreEqual(esperado, chegou,
                           $"a ficha de {p.Chave.EntryId} viajou com o corpo de outra")
                   Next
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>Enviar sem ter lido recusa</b>, em vez de mandar o lote anterior.
    '''
    ''' A ordem — ler e depois mandar — é garantida por
    ''' <c>ClassificarUmaPasta.Passar</c>. Depender disso em silêncio seria
    ''' depender de alguém não mudar de ideia; mandar a instrução de um lote com
    ''' a seleção de outro é divulgação que ninguém pediu.
    ''' </summary>
    ''' <summary>
    ''' <b>A ficha que o lote sorteia atravessa o pipeline.</b>
    '''
    ''' Parece tautológico e não é: são dois módulos, um monta e o outro
    ''' confere, e eles divergiram no primeiro dia. O conferidor esperava oito
    ''' caracteres; a ficha de verdade tem nove — um prefixo e oito.
    '''
    ''' O efeito não era um erro: o pipeline recusava <b>todas</b> as mensagens,
    ''' o lote saía vazio, era pulado, e a passagem terminava com zero rótulos
    ''' <i>sem nada quebrar</i>. A classificação em lote inteira teria ficado
    ''' inútil em silêncio.
    '''
    ''' As duas pontas passaram a ler as mesmas constantes, e este teste é o que
    ''' garante que elas continuem lendo.
    ''' </summary>
    <TestMethod>
    Public Sub A_ficha_que_o_LOTE_sorteia_atravessa_o_PIPELINE()
        Dim lote = LoteDeClassificacao.Preparar({Chave("a")}, Array.Empty(Of String)())
        Assert.IsNotNull(lote, "nao montou o lote")
        Dim f = lote.FichaDe(Chave("a"))
        Assert.IsTrue(LoteDeClassificacao.EhFichaValida(f), $"ficha invalida: [{f}]")
        Dim r = ContentPipeline.Preparar(Instantaneo(Chave("a"), "corpo de a"), f)
        Assert.IsTrue(r.Ok, $"pipeline recusou: {r.Recusa}")
    End Sub

    <TestMethod>
    Public Sub Enviar_sem_ter_lido_RECUSA()
        Dim provedor As New ProvedorQueClassifica()
        Dim broker = BrokerBom({"a"})
        Dim borda As New BordaEmLote(broker, Transmissor(provedor, ComAtivacao()),
                                     provedor.Destino, Pasta)

        Dim saiu = borda.Envio("instrução", {ParteQualquer()}, CancellationToken.None)

        Assert.IsFalse(saiu.Incerta, "recusa nao e incerteza: nada saiu")
        Assert.AreEqual(0, saiu.Texto.Length, "mandou um lote que ninguem leu")
        Assert.AreEqual(0, provedor.Chamadas)
    End Sub

    ''' <summary>
    ''' <b>Lista encolhida não é "ler o que deu": é não ler nada.</b>
    '''
    ''' A leitura em lote promete uma posição por item pedido. Uma implementação
    ''' que devolvesse só os que deram certo faria a ficha da mensagem 5 viajar
    ''' com o corpo da 6 — e nada na tela mostraria isso.
    '''
    ''' O teste finge exatamente esse defeito, que é o único jeito de saber que a
    ''' conferência existe.
    ''' </summary>
    <TestMethod>
    Public Sub Leitura_em_lote_que_ENCOLHE_a_lista_nao_produz_parte_nenhuma()
        Dim broker = BrokerBom({"a", "b"})
        broker.Instantaneos = Function(k) OperationResult(Of MessageSnapshot).
            Ok(Instantaneo(k, "corpo"))
        ' O defeito: devolve UM para dois pedidos.
        broker.Lote = Function(chaves) OperationResult(Of IReadOnlyList(Of MessageSnapshot)).
            Ok(New MessageSnapshot() {Instantaneo(chaves(0), "corpo de a")})

        Dim itens As IReadOnlyList(Of ItemKey) = {Chave("a"), Chave("b")}
        Dim contexto As New ContextoDoOutlook(
            broker, New AssistDestination("p", Endereco, "m"),
            Function() (Pasta, itens))

        Assert.AreEqual(0, contexto.Partes().Count,
            "MONTOU PARTES A PARTIR DE UMA LISTA DESALINHADA")
    End Sub

    ''' <summary>
    ''' O controle negativo do teste acima: com a lista <b>alinhada</b>, as duas
    ''' partes saem. Sem ele, uma conferência que recusasse sempre passaria.
    ''' </summary>
    <TestMethod>
    Public Sub Com_a_lista_alinhada_as_partes_saem()
        Dim broker = BrokerBom({"a", "b"})
        Dim itens As IReadOnlyList(Of ItemKey) = {Chave("a"), Chave("b")}
        Dim contexto As New ContextoDoOutlook(
            broker, New AssistDestination("p", Endereco, "m"),
            Function() (Pasta, itens))

        Assert.AreEqual(2, contexto.Partes().Count)
    End Sub

    ''' <summary>
    ''' <b>Ficha com forma errada recusa a mensagem</b> — e não vira campo livre
    ''' no envelope.
    '''
    ''' A ficha é o único identificador que sai desta máquina, e é o único campo
    ''' que não vem do Outlook: vem de quem montou o lote. Todos os outros campos
    ''' o pipeline confere.
    ''' </summary>
    <TestMethod>
    Public Sub Ficha_com_forma_errada_NAO_atravessa_o_pipeline()
        Dim retrato = Instantaneo(Chave("a"), "corpo de a")

        Assert.IsTrue(ContentPipeline.Preparar(retrato, "iabcdefgh").Ok,
            "controle: uma ficha bem formada tem de passar")

        Dim torta = ContentPipeline.Preparar(retrato, "ricardo@empresa.com")
        Assert.IsFalse(torta.Ok, "UM ENDERECO PASSOU COMO FICHA")
        Assert.AreEqual(ContentRefusal.FichaInvalida, torta.Recusa)

        Assert.IsTrue(ContentPipeline.Preparar(retrato, Nothing).Ok,
            "fora de lote não há ficha, e isso é legítimo")
    End Sub

    ''' <summary>
    ''' <b>A mensagem grande demais é recusada sozinha, e o lote segue.</b>
    '''
    ''' O pipeline aceita 200 mil caracteres de corpo e o envelope inteiro cabe em
    ''' 256 KiB. Uma mensagem grande sozinha estoura o envelope de um lote de
    ''' vinte: ele sai truncado, o cofre recusa — corretamente — e nada é mandado.
    '''
    ''' E como os lotes se formam sempre na mesma ordem, ela volta ao mesmo lote em
    ''' toda passagem: <b>aquelas vinte nunca seriam classificadas</b>. Mesma
    ''' família do defeito do anexo, por outra rota.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_GRANDE_DEMAIS_nao_envenena_o_lote()
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a", "b", "c"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim broker = BrokerBom({"a", "b", "c"})

                   ' A "b" tem um corpo que sozinho nao cabe num lote de vinte.
                   Dim enorme = New String("x"c, BordaEmLote.TetoDoCorpoNoLote + 1)
                   broker.Instantaneos =
                       Function(k) OperationResult(Of MessageSnapshot).Ok(
                           New MessageSnapshot(k, "CK-" & k.EntryId, "assunto",
                                               "de@x.invalido", {"para@x.invalido"},
                                               If(SufixoDe(k) = "b", enorme,
                                                  "corpo de " & SufixoDe(k)), False,
                                               corpoCompleto:=True, temAnexo:=False,
                                               pasta:=Pasta))

                   Dim r = Classificar(db, noCache, provedor, ComAtivacao(),
                                       broker:=broker)

                   Assert.AreEqual(1, provedor.Chamadas,
                       "A MENSAGEM GRANDE DERRUBOU O LOTE INTEIRO")
                   Assert.AreEqual(2, r.Classificados,
                       "as outras duas tinham de ser classificadas")
                   Assert.AreEqual(1, r.RecusadasPeloConteudo,
                       "a grande tinha de aparecer na conta, e sozinha")
               End Sub)
    End Sub
    ''' <summary>
    ''' <b>O teto é de BYTES, e o teste tem de usar bytes.</b>
    '''
    ''' O par de testes acima usava só <c>"x"</c> — ASCII, em que caractere e byte
    ''' coincidem. Ele passava com a produção contando <c>texto.Length</c> contra um
    ''' orçamento de bytes UTF-8, que é exatamente o defeito que existia. Um teste
    ''' cuja fixture apaga a distinção que ele deveria medir não mede nada.
    '''
    ''' Aqui o corpo tem <b>menos caracteres</b> que o teto e <b>mais bytes</b>. Um
    ''' emoji pesa quatro. Achado por revisão externa em 02/09/2026.
    ''' </summary>
    <TestMethod>
    Public Sub Corpo_curto_em_CARACTERES_e_longo_em_BYTES_e_recusado()
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a", "b"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim broker = BrokerBom({"a", "b"})

                   ' Metade dos caracteres do teto, e o dobro dos bytes dele.
                   Dim quantos = BordaEmLote.TetoDoCorpoNoLote \ 2
                   Dim pesado = String.Concat(Enumerable.Repeat("😀", quantos))
                   Assert.IsTrue(pesado.Length < BordaEmLote.TetoDoCorpoNoLote,
                       "controle: em CARACTERES o corpo tem de caber")
                   Assert.IsTrue(Text.Encoding.UTF8.GetByteCount(pesado) >
                                 BordaEmLote.TetoDoCorpoNoLote,
                       "controle: em BYTES o corpo tem de estourar")

                   broker.Instantaneos =
                       Function(k) OperationResult(Of MessageSnapshot).Ok(
                           New MessageSnapshot(k, "CK-" & k.EntryId, "assunto",
                                               "de@x.invalido", {"para@x.invalido"},
                                               If(SufixoDe(k) = "b", pesado,
                                                  "corpo de " & SufixoDe(k)), False,
                                               corpoCompleto:=True, temAnexo:=False,
                                               pasta:=Pasta))

                   Dim r = Classificar(db, noCache, provedor, ComAtivacao(),
                                       broker:=broker)

                   Assert.AreEqual(1, r.RecusadasPeloConteudo,
                       "O TETO CONTOU CARACTERES: o corpo pesado passou")
                   Assert.AreEqual(1, r.Classificados,
                       "a outra tinha de seguir")
               End Sub)
    End Sub


    ''' <summary>
    ''' O controle negativo: um corpo <b>logo abaixo</b> do teto atravessa. Sem
    ''' ele, um teto de zero passaria no teste acima e recusaria tudo.
    ''' </summary>
    <TestMethod>
    Public Sub Corpo_logo_ABAIXO_do_teto_atravessa()
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim broker = BrokerBom({"a"})
                   Dim quaseLa = New String("x"c, BordaEmLote.TetoDoCorpoNoLote - 1)
                   broker.Instantaneos =
                       Function(k) OperationResult(Of MessageSnapshot).Ok(
                           New MessageSnapshot(k, "CK-" & k.EntryId, "assunto",
                                               "de@x.invalido", {"para@x.invalido"},
                                               quaseLa, False,
                                               corpoCompleto:=True, temAnexo:=False,
                                               pasta:=Pasta))

                   Dim r = Classificar(db, noCache, provedor, ComAtivacao(),
                                       broker:=broker)

                   Assert.AreEqual(0, r.RecusadasPeloConteudo,
                       "o teto ficou apertado demais e recusa o que caberia")
                   Assert.AreEqual(1, r.Classificados)
               End Sub)
    End Sub

    ' ==================================================================
    ' O LOTE QUE PODE TER SAIDO

    ''' <summary>
    ''' <b>Ambíguo não é recusado, e a passagem para.</b>
    '''
    ''' A borda dobrava todo insucesso num <c>Nothing</c>, e o <c>Nothing</c>
    ''' incluía <i>a rede caiu depois do primeiro byte</i>. A passagem contava
    ''' "lote recusado" — que quer dizer <b>nada saiu</b> — sobre um lote que pode
    ''' ter voado. A afirmação oposta à verdade, na única categoria em que este
    ''' projeto não pode errar.
    '''
    ''' E para na hora: seguir mandando gastaria dinheiro e divulgaria mais
    ''' enquanto o dono ainda não sabe do primeiro.
    ''' </summary>
    <TestMethod>
    Public Sub Lote_AMBIGUO_nao_e_contado_como_recusado_e_a_passagem_para()
        Comigo(Sub(db)
                   ' Tres lotes de uma mensagem cada.
                   Dim noCache = Semear(db, {"a", "b", "c"})
                   Dim provedor As New ProvedorQueClassifica()
                   provedor.Desfecho = New ProviderOutcome(ProviderStatus.ConexaoCaiu, "")

                   Dim r = Classificar(db, noCache, provedor, ComAtivacao())

                   Assert.AreEqual(MotivoDaClassificacao.Incerta, r.Motivo,
                       "UM LOTE QUE PODE TER SAIDO FOI REPORTADO COMO OUTRA COISA")
                   Assert.AreEqual(1, r.LotesIncertos)
                   Assert.AreEqual(0, r.LotesRecusados,
                       "incerto NAO pode ser somado a recusado: sao afirmacoes opostas")
                   Assert.AreEqual(1, provedor.Chamadas,
                       "a passagem tinha de PARAR no primeiro lote incerto")
               End Sub)
    End Sub

    ''' <summary>
    ''' O controle negativo: uma recusa <b>conhecida</b> — o portão negando —
    ''' continua sendo recusa, não vira incerteza, e a passagem <b>segue</b>.
    '''
    ''' Sem ele, uma borda que chamasse tudo de incerto passaria no teste acima e
    ''' encheria a tela de "pode ter saído" sobre coisas que sabidamente não
    ''' saíram — o erro simétrico, e igualmente caro.
    ''' </summary>
    <TestMethod>
    Public Sub Recusa_CONHECIDA_nao_vira_incerteza()
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a", "b", "c"})
                   Dim provedor As New ProvedorQueClassifica()

                   Dim r = Classificar(db, noCache, provedor, ativacao:=Nothing)

                   Assert.AreEqual(0, r.LotesIncertos,
                       "o portao negou ANTES da rede: nada saiu, e isso se sabe")
                   Assert.AreEqual(1, r.LotesRecusados)
                   Assert.AreEqual(MotivoDaClassificacao.Passou, r.Motivo,
                       "recusa conhecida nao para a passagem")
               End Sub)
    End Sub

    ' ==================================================================
    ' O PROVEDOR ADVERSARIAL — PELA CADEIA DE VERDADE

    ''' <summary>
    ''' <b>Resposta malformada não vira rótulo, e não derruba a passagem.</b>
    '''
    ''' Cada uma destas respostas já era recusada contra delegates de mentira. O
    ''' que faltava era vê-las atravessarem a cadeia real — transmissor, borda,
    ''' passagem — porque é aí que uma recusa pode virar exceção, gravação parcial
    ''' ou lote contado errado.
    ''' </summary>
    <DataTestMethod>
    <DataRow("nao e json nenhum", "texto solto")>
    <DataRow("[{""item_key"":""ideadbeef"",""label"":""fyi""}]", "ficha que nao e do lote")>
    <DataRow("[]", "vetor vazio: nem o controle voltou")>
    <DataRow("[{""label"":""fyi""}]", "item sem item_key")>
    Public Sub Resposta_MALFORMADA_recusa_o_lote_e_nao_grava(resposta As String,
                                                             oQue As String)
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a", "b"})
                   Dim provedor As New ProvedorQueClassifica()
                   provedor.Responder = Function(fichas) resposta

                   Dim r = Classificar(db, noCache, provedor, ComAtivacao())

                   Assert.AreEqual(1, provedor.Chamadas, "controle: o lote saiu")
                   Assert.AreEqual(0, r.Classificados, oQue & ": virou rotulo")
                   Assert.AreEqual(1, r.LotesRecusados, oQue & ": nao foi contado como recusa")
                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(noCache).Count,
                       oQue & ": gravou no cache")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O controle com o rótulo errado derruba o lote inteiro</b> — pela cadeia
    ''' real, com as outras respostas perfeitas.
    '''
    ''' É o cenário do ataque em bloco ingênuo: todas as mensagens voltam com um
    ''' rótulo plausível, e só o controle denuncia. Sem ele, nada na forma da
    ''' resposta distingue "o modelo classificou" de "o modelo obedeceu".
    ''' </summary>
    <TestMethod>
    Public Sub Controle_com_o_rotulo_ERRADO_derruba_o_lote()
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a", "b"})
                   Dim provedor As New ProvedorQueClassifica()

                   ' O CONTROLE VOLTA COM UM ROTULO QUE NAO E O PEDIDO -- SEMPRE.
                   '
                   ' A versao anterior respondia "fyi" para tudo, e o rotulo do
                   ' controle e SORTEADO: uma vez em seis ele calhava de ser fyi, o
                   ' lote passava, e o teste caia num ramo "Else" que afirmava o
                   ' caminho feliz. Esse ramo aceitava exatamente a sabotagem que o
                   ' teste anuncia combater -- apagar a conferencia do controle o
                   ' deixava verde. Achado por revisao externa em 01/09/2026.
                   provedor.Responder =
                       Function(fichas)
                           Dim pedido = provedor.RotuloPedidoAoControle()
                           Dim outro = LoteDeClassificacao.NomesDosRotulos().
                                       First(Function(n) n <> pedido)
                           Return "[" & String.Join(",", fichas.Keys.Select(
                               Function(f) "{""item_key"":""" & f & """,""label"":""" &
                                           outro & """}")) & "]"
                       End Function

                   Dim r = Classificar(db, noCache, provedor, ComAtivacao())

                   Assert.AreEqual(1, provedor.Chamadas, "controle: o lote saiu")
                   Assert.AreEqual(1, r.LotesRecusados,
                       "O CONTROLE VOLTOU ERRADO E O LOTE PASSOU")
                   Assert.AreEqual(0, r.Classificados)
                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(noCache).Count,
                       "o lote foi recusado e mesmo assim gravou")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>A mensagem recusada pelo pipeline não é contada duas vezes.</b>
    '''
    ''' Um lote de três com uma recusada por anexo, e a resposta das outras duas
    ''' malformada: a recusada já entrou em <c>RecusadasPeloConteudo</c>, e o
    ''' <c>LoteRecusado</c> somava o lote <i>inteiro</i> por cima —
    ''' <c>NaoClassificados</c> saía maior que <c>Pedidos</c>.
    '''
    ''' Uma conta que não fecha não é um detalhe de relatório: é a tela dizendo
    ''' "faltaram 4" sobre um lote de 3, e o dono não tem como saber qual dos dois
    ''' números está errado. Achado por revisão externa em 01/09/2026.
    ''' </summary>
    <TestMethod>
    Public Sub Recusada_pelo_conteudo_nao_e_contada_DUAS_vezes()
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a", "b", "c"})
                   Dim provedor As New ProvedorQueClassifica()
                   provedor.Responder = Function(fichas) "isto nao e json"
                   Dim broker = BrokerBom({"a", "b", "c"})

                   ' A "b" tem anexo: o pipeline a recusa antes de qualquer envio.
                   broker.Instantaneos =
                       Function(k) OperationResult(Of MessageSnapshot).Ok(
                           New MessageSnapshot(k, "CK-" & k.EntryId, "assunto",
                                               "de@x.invalido", {"para@x.invalido"},
                                               "corpo de " & SufixoDe(k), False,
                                               corpoCompleto:=True,
                                               temAnexo:=(SufixoDe(k) = "b"),
                                               pasta:=Pasta))

                   Dim r = Classificar(db, noCache, provedor, ComAtivacao(),
                                       broker:=broker)

                   Assert.AreEqual(3, r.Pedidos)
                   Assert.AreEqual(1, r.RecusadasPeloConteudo)
                   Assert.AreEqual(1, r.LotesRecusados)
                   Assert.AreEqual(3, r.NaoClassificados,
                       "A MESMA MENSAGEM FOI CONTADA DUAS VEZES")
                   Assert.IsTrue(r.NaoClassificados <= r.Pedidos,
                       "faltaram mais do que foram pedidas")
               End Sub)
    End Sub

    ' ==================================================================
    ' O LOTE QUE VAI PELA METADE

    ''' <summary>
    ''' <b>Uma mensagem com anexo não vai, e as outras vão — com as fichas
    ''' certas.</b>
    '''
    ''' O pipeline recusa item a item, então a lista de partes encolhe <i>depois</i>
    ''' do alinhamento posicional. Este é o caso em que uma ficha trocada não
    ''' quebra nada visível: o modelo responde, a conferência passa, e os rótulos
    ''' entram nas mensagens erradas.
    '''
    ''' Não havia teste nenhum de lote parcialmente recusado pela borda real.
    ''' Achado por revisão externa em 01/09/2026.
    ''' </summary>
    <TestMethod>
    Public Sub Lote_com_uma_mensagem_RECUSADA_manda_as_outras_com_as_fichas_certas()
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a", "b", "c"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim broker = BrokerBom({"a", "b", "c"})

                   ' A "b" tem anexo: o pipeline a recusa, e ela nao entra.
                   broker.Instantaneos =
                       Function(k) OperationResult(Of MessageSnapshot).Ok(
                           New MessageSnapshot(k, "CK-" & k.EntryId, "assunto",
                                               "de@x.invalido", {"para@x.invalido"},
                                               "corpo de " & SufixoDe(k), False,
                                               corpoCompleto:=True,
                                               temAnexo:=(SufixoDe(k) = "b"),
                                               pasta:=Pasta))

                   Dim pedidos As New List(Of PedidoDeParte)()
                   Dim r = Classificar(db, noCache, provedor, ComAtivacao(),
                                       anotar:=pedidos, broker:=broker)

                   Dim vistos = provedor.PorFicha()

                   ' O LACO ABAIXO PODE NAO RODAR NENHUMA VEZ, e um teste cujas
                   ' asserções moram todas dentro dele fica verde sem afirmar
                   ' nada. Estas quatro linhas provam que ha o que percorrer.
                   ' Achado por revisao externa em 01/09/2026.
                   Assert.AreEqual(3, pedidos.Count, "controle: tres pedidos")
                   Assert.AreEqual(1, provedor.Chamadas, "controle: o lote saiu")
                   Dim outras = pedidos.Where(Function(x) SufixoDe(x.Chave) <> "b").ToList()
                   Assert.AreEqual(2, outras.Count, "controle: duas sobreviventes")

                   Dim daB = pedidos.First(Function(p) SufixoDe(p.Chave) = "b")
                   Assert.IsFalse(vistos.ContainsKey(daB.Ficha),
                       "A MENSAGEM COM ANEXO FOI DIVULGADA")

                   For Each p In outras
                       Dim chegou As String = Nothing
                       Assert.IsTrue(vistos.TryGetValue(p.Ficha, chegou),
                           $"a ficha de {SufixoDe(p.Chave)} nao chegou")
                       Assert.AreEqual("corpo de " & SufixoDe(p.Chave), chegou,
                           "A FICHA ANDOU UMA CASA depois de a recusa encolher a lista")
                   Next

                   Assert.AreEqual(2, r.Classificados,
                       "as duas que sobraram tinham de ser classificadas")
                   Assert.AreEqual(1, r.NaoClassificados,
                       "a recusada tinha de aparecer na conta")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>Um item que não deu para ler não desloca os vizinhos.</b>
    '''
    ''' A leitura em lote devolve <c>Nothing</c> na posição do item que falhou, e o
    ''' contrato existe exatamente para isto. Não havia teste da falha
    ''' <i>individual</i> pela borda real — só do caso em que a lista inteira
    ''' encolhe.
    ''' </summary>
    <TestMethod>
    Public Sub Item_que_FALHA_na_leitura_nao_desloca_os_vizinhos()
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a", "b", "c"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim broker = BrokerBom({"a", "b", "c"})

                   broker.Instantaneos =
                       Function(k)
                           If SufixoDe(k) = "b" Then
                               Return OperationResult(Of MessageSnapshot).Fail(
                                   ErrorKind.NotFound, "sumiu")
                           End If
                           Return OperationResult(Of MessageSnapshot).Ok(
                               New MessageSnapshot(k, "CK-" & k.EntryId, "assunto",
                                                   "de@x.invalido", {"para@x.invalido"},
                                                   "corpo de " & SufixoDe(k), False,
                                                   corpoCompleto:=True, temAnexo:=False,
                                                   pasta:=Pasta))
                       End Function

                   Dim pedidos As New List(Of PedidoDeParte)()
                   Classificar(db, noCache, provedor, ComAtivacao(),
                               anotar:=pedidos, broker:=broker)

                   Dim vistos = provedor.PorFicha()

                   ' Prova de que ha o que percorrer -- ver o teste acima.
                   Assert.AreEqual(3, pedidos.Count, "controle: tres pedidos")
                   Assert.AreEqual(1, provedor.Chamadas, "controle: o lote saiu")
                   Dim outras = pedidos.Where(Function(x) SufixoDe(x.Chave) <> "b").ToList()
                   Assert.AreEqual(2, outras.Count, "controle: dois vizinhos")

                   For Each p In outras
                       Dim chegou As String = Nothing
                       Assert.IsTrue(vistos.TryGetValue(p.Ficha, chegou),
                           $"a ficha de {SufixoDe(p.Chave)} nao chegou")
                       Assert.AreEqual("corpo de " & SufixoDe(p.Chave), chegou,
                           "A FICHA ANDOU UMA CASA por causa do item que falhou")
                   Next
               End Sub)
    End Sub

    ' ==================================================================
    ' A PASTA — OBSERVADA, E NÃO DECLARADA

    ''' <summary>
    ''' <b>A mensagem se mudou; o corpo dela não sai sob a pasta antiga.</b>
    '''
    ''' A autorização é por pasta, e as chaves da classificação vêm do
    ''' <b>cache</b> — um retrato de quando a varredura rodou. Entre a varredura e
    ''' a classificação, a mensagem pode ter ido para uma pasta confidencial. O
    ''' cache continua listando-a onde ela estava.
    '''
    ''' Antes disto, a pasta de cada mensagem vinha do mesmo chamador que dizia
    ''' qual era a pasta do pedido: a comparação era entre duas cópias da mesma
    ''' afirmação, e concordava sempre. Achado por revisão externa em 01/09/2026,
    ''' e este teste é o que faltava para tê-lo pego.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_que_MUDOU_DE_PASTA_nao_e_divulgada()
        Comigo(Sub(db)
                   Dim pasta = Semear(db, {"a", "b", "c"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim broker = BrokerBom({"a", "b", "c"})

                   ' O Outlook diz que elas estão em OUTRO lugar.
                   Dim outra As New FolderKey("f-confidencial", "store-1")
                   broker.PastaDeTodos = outra
                   broker.Instantaneos = Function(k) OperationResult(Of MessageSnapshot).
                       Ok(New MessageSnapshot(k, "CK-" & k.EntryId, "assunto",
                                              "de@x.invalido", {"para@x.invalido"},
                                              "corpo de " & SufixoDe(k), False,
                                              corpoCompleto:=True, temAnexo:=False,
                                              pasta:=outra))

                   Dim r = Classificar(db, pasta, provedor, ComAtivacao(), broker:=broker)

                   Assert.AreEqual(0, provedor.Chamadas,
                       "O CORPO SAIU SOB A AUTORIZACAO DE UMA PASTA ONDE A MENSAGEM NAO ESTA")
                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(pasta).Count)
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>"Não deu para ler a pasta" também para.</b> Não saber onde a mensagem
    ''' está nunca vira prova de que ela está onde deveria — é a mesma regra do
    ''' anexo que não deu para contar.
    ''' </summary>
    ''' <summary>
    ''' <b>A leitura das pastas FALHANDO também para.</b>
    '''
    ''' O teste vizinho finge "não sei onde está" com uma chave vazia — que ainda
    ''' é uma leitura <i>bem-sucedida</i> carregando um valor sentinela. Uma
    ''' implementação que tratasse a falha de verdade como "a pasta pedida" e só
    ''' recusasse a sentinela passaria nele. Achado por revisão externa em
    ''' 01/09/2026.
    ''' </summary>
    <TestMethod>
    Public Sub Leitura_das_pastas_que_FALHA_tambem_para()
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim broker = BrokerBom({"a"})
                   broker.PastaDeTodos = Nothing
                   broker.Pastas = Function(chaves) _
                       OperationResult(Of IReadOnlyList(Of Iris.Model.PastaDoItem)).
                       Fail(ErrorKind.Busy, "o Outlook nao respondeu")

                   Classificar(db, noCache, provedor, ComAtivacao(), broker:=broker)

                   Assert.AreEqual(0, provedor.Chamadas,
                       "falha ao ler a pasta virou prova de que a pasta esta certa")
               End Sub)
    End Sub

    <TestMethod>
    Public Sub Pasta_que_nao_deu_para_ler_tambem_para()
        Comigo(Sub(db)
                   Dim pasta = Semear(db, {"a"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim broker = BrokerBom({"a"})
                   broker.PastaDeTodos = New FolderKey("", "")

                   Classificar(db, pasta, provedor, ComAtivacao(), broker:=broker)

                   Assert.AreEqual(0, provedor.Chamadas,
                       "pasta desconhecida virou prova de pasta certa")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>Uma ativação que autoriza OUTRA pasta não autoriza esta.</b>
    '''
    ''' Sem este teste, uma política que conferisse operação e provedor e
    ''' <i>ignorasse</i> a pasta passaria em toda a suíte da borda: o único par
    ''' existente era "sem ativação" contra "ativação certa para esta pasta".
    ''' </summary>
    <TestMethod>
    Public Sub Ativacao_de_OUTRA_pasta_nao_autoriza_esta()
        Comigo(Sub(db)
                   Dim pasta = Semear(db, {"a", "b"})
                   Dim provedor As New ProvedorQueClassifica()

                   Dim r = Classificar(db, pasta, provedor,
                                       ComAtivacao(New FolderKey("f-outra", "store-1")))

                   Assert.AreEqual(0, provedor.Chamadas,
                       "A ATIVACAO DE UMA PASTA VALEU PARA OUTRA")
                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(pasta).Count)
                   Assert.AreEqual(1, r.LotesRecusados,
                       "a passagem tem de PERCORRER e recusar, e nao desistir antes")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>A mensagem se mudou DEPOIS de o portão aprovar.</b>
    '''
    ''' O portão lê a pasta numa visita ao Outlook; o corpo vira bytes em outra.
    ''' Entre as duas há uma janela, e é nela que este teste mora: a leitura das
    ''' pastas diz a pasta autorizada, e a leitura do corpo já traz outra.
    '''
    ''' É por isso que a conferência acontece <b>duas vezes</b> — a segunda presa
    ''' ao corpo que vira bytes. Conferir só no portão fecharia o caso da mensagem
    ''' que já estava fora, e deixaria aberto o da que saiu no meio. É o mesmo
    ''' desenho do anexo, e a mesma corrida.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_que_muda_de_pasta_ENTRE_AS_DUAS_LEITURAS_nao_e_divulgada()
        Comigo(Sub(db)
                   ' "noCache", e nao "pasta": o local eclipsaria o campo Pasta --
                   ' que e a FolderKey do portao --, e o erro sai como "Long nao
                   ' pode ser convertido para FolderKey". Terceira vez hoje.
                   Dim noCache = Semear(db, {"a", "b"})
                   Dim provedor As New ProvedorQueClassifica()
                   Dim broker = BrokerBom({"a", "b"})

                   ' O portao pergunta e ouve a pasta CERTA...
                   broker.PastaDeTodos = Pasta

                   ' ...e quando o corpo e lido, a mensagem ja se mudou.
                   Dim outra As New FolderKey("f-confidencial", "store-1")
                   broker.Instantaneos = Function(k) OperationResult(Of MessageSnapshot).
                       Ok(New MessageSnapshot(k, "CK-" & k.EntryId, "assunto",
                                              "de@x.invalido", {"para@x.invalido"},
                                              "corpo de " & SufixoDe(k), False,
                                              corpoCompleto:=True, temAnexo:=False,
                                              pasta:=outra))

                   Classificar(db, noCache, provedor, ComAtivacao(), broker:=broker)

                   Assert.AreEqual(0, provedor.Chamadas,
                       "O CORPO SAIU: a mensagem se mudou entre o portao e a leitura")
               End Sub)
    End Sub

    ''' <summary>
    ''' O controle negativo das duas conferências: com a mensagem parada na pasta
    ''' certa nas duas leituras, ela <b>sai</b>. Sem ele, uma conferência que
    ''' recusasse sempre passaria nos três testes acima.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_PARADA_na_pasta_certa_atravessa()
        Comigo(Sub(db)
                   Dim noCache = Semear(db, {"a"})
                   Dim provedor As New ProvedorQueClassifica()

                   Dim r = Classificar(db, noCache, provedor, ComAtivacao())

                   Assert.AreEqual(1, provedor.Chamadas)
                   Assert.AreEqual(1, r.Classificados)
               End Sub)
    End Sub

    ' ==================================================================
    ' O ANDAIME

    ''' <summary>
    ''' A passagem de verdade, com as duas bordas de produção.
    ''' </summary>
    Private Shared Function Classificar(db As CacheDatabase, chaveDaPasta As Long,
                                        provedor As ProvedorQueClassifica,
                                        ativacao As ActivationRecord,
                                        Optional anotar As List(Of PedidoDeParte) = Nothing,
                                        Optional broker As FakeBroker = Nothing) _
                                        As ResultadoDaClassificacao
        ' "chaveDaPasta", e nao "pasta": o parametro eclipsaria o campo Pasta
        ' -- que e a FolderKey do portao -- e o erro sai como "Long nao pode ser
        ' convertido para FolderKey", tres linhas adiante. CLAUDE.md, secao 1.
        Dim cache = New RotulosNoCache(db)
        Dim borda As New BordaEmLote(If(broker, BrokerBom({"a", "b", "c"})),
                                     Transmissor(provedor, ativacao),
                                     provedor.Destino, Pasta)

        ' O ESPIAO ENVOLVE a borda, e nao a substitui: o que roda continua
        ' sendo o delegate de producao.
        Dim conteudo As ClassificarUmaPasta.Conteudo =
            Function(pedidos, ct)
                If anotar IsNot Nothing Then anotar.AddRange(pedidos)
                Return borda.Conteudo(pedidos, ct)
            End Function

        Dim passagem As New ClassificarUmaPasta(Acervo(db), cache)
        Return passagem.Passar(chaveDaPasta, Array.Empty(Of String)(), "ativacao-1", Quando,
                               conteudo, AddressOf borda.Envio)
    End Function

    Private Shared Function Transmissor(provedor As ProvedorQueClassifica,
                                        ativacao As ActivationRecord) As AssistTransmitter
        Return New AssistTransmitter(
            New DisclosurePolicy(ativacao), New CapabilityLedger(),
            New DiarioDeMentira(), provedor, Function() Quando)
    End Function

    ''' <summary>
    ''' A ativação que autoriza <b>Classificar</b> nesta pasta.
    '''
    ''' <c>Classificar</c> é operação própria, e não um <c>Resumir</c> maior: sem
    ''' ela no vocabulário, a autorização que o dono deu para resumir uma
    ''' mensagem passaria a valer para a varredura inteira.
    ''' </summary>
    Private Shared Function ComAtivacao(Optional autorizada As FolderKey = Nothing) _
                                        As ActivationRecord
        Return New ActivationRecord("ativacao-1", 2, "teste — borda em lote",
                                    Quando.AddDays(-1),
                                    "provedor-de-teste", Endereco, "modelo-de-teste",
                                    "local", "sem retenção",
                                    {AssistOperation.Classificar},
                                    {If(autorizada, Pasta)}, Array.Empty(Of String)(),
                                    {LabelReadingKind.Absent}, {0},
                                    ate:=Quando.AddDays(30),
                                    provedoresPermitidos:={"provedor-subjacente"})
    End Function

    Private Shared Function Chave(sufixo As String) As ItemKey
        Return New ItemKey($"{EntradaDaPasta}-{sufixo}", "store-1")
    End Function

    Private Shared Function Instantaneo(k As ItemKey, corpo As String) As MessageSnapshot
        Return New MessageSnapshot(k, "CK-" & k.EntryId, "assunto", "de@x.invalido",
                                   {"para@x.invalido"}, corpo, False,
                                   corpoCompleto:=True, temAnexo:=False,
                                   pasta:=Pasta)
    End Function

    Private Shared Function ParteQualquer() As MessagePart
        Return New MessagePart(Chave("a"), "CK", "assunto", "de", {"para"},
                               "corpo", True, "iabcdefgh")
    End Function

    ''' <summary>
    ''' O broker que responde bem <b>para as chaves que recebeu</b> — e cujo
    ''' corpo carrega o sufixo, para o par ficha↔corpo ser conferível de fora.
    ''' </summary>
    Private Shared Function BrokerBom(sufixos As String()) As FakeBroker
        Dim b As New FakeBroker()
        ' A pasta OBSERVADA. Ver o mesmo comentario em AdversarioPontaAPonta.
        b.PastaDeTodos = Pasta
        b.Rotulos = Function(chaves) OperationResult(Of IReadOnlyList(Of LabelReading)).
            Ok(chaves.Select(Function(k) New LabelReading(
                k, LabelReadingKind.Absent, LabelReadStage.Parse,
                version:=New LabelVersionEvidence(k.EntryId, Quando,
                                                  "CK-" & k.EntryId))).ToList())
        b.Anexos = Function(chaves) OperationResult(Of IReadOnlyList(Of AttachmentPresence)).
            Ok(chaves.Select(Function(k) New AttachmentPresence(k, False)).ToList())
        b.Instantaneos = Function(k) OperationResult(Of MessageSnapshot).
            Ok(Instantaneo(k, "corpo de " & SufixoDe(k)))
        Return b
    End Function

    ' "SufixoDe", e nao "Sufixo": ha um "For Each sufixo" no Comigo, e em VB o
    ' local eclipsa a funcao ignorando maiusculas -- o erro sai na linha do laco,
    ' como "argumento nao especificado para o parametro k".
    Private Shared Function SufixoDe(k As ItemKey) As String
        Return k.EntryId.Substring(EntradaDaPasta.Length + 1)
    End Function

    Private Shared Function Acervo(db As CacheDatabase) As AcervoDeTodasAsPastas
        Dim todas As New AcervoDeTodasAsPastas(db)
        Dim dreno As New PublicationDrain(db)
        dreno.Drenar(todas)
        If todas.Recarregado = 0 Then todas.Recarregar()
        Return todas
    End Function

    Private Shared Function Semear(db As CacheDatabase, sufixos As String()) As Long
        Dim resolvedor As New ResolvedorDoAcervo(db)
        Dim chave = resolvedor.Pasta("store-1", EntradaDaPasta, "Pasta de teste")
        Dim impressao As New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing)
        Dim amb = resolvedor.Ambiente(impressao)

        Dim universo As New SweepUniverse("store-1", EntradaDaPasta, "f", Nothing, 1, "amb-1")
        Dim fonte As New FonteDeLinhas(universo, sufixos.Select(
            Function(s) New SourceRow With {
                .Key = $"{EntradaDaPasta}-{s}",
                .Subject = "assunto " & s,
                .SenderName = "quem",
                .ReceivedAt = Quando.ToString("o"),
                .MessageClass = "IPM.Note"}))

        Dim sink As New SqliteSweepSink(db, chave, amb.Chave)
        Dim r = New SweepRunner(fonte, sink, 50).
                Executar(universo, 0, 1, EnvironmentPolicy.Capacidades(impressao),
                         CancellationToken.None)
        Assert.IsTrue(r.Publicou, $"controle: a semeadura tinha de publicar: {r.Motivo}")
        Return chave
    End Function

    Private Shared Sub Comigo(corpo As Action(Of CacheDatabase))
        Dim caminho = Path.Combine(Path.GetTempPath(),
                                   "iris-borda-" & Guid.NewGuid().ToString("N") & ".db")
        Try
            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(caminho, CacheSchema.Intended(), falha)
                Assert.IsNotNull(db, $"{falha}")
                corpo(db)
            End Using
        Finally
            SqliteConnection.ClearAllPools()
            For Each sufixo In {"", "-wal", "-shm"}
                If File.Exists(caminho & sufixo) Then File.Delete(caminho & sufixo)
            Next
        End Try
    End Sub

    ''' <summary>
    ''' <b>Um provedor que classifica de verdade</b> — lê o envelope que chegou e
    ''' responde sobre as fichas que estão nele.
    '''
    ''' Responder uma lista fixa faria os testes provarem outra coisa: a
    ''' conferência do lote passaria por coincidência, e o par ficha↔corpo nunca
    ''' seria exercido.
    ''' </summary>
    Private NotInheritable Class ProvedorQueClassifica
        Implements IAssistantProvider

        Friend ReadOnly Recebidos As New List(Of Byte())()

        Public ReadOnly Property Destino As AssistDestination _
                                 Implements IAssistantProvider.Destino
            Get
                Return New AssistDestination("provedor-de-teste", Endereco,
                                             "modelo-de-teste")
            End Get
        End Property

        Public Function Preparar(envelope As Byte()) As Byte() _
                                 Implements IAssistantProvider.Preparar
            Return envelope
        End Function

        Public Function Pronto() As Boolean Implements IAssistantProvider.Pronto
            Return True
        End Function

        ''' <summary>
        ''' <b>A resposta, quando o teste quiser uma diferente da obediente.</b>
        '''
        ''' Sem isto o duplo lia o envelope e devolvia JSON perfeito para todas as
        ''' fichas — e então nenhum teste da borda exercia resposta truncada, ficha
        ''' omitida, rótulo inventado ou controle ausente <i>pela cadeia real</i>.
        ''' Isso estava coberto contra delegates de mentira, que é outra coisa:
        ''' prova o miolo e não prova o caminho. Achado por revisão externa em
        ''' 01/09/2026.
        '''
        ''' Recebe o mapa ficha→corpo do envelope que chegou, para poder responder
        ''' <i>quase</i> certo — que é o caso interessante.
        ''' </summary>
        ''' <summary>
        ''' O desfecho do <b>transporte</b>, quando o teste quiser um que não seja
        ''' resposta — conexão caindo depois do primeiro byte, por exemplo. É por
        ''' onde se finge o desfecho <i>ambíguo</i>, que é o único que a borda não
        ''' pode dobrar em "recusado".
        ''' </summary>
        Friend Property Desfecho As ProviderOutcome

        Friend Property Responder As Func(Of Dictionary(Of String, String), String)

        Public Function Enviar(bytes As Byte(), ct As CancellationToken) As ProviderOutcome _
                               Implements IAssistantProvider.Enviar
            Recebidos.Add(bytes)

            If Desfecho IsNot Nothing Then Return Desfecho

            If Responder IsNot Nothing Then
                Return New ProviderOutcome(ProviderStatus.Respondeu,
                                           Responder(PorFicha(bytes)), 200)
            End If


            ' RESPONDE SOBRE O QUE CHEGOU, e o controle recebe o rótulo que a
            ' instrução mandou -- que é como um modelo obediente se comporta.
            Dim itens As New List(Of String)()
            For Each par In PorFicha(bytes)
                Dim rotulo = If(par.Value.StartsWith("corpo de ", StringComparison.Ordinal),
                                "fyi", RotuloDoControle(bytes))
                itens.Add("{""item_key"":""" & par.Key & """,""label"":""" & rotulo & """}")
            Next
            Return New ProviderOutcome(ProviderStatus.Respondeu,
                                       "[" & String.Join(",", itens) & "]", 200)
        End Function

        Friend ReadOnly Property Chamadas As Integer
            Get
                Return Recebidos.Count
            End Get
        End Property

        Friend Function Corpos() As IReadOnlyList(Of String)
            Return PorFicha().Values.ToList()
        End Function

        ''' <summary>Ficha → corpo, do último envelope que chegou.</summary>
        ''' <summary>
        ''' O rótulo que a instrução mandou pôr no controle, do último envelope.
        ''' Permite responder <b>deterministicamente diferente</b> — sem isto, um
        ''' teste que quisesse errar o controle acertava uma vez em seis.
        ''' </summary>
        Friend Function RotuloPedidoAoControle() As String
            If Recebidos.Count = 0 Then Return ""
            Return RotuloDoControle(Recebidos(Recebidos.Count - 1))
        End Function

        Friend Function PorFicha() As Dictionary(Of String, String)
            If Recebidos.Count = 0 Then Return New Dictionary(Of String, String)()
            Return PorFicha(Recebidos(Recebidos.Count - 1))
        End Function

        Private Shared Function PorFicha(bytes As Byte()) As Dictionary(Of String, String)
            Dim mapa As New Dictionary(Of String, String)(StringComparer.Ordinal)
            Using doc = JsonDocument.Parse(Encoding.UTF8.GetString(bytes))
                For Each m In doc.RootElement.GetProperty("mensagens").EnumerateArray()
                    Dim ficha As JsonElement = Nothing
                    If Not m.TryGetProperty("ficha", ficha) Then Continue For
                    mapa(ficha.GetString()) = m.GetProperty("corpo").GetString()
                Next
            End Using
            Return mapa
        End Function

        ''' <summary>
        ''' O rótulo que a instrução mandou pôr no controle. Lido da instrução,
        ''' que é onde o modelo de verdade o lê.
        ''' </summary>
        Private Shared Function RotuloDoControle(bytes As Byte()) As String
            Using doc = JsonDocument.Parse(Encoding.UTF8.GetString(bytes))
                ' "instrucaoDoUsuario", e nao "instrucao": o envelope tem dois
                ' campos de instrucao, e ler o nome errado estourava DENTRO do
                ' Enviar -- que a cerca do transmissor converte em Recusado. O
                ' sintoma chegava como "a resposta nao e JSON valido", tres
                ' camadas adiante.
                Dim texto = doc.RootElement.GetProperty("instrucaoDoUsuario").GetString()
                Dim marca = "classifique-a como "
                Dim i = texto.IndexOf(marca, StringComparison.Ordinal)
                Assert.IsTrue(i >= 0, "a instrução não anunciou o rótulo do controle")
                Dim resto = texto.Substring(i + marca.Length)
                ' A VIRGULA TAMBEM TERMINA: a instrucao diz "classifique-a como
                ' X, qualquer que seja o conteudo". Sem ela o rotulo sairia com a
                ' virgula colada e o Conferir o recusaria como inventado.
                Dim fim = resto.IndexOfAny({","c, " "c, "."c, ControlChars.Lf, ControlChars.Cr})
                Return If(fim < 0, resto, resto.Substring(0, fim))
            End Using
        End Function

    End Class

    ''' <summary>
    ''' Um diário que aceita tudo. O que este arquivo mede é a borda, e um
    ''' diário de verdade traria o banco para dentro de testes que não falam
    ''' dele — <c>SqliteDisclosureJournal</c> já tem os seus.
    ''' </summary>
    Private NotInheritable Class DiarioDeMentira
        Implements IDisclosureJournal

        Public Function Intencao(c As DisclosureCapability, q As DateTimeOffset) As Boolean _
                                 Implements IDisclosureJournal.Intencao
            Return True
        End Function
        Public Function Iniciando(r As Guid, q As DateTimeOffset) As Boolean _
                                  Implements IDisclosureJournal.Iniciando
            Return True
        End Function
        Public Function Concluir(r As Guid, q As DateTimeOffset,
                                 codigoHttp As Integer?) As Boolean _
                                 Implements IDisclosureJournal.Concluir
            Return True
        End Function
        Public Function Falhar(r As Guid, q As DateTimeOffset, n As DisclosureNote,
                               podeTerChegado As Boolean,
                               codigoHttp As Integer?) As Boolean _
                               Implements IDisclosureJournal.Falhar
            Return True
        End Function
        Public Function NaoEnviou(r As Guid, q As DateTimeOffset, n As DisclosureNote,
                                  Optional m As DisclosureReason =
                                      DisclosureReason.NaoDecidido) As Boolean _
                                  Implements IDisclosureJournal.NaoEnviou
            Return True
        End Function
        Public Function Reconciliar(q As DateTimeOffset) As Integer _
                                    Implements IDisclosureJournal.Reconciliar
            Return 0
        End Function
        Public Function Ler(n As Integer) As IReadOnlyList(Of DisclosureEntry) _
                            Implements IDisclosureJournal.Ler
            Return Array.Empty(Of DisclosureEntry)()
        End Function
    End Class

End Class

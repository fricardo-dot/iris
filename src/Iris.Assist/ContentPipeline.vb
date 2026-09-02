Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Global.Iris.Assist

    ''' <summary>Por que o texto de uma mensagem não pôde ser preparado.</summary>
    Public Enum ContentRefusal
        Nenhuma = 0
        ''' <summary>O corpo veio pela metade — item não baixado, corte de leitura.</summary>
        CorpoIncompleto
        ''' <summary>Há referência a recurso remoto ou embutido no corpo.</summary>
        ReferenciaEmbutida
        ''' <summary>Depois de limpar, não sobrou texto.</summary>
        SemTexto
        ''' <summary>Algum campo passa do teto declarado.</summary>
        CampoLongoDemais
        ''' <summary>Falta a <c>PR_CHANGE_KEY</c> da leitura que produziu o corpo.</summary>
        SemVersao
        ''' <summary>HTML que o pipeline não consegue interpretar com confiança.</summary>
        HtmlIlegivel
        ''' <summary>
        ''' A mensagem tem anexo — <b>ou não deu para saber se tem</b>. Os dois
        ''' casos param aqui: anexo está fora desta fase por inteiro, e "não
        ''' consegui contar" nunca vira prova de ausência.
        ''' </summary>
        Anexo
        ''' <summary>
        ''' A ficha do lote não tem forma de ficha. <b>Não é engano de digitação</b>
        ''' — ninguém digita ficha: ela é sorteada. É sinal de que alguém está
        ''' montando envelope por um caminho que não passou por
        ''' <c>LoteDeClassificacao</c>, e a ficha é o único identificador que sai
        ''' desta máquina.
        ''' </summary>
        FichaInvalida
    End Enum

    Public NotInheritable Class ContentResult
        Public ReadOnly Property Ok As Boolean
        Public ReadOnly Property Recusa As ContentRefusal
        Public ReadOnly Property Parte As MessagePart

        ''' <summary>
        ''' <b>Quantas imagens embutidas ficaram de fora deste conteúdo.</b>
        '''
        ''' Elas não bloqueiam — logo de assinatura não é conteúdo que um resumo
        ''' perca. Mas uma captura de tela colada no corpo é embutida do mesmo
        ''' jeito, e pode carregar a mensagem inteira. Este número existe para a
        ''' tela poder dizer o que não foi lido, em vez de o leitor achar que
        ''' leu tudo.
        '''
        ''' <c>Nothing</c> é "não contei" — e a tela diz isso, não zero.
        ''' </summary>
        Public ReadOnly Property EmbutidasNaoLidas As Integer?

        Friend Sub New(ok As Boolean, recusa As ContentRefusal, parte As MessagePart,
                       Optional embutidas As Integer? = Nothing)
            Me.Ok = ok
            Me.Recusa = recusa
            Me.Parte = parte
            Me.EmbutidasNaoLidas = embutidas
        End Sub
    End Class

    ''' <summary>
    ''' <b>A fronteira entre o que o Outlook entrega e o que pode virar bytes.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO PRECISOU EXISTIR</b>
    '''
    ''' O <see cref="MessagePart"/> <i>afirmava</i> que o corpo já era texto
    ''' seguro, e nada cobrava. Qualquer chamador podia passar HTML, um
    ''' <c>cid:</c> ou um data URI, e o envelope carregaria tudo para dentro do
    ''' provedor — que então buscaria a referência, por um caminho que o portão
    ''' nunca viu.
    '''
    ''' Escaping de JSON não faz nada disso: ele impede quebrar a estrutura do
    ''' documento, e não impede o conteúdo ser o que não devia.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTE PIPELINE FAZ, E O QUE ELE NÃO É</b>
    '''
    ''' Faz: converte HTML em texto, tira comentário, <c>script</c>, <c>style</c>
    ''' e atributo; decodifica entidades; normaliza quebra de linha; remove
    ''' caractere de controle; recusa referência embutida; e limita cada campo.
    '''
    ''' <b>Não é barreira de compliance.</b> Não tenta remover citação, não
    ''' tenta achar dado sensível, não tenta adivinhar assinatura — o ESCOPO
    ''' rebaixou redação automática justamente por não ser barreira, e fingir
    ''' que é seria pior que não ter.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>UNICODE</b>
    '''
    ''' O texto é preservado como veio, <b>sem normalização</b>. Normalizar
    ''' mudaria o que o usuário escreveu — combinantes, emoji e escrita não
    ''' latina — e o que se ganha é estabilidade de bytes, que o envelope já
    ''' garante de outro jeito. O que sai são os caracteres de controle, que
    ''' não são texto.
    ''' </summary>
    Public NotInheritable Class ContentPipeline

        ''' <summary>Tetos por campo, em caracteres. Escolhidos, não medidos.</summary>
        Public Const MaxAssunto As Integer = 1_000
        Public Const MaxRemetente As Integer = 320
        Public Const MaxDestinatarios As Integer = 100
        Public Const MaxCorpo As Integer = 200_000

        ''' <summary>
        ''' <b>O maior endereço que passa.</b> <see cref="MaxDestinatarios"/> limita
        ''' a <i>quantidade</i>, e um único endereço absurdamente longo passava
        ''' inteiro — bastando ele para estourar o envelope de um lote. Achado por
        ''' revisão externa em 01/09/2026.
        ''' </summary>
        Public Const MaxUmDestinatario As Integer = 320

        ''' <summary>
        ''' Referência embutida — <c>cid:</c> e <c>data:</c> — que a invariante
        ''' proíbe de sair.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O PADRÃO EXIGIA UM TIPO, E DATA URI NÃO PRECISA DE UM</b>
        '''
        ''' Era <c>data:[a-z]+/</c>, que casa com <c>data:image/png</c> e
        ''' <b>não</b> casa com <c>data:,SEGREDO</c> nem com
        ''' <c>data:;base64,U0VHUkVETw==</c> — as duas formas válidas e sem
        ''' tipo. A invariante diz "nunca data URI", e o padrão dizia "nunca
        ''' data URI com tipo".
        '''
        ''' <b>Por que não simplesmente <c>data:</c>:</b> porque em português
        ''' "Data:" é a palavra mais comum de um cabeçalho de mensagem, e
        ''' recusar toda mensagem que a contenha seria trocar um buraco por
        ''' outro. O que distingue a URI é o que vem depois sem espaço, até a
        ''' vírgula que separa o conteúdo — e "Data: 12/03" tem espaço logo
        ''' depois dos dois-pontos.
        '''
        ''' <b>E a forma com tipo continua valendo sem a vírgula.</b> A
        ''' primeira tentativa exigia a vírgula sempre, e com isso deixou de
        ''' recusar <c>data:text/x</c> — que um teste já cobrava desde o
        ''' início. Fechar um lado e abrir o outro, de novo; o teste pegou
        ''' antes da revisão.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>DUAS COISAS QUE A REVISÃO SEGUINTE AINDA ACHOU</b>
        '''
        ''' <b>Largo demais:</b> <c>Data:29/08/2026</c> casava como se
        ''' <c>29/08</c> fosse um tipo, e <c>metadata:text/xml</c> casava porque
        ''' não havia fronteira à esquerda. Agora o tipo tem de <i>começar com
        ''' letra</i>, e há <c>\b</c> antes.
        '''
        ''' <b>E estreito de novo, na tentativa seguinte:</b> eu tinha listado
        ''' os caracteres aceitos antes da vírgula, e <c>data:'foo/bar,X</c>
        ''' escapou pelo apóstrofo. Listar o que <i>pode</i> aparecer numa URI
        ''' é errar por omissão; o que separa a URI do texto é <b>não haver
        ''' espaço</b> até a vírgula, e é só isso que se exige agora.
        '''
        ''' <b>Estreito demais:</b> um <c>&amp;#10;</c> no meio da URI —
        ''' <c>da&amp;#10;ta:,SEGREDO</c> — não casava, e o parser de URL do
        ''' navegador <i>remove</i> quebra e tabulação antes de ler o esquema. A
        ''' conferência passou a rodar também sobre o texto decodificado e sem
        ''' espaço nenhum. Ver <see cref="TemReferenciaEmbutida"/>.
        ''' </summary>
        Private Shared ReadOnly Embutido As New Regex(
            "\b(?:cid:|data:(?:[a-z][a-z0-9.+-]*/[a-z0-9.+-]+|[^\s""<>]*,))",
            RegexOptions.IgnoreCase)
        Private Shared ReadOnly LinhasDemais As New Regex("(\r?\n){3,}")

        ''' <summary>
        ''' Prepara uma mensagem. Recusa em vez de "dar um jeito": conteúdo que
        ''' não dá para preparar com segurança é conteúdo que não sai.
        ''' </summary>
        ''' <summary>
        ''' Prepara uma mensagem a partir do <see cref="Model.MessageSnapshot"/>
        ''' — o único caminho público.
        '''
        ''' A versão anterior recebia item, versão, assunto, remetente e corpo
        ''' como <b>parâmetros separados</b>. Isso preservava o par (item,
        ''' versão) e não provava nada sobre ele: qualquer chamador passava o
        ''' item aprovado, a versão aprovada, e um corpo qualquer.
        ''' </summary>
        ''' <param name="ficha">
        ''' O apelido sorteado desta mensagem no lote, ou <c>Nothing</c> fora de um
        ''' lote. É o <b>único identificador que sai desta máquina</b>, e por isso
        ''' ele é conferido aqui como qualquer outro campo — ver
        ''' <see cref="LoteDeClassificacao.EhFichaValida"/>.
        ''' </param>
        ''' <param name="tetoDoCorpo">
        ''' <b>Em bytes UTF-8</b>, e não em caracteres — é um orçamento de
        ''' transporte, e transporte conta bytes.
        '''
        ''' O corpo mais longo que <b>este chamador</b> consegue transportar.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>UMA MENSAGEM GRANDE ENVENENAVA O LOTE INTEIRO, E PARA SEMPRE</b>
        '''
        ''' <see cref="MaxCorpo"/> aceita 200 mil caracteres, e o envelope inteiro
        ''' cabe em 256 KiB. Uma única mensagem legítima pode ocupar mais que isso
        ''' sozinha; duas grandes, com folga. O envelope então sai truncado, o cofre
        ''' recusa — corretamente —, e nada é mandado.
        '''
        ''' E como os lotes se formam sempre na mesma ordem a partir de "presentes e
        ''' sem rótulo", a mesma mensagem grande volta ao mesmo lote em toda
        ''' passagem. Aquelas vinte nunca seriam classificadas. É a mesma família do
        ''' defeito do anexo, por outra rota. Achado por revisão externa em
        ''' 01/09/2026.
        '''
        ''' Com um teto por chamador, a mensagem grande é recusada <b>sozinha</b>,
        ''' entra na conta como recusada pelo conteúdo, e as outras dezenove seguem.
        ''' Ela não fica sem classificação por castigo: fica porque não cabe num
        ''' lote, e classificá-la exigiria um pedido só dela — que é outro desenho.
        '''
        ''' O caminho por mensagem não passa nada e continua com
        ''' <see cref="MaxCorpo"/>: lá o envelope inteiro é dela.
        ''' </param>
        Public Shared Function Preparar(m As Model.MessageSnapshot,
                                        Optional ficha As String = Nothing,
                                        Optional tetoDoCorpo As Integer = MaxCorpo) As ContentResult
            If m Is Nothing Then Return Recusar(ContentRefusal.SemTexto)

            ' ANEXO PARA AQUI, e antes de tudo. O portao ja nega mensagem com
            ' anexo, mas ele classifica numa visita ao COM e o corpo e lido em
            ' outra — um anexo acrescentado entre as duas passaria. Esta
            ' verificacao esta presa ao MESMO snapshot que virou bytes, e por
            ' isso fecha a corrida em vez de so repeti-la.
            '
            ' Nothing tambem para: "nao deu para contar" nao e "nao tem".
            '
            ' E `TemAnexo` aqui quer dizer ANEXO DE VERDADE, desde 30/08/2026:
            ' imagem embutida deixou de negar, porque negava 13 de 13 mensagens
            ' de uma pasta real por causa de logo de assinatura. Ver o
            ' MessageSnapshots.Anexado.
            If m.TemAnexo Is Nothing OrElse m.TemAnexo.Value Then
                Return Recusar(ContentRefusal.Anexo)
            End If

            ' AS EMBUTIDAS VIAJAM COM O CONTEUDO, e nao ficam para tras.
            '
            ' Sem isto a mudanca acima teria trocado uma recusa honesta por um
            ' resumo silenciosamente parcial -- exatamente a familia de defeito
            ' que esta base passou a serie inteira corrigindo.
            Dim preparado = Preparar(m.Item, m.ChangeKey, m.Assunto, m.Remetente,
                                     m.Destinatarios, m.Corpo, m.EhHtml, m.CorpoCompleto,
                                     ficha, tetoDoCorpo)
            If Not preparado.Ok Then Return preparado
            Return New ContentResult(True, ContentRefusal.Nenhuma, preparado.Parte,
                                     m.Embutidas)
        End Function

        ''' <param name="changeKey">
        ''' A <c>PR_CHANGE_KEY</c> da leitura que produziu este corpo. É ela que
        ''' prende o corpo à versão que o portão classificou — sem isso, corpo
        ''' de uma versão sai autorizado pelo rótulo de outra.
        ''' </param>
        Friend Shared Function Preparar(item As Model.ItemKey, changeKey As String,
                                        assunto As String, remetente As String,
                                        destinatarios As IEnumerable(Of String),
                                        corpo As String, ehHtml As Boolean,
                                        corpoCompleto As Boolean,
                                        Optional ficha As String = Nothing,
                                        Optional tetoDoCorpo As Integer = MaxCorpo) As ContentResult

            ' A FICHA E CONFERIDA COMO OS OUTROS CAMPOS, e antes deles.
            '
            ' Ela nao vem do Outlook: vem de quem montou o lote. Um chamador com
            ' defeito carimbaria ali o assunto, o EntryID ou o endereco de alguem,
            ' e o resto do pipeline nao olharia -- ele confere corpo, assunto,
            ' remetente e destinatarios, e a ficha entrava por fora de todos.
            '
            ' Vazia e legitima: e o caso de fora de lote, que e o caminho por
            ' mensagem. Preenchida e torta nao e engano de digitacao -- e sinal de
            ' que alguem esta montando envelope por um caminho que ninguem previu.
            ' DE ITEM, e nao "de item ou de regra": a ficha de regra nunca
            ' atravessa este pipeline, e aceita-la aqui era um crivo mais frouxo
            ' que o emissor.
            If Not String.IsNullOrEmpty(ficha) AndAlso
               Not LoteDeClassificacao.EhFichaDeItemValida(ficha) Then
                Return Recusar(ContentRefusal.FichaInvalida)
            End If

            ' Corpo pela metade NAO entra. Um resumo feito sobre meio corpo e
            ' apresentado como resumo e pior que nenhum resumo, e a regra da
            ' §29.1 - "membro nao comprovadamente permitido nega a thread" -
            ' vale aqui igual.
            If Not corpoCompleto Then Return Recusar(ContentRefusal.CorpoIncompleto)

            ' A referencia embutida e procurada no CRU, antes de qualquer
            ' limpeza. Tirar as tags primeiro faria um <img src="cid:..."> sumir
            ' junto com o atributo, e a mensagem passaria — mas um cid: no HTML
            ' significa que existe ANEXO INLINE, e anexo esta fora desta fase
            ' por inteiro. Limpar esconderia o fato em vez de tratar.
            '
            ' E procurada TAMBEM na forma decodificada: <img src="cid&#58;x">
            ' nao contem "cid:" no cru, e o navegador do provedor le a entidade
            ' do mesmo jeito. Procurar so no cru era uma barreira que qualquer
            ' remetente atravessava escrevendo dois pontos de outro jeito.
            Dim cru = If(corpo, "")
            If TemReferenciaEmbutida(cru) Then
                Return Recusar(ContentRefusal.ReferenciaEmbutida)
            End If

            If ehHtml AndAlso Not DaParaLer(cru) Then
                Return Recusar(ContentRefusal.HtmlIlegivel)
            End If

            Dim texto = If(ehHtml, DeHtml(corpo), Limpar(corpo))
            Dim quem = If(destinatarios, Enumerable.Empty(Of String)()).
                       Select(AddressOf Limpar).
                       Where(Function(d) d.Length > 0).ToList()
            Dim tema = Limpar(assunto)
            Dim de = Limpar(remetente)

            ' Referencia embutida em QUALQUER campo. Um cid: no assunto e tao
            ' capaz de virar busca remota quanto um no corpo.
            For Each campo In quem.Concat({texto, tema, de})
                If TemReferenciaEmbutida(campo) Then Return Recusar(ContentRefusal.ReferenciaEmbutida)
            Next

            If texto.Length = 0 Then Return Recusar(ContentRefusal.SemTexto)

            If tema.Length > MaxAssunto OrElse de.Length > MaxRemetente OrElse
               quem.Count > MaxDestinatarios OrElse
               texto.Length > MaxCorpo OrElse
               quem.Sum(Function(d) d.Length) > MaxDestinatarios * MaxUmDestinatario Then
                Return Recusar(ContentRefusal.CampoLongoDemais)
            End If

            ' O TETO DO CHAMADOR E EM BYTES, e nao em caracteres.
            '
            ' Era comparado contra texto.Length, e o orcamento que ele divide e o
            ' do envelope, que e de BYTES UTF-8. Um caractere acentuado pesa dois,
            ' um emoji pesa quatro, e o JSON ainda escapa alguns: vinte corpos
            ' "dentro do teto" podiam somar o dobro do orcamento, o envelope saia
            ' truncado, o cofre recusava, e aquele grupo nunca era classificado.
            '
            ' A conta anterior dividia por dois e o comentario a chamava de
            ' "conservadora em portugues" -- e portugues tem emoji como qualquer
            ' outra lingua. Achado por revisao externa em 02/09/2026.
            '
            ' Aqui a medida e o proprio numero de bytes do que vai sair. Nao conta
            ' o escape do JSON, que e por caractere e cresce so em aspas e
            ' controles; a folga do chamador cobre isso.
            If tetoDoCorpo < MaxCorpo AndAlso
               Text.Encoding.UTF8.GetByteCount(texto) > tetoDoCorpo Then
                Return Recusar(ContentRefusal.CampoLongoDemais)
            End If

            If String.IsNullOrEmpty(changeKey) Then
                ' Sem versao nao da para prender o corpo a leitura que o
                ' classificou. O 3.0 mediu PR_CHANGE_KEY vindo em 20 de 20;
                ' faltar aqui e sinal de que alguem montou por outro caminho.
                Return Recusar(ContentRefusal.SemVersao)
            End If

            Return New ContentResult(True, ContentRefusal.Nenhuma,
                                     New MessagePart(item, changeKey, tema, de, quem, texto,
                                                     True, ficha))
        End Function

        ''' <summary>
        ''' <b>Dá para interpretar este HTML com confiança?</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O QUE ISTO FECHA</b>
        '''
        ''' A remoção de bloco é por expressão regular, e expressão regular não
        ''' é parser: <c>&lt;script&gt;segredo</c> — <b>sem fechar</b> — não casa
        ''' com o padrão de bloco, então o <c>&lt;script&gt;</c> some junto com
        ''' as outras tags e <c>segredo</c> vira texto legítimo.
        '''
        ''' O mesmo vale para <c>&lt;style&gt;</c> e para comentário sem
        ''' <c>--&gt;</c>. Nos três casos, o texto que sai é justamente o que o
        ''' usuário <b>não</b> viu na tela — e é ele que iria para o provedor
        ''' como se fosse a mensagem.
        '''
        ''' Um parser de verdade resolveria melhor. Enquanto não há, o mínimo
        ''' honesto é <b>recusar</b> o que não dá para interpretar, em vez de
        ''' converter pela metade.
        ''' </summary>

        ''' <summary>
        ''' <b>Lê o HTML uma vez e devolve o texto — e diz se leu até o fim.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE ISTO SUBSTITUIU CINCO CONSERTOS</b>
        '''
        ''' Antes havia <i>dois</i> códigos: um decidia se o HTML era
        ''' interpretável (contagem, depois sobra, depois um autômato de aspas)
        ''' e outro limpava (uma fila de expressões regulares). <b>Todo defeito
        ''' desta família foi um desacordo entre os dois.</b> Cinco passadas de
        ''' revisão externa acharam contraexemplo, e cada conserto alinhava um
        ''' caso e deixava um irmão:
        '''
        ''' <list type="bullet">
        ''' <item><c>&lt;/script&gt;</c> exato, depois com espaço, depois com
        ''' lixo — três versões, e a contagem aceitava o que o removedor não
        ''' removia: o segredo saía como mensagem.</item>
        ''' <item>Fechamento falso em comentário ou atributo equilibrava a
        ''' conta.</item>
        ''' <item>O autômato comia texto <b>visível</b> entre marcadores falsos,
        ''' e recusava um <c>&lt;</c> dentro de string JavaScript.</item>
        ''' </list>
        '''
        ''' <b>Um leitor só não tem com quem discordar.</b> Ele percorre o texto
        ''' como um tokenizador: dado, tag, atributo com e sem aspas, comentário,
        ''' e o <i>texto cru</i> de <c>script</c> e <c>style</c>. O que ele emite
        ''' é o que sai; o que ele pula é o que o navegador também não mostra.
        '''
        ''' <b>E ele acerta os casos que derrubavam a aproximação</b>, sem regra
        ''' especial: <c>a &lt; b</c> é texto (um <c>&lt;</c> só abre tag antes
        ''' de letra ASCII, <c>/</c> ou <c>!</c>); <c>&lt;é&gt;</c> é texto;
        ''' <c>&lt;script-note&gt;</c> não é <c>script</c>; e
        ''' <c>&lt;/scripture&gt;</c> não fecha <c>&lt;script&gt;</c>, porque
        ''' fechamento exige o nome inteiro seguido de espaço, <c>/</c> ou
        ''' <c>&gt;</c>.
        '''
        ''' <b>Recusa em dois casos, e os dois são "não sei":</b> o
        ''' <i>truncado</i> — acabar dentro de uma tag, de um comentário ou de
        ''' um bloco cru — e o <i>não modelado</i>. Esta frase já disse "o único
        ''' desfecho que recusa é o truncado", e ficou falsa no mesmo commit em
        ''' que a lista de não modelados nasceu.
        ''' </summary>
        ''' <summary>
        ''' <b>Dá para ler este HTML?</b> É o mesmo percurso que produz o texto.
        '''
        ''' Recusa em dois casos, e os dois são "não sei":
        ''' <b>truncado</b> — acaba dentro de uma construção — e
        ''' <b>não modelado</b> — usa uma parte do HTML que este leitor não
        ''' implementa.
        ''' </summary>
        Private Shared Function DaParaLer(bruto As String) As Boolean
            Dim incerto = False
            LerHtml(bruto, incerto)
            Return Not incerto
        End Function

        ''' <summary>
        ''' <b>Lê o HTML uma vez e devolve o texto — e diz quando não sabe.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>CONGELADO EM 29/08/2026, E SEM CONSUMIDOR</b>
        '''
        ''' Isto não roda. A captura lê <c>mail.Body</c> — texto puro — e monta
        ''' o snapshot com <c>ehHtml:=False</c>, então nada daqui alcança a tela
        ''' nem o provedor.
        '''
        ''' Sete passadas de revisão externa seguidas foram neste arquivo, e
        ''' cada uma achou mais um estado do HTML. É o esperado: transformar
        ''' HTML em texto sem escrever um parser não fecha por remendo. Parar
        ''' foi decisão, para não endurecer código morto.
        '''
        ''' <b>Quem for ligar o caminho HTML</b> — passar <c>ehHtml:=True</c>,
        ''' ou capturar <c>HTMLBody</c> — <b>precisa de uma revisão nova deste
        ''' arquivo antes do primeiro byte sair</b>. Ele está bom o bastante
        ''' para ficar guardado, e não para ser confiado sem outra passada. O
        ''' risco residual conhecido está no ESCOPO.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>PRIMEIRO: POR QUE UM LEITOR SÓ</b>
        '''
        ''' Antes havia <i>dois</i> códigos: um decidia se o HTML era
        ''' interpretável e outro limpava, por expressões regulares. Todo defeito
        ''' desta família foi um desacordo entre os dois, e foram cinco passadas
        ''' de revisão externa alinhando um caso e deixando um irmão. Um leitor
        ''' só não tem com quem discordar.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>SEGUNDO, E MAIS IMPORTANTE: ELE NÃO É UM PARSER DE HTML</b>
        '''
        ''' A primeira versão deste leitor foi escrita como se fosse, e a revisão
        ''' seguinte achou cinco estados do HTML real que faltavam —
        ''' <i>double-escaped script</i>, RCDATA, fechamento abrupto de
        ''' comentário, espaço não-ASCII no fechamento, estados de atributo. Isso
        ''' não acaba: HTML é maior do que cabe aqui, e cada passada acharia mais.
        '''
        ''' <b>Então ele modela um subconjunto e RECUSA o resto.</b> É a mesma
        ''' regra que o projeto aplica em todo lugar: recusa declarada é mais
        ''' forte que conversão pela metade. O que ele modela:
        '''
        ''' <list type="bullet">
        ''' <item>texto, e um <c>&lt;</c> que não abre tag é texto;</item>
        ''' <item>tag, com atributo com aspas, sem aspas, e sem valor;</item>
        ''' <item>comentário, incluindo os fechamentos <c>--&gt;</c>,
        ''' <c>--!&gt;</c> e os abruptos <c>&lt;!--&gt;</c> e
        ''' <c>&lt;!---&gt;</c>;</item>
        ''' <item>o texto cru de <c>script</c> e <c>style</c>.</item>
        ''' </list>
        '''
        ''' O que ele <b>recusa por não saber</b> está em
        ''' <see cref="NaoModelados"/>, no comentário condicional nas duas
        ''' formas, no nome de tag com sufixo estranho, no <c>=</c> antes do
        ''' nome do atributo — e no <c>script</c> que contém <c>&lt;!--</c>, que
        ''' é a porta do estado <i>double-escaped</i>, onde um
        ''' <c>&lt;/script&gt;</c> não fecha nada.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>TERCEIRO: O QUE ELE NÃO PROMETE, E EU PROMETI POR DUAS PASSADAS</b>
        '''
        ''' Eu escrevi, no commit e no relatório, que a propriedade era <i>"o
        ''' texto visível sai inteiro, e o invisível não sai"</i>. <b>Isso é
        ''' falso, e não dá para consertar por enumeração.</b> Visibilidade é
        ''' renderização:
        '''
        ''' <code>&lt;style&gt;.x{display:none}&lt;/style&gt;&lt;p class=x&gt;SEGREDO&lt;/p&gt;</code>
        '''
        ''' é HTML perfeitamente comum, o leitor o aceita, e o <c>SEGREDO</c>
        ''' sai — sem que ninguém o tenha visto na tela. Para prometer o que eu
        ''' prometi seria preciso aplicar CSS, e isso é um navegador.
        '''
        ''' <b>O que ele entrega é o TEXTO DO DOCUMENTO</b>: o texto que está
        ''' escrito no HTML, menos o conteúdo das construções que não são texto
        ''' — <c>script</c>, <c>style</c> e comentário. Nada sobre o que a tela
        ''' mostra.
        '''
        ''' Isso deixa um risco residual real, e ele está declarado no ESCOPO:
        ''' texto escondido por CSS pode ir para o provedor. Declarar é o que dá
        ''' para fazer aqui; consertar seria outro projeto.
        ''' </summary>
        Friend Shared Function LerHtml(bruto As String, ByRef incerto As Boolean) As String
            incerto = False
            Dim texto = If(bruto, "")
            Dim sb As New StringBuilder(texto.Length)
            Dim i As Integer = 0

            While i < texto.Length
                Dim c = texto(i)

                If c <> "<"c Then
                    sb.Append(c)
                    i += 1
                    Continue While
                End If

                If Casa(texto, i, "<!--") Then
                    Dim depoisDoComentario = PularComentario(texto, i, incerto)
                    If incerto Then Return Fechar(sb)
                    i = depoisDoComentario
                    Continue While
                End If

                ' Declaracao, instrucao e fechamento sem nome: o navegador
                ' trata como comentario torto e nao mostra.
                ' CONDICIONAL REVELADO: "<![if !mso]>...<![endif]>". A mesma
                ' ambiguidade do "<!--[if mso]>" -- o conteudo aparece num
                ' leitor e some no outro -- e eu tinha recusado so a outra
                ' forma.
                If Casa(texto, i, "<![") Then
                    incerto = True
                    Return Fechar(sb)
                End If

                If Casa(texto, i, "<!") OrElse Casa(texto, i, "<?") OrElse
                   (Casa(texto, i, "</") AndAlso Not ComecaNome(texto, i + 2)) Then
                    ' ASPAS PROTEGEM SO NO DOCTYPE.
                    '
                    ' La um ">" pode morar dentro do identificador publico --
                    ' <!DOCTYPE html PUBLIC "A>B"> -- e parar no primeiro ">"
                    ' fazia o resto sair como texto. Mas "<!x" e "<?x" sao
                    ' comentario torto para o HTML, e ali a aspa NAO protege:
                    ' tratar igual fazia o leitor engolir texto de verdade.
                    ' Sem caixa: "<!doctype html PUBLIC ""A>B"">" e valido, e a
                    ' comparacao ordinal deixava a aspa desprotegida nele.
                    Dim fim = FimDaDeclaracao(
                        texto, i,
                        i + 9 <= texto.Length AndAlso
                        String.Compare(texto, i, "<!DOCTYPE", 0, 9,
                                       StringComparison.OrdinalIgnoreCase) = 0)
                    If fim < 0 Then
                        incerto = True
                        Return Fechar(sb)
                    End If
                    i = fim
                    Continue While
                End If

                ' "<" QUE NAO ABRE TAG E TEXTO -- e o navegador faz igual.
                Dim ehFechamento = Casa(texto, i, "</")
                Dim inicioDoNome = If(ehFechamento, i + 2, i + 1)
                If Not ComecaNome(texto, inicioDoNome) Then
                    sb.Append(c)
                    i += 1
                    Continue While
                End If

                Dim nome = LerNome(texto, inicioDoNome)

                ' NAO MODELADO: recusa em vez de adivinhar.
                If NaoModelados.Contains(nome) Then
                    incerto = True
                    Return Fechar(sb)
                End If

                ' O NOME TEM DE TERMINAR EM SEPARADOR.
                '
                ' Sem isto, "<script.foo>" era lido como "script" com um
                ' atributo ".foo" -- e o conteudo visivel do elemento
                ' desconhecido script.foo era apagado como se fosse texto cru.
                ' Um prefixo reconhecido nao e o nome.
                Dim depoisDoNome = inicioDoNome + nome.Length
                If Not TerminaNome(texto, depoisDoNome) Then
                    incerto = True
                    Return Fechar(sb)
                End If

                ' -1 = tag que nao fecha (truncado); -2 = estado de atributo
                ' que eu nao modelo. Os dois recusam, e por motivos distintos.
                Dim depois = PularAtributos(texto, depoisDoNome)
                If depois < 0 Then
                    incerto = True
                    Return Fechar(sb)
                End If

                ' TAG NAO INVENTA ESPACO.
                '
                ' Cada tag acrescentava um " ", e isso PARTE palavra:
                ' "co<strong>ntra</strong>to" saia "co ntra to". Nao e
                ' normalizacao de apresentacao -- e texto diferente do que esta
                ' escrito, e a promessa e o texto do documento. Estrutura vira
                ' quebra de linha; o resto nao vira nada.
                If EhQuebra(nome, ehFechamento) Then sb.Append(vbLf)
                i = depois

                If Not ehFechamento AndAlso (nome = "script" OrElse nome = "style") Then
                    i = PularTextoCru(texto, i, nome, incerto)
                    If incerto Then Return Fechar(sb)
                End If
            End While

            Return Fechar(sb)
        End Function

        ''' <summary>
        ''' Partes do HTML que este leitor <b>não</b> modela, e por isso recusa.
        '''
        ''' <c>textarea</c> e <c>title</c> são RCDATA: o conteúdo é texto
        ''' literal, então tratá-lo como marcação <i>apaga</i> o que o usuário vê.
        ''' <c>iframe</c>, <c>noscript</c>, <c>noembed</c>, <c>noframes</c> e
        ''' <c>template</c> guardam conteúdo que o navegador em geral <b>não</b>
        ''' mostra — tratá-lo como texto o faz <i>vazar</i>. <c>xmp</c> e
        ''' <c>plaintext</c> mudam a leitura de tudo o que vem depois.
        ''' <c>svg</c> e <c>math</c> são <i>conteúdo estrangeiro</i>: dentro
        ''' deles o HTML muda de regras — <c>CDATA</c> passa a valer, e o texto
        ''' de um <c>&lt;text&gt;</c> é desenhado na tela.
        '''
        ''' Os dois erros já aconteceram aqui, em versões anteriores. A lista é
        ''' curta de propósito: é mais honesto recusar uma mensagem rara do que
        ''' entregar uma mensagem errada.
        ''' </summary>
        Private Shared ReadOnly NaoModelados As New HashSet(Of String)(
            {"textarea", "title", "iframe", "noscript", "noembed", "noframes",
             "xmp", "plaintext", "template", "svg", "math"}, StringComparer.Ordinal)

        ''' <summary>
        ''' Do <c>&lt;!--</c> até depois do fechamento. Conhece os quatro
        ''' fechamentos do HTML, e <b>recusa o comentário condicional</b>.
        '''
        ''' O condicional (<c>&lt;!--[if mso]&gt;</c>) é o caso em que o
        ''' conteúdo do comentário <i>é</i> visível — justamente no Outlook. Um
        ''' leitor que o apaga perde texto que o usuário viu; um que o mostra
        ''' inventa texto onde o navegador não mostra nada. Não dá para acertar
        ''' os dois sem saber quem vai ler, então recusa.
        ''' </summary>
        Private Shared Function PularComentario(texto As String, i As Integer,
                                                ByRef incerto As Boolean) As Integer
            ' Fechamento abrupto: <!--> e <!--->
            If Casa(texto, i, "<!-->") Then Return i + 5
            If Casa(texto, i, "<!--->") Then Return i + 6

            If Casa(texto, i, "<!--[") Then
                incerto = True
                Return -1
            End If

            Dim a = texto.IndexOf("-->", i + 4, StringComparison.Ordinal)
            Dim b = texto.IndexOf("--!>", i + 4, StringComparison.Ordinal)
            If a < 0 AndAlso b < 0 Then
                incerto = True
                Return -1
            End If
            If a < 0 OrElse (b >= 0 AndAlso b < a) Then Return b + 4
            Return a + 3
        End Function

        ''' <summary>
        ''' Do fim da tag de abertura até depois da tag de fechamento do bloco
        ''' cru, sem emitir nada — o conteúdo não é texto.
        '''
        ''' <b>Recusa se houver <c>&lt;!--</c> dentro</b>: é a porta do estado
        ''' <i>double-escaped</i> do HTML, onde <c>&lt;/script&gt;</c> deixa de
        ''' fechar o bloco. Modelar aquilo é reimplementar o tokenizador; não
        ''' modelar e seguir em frente entrega como texto o que estava escondido.
        ''' </summary>
        Private Shared Function PularTextoCru(texto As String, i As Integer,
                                              nome As String,
                                              ByRef incerto As Boolean) As Integer
            Dim fim = AcharFechamento(texto, i, nome)
            If fim < 0 Then
                incerto = True
                Return -1
            End If

            ' SO NO SCRIPT. O estado double-escaped nao existe em style: la
            ' "<!--" e texto cru como qualquer outro, e "<style><!-- .x{} --></style>"
            ' e CSS comum e antigo. Eu aplicava aos dois e recusava e-mail
            ' legitimo -- o comentario ao lado dizia "script" e o codigo nao.
            If nome = "script" AndAlso
               texto.IndexOf("<!--", i, fim - i, StringComparison.Ordinal) >= 0 Then
                incerto = True
                Return -1
            End If

            ' Consome tambem a tag de fechamento, para o bloco inteiro custar
            ' UM espaco -- como custava quando era uma expressao regular so.
            Dim depois = PularAtributos(texto, fim + 2 + nome.Length)
            If depois < 0 Then
                incerto = True
                Return -1
            End If
            Return depois
        End Function

        ''' <summary>
        ''' Fim de uma declaração, respeitando aspas. Devolve a posição
        ''' <b>depois</b> do <c>&gt;</c>, ou <c>-1</c> se ela não fecha.
        ''' </summary>
        Private Shared Function FimDaDeclaracao(texto As String, i As Integer,
                                                ehDoctype As Boolean) As Integer
            Dim aspa As Char = ChrW(0)
            Dim j = i
            While j < texto.Length
                Dim c = texto(j)
                If aspa <> ChrW(0) Then
                    If c = aspa Then aspa = ChrW(0)
                ElseIf ehDoctype AndAlso (c = ChrW(34) OrElse c = ChrW(39)) Then
                    aspa = c
                ElseIf c = ">"c Then
                    Return j + 1
                End If
                j += 1
            End While
            Return -1
        End Function

        ''' <summary>
        ''' <b>Há <c>cid:</c> ou <c>data:</c> aqui — inclusive escondido?</b>
        '''
        ''' Olha o texto como está <b>e</b> como um leitor de URL o veria: sem
        ''' entidade e sem espaço em branco. O parser de URL do navegador
        ''' descarta quebra de linha e tabulação <i>antes</i> de ler o esquema,
        ''' então <c>da&amp;#10;ta:,SEGREDO</c> é uma data URI para ele e não
        ''' era para nós.
        '''
        ''' Depois disso o atributo some na leitura, e some junto a evidência —
        ''' por isso a conferência tem de acontecer aqui, no bruto.
        ''' </summary>
        Friend Shared Function TemReferenciaEmbutida(bruto As String) As Boolean
            If String.IsNullOrEmpty(bruto) Then Return False
            If Embutido.IsMatch(bruto) Then Return True

            Dim decodificado = Net.WebUtility.HtmlDecode(bruto)
            Dim sb As New StringBuilder(decodificado.Length)
            ' SO CONTROLE, e NAO espaco em branco.
            '
            ' O parser de URL descarta tabulacao e quebra de linha, e nao o
            ' espaco. Tirar espaco tambem fazia "Data: 12/03/2026, as 10h"
            ' virar "Data:12/03/2026,as10h" -- uma data URI perfeita, e uma
            ' recusa de texto comum em portugues. Foi o teste do outro lado
            ' que pegou, no mesmo minuto.
            For Each c In decodificado
                If Not Char.IsControl(c) Then sb.Append(c)
            Next
            Return Embutido.IsMatch(sb.ToString())
        End Function

        Private Shared Function Fechar(sb As StringBuilder) As String
            Return Limpar(Net.WebUtility.HtmlDecode(LegadoC1(sb.ToString())))
        End Function

        ''' <summary>
        ''' As referências numéricas de <c>&amp;#128;</c> a <c>&amp;#159;</c>,
        ''' que o HTML manda ler como windows-1252.
        '''
        ''' <c>&amp;#128;</c> é o símbolo do euro no HTML, e o
        ''' <c>WebUtility.HtmlDecode</c> devolve <c>U+0080</c> — um controle,
        ''' que o <see cref="Limpar"/> apaga. Resultado: <c>A&amp;#128;B</c>
        ''' virava <c>AB</c>, e o euro sumia de uma cotação. É texto do
        ''' documento desaparecendo, que é metade do que este leitor existe
        ''' para não fazer.
        '''
        ''' A tabela é a do padrão, e cabe em vinte e sete linhas.
        ''' </summary>
        Private Shared ReadOnly C1 As Char() = {
            ChrW(&H20AC), ChrW(&H81), ChrW(&H201A), ChrW(&H192),
            ChrW(&H201E), ChrW(&H2026), ChrW(&H2020), ChrW(&H2021),
            ChrW(&H2C6), ChrW(&H2030), ChrW(&H160), ChrW(&H2039),
            ChrW(&H152), ChrW(&H8D), ChrW(&H17D), ChrW(&H8F),
            ChrW(&H90), ChrW(&H2018), ChrW(&H2019), ChrW(&H201C),
            ChrW(&H201D), ChrW(&H2022), ChrW(&H2013), ChrW(&H2014),
            ChrW(&H2DC), ChrW(&H2122), ChrW(&H161), ChrW(&H203A),
            ChrW(&H153), ChrW(&H9D), ChrW(&H17E), ChrW(&H178)}

        ''' <summary>
        ''' Qualquer referência numérica, decimal ou hexadecimal, com zeros à
        ''' esquerda. A faixa é conferida <b>depois</b>, no código.
        '''
        ''' A primeira versão embutia a faixa no padrão — <c>x0*8([0-9a-f])</c>
        ''' — e com isso pegava só <c>80</c>–<c>8F</c>: <c>&amp;#x91;</c>
        ''' continuava perdendo a aspa curva. Faixa em expressão regular é
        ''' fácil de escrever pela metade.
        ''' </summary>
        ' O ";" E OPCIONAL, e o HTML aceita assim: "A&#128B" e "A€B". Sem
        ' isto o euro continuava sumindo -- so que num caso mais raro.
        Private Shared ReadOnly RefNumerica As New Regex("&#(?:[xX]([0-9a-fA-F]+)|([0-9]+));?")

        Private Shared Function LegadoC1(bruto As String) As String
            If bruto.IndexOf("&#", StringComparison.Ordinal) < 0 Then Return bruto

            Return RefNumerica.Replace(bruto,
                Function(m)
                    Dim n As Integer
                    Try
                        If m.Groups(1).Success Then
                            n = Convert.ToInt32(m.Groups(1).Value, 16)
                        Else
                            n = Integer.Parse(m.Groups(2).Value, CultureInfo.InvariantCulture)
                        End If
                    Catch
                        ' Numero grande demais para caber: nao e C1, e o
                        ' HtmlDecode resolve o que der.
                        Return m.Value
                    End Try
                    If n < &H80 OrElse n > &H9F Then Return m.Value
                    Return C1(n - &H80)
                End Function)
        End Function

        Private Shared Function Casa(texto As String, i As Integer, agulha As String) As Boolean
            Return i + agulha.Length <= texto.Length AndAlso
                   String.CompareOrdinal(texto, i, agulha, 0, agulha.Length) = 0
        End Function

        ''' <summary>
        ''' Nome de tag começa com letra <b>ASCII</b>. O <c>Char.IsLetter</c>
        ''' aceitava <c>&lt;é&gt;</c> como tag, e a limpeza comia o que vinha
        ''' junto.
        ''' </summary>
        Private Shared Function ComecaNome(texto As String, i As Integer) As Boolean
            If i >= texto.Length Then Return False
            Dim c = texto(i)
            Return (c >= "a"c AndAlso c <= "z"c) OrElse (c >= "A"c AndAlso c <= "Z"c)
        End Function

        ''' <summary>
        ''' O nome em minúsculas. Inclui <c>-</c> e dígito porque
        ''' <c>&lt;script-note&gt;</c> é um elemento válido, e a versão que lia
        ''' só <c>script</c> nele apagava o conteúdo visível inteiro.
        '''
        ''' E inclui <c>:</c> por um motivo prático: o HTML que o Outlook gera é
        ''' cheio de <c>&lt;o:p&gt;</c> e <c>&lt;v:shape&gt;</c>. Sem isso, a
        ''' regra do separador recusaria <b>quase toda mensagem vinda do
        ''' Outlook</b> — e uma fronteira que recusa o caso comum não é
        ''' conservadora, é inútil.
        ''' </summary>
        ''' <summary>
        ''' Depois do nome tem de vir separador — espaço ASCII, <c>/</c> ou
        ''' <c>&gt;</c>. Qualquer outra coisa quer dizer que o nome não é o que
        ''' eu li, e aí eu não sei que elemento é este.
        ''' </summary>
        Private Shared Function TerminaNome(texto As String, i As Integer) As Boolean
            If i >= texto.Length Then Return False
            Dim c = texto(i)
            Return EspacoAscii(c) OrElse c = "/"c OrElse c = ">"c
        End Function

        Private Shared Function LerNome(texto As String, i As Integer) As String
            Dim sb As New StringBuilder()
            Dim j = i
            While j < texto.Length
                Dim c = texto(j)
                If (c >= "a"c AndAlso c <= "z"c) OrElse (c >= "A"c AndAlso c <= "Z"c) OrElse
                   (c >= "0"c AndAlso c <= "9"c) OrElse c = "-"c OrElse c = ":"c Then
                    sb.Append(Char.ToLowerInvariant(c))
                    j += 1
                Else
                    Exit While
                End If
            End While
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Anda até depois do <c>&gt;</c> da tag. Devolve <c>-1</c> se ela não
        ''' fecha até o fim do texto.
        '''
        ''' <b>Aspa só delimita valor depois do <c>=</c>.</b> A primeira versão
        ''' abria aspas onde quer que elas aparecessem, e com isso
        ''' <c>&lt;p a=x"&gt;VISIVEL"&lt;span&gt;</c> engolia o
        ''' <c>&lt;span&gt;</c> inteiro como se fosse parte da tag — perdendo
        ''' texto que o navegador mostra.
        ''' </summary>
        Private Shared Function PularAtributos(texto As String, i As Integer) As Integer
            Const EsperandoNome = 0
            Const NoNome = 1
            Const AntesDoValor = 2
            Const ValorComAspas = 3
            Const ValorSemAspas = 4

            Dim estado = EsperandoNome
            Dim aspa As Char = ChrW(0)
            Dim j = i

            While j < texto.Length
                Dim c = texto(j)

                Select Case estado
                    Case EsperandoNome
                        ' "=" AQUI NAO ABRE VALOR. No HTML ele comeca um
                        ' atributo cujo NOME comeca com "=", e a aspa seguinte
                        ' nao delimita nada -- entao "<p =">VISIVEL</p>" mostra
                        ' VISIVEL. Eu lia como valor com aspas e engolia a
                        ' tag seguinte inteira. Nao modelo isso: recuso.
                        If c = "="c Then Return -2
                        If c = ">"c Then Return j + 1
                        If Not EspacoAscii(c) AndAlso c <> "/"c Then estado = NoNome

                    Case NoNome
                        If c = ">"c Then Return j + 1
                        If c = "="c Then estado = AntesDoValor
                        If EspacoAscii(c) Then estado = EsperandoNome

                    Case AntesDoValor
                        If c = ">"c Then Return j + 1
                        If Not EspacoAscii(c) Then
                            If c = ChrW(34) OrElse c = ChrW(39) Then
                                aspa = c
                                estado = ValorComAspas
                            Else
                                estado = ValorSemAspas
                            End If
                        End If

                    Case ValorComAspas
                        If c = aspa Then estado = EsperandoNome

                    Case Else   ' ValorSemAspas
                        If c = ">"c Then Return j + 1
                        If EspacoAscii(c) Then estado = EsperandoNome
                End Select

                j += 1
            End While

            Return -1
        End Function

        ''' <summary>
        ''' Onde começa o fechamento de um bloco cru.
        '''
        ''' Exige o nome <b>inteiro</b> seguido de espaço ASCII, <c>/</c> ou
        ''' <c>&gt;</c>. Procurar só o prefixo fazia <c>&lt;/scripture&gt;</c>
        ''' passar por fechamento; aceitar <c>Char.IsWhiteSpace</c> fazia um
        ''' espaço U+00A0 fechar o bloco que o navegador deixa aberto — e nos
        ''' dois casos o que estava escondido saía como texto.
        ''' </summary>
        Private Shared Function AcharFechamento(texto As String, i As Integer,
                                                nome As String) As Integer
            Dim alvo = "</" & nome
            Dim j = i
            While True
                j = texto.IndexOf(alvo, j, StringComparison.OrdinalIgnoreCase)
                If j < 0 Then Return -1
                Dim k = j + alvo.Length
                If k >= texto.Length Then Return -1
                Dim c = texto(k)
                If EspacoAscii(c) OrElse c = "/"c OrElse c = ">"c Then Return j
                j = k
            End While
            Return -1
        End Function

        ''' <summary>
        ''' O espaço que o HTML reconhece: tab, LF, FF, CR e espaço. Nada mais —
        ''' um U+00A0 <b>não</b> separa nome de atributo nem fecha tag.
        ''' </summary>
        Private Shared Function EspacoAscii(c As Char) As Boolean
            Return c = " "c OrElse c = ChrW(9) OrElse c = ChrW(10) OrElse
                   c = ChrW(12) OrElse c = ChrW(13)
        End Function

        ''' <summary>
        ''' Elemento de bloco: <b>abrir e fechar</b> quebram linha.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>SÓ O FECHAMENTO NÃO BASTA, PORQUE ELE É OPCIONAL</b>
        '''
        ''' A versão anterior quebrava em alguns fechamentos, e depois que a tag
        ''' deixou de emitir espaço isso colou texto que estava separado:
        ''' <c>&lt;h1&gt;Resumo&lt;/h1&gt;&lt;p&gt;Agora&lt;/p&gt;</c> virava
        ''' <c>ResumoAgora</c>; <c>&lt;li&gt;um&lt;li&gt;dois</c> — com o
        ''' fechamento omitido, que o HTML permite — virava <c>umdois</c>; e
        ''' bloco dentro de bloco também colava.
        '''
        ''' Quebrar nos <b>dois</b> lados resolve os três de uma vez, e o custo é
        ''' uma linha em branco a mais, que o <see cref="LinhasDemais"/> já
        ''' colapsa.
        ''' </summary>
        Private Shared ReadOnly Blocos As New HashSet(Of String)(
            {"p", "div", "tr", "td", "th", "li", "ul", "ol", "dl", "dt", "dd",
             "table", "thead", "tbody", "tfoot", "blockquote", "pre",
             "h1", "h2", "h3", "h4", "h5", "h6", "hgroup",
             "section", "article", "header", "footer", "nav", "aside", "main",
             "figure", "figcaption", "hr", "form", "fieldset", "legend",
             "address", "details", "summary", "dialog", "menu", "search",
             "center"},
            StringComparer.Ordinal)

        Private Shared Function EhQuebra(nome As String, ehFechamento As Boolean) As Boolean
            ' "</br>" nao existe no HTML como fechamento, mas o navegador o
            ' trata como quebra -- entao os dois lados de br tambem.
            If nome = "br" Then Return True
            Return Blocos.Contains(nome)
        End Function
        ''' <summary>
        ''' HTML vira texto — pelo <see cref="LerHtml"/>, que é o mesmo código
        ''' que decidiu se dava para interpretar.
        '''
        ''' Isto era uma fila de quatro expressões regulares, e o decisor era
        ''' outro código. Ver o comentário do <c>LerHtml</c> para os cinco
        ''' desacordos que essa separação produziu.
        ''' </summary>
        Private Shared Function DeHtml(bruto As String) As String
            Dim truncado = False
            Return LerHtml(bruto, truncado)
        End Function

        ''' <summary>
        ''' Quebra de linha determinística, controle fora, espaço aparado.
        '''
        ''' Caractere de controle sai porque não é texto: um <c>U+0000</c> ou um
        ''' marcador de direção no meio de uma frase muda o que se lê sem mudar
        ''' o que está escrito. Tabulação e quebra ficam.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>MAS "CATEGORIA FORMAT" ERA LARGO DEMAIS, E APAGAVA EMOJI</b>
        '''
        ''' A regra derrubava toda a categoria <c>Format</c>, e ali dentro moram
        ''' o ZWJ (<c>U+200D</c>) e o ZWNJ (<c>U+200C</c>) — que <b>ligam</b>
        ''' caracteres em vez de reordená-los. Com isso <c>👩‍💻</c> virava
        ''' <c>👩💻</c>: duas pessoas onde havia uma. E o teste que dizia
        ''' preservar emoji usava um emoji <i>sem</i> junção, então passava.
        '''
        ''' Agora sai só o que <b>reordena</b>: os marcadores bidirecionais.
        ''' </summary>
        ''' <summary>
        ''' Os marcadores que <b>reordenam</b> a leitura — e só eles.
        '''
        ''' <c>U+061C</c> (árabe), <c>U+200E</c>/<c>U+200F</c> (marcas de
        ''' direção), <c>U+202A</c>–<c>U+202E</c> (embutir e sobrepor) e
        ''' <c>U+2066</c>–<c>U+2069</c> (isolar). São eles que fazem uma frase
        ''' ser lida diferente do que está escrita.
        '''
        ''' O ZWJ e o ZWNJ ficam: eles montam emoji e escrita real.
        ''' </summary>
        Private Shared Function Bidirecional(c As Char) As Boolean
            Dim n = AscW(c)
            Return n = &H61C OrElse n = &H200E OrElse n = &H200F OrElse
                   (n >= &H202A AndAlso n <= &H202E) OrElse
                   (n >= &H2066 AndAlso n <= &H2069)
        End Function

        Private Shared Function Limpar(bruto As String) As String
            Dim s = If(bruto, "")
            s = s.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)

            Dim sb As New StringBuilder(s.Length)
            For Each c In s
                If c = vbLf(0) OrElse c = vbTab(0) Then
                    sb.Append(c)
                ElseIf Char.GetUnicodeCategory(c) = UnicodeCategory.Control OrElse
                       Bidirecional(c) Then
                    ' Fora. Inclui os marcadores de direcao, que reordenam a
                    ' leitura sem mudar o texto.
                ElseIf c = " "c Then
                    sb.Append(" "c)
                Else
                    sb.Append(c)
                End If
            Next

            Return LinhasDemais.Replace(sb.ToString(), vbLf & vbLf).Trim()
        End Function

        Private Shared Function Recusar(r As ContentRefusal) As ContentResult
            Return New ContentResult(False, r, Nothing)
        End Function

    End Class

End Namespace

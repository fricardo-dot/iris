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
    End Enum

    Public NotInheritable Class ContentResult
        Public ReadOnly Property Ok As Boolean
        Public ReadOnly Property Recusa As ContentRefusal
        Public ReadOnly Property Parte As MessagePart

        Friend Sub New(ok As Boolean, recusa As ContentRefusal, parte As MessagePart)
            Me.Ok = ok
            Me.Recusa = recusa
            Me.Parte = parte
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

        Private Shared ReadOnly Embutido As New Regex("(cid:|data:[a-z]+/)",
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
        Public Shared Function Preparar(m As Model.MessageSnapshot) As ContentResult
            If m Is Nothing Then Return Recusar(ContentRefusal.SemTexto)

            ' ANEXO PARA AQUI, e antes de tudo. O portao ja nega mensagem com
            ' anexo, mas ele classifica numa visita ao COM e o corpo e lido em
            ' outra — um anexo acrescentado entre as duas passaria. Esta
            ' verificacao esta presa ao MESMO snapshot que virou bytes, e por
            ' isso fecha a corrida em vez de so repeti-la.
            '
            ' Nothing tambem para: "nao deu para contar" nao e "nao tem".
            If m.TemAnexo Is Nothing OrElse m.TemAnexo.Value Then
                Return Recusar(ContentRefusal.Anexo)
            End If

            Return Preparar(m.Item, m.ChangeKey, m.Assunto, m.Remetente,
                            m.Destinatarios, m.Corpo, m.EhHtml, m.CorpoCompleto)
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
                                        corpoCompleto As Boolean) As ContentResult

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
            If Embutido.IsMatch(cru) OrElse
               Embutido.IsMatch(Net.WebUtility.HtmlDecode(cru)) Then
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
                If Embutido.IsMatch(campo) Then Return Recusar(ContentRefusal.ReferenciaEmbutida)
            Next

            If texto.Length = 0 Then Return Recusar(ContentRefusal.SemTexto)

            If tema.Length > MaxAssunto OrElse de.Length > MaxRemetente OrElse
               quem.Count > MaxDestinatarios OrElse texto.Length > MaxCorpo Then
                Return Recusar(ContentRefusal.CampoLongoDemais)
            End If

            If String.IsNullOrEmpty(changeKey) Then
                ' Sem versao nao da para prender o corpo a leitura que o
                ' classificou. O 3.0 mediu PR_CHANGE_KEY vindo em 20 de 20;
                ' faltar aqui e sinal de que alguem montou por outro caminho.
                Return Recusar(ContentRefusal.SemVersao)
            End If

            Return New ContentResult(True, ContentRefusal.Nenhuma,
                                     New MessagePart(item, changeKey, tema, de, quem, texto, True))
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
        ''' <b>O único desfecho que recusa é o TRUNCADO</b> — acabar dentro de
        ''' uma tag, de um comentário ou de um bloco cru. Ali o navegador
        ''' descarta o resto, e nós não temos como saber o que ficou faltando.
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
        ''' <see cref="NaoModelados"/> e no comentário condicional — e o
        ''' <c>script</c> que contém <c>&lt;!--</c>, que é a porta do estado
        ''' <i>double-escaped</i>, onde um <c>&lt;/script&gt;</c> não fecha nada.
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
                    sb.Append(" "c)
                    i = depoisDoComentario
                    Continue While
                End If

                ' Declaracao, instrucao e fechamento sem nome: o navegador
                ' trata como comentario torto e nao mostra.
                If Casa(texto, i, "<!") OrElse Casa(texto, i, "<?") OrElse
                   (Casa(texto, i, "</") AndAlso Not ComecaNome(texto, i + 2)) Then
                    Dim fim = texto.IndexOf(">"c, i)
                    If fim < 0 Then
                        incerto = True
                        Return Fechar(sb)
                    End If
                    sb.Append(" "c)
                    i = fim + 1
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

                Dim depois = PularAtributos(texto, inicioDoNome + nome.Length)
                If depois < 0 Then
                    incerto = True
                    Return Fechar(sb)
                End If

                If EhQuebra(nome, ehFechamento) Then sb.Append(vbLf) Else sb.Append(" "c)
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
        '''
        ''' Os dois erros já aconteceram aqui, em versões anteriores. A lista é
        ''' curta de propósito: é mais honesto recusar uma mensagem rara do que
        ''' entregar uma mensagem errada.
        ''' </summary>
        Private Shared ReadOnly NaoModelados As New HashSet(Of String)(
            {"textarea", "title", "iframe", "noscript", "noembed", "noframes",
             "xmp", "plaintext", "template"}, StringComparer.Ordinal)

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

            If texto.IndexOf("<!--", i, fim - i, StringComparison.Ordinal) >= 0 Then
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

        Private Shared Function Fechar(sb As StringBuilder) As String
            Return Limpar(Net.WebUtility.HtmlDecode(sb.ToString()))
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
        ''' </summary>
        Private Shared Function LerNome(texto As String, i As Integer) As String
            Dim sb As New StringBuilder()
            Dim j = i
            While j < texto.Length
                Dim c = texto(j)
                If (c >= "a"c AndAlso c <= "z"c) OrElse (c >= "A"c AndAlso c <= "Z"c) OrElse
                   (c >= "0"c AndAlso c <= "9"c) OrElse c = "-"c Then
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
            Const ForaDeValor = 0
            Const AntesDoValor = 1
            Const ValorComAspas = 2
            Const ValorSemAspas = 3

            Dim estado = ForaDeValor
            Dim aspa As Char = ChrW(0)
            Dim j = i

            While j < texto.Length
                Dim c = texto(j)

                Select Case estado
                    Case ForaDeValor
                        If c = ">"c Then Return j + 1
                        If c = "="c Then estado = AntesDoValor

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
                        If c = aspa Then estado = ForaDeValor

                    Case Else   ' ValorSemAspas
                        If c = ">"c Then Return j + 1
                        If EspacoAscii(c) Then estado = ForaDeValor
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
        ''' As tags que viram quebra de linha: <c>br</c> abrindo, e <c>p</c>,
        ''' <c>div</c>, <c>tr</c>, <c>li</c> fechando.
        ''' </summary>
        Private Shared Function EhQuebra(nome As String, ehFechamento As Boolean) As Boolean
            If Not ehFechamento Then Return nome = "br"
            Return nome = "p" OrElse nome = "div" OrElse nome = "tr" OrElse nome = "li"
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
        ''' </summary>
        Private Shared Function Limpar(bruto As String) As String
            Dim s = If(bruto, "")
            s = s.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)

            Dim sb As New StringBuilder(s.Length)
            For Each c In s
                If c = vbLf(0) OrElse c = vbTab(0) Then
                    sb.Append(c)
                ElseIf Char.GetUnicodeCategory(c) = UnicodeCategory.Control OrElse
                       Char.GetUnicodeCategory(c) = UnicodeCategory.Format Then
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

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

        Private Shared ReadOnly Comentario As New Regex("<!--.*?-->", RegexOptions.Singleline)
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
        ''' <b>Dá para ler este HTML até o fim?</b>
        '''
        ''' É o mesmo leitor que produz o texto — de propósito. A pergunta
        ''' "posso interpretar" e a resposta "aqui está o texto" saem do
        ''' mesmo percurso, então não têm como discordar; e discordar era a
        ''' raiz de todos os defeitos desta família.
        '''
        ''' Recusa só o <b>truncado</b>: HTML que acaba dentro de uma tag, de
        ''' um comentário ou de um bloco cru. Ali o navegador descarta o resto
        ''' e nós não sabemos o que ficou faltando.
        ''' </summary>
        Private Shared Function DaParaLer(bruto As String) As Boolean
            Dim truncado = False
            LerHtml(bruto, truncado)
            Return Not truncado
        End Function

        Friend Shared Function LerHtml(bruto As String, ByRef truncado As Boolean) As String
            truncado = False
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

                ' COMENTARIO: some inteiro, como no navegador.
                If Casa(texto, i, "<!--") Then
                    Dim fim = texto.IndexOf("-->", i + 4, StringComparison.Ordinal)
                    If fim < 0 Then
                        truncado = True
                        Return Fechar(sb)
                    End If
                    sb.Append(" "c)
                    i = fim + 3
                    Continue While
                End If

                ' Declaracao, instrucao e fechamento sem nome: o navegador
                ' trata como comentario torto e nao mostra. Some ate o ">".
                If Casa(texto, i, "<!") OrElse Casa(texto, i, "<?") OrElse
                   (Casa(texto, i, "</") AndAlso Not ComecaNome(texto, i + 2)) Then
                    Dim fim = texto.IndexOf(">"c, i)
                    If fim < 0 Then
                        truncado = True
                        Return Fechar(sb)
                    End If
                    sb.Append(" "c)
                    i = fim + 1
                    Continue While
                End If

                ' "<" QUE NAO ABRE TAG E TEXTO. E o que o navegador faz, e o
                ' que a versao anterior nao fazia: ela recusava a mensagem.
                Dim ehFechamento = Casa(texto, i, "</")
                Dim inicioDoNome = If(ehFechamento, i + 2, i + 1)
                If Not ComecaNome(texto, inicioDoNome) Then
                    sb.Append(c)
                    i += 1
                    Continue While
                End If

                Dim nome = LerNome(texto, inicioDoNome)
                Dim depois = PularAtributos(texto, inicioDoNome + nome.Length)
                If depois < 0 Then
                    truncado = True
                    Return Fechar(sb)
                End If

                ' Estrutura vira quebra de linha. O resto da tag some.
                If EhQuebra(nome, ehFechamento) Then sb.Append(vbLf) Else sb.Append(" "c)
                i = depois

                ' TEXTO CRU: dentro de script/style nada e marcacao.
                If Not ehFechamento AndAlso (nome = "script" OrElse nome = "style") Then
                    Dim fim = AcharFechamento(texto, i, nome)
                    If fim < 0 Then
                        truncado = True
                        Return Fechar(sb)
                    End If
                    i = fim
                End If
            End While

            Return Fechar(sb)
        End Function

        Private Shared Function Fechar(sb As StringBuilder) As String
            Return Limpar(Net.WebUtility.HtmlDecode(sb.ToString()))
        End Function

        Private Shared Function Casa(texto As String, i As Integer, agulha As String) As Boolean
            Return i + agulha.Length <= texto.Length AndAlso
                   String.CompareOrdinal(texto, i, agulha, 0, agulha.Length) = 0
        End Function

        ''' <summary>
        ''' Nome de tag começa com letra <b>ASCII</b>. O tokenizador do HTML é
        ''' assim, e o <c>Char.IsLetter</c> não era: ele aceitava
        ''' <c>&lt;é&gt;</c> como tag, e a limpeza comia o que vinha junto.
        ''' </summary>
        Private Shared Function ComecaNome(texto As String, i As Integer) As Boolean
            If i >= texto.Length Then Return False
            Dim c = texto(i)
            Return (c >= "a"c AndAlso c <= "z"c) OrElse (c >= "A"c AndAlso c <= "Z"c)
        End Function

        ''' <summary>
        ''' O nome em minúsculas. Inclui <c>-</c> e dígito porque
        ''' <c>&lt;script-note&gt;</c> é um elemento válido, e a versão anterior
        ''' lia só <c>script</c> nele — e apagava o conteúdo visível inteiro.
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
        ''' Anda até depois do <c>&gt;</c> da tag, respeitando aspas. Devolve
        ''' <c>-1</c> se a tag não fecha até o fim do texto — o caso truncado.
        ''' </summary>
        Private Shared Function PularAtributos(texto As String, i As Integer) As Integer
            Dim j = i
            Dim aspa As Char = ChrW(0)
            While j < texto.Length
                Dim c = texto(j)
                If aspa <> ChrW(0) Then
                    If c = aspa Then aspa = ChrW(0)
                ElseIf c = ChrW(34) OrElse c = ChrW(39) Then
                    aspa = c
                ElseIf c = ">"c Then
                    Return j + 1
                End If
                j += 1
            End While
            Return -1
        End Function

        ''' <summary>
        ''' Onde começa o fechamento de um bloco cru.
        '''
        ''' <b>Exige o nome INTEIRO</b> seguido de espaço, <c>/</c> ou
        ''' <c>&gt;</c> — que é a regra do HTML. Procurar só o prefixo
        ''' <c>&lt;/script</c> fazia <c>&lt;/scripture&gt;</c> passar por
        ''' fechamento, e o que vinha depois — invisível no navegador — saía
        ''' como texto.
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
                If Char.IsWhiteSpace(c) OrElse c = "/"c OrElse c = ">"c Then Return j
                j = k
            End While
            Return -1
        End Function

        ''' <summary>
        ''' As tags que viram quebra de linha, como na versão por expressão
        ''' regular: <c>br</c> abrindo, e <c>p</c>, <c>div</c>, <c>tr</c>,
        ''' <c>li</c> fechando.
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

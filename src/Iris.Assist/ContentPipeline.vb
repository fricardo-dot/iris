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
        Private Shared ReadOnly ScriptOuEstilo As New Regex(
            "<(script|style)\b[^>]*>.*?</\1>",
            RegexOptions.Singleline Or RegexOptions.IgnoreCase)
        Private Shared ReadOnly Quebra As New Regex("<(br|/p|/div|/tr|/li)\b[^>]*>",
                                                    RegexOptions.IgnoreCase)
        Private Shared ReadOnly Tag As New Regex("<[^>]*>", RegexOptions.Singleline)
        Private Shared ReadOnly Embutido As New Regex("(cid:|data:[a-z]+/)",
                                                      RegexOptions.IgnoreCase)
        Private Shared ReadOnly LinhasDemais As New Regex("(\r?\n){3,}")

        ''' <summary>
        ''' Prepara uma mensagem. Recusa em vez de "dar um jeito": conteúdo que
        ''' não dá para preparar com segurança é conteúdo que não sai.
        ''' </summary>
        ''' <param name="changeKey">
        ''' A <c>PR_CHANGE_KEY</c> da leitura que produziu este corpo. É ela que
        ''' prende o corpo à versão que o portão classificou — sem isso, corpo
        ''' de uma versão sai autorizado pelo rótulo de outra.
        ''' </param>
        Public Shared Function Preparar(item As Model.ItemKey, changeKey As String,
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
        ''' HTML vira texto. <b>Estrutura vira quebra de linha</b>, o resto some.
        '''
        ''' A ordem importa: comentário e <c>script</c>/<c>style</c> saem
        ''' <b>antes</b> das tags, senão o conteúdo deles vira texto visível —
        ''' e é justamente ali que mora texto que o usuário nunca viu na tela.
        ''' </summary>
        Private Shared Function DeHtml(bruto As String) As String
            Dim s = If(bruto, "")
            s = Comentario.Replace(s, " ")
            s = ScriptOuEstilo.Replace(s, " ")
            s = Quebra.Replace(s, vbLf)
            s = Tag.Replace(s, " ")
            s = Net.WebUtility.HtmlDecode(s)
            Return Limpar(s)
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

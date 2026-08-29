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
        ' O QUE A CONTAGEM ACEITA COMO FECHAMENTO, ESTE PADRAO TEM DE REMOVER.
        '
        ' O HtmlInterpretavel conta "</script" -- sem o ">" -- entao qualquer
        ' coisa que comece assim conta como fechamento e o HTML passa. Este
        ' padrao exigia "</script>" EXATO, e depois "</script\s*>": as duas
        ' versoes deixavam frestas que a contagem aceita e o parser HTML
        ' tambem trata como fechamento -- "</script >", "</script x>",
        ' "</script/>". Em todas elas o bloco NAO era removido, a limpeza
        ' generica comia so as tags, e o conteudo do script ia para o provedor
        ' COMO SE FOSSE A MENSAGEM.
        '
        ' Agora os dois usam o mesmo criterio: "</nome" seguido de qualquer
        ' coisa que nao seja ">", ate o ">".
        Private Shared ReadOnly ScriptOuEstilo As New Regex(
            "<(script|style)\b[^>]*>.*?</\1[^>]*>",
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

            If ehHtml AndAlso Not HtmlInterpretavel(cru) Then
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
        ''' <b>Fechamento que conta é fechamento TERMINADO.</b>
        '''
        ''' Contar a substring <c>"&lt;/script"</c> aceitava
        ''' <c>"&lt;/script"</c> <i>sem o <c>&gt;</c></i> como fechamento: o
        ''' HTML passava como interpretável, o padrão de bloco — que precisa do
        ''' <c>&gt;</c> — não removia nada, e sobrava
        ''' <c>SEGREDO&lt;/script</c> no texto que vai para o provedor.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>CONTAR ERA A ABORDAGEM ERRADA, E EU INSISTI NELA TRÊS VEZES</b>
        '''
        ''' Contar aberturas e fechamentos no HTML <b>bruto</b> aceita
        ''' fechamento falso vindo de comentário ou de atributo:
        ''' <c>&lt;!-- &lt;/script&gt; --&gt;&lt;script&gt;SEGREDO</c> equilibra
        ''' e passa. E a minha correção anterior — contar só fechamento
        ''' terminado — <b>abriu um caso novo</b>: o fechamento dentro do
        ''' comentário equilibrava a abertura real, e
        ''' <c>SEGREDO&lt;/script</c> sobrava. Consertar contagem com contagem
        ''' estava sempre a um contraexemplo de distância.
        '''
        ''' <b>A pergunta certa não é "está balanceado", é "sobrou alguma coisa
        ''' que eu não soube remover".</b> Então o teste é o próprio removedor:
        ''' tira comentário, tira bloco, e se ainda restar <c>&lt;script</c> ou
        ''' <c>&lt;/script</c> em qualquer forma, é porque a limpeza não deu
        ''' conta — e o que ela não removeu viraria texto para o provedor.
        '''
        ''' Isso recusa também um atributo que contenha <c>&lt;/script&gt;</c>,
        ''' que é HTML esquisito e legítimo. É o lado certo de errar: o mínimo
        ''' honesto é <b>recusar</b> o que não dá para interpretar.
        ''' </summary>
        Private Shared Function HtmlInterpretavel(bruto As String) As Boolean
            ' MARCADOR DENTRO DE ATRIBUTO DERRUBA ANTES DE QUALQUER COISA.
            '
            ' O removedor nao sabe que esta dentro de aspas, e isso corta dos
            ' DOIS lados: <p title="<script>">VISIVEL</p><p title="</script>">
            ' faz a regex comer o VISIVEL inteiro -- texto que o usuario VE, e
            ' que sumiria do que vai para o provedor sem ninguem notar. E do
            ' outro lado um </script> em atributo ja servia de fechamento
            ' falso. Nao da para interpretar isso sem um parser; entao recusa.
            If MarcadorForaDeLugar(bruto) Then Return False

            ' NA MESMA ORDEM DA LIMPEZA: comentario primeiro, bloco depois.
            Dim resto = ScriptOuEstilo.Replace(Comentario.Replace(bruto, " "), " ")
            For Each nome In {"script", "style"}
                If resto.IndexOf("<" & nome, StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
                If resto.IndexOf("</" & nome, StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
            Next
            ' O COMENTARIO PELA MESMA REGRA, e ele tinha ficado na contagem.
            '
            ' "<p>VISIVEL</p>--><!-- marcador >SEGREDO" tem um "<!--" e um
            ' "-->", entao a contagem fechava -- so que o "-->" vem ANTES, o
            ' Comentario.Replace nao acha par nenhum, e a limpeza generica
            ' entregava o SEGREDO que o navegador trata como comentario aberto.
            ' Contagem nao distingue ordem; sobra distingue.
            ' SO O "<!--" QUE SOBROU DERRUBA, e o "-->" sozinho nao.
            '
            ' A primeira versao desta regra recusava os dois, e com isso
            ' recusava "Use a seta --> para continuar" -- texto visivel, que a
            ' limpeza generica preserva inteiro, porque "-->" nao casa com o
            ' padrao de tag (nao tem "<"). Falso positivo puro.
            '
            ' O caso perigoso continua pego: "--> ... <!--" deixa o "<!--" na
            ' sobra, porque nao houve par para remover.
            If resto.IndexOf("<!--", StringComparison.Ordinal) >= 0 Then Return False
            Return True
        End Function

        ''' <summary>
        ''' <b>Um <c>&lt;</c> que a limpeza genérica vai comer errado.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>A PRIMEIRA VERSÃO OLHAVA SÓ ATRIBUTO, E ERRAVA DOS DOIS LADOS</b>
        '''
        ''' Ela chamava-se <c>MarcadorDentroDeAtributo</c> e a revisão externa
        ''' achou um contraexemplo em cada direção:
        '''
        ''' <b>Falso negativo</b> — <c>&lt;p&gt;a &lt; b e c &gt; d&lt;/p&gt;</c>
        ''' passava, e o padrão de tag comia <c>&lt; b e c &gt;</c>: sobrava
        ''' <c>a d</c>. E o irmão direto do defeito original, com atributo
        ''' <i>sem aspas</i>: <c>&lt;p title=&lt;script&gt;&gt;VISIVEL</c> fazia o
        ''' removedor comer o VISIVEL.
        '''
        ''' <b>Falso positivo</b> — <c>&lt;script&gt;const lt = "&lt;";&lt;/script&gt;</c>:
        ''' o <c>&lt;</c> de dentro da string JS abria uma tag fictícia, a aspa
        ''' seguinte abria um atributo fictício, e o <c>&lt;/script&gt;</c>
        ''' derrubava. HTML legítimo recusado — e o bloco seria removido
        ''' corretamente pela expressão regular.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>ENTÃO ELE PRECISA SABER TRÊS COISAS, E NÃO UMA</b>
        '''
        ''' <b>1. Onde uma tag começa.</b> Um <c>&lt;</c> só abre marcação se
        ''' vier letra, <c>/</c> ou <c>!</c> depois. Qualquer outro é texto
        ''' solto, e o padrão de tag vai comer dele até o próximo <c>&gt;</c>.
        '''
        ''' <b>2. Que dentro de <c>script</c> e <c>style</c> não há marcação.</b>
        ''' São elementos de texto cru: o conteúdo é pulado até o fechamento,
        ''' que é o que um tokenizador de verdade faz.
        '''
        ''' <b>3. Que aspas protegem, e a falta delas não.</b> Um <c>&lt;</c>
        ''' dentro da tag — com ou sem aspas — é marcador fora de lugar.
        '''
        ''' O que ele NÃO decide fica para a regra da sobra: tag ou bloco que
        ''' não fecham devolvem <c>False</c> aqui e são recusados lá, por não
        ''' terem sido removidos.
        ''' </summary>
        Friend Shared Function MarcadorForaDeLugar(bruto As String) As Boolean
            If String.IsNullOrEmpty(bruto) Then Return False

            Dim i As Integer = 0
            While i < bruto.Length
                If bruto(i) <> "<"c Then
                    i += 1
                    Continue While
                End If

                ' 1. TEXTO SOLTO: o padrao de tag come daqui ate o proximo >.
                If Not AbreMarcacao(bruto, i) Then Return True

                ' Comentario: pula o par inteiro. Sem par, a sobra decide.
                If i + 4 <= bruto.Length AndAlso
                   bruto.Substring(i, 4) = "<!--" Then
                    Dim fim = bruto.IndexOf("-->", i, StringComparison.Ordinal)
                    If fim < 0 Then Return False
                    i = fim + 3
                    Continue While
                End If

                Dim nome = NomeDaTag(bruto, i)
                Dim j As Integer = i + 1
                Dim aspa As Char = ChrW(0)
                Dim fechou = False

                While j < bruto.Length
                    Dim d = bruto(j)
                    If aspa <> ChrW(0) Then
                        If d = aspa Then
                            aspa = ChrW(0)
                        ElseIf d = "<"c Then
                            Return True
                        End If
                    ElseIf d = ChrW(34) OrElse d = ChrW(39) Then
                        aspa = d
                    ElseIf d = "<"c Then
                        ' 3. Atributo SEM aspas com marcador dentro.
                        Return True
                    ElseIf d = ">"c Then
                        fechou = True
                        Exit While
                    End If
                    j += 1
                End While

                If Not fechou Then Return False
                i = j + 1

                ' 2. TEXTO CRU: dentro de script/style nao ha marcacao.
                If nome = "script" OrElse nome = "style" Then
                    Dim fim = bruto.IndexOf("</" & nome, i,
                                            StringComparison.OrdinalIgnoreCase)
                    If fim < 0 Then Return False
                    i = fim
                End If
            End While

            Return False
        End Function

        ''' <summary>
        ''' Um <c>&lt;</c> abre marcação se vier letra, <c>/</c> ou <c>!</c>.
        ''' Qualquer outra coisa é o caractere literal num texto que ninguém
        ''' escapou — e é ele que faz a limpeza genérica comer o que vem junto.
        ''' </summary>
        Private Shared Function AbreMarcacao(bruto As String, i As Integer) As Boolean
            If i + 1 >= bruto.Length Then Return False
            Dim d = bruto(i + 1)
            Return Char.IsLetter(d) OrElse d = "/"c OrElse d = "!"c
        End Function

        ''' <summary>
        ''' O nome da tag em minúsculas, ou vazio para fechamento e declaração.
        ''' Serve só para reconhecer <c>script</c> e <c>style</c>, que carregam
        ''' texto cru.
        ''' </summary>
        Private Shared Function NomeDaTag(bruto As String, i As Integer) As String
            Dim j = i + 1
            Dim sb As New Text.StringBuilder()
            While j < bruto.Length AndAlso Char.IsLetter(bruto(j))
                sb.Append(Char.ToLowerInvariant(bruto(j)))
                j += 1
            End While
            Return sb.ToString()
        End Function

        Private Shared Function Contagem(texto As String, agulha As String) As Integer
            Dim n = 0
            Dim i = texto.IndexOf(agulha, StringComparison.OrdinalIgnoreCase)
            While i >= 0
                n += 1
                i = texto.IndexOf(agulha, i + 1, StringComparison.OrdinalIgnoreCase)
            End While
            Return n
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

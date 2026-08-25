Imports System.Net
Imports System.Net.Http
Imports System.Threading
Imports Iris.Assist

Namespace Global.Iris.Integration.Assist.Http

    ''' <summary>
    ''' <b>O transporte HTTP — genérico, e o único lugar com egress de IA.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ELE NÃO É</b>
    '''
    ''' Não é adaptador de fornecedor. Não sabe o formato de request de
    ''' ninguém, não faz autenticação de ninguém, não interpreta resposta de
    ''' ninguém: manda os bytes do envelope e devolve o texto que voltou.
    '''
    ''' Isso é deliberado, e é a §28.2: o usuário não escolheu provedor, então
    ''' escrever o contrato de uma API específica — autenticação, formato,
    ''' streaming, códigos — seria <b>inventar requisito</b>. O adaptador de
    ''' fornecedor fica pendente, declarado.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>AS REGRAS DA §30, UMA A UMA</b>
    '''
    ''' • <b>Redirects desativados.</b> Um 302 mandaria o corpo para um endereço
    '''   que ninguém autorizou — e a capability se prendeu ao endpoint, não a
    '''   "onde ele apontar".
    ''' • <b>Endpoint fixo</b>, vindo da configuração autorizada. Nunca do
    '''   prompt, nunca da resposta.
    ''' • <b>Só HTTPS.</b> A exceção de loopback existe para o servidor falso
    '''   dos testes e tem de ser ligada explicitamente — a produção não liga.
    ''' • <b>Nenhum retry depois de começar.</b> É a regra "leitura tem retry,
    '''   mutação não" do CLAUDE.md: egress é mutação do mundo, e repetir manda
    '''   o mesmo conteúdo duas vezes.
    ''' • <b>Timeout e cancelamento não são "não chegou".</b> Viram desfecho que
    '''   o diário registra como ambíguo.
    ''' • <b>O corpo do erro não sai daqui.</b> Ele pode ecoar o que foi
    '''   enviado; o que atravessa é o código HTTP.
    ''' • <b>A credencial não aparece em lugar nenhum</b> — nem em log, nem em
    '''   exceção, nem em query string. Ela é lida na hora de montar o
    '''   cabeçalho e não é guardada.
    ''' </summary>
    Public NotInheritable Class HttpAssistantProvider
        Implements IAssistantProvider
        Implements IDisposable

        ''' <summary>Teto da resposta. Passou disso, não se lê o resto.</summary>
        Public Const MaxResposta As Integer = 1024 * 1024

        Private ReadOnly _cliente As HttpClient
        Private ReadOnly _destino As AssistDestination
        Private ReadOnly _credencial As Func(Of String)
        Private ReadOnly _cabecalho As String
        Private ReadOnly _permitirLoopbackSemTls As Boolean

        ''' <param name="credencial">
        ''' Lida <b>na hora</b> de montar o cabeçalho, e não guardada. Um campo
        ''' com a credencial dentro é um campo que vaza em dump, em log de
        ''' exceção e em serialização acidental.
        ''' </param>
        ''' <param name="permitirLoopbackSemTls">
        ''' <b>Só para teste.</b> Existe porque o servidor falso local não tem
        ''' certificado. A produção nunca liga — e há teste provando que o
        ''' padrão recusa <c>http://</c>.
        ''' </param>
        Public Sub New(destino As AssistDestination, credencial As Func(Of String),
                       Optional cabecalho As String = "Authorization",
                       Optional tempoLimite As TimeSpan = Nothing,
                       Optional permitirLoopbackSemTls As Boolean = False)
            _destino = destino
            _credencial = credencial
            _cabecalho = cabecalho
            _permitirLoopbackSemTls = permitirLoopbackSemTls

            Dim h As New HttpClientHandler With {
                .AllowAutoRedirect = False,
                .UseCookies = False,
                .MaxAutomaticRedirections = 1}

            _cliente = New HttpClient(h) With {
                .Timeout = If(tempoLimite = Nothing, TimeSpan.FromSeconds(60), tempoLimite),
                .MaxResponseContentBufferSize = MaxResposta}
        End Sub

        Public ReadOnly Property Destino As AssistDestination _
                                 Implements IAssistantProvider.Destino
            Get
                Return _destino
            End Get
        End Property

        ' ==============================================================

        ''' <summary>
        ''' Endereço parseável, HTTPS e credencial presente — <b>sem tocar na
        ''' rede</b>.
        '''
        ''' São as recusas que se sabem antes de qualquer byte, e é por isso que
        ''' elas não precisam virar ambíguas no diário.
        ''' </summary>
        Public Function Pronto() As Boolean Implements IAssistantProvider.Pronto
            Dim alvo As Uri = Nothing
            If Not Uri.TryCreate(_destino.Endpoint, UriKind.Absolute, alvo) Then Return False
            If Not Seguro(alvo) Then Return False
            Return Not String.IsNullOrEmpty(If(_credencial Is Nothing, Nothing, _credencial()))
        End Function

        Public Function Enviar(bytes As Byte(), ct As CancellationToken) As ProviderOutcome _
                               Implements IAssistantProvider.Enviar

            ' As mesmas conferencias do Pronto(), e nao um atalho confiando
            ' nele: entre uma chamada e outra a credencial pode ter sumido, e
            ' quem transmite nao pode depender de alguem ter perguntado antes.
            Dim alvo As Uri = Nothing
            If Not Uri.TryCreate(_destino.Endpoint, UriKind.Absolute, alvo) Then
                Return New ProviderOutcome(ProviderStatus.NaoComecou, "")
            End If

            If Not Seguro(alvo) Then
                Return New ProviderOutcome(ProviderStatus.NaoComecou, "")
            End If

            Dim chave = If(_credencial Is Nothing, Nothing, _credencial())
            If String.IsNullOrEmpty(chave) Then
                Return New ProviderOutcome(ProviderStatus.NaoComecou, "")
            End If

            Try
                Using pedido As New HttpRequestMessage(HttpMethod.Post, alvo)
                    pedido.Content = New ByteArrayContent(bytes)
                    pedido.Content.Headers.ContentType =
                        New Headers.MediaTypeHeaderValue("application/json") With {
                            .CharSet = "utf-8"}
                    pedido.Headers.TryAddWithoutValidation(_cabecalho, chave)

                    ' UMA chamada. Sem laco, sem politica de retry, sem
                    ' HttpClient com handler que repete: comecou, acabou.
                    Using resposta = _cliente.Send(pedido,
                                                   HttpCompletionOption.ResponseHeadersRead, ct)
                        If Not resposta.IsSuccessStatusCode Then
                            ' O CORPO do erro nao sai daqui: ele pode ecoar o
                            ' que foi enviado. So o codigo atravessa.
                            Return New ProviderOutcome(ProviderStatus.Recusou, "",
                                                       CInt(resposta.StatusCode))
                        End If

                        Return New ProviderOutcome(ProviderStatus.Respondeu,
                                                   LerLimitado(resposta, ct),
                                                   CInt(resposta.StatusCode))
                    End Using
                End Using

            Catch ex As OperationCanceledException When ct.IsCancellationRequested
                ' Cancelado DEPOIS de comecar. Nao quer dizer que nao chegou.
                Return New ProviderOutcome(ProviderStatus.Cancelado, "")
            Catch ex As TaskCanceledException
                Return New ProviderOutcome(ProviderStatus.Timeout, "")
            Catch ex As OperationCanceledException
                Return New ProviderOutcome(ProviderStatus.Timeout, "")
            Catch ex As HttpRequestException
                ' A mensagem NAO atravessa: ela pode carregar host, caminho e,
                ' em alguns casos, pedaco do que foi enviado.
                Return New ProviderOutcome(ProviderStatus.ConexaoCaiu, "")
            End Try
        End Function

        ''' <summary>
        ''' HTTPS, ou loopback com a exceção de teste explicitamente ligada.
        '''
        ''' Sem HTTPS o conteúdo corporativo vai em claro, e nenhuma autorização
        ''' cobre isso — o portão já recusa endpoint não-HTTPS, e aqui é a
        ''' segunda barreira, porque uma só é uma barreira que alguém contorna.
        ''' </summary>
        Private Function Seguro(alvo As Uri) As Boolean
            If String.Equals(alvo.Scheme, Uri.UriSchemeHttps,
                             StringComparison.OrdinalIgnoreCase) Then Return True
            Return _permitirLoopbackSemTls AndAlso alvo.IsLoopback
        End Function

        ''' <summary>
        ''' Lê a resposta até o teto e para. Uma resposta sem fim travaria o
        ''' processo esperando bytes de um servidor que ninguém controla.
        ''' </summary>
        Private Shared Function LerLimitado(r As HttpResponseMessage,
                                            ct As CancellationToken) As String
            Using fonte = r.Content.ReadAsStream(ct)
                Dim buffer(MaxResposta - 1) As Byte
                Dim lidos = 0
                While lidos < buffer.Length
                    Dim n = fonte.Read(buffer, lidos, buffer.Length - lidos)
                    If n = 0 Then Exit While
                    lidos += n
                End While
                Return Text.Encoding.UTF8.GetString(buffer, 0, lidos)
            End Using
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            _cliente.Dispose()
        End Sub

    End Class

End Namespace

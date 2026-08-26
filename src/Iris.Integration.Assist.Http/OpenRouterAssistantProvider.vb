Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports Iris.Assist

Namespace Global.Iris.Integration.Assist.Http

    ''' <summary>
    ''' <b>O OpenRouter — tradução do envelope, e leitura da resposta.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ELE MORA NESTE ASSEMBLY</b>
    '''
    ''' Pedido, autenticação, HTTP e interpretação da resposta são uma
    ''' <b>borda só</b>. Um assembly separado que dependesse deste adquiriria
    ''' egress "de segunda mão" e furaria a regra que o
    ''' <c>EgressArquiteturaTests</c> cobra. Em troca, só o composition root do
    ''' <c>Iris.App</c> pode referenciar e instanciar isto.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE VAI NO CORPO, E POR QUE NADA MAIS</b>
    '''
    ''' <code>
    ''' { "model": …, "messages": [ { "role": "user", "content": &lt;envelope&gt; } ],
    '''   "temperature": 0, "provider": { … } }
    ''' </code>
    '''
    ''' O envelope inteiro vai <b>verbatim</b> como conteúdo da única mensagem.
    ''' <b>Não há mensagem <c>system</c> separada</b>: o envelope já carrega
    ''' <c>instrucaoDoSistema</c> dentro dele, e foi para isso que ela foi posta
    ''' lá. Criar um papel <c>system</c> com texto que a capability não cobre
    ''' seria acrescentar conteúdo à revelia da autorização.
    '''
    ''' Nada mais entra: nem prompt extra, nem metadado, nem nome de pasta, nem
    ''' identificador de usuário ou de sessão.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O BLOCO <c>provider</c> É O QUE TORNA A ATIVAÇÃO VERDADEIRA</b>
    '''
    ''' Gateway <b>roteia</b>: o endpoint autorizado é o dele, e quem processa o
    ''' conteúdo pode ser outro, com retenção e jurisdição próprias. Sem este
    ''' bloco, a ativação diria "retenção zero exigida" e o pedido não exigiria
    ''' nada — uma frase bonita, que é pior que silêncio porque alguém confia
    ''' nela.
    '''
    ''' <list type="bullet">
    ''' <item><c>zdr</c> e <c>data_collection</c> vêm de
    ''' <see cref="ActivationRecord.ExigirRetencaoZero"/>;</item>
    ''' <item><c>only</c> vem de
    ''' <see cref="ActivationRecord.ProvedoresPermitidos"/>;</item>
    ''' <item><c>allow_fallbacks</c> é <b>sempre falso</b>. Sem ele o OpenRouter
    ''' pode cair para um provedor fora da lista quando o fixado não atende — e
    ''' aí a autorização declara um conjunto que o pedido não impõe. Falhar é o
    ''' desfecho certo: <b>degradar em silêncio</b> é o que não pode
    ''' acontecer.</item>
    ''' </list>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE FICA PENDENTE, DECLARADO</b>
    '''
    ''' <b>Qual provedor de fato atendeu</b> não é registrado. A resposta traz
    ''' esse dado, e conferi-lo depois do voo não impediria nada — o conteúdo já
    ''' saiu. A imposição está no pedido, com <c>only</c> e
    ''' <c>allow_fallbacks:false</c>; registrar o que atendeu seria a segunda
    ''' prova, e ela exige um campo no diário que hoje não existe.
    ''' </summary>
    Public NotInheritable Class OpenRouterAssistantProvider
        Implements IAssistantProvider

        Private ReadOnly _transporte As HttpAssistantProvider
        Private ReadOnly _modelo As String
        Private ReadOnly _exigirRetencaoZero As Boolean
        Private ReadOnly _provedoresPermitidos As IReadOnlyList(Of String)

        ''' <summary>
        ''' O nome que a ativação tem de declarar para este adaptador servir.
        ''' </summary>
        Public Const Provedor As String = "openrouter"

        ''' <summary>
        ''' <b>Esta ativação é para este adaptador?</b>
        '''
        ''' Existe porque a composição instanciava o OpenRouter para
        ''' <b>qualquer</b> ativação válida, mesmo uma que declarasse outro
        ''' provedor. O resultado seria mandar o protocolo do OpenRouter — e a
        ''' credencial guardada sob o nome dele — para um endereço arbitrário
        ''' que alguém escreveu no arquivo.
        '''
        ''' A conferência mora <b>aqui</b> e não só em quem monta: um adaptador
        ''' que aceita ativação de outro provedor é um adaptador que confia em
        ''' quem o construiu.
        ''' </summary>
        Public Shared Function Atende(ativacao As ActivationRecord) As Boolean
            If ativacao Is Nothing Then Return False
            Return String.Equals(ativacao.Provedor.Trim(), Provedor,
                                 StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <param name="ativacao">
        ''' <b>Tudo o que este adaptador decide vem daqui.</b> Modelo, retenção e
        ''' provedores permitidos são campos da autorização, e não configuração
        ''' do programa — configuração é coisa que alguém muda sem refazer a
        ''' cerimônia.
        ''' </param>
        ''' <param name="credencial">
        ''' Lida <b>na hora</b>, nunca guardada. Ver o construtor do transporte.
        ''' </param>
        Public Sub New(ativacao As ActivationRecord, credencial As Func(Of String),
                       Optional tempoLimite As TimeSpan = Nothing)
            If ativacao Is Nothing Then Throw New ArgumentNullException(NameOf(ativacao))
            If Not Atende(ativacao) Then
                Throw New ArgumentException(
                    "esta ativação não é para o OpenRouter", NameOf(ativacao))
            End If

            ' LISTA VAZIA NAO CHEGA AQUI pelo carregador, que a recusa. Mas o
            ' ActivationRecord e construivel dentro do assembly, e um adaptador
            ' que aceitasse a lista vazia omitiria `only` e deixaria o gateway
            ' rotear para qualquer hospedeiro. Barreira que so existe a montante
            ' e barreira que some quando alguem constroi por outro caminho.
            If ativacao.ProvedoresPermitidos.Count = 0 Then
                Throw New ArgumentException(
                    "a ativação não lista provedor subjacente nenhum", NameOf(ativacao))
            End If

            _modelo = ativacao.Modelo
            _exigirRetencaoZero = ativacao.ExigirRetencaoZero
            _provedoresPermitidos = ativacao.ProvedoresPermitidos

            _transporte = New HttpAssistantProvider(
                New AssistDestination(ativacao.Provedor, ativacao.Endpoint, ativacao.Modelo),
                credencial, "Authorization", tempoLimite)
        End Sub

        Public ReadOnly Property Destino As AssistDestination _
                                 Implements IAssistantProvider.Destino
            Get
                Return _transporte.Destino
            End Get
        End Property

        Public Function Pronto() As Boolean Implements IAssistantProvider.Pronto
            Return _modelo.Trim().Length > 0 AndAlso _transporte.Pronto()
        End Function

        ' ==============================================================

        Public Function Preparar(envelope As Byte()) As Byte() _
                                 Implements IAssistantProvider.Preparar
            If envelope Is Nothing OrElse envelope.Length = 0 Then Return Nothing
            If _modelo.Trim().Length = 0 Then Return Nothing

            ' O envelope e UTF-8 valido por construcao; se nao for, alguem o
            ' produziu por outro caminho e nada deve sair.
            Dim texto As String
            Try
                texto = New UTF8Encoding(False, True).GetString(envelope)
            Catch
                Return Nothing
            End Try

            Using fluxo As New IO.MemoryStream()
                Using w As New Utf8JsonWriter(fluxo, New JsonWriterOptions With {.Indented = False})
                    w.WriteStartObject()
                    w.WriteString("model", _modelo)

                    w.WriteStartArray("messages")
                    w.WriteStartObject()
                    w.WriteString("role", "user")
                    w.WriteString("content", texto)
                    w.WriteEndObject()
                    w.WriteEndArray()

                    ' Zero, e nao o padrao do provedor: resumo de e-mail nao
                    ' ganha nada com variacao, e variacao faz a mesma mensagem
                    ' produzir resumos diferentes em execucoes diferentes.
                    w.WriteNumber("temperature", 0)

                    w.WriteStartObject("provider")
                    ' SEMPRE falso. Ver o doc da classe: cair para fora da lista
                    ' faria a autorizacao declarar o que o pedido nao impoe.
                    w.WriteBoolean("allow_fallbacks", False)

                    If _exigirRetencaoZero Then
                        w.WriteBoolean("zdr", True)
                        w.WriteString("data_collection", "deny")
                    End If

                    If _provedoresPermitidos.Count > 0 Then
                        w.WriteStartArray("only")
                        For Each slug In _provedoresPermitidos
                            w.WriteStringValue(slug)
                        Next
                        w.WriteEndArray()
                    End If
                    w.WriteEndObject()

                    w.WriteEndObject()
                End Using
                Return fluxo.ToArray()
            End Using
        End Function

        ' ==============================================================

        ''' <summary>
        ''' Manda o corpo já preparado, e <b>interpreta a resposta</b>.
        '''
        ''' A interpretação é estrita, e a falha dela é
        ''' <see cref="ProviderStatus.RespostaIlegivel"/> — nunca
        ''' <c>NaoComecou</c>: o conteúdo saiu, e um desfecho que diga o
        ''' contrário estaria mentindo sobre a única coisa que importa.
        ''' </summary>
        Public Function Enviar(bytes As Byte(), ct As CancellationToken) As ProviderOutcome _
                               Implements IAssistantProvider.Enviar

            Dim r = _transporte.Enviar(bytes, ct)
            If r.Status <> ProviderStatus.Respondeu Then Return r

            Dim texto = Extrair(r.Texto)
            If texto Is Nothing Then
                Return New ProviderOutcome(ProviderStatus.RespostaIlegivel, "", r.Codigo)
            End If
            Return New ProviderOutcome(ProviderStatus.Respondeu, texto, r.Codigo)
        End Function

        ''' <summary>
        ''' <c>choices[0].message.content</c>, e só isso.
        '''
        ''' <c>Nothing</c> quer dizer <b>ilegível</b>. Só uma string em
        ''' <c>content</c> conta — inclusive <c>""</c>, que é resposta vazia e
        ''' tem tratamento próprio mais acima. Ausente, <c>null</c>, tipo
        ''' errado, <c>choices</c> vazio ou JSON inválido param aqui.
        '''
        ''' Nada mais da resposta é lido: <c>usage</c>, <c>provider</c>,
        ''' <c>id</c> e o que mais vier são <b>dado de fora</b>, e a §29.5 diz
        ''' onde o dado de fora para.
        ''' </summary>
        Private Shared Function Extrair(bruto As String) As String
            If String.IsNullOrEmpty(bruto) Then Return Nothing

            Dim doc As JsonDocument = Nothing
            Try
                doc = JsonDocument.Parse(bruto, New JsonDocumentOptions With {
                    .AllowTrailingCommas = False,
                    .CommentHandling = JsonCommentHandling.Disallow,
                    .MaxDepth = 32})
            Catch ex As JsonException
                Return Nothing
            End Try

            Try
                Dim raiz = doc.RootElement
                If raiz.ValueKind <> JsonValueKind.Object Then Return Nothing

                Dim escolhas As JsonElement = Nothing
                If Not raiz.TryGetProperty("choices", escolhas) OrElse
                   escolhas.ValueKind <> JsonValueKind.Array OrElse
                   escolhas.GetArrayLength() = 0 Then
                    Return Nothing
                End If

                Dim primeira = escolhas(0)
                If primeira.ValueKind <> JsonValueKind.Object Then Return Nothing

                Dim mensagem As JsonElement = Nothing
                If Not primeira.TryGetProperty("message", mensagem) OrElse
                   mensagem.ValueKind <> JsonValueKind.Object Then
                    Return Nothing
                End If

                Dim conteudo As JsonElement = Nothing
                If Not mensagem.TryGetProperty("content", conteudo) OrElse
                   conteudo.ValueKind <> JsonValueKind.String Then
                    Return Nothing
                End If

                Return conteudo.GetString()
            Finally
                doc.Dispose()
            End Try
        End Function

    End Class

End Namespace

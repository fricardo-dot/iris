Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports Iris.Model
Imports Iris.Sync

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>A fila de respostas, lida do acervo.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTA CLASSE DECIDE, E O QUE ELA NÃO DECIDE</b>
    '''
    ''' Ela decide <b>o que existe</b>: quais pastas foram varridas, quais itens
    ''' ainda estão lá, e como uma linha do cache vira uma mensagem. As regras de
    ''' <i>quem falou por último</i> ficam inteiras em
    ''' <see cref="FilaDeRespostas"/>, que é puro e não sabe o que é um banco.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>SÓ O QUE ESTÁ PRESENTE ENTRA</b>
    '''
    ''' A associação tem quatro estados. <c>Presente</c> entra; os outros três
    ''' não, e cada um por um motivo diferente:
    '''
    ''' <list type="bullet">
    ''' <item><c>AusenteDaPasta</c> — a mensagem saiu. Uma pendência sobre uma
    ''' mensagem que não está mais lá manda o dono abrir o que não existe.</item>
    ''' <item><c>Suspeito</c> — o Iris <b>acha</b> que sumiu. Trazê-la de volta
    ''' seria ressuscitar por dúvida.</item>
    ''' <item><c>NaoVerificado</c> — ninguém olhou. Não é presença.</item>
    ''' </list>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>"VIU OS ENVIADOS" É UMA PERGUNTA COM DUAS PARTES</b>
    '''
    ''' Não basta a pasta existir no acervo: ela precisa ter <b>geração
    ''' publicada</b>. Uma pasta conhecida e nunca varrida é o caso que a fila
    ''' mais precisa distinguir — é ela que faz toda conversa já respondida
    ''' parecer pendente.
    '''
    ''' <b>Frescor fica de fora, e é decisão.</b> "Varrida há seis meses" é pior
    ''' que "nunca varrida"? Não: é melhor, e o que falta aparece como conversa
    ''' velha demais, que a própria fila mostra. Inventar um prazo aqui seria
    ''' escolher um número sem medida — e o dono vê a data da última varredura no
    ''' acervo.
    ''' </summary>
    Public NotInheritable Class FilaDoAcervo

        Private ReadOnly _acervo As AcervoDeTodasAsPastas

        Public Sub New(acervo As AcervoDeTodasAsPastas)
            If acervo Is Nothing Then Throw New ArgumentNullException(NameOf(acervo))
            _acervo = acervo
        End Sub

        ''' <summary>
        ''' Monta a fila com o que o acervo tem publicado.
        ''' </summary>
        ''' <param name="pastaDosEnviados">
        ''' A chave da pasta de itens enviados. <c>Nothing</c> — o Iris não sabe
        ''' qual é — vale como <b>não varrida</b>: sem ela a fila recusa, que é o
        ''' mesmo que ela faz quando a pasta existe e nunca foi lida.
        ''' </param>
        Public Function Montar(eu As MinhasIdentidades,
                               agora As DateTimeOffset,
                               fuso As TimeZoneInfo,
                               pastaDosEnviados As Long?,
                               dispensadas As IEnumerable(Of String),
                               Optional remetentesIgnorados As MinhasIdentidades = Nothing) As ResultadoDaFila

            Dim mensagens As New List(Of MensagemNaFila)()
            Dim viuOsEnviados = False

            For Each pasta In _acervo.Pastas
                Dim manifesto = pasta.Manifesto

                ' PASTA SEM GERAÇÃO PUBLICADA NÃO É PASTA VAZIA. É a mesma regra
                ' da busca no acervo, e aqui ela decide mais: uma pasta de
                ' enviados conhecida e nunca varrida faria toda conversa já
                ' respondida parecer pendente.
                If Not manifesto.GenerationKey.HasValue Then Continue For

                If pastaDosEnviados.HasValue AndAlso pasta.Chave = pastaDosEnviados.Value Then
                    viuOsEnviados = True
                End If

                For Each item In manifesto.Items
                    If item.Presence <> PresenceState.Presente Then Continue For

                    mensagens.Add(New MensagemNaFila(
                        New ItemKey(item.ProviderEntryId, pasta.Store),
                        item.ConversationId,
                        item.Subject,
                        item.SenderName,
                        item.SenderAddress,
                        Instante(item.ReceivedAt)))
                Next
            Next

            Return FilaDeRespostas.Montar(mensagens, eu, agora, fuso,
                                          viuOsEnviados, dispensadas,
                                          remetentesIgnorados)
        End Function

        ''' <summary>
        ''' <b>Qual pasta do acervo é a de itens enviados</b> — pelo nome, que é
        ''' um palpite, e o único disponível aqui.
        '''
        ''' O acervo guarda nome e não papel: <c>FolderContentKind</c> distingue
        ''' correio de agenda, e não "enviados" de "entrada". Errar aqui erra
        ''' para o lado seguro — a fila recusa em vez de mentir —, e o dono
        ''' varre a pasta certa e ela aparece.
        ''' </summary>
        Public Function AcharOsEnviados() As Long?
            For Each pasta In _acervo.Pastas
                For Each nome In {"Itens Enviados", "Sent Items"}
                    If pasta.Nome.StartsWith(nome, StringComparison.OrdinalIgnoreCase) Then
                        Return pasta.Chave
                    End If
                Next
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' A data como o cache a guardou: ISO 8601 com deslocamento.
        '''
        ''' <b>Nada de <c>DateTime.Parse</c> com a cultura da máquina.</b> A
        ''' cadeia foi escrita com <c>ToString("o")</c>, e lê-la com a cultura do
        ''' usuário faz dia virar mês em metade do calendário — sem erro, e com o
        ''' resultado errado só em alguns dias do mês.
        '''
        ''' Ilegível vira <c>Nothing</c>, e a fila conta como "sem data".
        ''' </summary>
        Friend Shared Function Instante(iso As String) As DateTimeOffset?
            If String.IsNullOrWhiteSpace(iso) Then Return Nothing

            Dim quando As DateTimeOffset
            If DateTimeOffset.TryParse(iso, CultureInfo.InvariantCulture,
                                       DateTimeStyles.RoundtripKind, quando) Then
                Return quando
            End If
            Return Nothing
        End Function

    End Class

End Namespace

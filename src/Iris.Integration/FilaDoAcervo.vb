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
    ''' <b>E o FRESCOR entra, mas sem prazo inventado.</b>
    '''
    ''' Eu tinha escrito que "varrida há seis meses" era melhor que "nunca
    ''' varrida", porque o que faltasse apareceria como conversa velha. Era
    ''' falso, e a revisão externa mostrou o contrário: Enviados varrida em 1º,
    ''' pergunta no dia 29, resposta pelo OWA no dia 30, Entrada varrida no dia
    ''' 31 — e a tela dizia "esperando há 2 dias". O que faltava era a resposta;
    ''' o que aparecia era uma pergunta <b>nova</b>, e a idade da linha não
    ''' denunciava a assimetria.
    '''
    ''' A regra não é um prazo: é o próprio instante da varredura dos enviados
    ''' daquela caixa, que é medido. Conversa mais nova que ele não é afirmada.
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
                               dispensadas As IEnumerable(Of String),
                               Optional remetentesIgnorados As MinhasIdentidades = Nothing) As ResultadoDaFila

            Return FilaDeRespostas.Montar(Mensagens(), eu, agora, fuso,
                                          CoberturaDosEnviados(), dispensadas,
                                          remetentesIgnorados)
        End Function

        ''' <summary>
        ''' <b>As mensagens presentes, de todas as pastas publicadas.</b>
        '''
        ''' Saiu de dentro do <see cref="Montar"/> porque a caixa dividida precisa
        ''' exatamente disto e não precisa da fila. Duas listas montadas por dois
        ''' laços iguais divergiriam — e a divergência apareceria como uma mensagem
        ''' que está numa tela e não na outra, sem explicação possível.
        ''' </summary>
        Public Function Mensagens() As IReadOnlyList(Of MensagemNaFila)
            Dim todas As New List(Of MensagemNaFila)()

            For Each pasta In _acervo.Pastas
                Dim manifesto = pasta.Manifesto

                ' PASTA SEM GERACAO PUBLICADA NAO E PASTA VAZIA. E a mesma regra
                ' da busca no acervo.
                If Not manifesto.GenerationKey.HasValue Then Continue For

                For Each item In manifesto.Items
                    If item.Presence <> PresenceState.Presente Then Continue For

                    todas.Add(New MensagemNaFila(
                        New ItemKey(item.ProviderEntryId, pasta.Store),
                        item.ConversationId,
                        item.Subject,
                        item.SenderName,
                        item.SenderAddress,
                        Instante(item.ReceivedAt),
                        pasta.Store))
                Next
            Next

            Return todas
        End Function

        ''' <summary>
        ''' <b>Pastas que encolheram na última varredura.</b>
        '''
        ''' Uma geração publicada marca como <i>suspeito</i> tudo que estava
        ''' presente e não apareceu — e isso é o modelo funcionando, e não um
        ''' defeito: geração publicada autoriza positivos e suspeitas, e é a
        ''' verificação, não ela, que exige cobertura completa.
        '''
        ''' Só que <see cref="Mensagens"/> exclui suspeito, e o Outlook em modo
        ''' cache encolhe a janela sozinho. O efeito é uma fila que esvazia sem
        ''' nada ter sido resolvido, e o dono lendo isso como "acabou".
        '''
        ''' A ressalva da contração já existia — e morava no painel do acervo,
        ''' que é outra tela e, na prática, outro dia. Ela passa a sair também
        ''' onde o estrago aparece. Achado por revisão externa em 01/09/2026.
        ''' </summary>
        Public Function PastasQueEncolheram() As IReadOnlyList(Of String)
            Dim achadas As New List(Of String)()

            For Each pasta In _acervo.Pastas
                If Not pasta.Manifesto.GenerationKey.HasValue Then Continue For
                Dim c = pasta.Manifesto.Contracao
                If c Is Nothing Then Continue For
                If c.Verdict <> ContractionVerdict.Encolheu Then Continue For
                achadas.Add(pasta.Nome)
            Next

            Return achadas
        End Function

        ''' <summary>
        ''' <b>Por caixa, até quando as respostas do dono são conhecidas.</b>
        '''
        ''' É o instante em que a pasta de enviados <i>daquela caixa</i> foi
        ''' publicada. Caixa que não aparece aqui não entra na fila, e conversa
        ''' cuja última mensagem é posterior a este instante não é afirmada.
        '''
        ''' <b>Uma pasta por caixa, e não uma para todas.</b> Antes, uma pasta de
        ''' enviados varrida em qualquer lugar liberava a fila inteira — inclusive
        ''' as caixas cujas respostas ninguém tinha visto, onde toda conversa já
        ''' respondida virava pendência. Achado por revisão externa em
        ''' 31/08/2026.
        ''' </summary>
        ''' <summary>
        ''' <b>Endereços que enviaram de dentro da SUA pasta de Enviados e que não
        ''' estão em <c>identidades.txt</c>.</b>
        '''
        ''' Existe por causa de um erro que não tinha como ser notado. A direção de
        ''' uma conversa sai do remetente contra o conjunto das suas identidades, e
        ''' a proteção de "não sei" só funciona quando o conjunto está <i>vazio</i>.
        ''' Com ele parcialmente preenchido — um alias faltando, um X.500 novo —,
        ''' uma mensagem que <b>você</b> enviou é classificada como
        ''' <c>DoOutro</c> com toda a confiança: a conversa entra na fila como
        ''' possível resposta sua, e agora também pode disparar um rascunho
        ''' automático para algo que você já respondeu.
        '''
        ''' Não dá para adivinhar qual endereço é seu. Dá para notar o sintoma:
        ''' <b>numa pasta de enviados, quem envia é você</b>. Um endereço que
        ''' aparece ali e não está no arquivo é, quase sempre, uma identidade sua
        ''' que falta cadastrar.
        '''
        ''' <b>É indício, e não prova</b> — encaminhamento automático e caixa
        ''' compartilhada põem endereços de terceiros ali. Por isso isto vira uma
        ''' <i>ressalva</i> na tela, e não uma correção automática: o programa
        ''' cadastrando identidade sozinho mudaria quem está na sua fila sem você
        ''' saber. Achado por revisão externa em 01/09/2026.
        ''' </summary>
        Public Function EnderecosSeusQueFaltam(eu As MinhasIdentidades) As IReadOnlyList(Of String)
            If eu Is Nothing Then Return Array.Empty(Of String)()

            Dim achados As New List(Of String)()
            Dim vistos As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each pasta In _acervo.Pastas
                If Not EhDeEnviados(pasta.Nome) Then Continue For
                If Not pasta.Manifesto.GenerationKey.HasValue Then Continue For

                For Each item In pasta.Manifesto.Items
                    If item.Presence <> PresenceState.Presente Then Continue For

                    Dim de = If(item.SenderAddress, "").Trim()
                    If de.Length = 0 Then Continue For
                    If eu.DirecaoDe(de) <> Direcao.DoOutro Then Continue For

                    If vistos.Add(de) Then achados.Add(de)
                Next
            Next

            Return achados
        End Function

        Friend Function CoberturaDosEnviados() As IReadOnlyDictionary(Of String, DateTimeOffset)
            Dim mapa As New Dictionary(Of String, DateTimeOffset)(StringComparer.Ordinal)

            For Each pasta In _acervo.Pastas
                If Not EhDeEnviados(pasta.Nome) Then Continue For
                If Not pasta.Manifesto.GenerationKey.HasValue Then Continue For

                Dim quando = Instante(pasta.Manifesto.PublishedAt)
                If Not quando.HasValue Then Continue For

                ' A MAIS RECENTE MANDA. Uma caixa pode ter mais de uma pasta que
                ' casa pelo nome, e a cobertura e a melhor delas.
                Dim atual As DateTimeOffset
                If Not mapa.TryGetValue(pasta.Store, atual) OrElse quando.Value > atual Then
                    mapa(pasta.Store) = quando.Value
                End If
            Next

            Return mapa
        End Function

        ''' <summary>
        ''' O nome parece de pasta de enviados? É palpite, e o único disponível: o
        ''' acervo guarda nome e não papel — <c>FolderContentKind</c> distingue
        ''' correio de agenda, e não "enviados" de "entrada".
        ''' </summary>
        Friend Shared Function EhDeEnviados(nome As String) As Boolean
            If String.IsNullOrEmpty(nome) Then Return False
            For Each candidato In {"Itens Enviados", "Sent Items"}
                If nome.StartsWith(candidato, StringComparison.OrdinalIgnoreCase) Then Return True
            Next
            Return False
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

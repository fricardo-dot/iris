Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text.Json
Imports System.Linq
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Cache

    ''' <summary>
    ''' <b>Onde os rótulos da IA moram.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O RÓTULO É UMA OBSERVAÇÃO, COMO O METADADO</b>
    '''
    ''' Ele pende da <b>encarnação</b> — a mensagem naquela pasta — e da
    ''' <b>geração</b> em que foi observado. Geração nova invalida a anterior de
    ''' graça, que é o comportamento que o resto do acervo já tem: quem lê pede
    ''' os rótulos da geração publicada, e os antigos ficam no banco sem
    ''' aparecer.
    '''
    ''' <b>Não há justificativa gravada</b>, e não é economia: "por que este
    ''' rótulo" cita o corpo, e o D1 diz que o cache guarda metadado. Rótulo e
    ''' confiança são metadado; a frase que os explica não é.
    '''
    ''' <b>A ativação fica junto.</b> Sem ela, um rótulo gravado sob uma
    ''' autorização vencida seria indistinguível de um recente, e o dono não
    ''' teria como saber sob que regra a classificação dele foi feita.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>GRAVAR É POR LOTE, E O LOTE É ATÔMICO — COM UMA RESSALVA NOMEADA</b>
    '''
    ''' Os rótulos cujas mensagens <b>ainda estão na pasta</b> entram todos ou
    ''' nenhum. Meia gravação deixaria a pasta com uma parte classificada e
    ''' outra não, sem nada dizendo qual é qual.
    '''
    ''' Os que <b>não estão</b> saem do lote <i>antes</i> da gravação, e a conta
    ''' deles volta. O texto anterior dizia "todos ou nenhum" e o código
    ''' descartava os desconhecidos em silêncio: duas coisas diferentes com o
    ''' mesmo nome. Achado por revisão externa em 31/08/2026.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>UMA CHAMADA POR VEZ</b>
    '''
    ''' A conexão é a do <c>CacheDatabase</c>, e <c>SqliteConnection</c> não
    ''' tem contrato de uso simultâneo — o mesmo que o acervo já declara. Quem
    ''' chama serializa; nada aqui protege contra duas gravações ao mesmo
    ''' tempo, e o WAL não resolve isso (ele coordena conexões diferentes, não
    ''' torna uma conexão reentrante).
    ''' </summary>
    Public NotInheritable Class RotulosNoCache

        Private ReadOnly _conn As SqliteConnection

        Public Sub New(db As CacheDatabase)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _conn = db.Connection
        End Sub

        ''' <summary>
        ''' Grava os rótulos de um lote, todos ou nenhum.
        '''
        ''' <paramref name="rotulos"/> vem por <c>provider_entry_id</c>, que é
        ''' como o cache identifica a encarnação dentro da pasta. Item que não
        ''' está na pasta é <b>ignorado em silêncio</b> — ele saiu entre a
        ''' classificação e a gravação, e insistir criaria encarnação para uma
        ''' mensagem que não está mais lá.
        ''' </summary>
        ''' <returns>
        ''' Quantos entraram e quantos ficaram de fora por já não estarem na
        ''' pasta. <c>Gravou = False</c> quando a geração não é a publicada.
        ''' </returns>
        Public Function Gravar(folderKey As Long, geracao As Long,
                               ativacao As String,
                               quando As DateTimeOffset,
                               rotulos As IReadOnlyDictionary(Of String, String),
                               confiancas As IReadOnlyDictionary(Of String, Double?),
                               Optional regras As IReadOnlyDictionary(Of String, IReadOnlyList(Of String)) = Nothing) _
                               As ResultadoDaGravacao

            Dim gravados = 0
            Dim foraDaPasta = 0

            ' A TRANSACAO COMECA ANTES DA CONFERENCIA, e IMEDIATA.
            '
            ' A conferencia da geracao ficava fora da transacao, e entre ela e o
            ' commit cabia uma publicacao de outra conexao: a chamada confirmava
            ' G1, alguem publicava G2, e os rotulos entravam em G1 e ficavam
            ' invisiveis -- com Gravou = True. Era exatamente o cenario que o
            ' comentario abaixo dizia impedir. Achado por revisao externa em
            ' 31/08/2026.
            '
            ' Deferred = False emite BEGIN IMMEDIATE: a trava de escrita e tomada
            ' agora, e nao no primeiro INSERT. Com BEGIN comum, a leitura da
            ' geracao aconteceria sob trava compartilhada e outra conexao ainda
            ' poderia publicar entre ela e a escrita.
            Using tx = _conn.BeginTransaction(Data.IsolationLevel.Serializable, deferred:=False)

                ' SO NA GERACAO PUBLICADA DESTA PASTA.
                '
                ' As duas chaves estrangeiras eram conferidas em separado, e nada
                ' ligava uma a outra: dava para gravar a encarnacao da pasta A na
                ' geracao da pasta B, e o registro ficava estruturalmente valido e
                ' invisivel.
                '
                ' E ISTO VEM ANTES DO LOTE VAZIO. Antes o lote vazio devolvia
                ' Nada(), que diz Gravou = True, sem nem olhar a geracao -- e ai
                ' "gravei zero rotulos na geracao certa" e "recusei porque a
                ' geracao esta velha" saiam com a mesma cara.
                If Not EhAPublicada(tx, folderKey, geracao) Then
                    Return ResultadoDaGravacao.GeracaoErrada()
                End If

                If rotulos Is Nothing OrElse rotulos.Count = 0 Then
                    Return ResultadoDaGravacao.Nada()
                End If

                For Each par In rotulos
                    Dim incarnation = IncarnationDe(tx, folderKey, par.Key)
                    If Not incarnation.HasValue Then
                        foraDaPasta += 1
                        Continue For
                    End If

                    ' CONFIANCA AUSENTE VAI NULA, e nao zero: "o modelo disse zero"
                    ' e "o modelo nao disse" sao coisas diferentes, e um zero no
                    ' lugar da ausencia faria a leitura tratar silencio como
                    ' certeza minima.
                    Dim confianca As Double? = Nothing
                    Dim lida As Double?
                    If confiancas IsNot Nothing AndAlso
                       confiancas.TryGetValue(par.Key, lida) Then confianca = lida

                    ' NULO E "ninguem respondeu sobre as regras nesta mensagem".
                    ' Vetor vazio e "respondeu, e nenhuma casou" -- que e uma
                    ' informacao, e nao a ausencia dela.
                    '
                    ' A DISTINCAO E POR ITEM, e nao pelo lote inteiro. A versao
                    ' anterior punha vetor vazio em todo item ausente do mapa, e
                    ' entao a mensagem cuja resposta sobre regras nao deu para usar
                    ' -- o que um "nao responda as regras do dono" produz -- ficava
                    ' gravada como "perguntei e nao casou nada". Item fora do mapa
                    ' agora e nulo; quem quer dizer "nenhuma casou" poe a lista
                    ' vazia, explicitamente.
                    Dim casadas As String = Nothing
                    Dim minhas As IReadOnlyList(Of String) = Nothing
                    If regras IsNot Nothing AndAlso
                       regras.TryGetValue(par.Key, minhas) AndAlso
                       minhas IsNot Nothing Then
                        casadas = JsonSerializer.Serialize(minhas)
                    End If

                    ' SUBSTITUI, e nao acrescenta: o indice e unico por
                    ' (encarnacao, geracao), e reclassificar a mesma geracao e
                    ' uma correcao, nao um segundo fato.
                    Executar(tx,
                        "INSERT INTO label_observation (incarnation_key, generation_key, " &
                        "  label, confidence, activation_id, observed_at, matched_rules) " &
                        "VALUES ($i,$g,$l,$c,$a,$q,$r) " &
                        "ON CONFLICT (incarnation_key, generation_key) DO UPDATE SET " &
                        "  label = excluded.label, confidence = excluded.confidence, " &
                        "  activation_id = excluded.activation_id, " &
                        "  observed_at = excluded.observed_at, " &
                        "  matched_rules = excluded.matched_rules",
                        ("$i", CObj(incarnation.Value)), ("$g", CObj(geracao)),
                        ("$l", CObj(par.Value)),
                        ("$c", If(confianca.HasValue, CObj(confianca.Value), Nothing)),
                        ("$a", CObj(ativacao)),
                        ("$q", CObj(quando.ToString("o", CultureInfo.InvariantCulture))),
                        ("$r", If(casadas Is Nothing, Nothing, CObj(casadas))))
                    gravados += 1
                Next
                tx.Commit()
            End Using

            Return ResultadoDaGravacao.Feita(gravados, foraDaPasta)
        End Function

        ''' <summary>
        ''' As regras casadas de volta. <b>Texto ilegível vale como vetor vazio</b>,
        ''' e não como nulo: o que está gravado ali é a prova de que a pergunta foi
        ''' feita naquela varredura, e essa parte continua verdadeira mesmo quando
        ''' o resto não puder ser lido. Nulo diria que ninguém perguntou nada.
        ''' </summary>
        Private Shared Function Desserializar(bruto As String) As IReadOnlyList(Of String)
            Try
                Dim lidas = JsonSerializer.Deserialize(Of List(Of String))(bruto)
                If lidas Is Nothing Then Return Array.Empty(Of String)()
                Return lidas.Where(Function(r) Not String.IsNullOrEmpty(r)).ToList()
            Catch ex As JsonException
                Return Array.Empty(Of String)()
            End Try
        End Function

        ''' <summary>Esta geração é a publicada <b>desta</b> pasta?</summary>
        Private Function EhAPublicada(tx As SqliteTransaction,
                                      folderKey As Long, geracao As Long) As Boolean
            Using cmd = _conn.CreateCommand()
                cmd.Transaction = tx
                cmd.CommandText =
                    "SELECT 1 FROM folder " &
                    "WHERE folder_key = $f AND published_generation_key = $g"
                cmd.Parameters.AddWithValue("$f", folderKey)
                cmd.Parameters.AddWithValue("$g", geracao)
                Return cmd.ExecuteScalar() IsNot Nothing
            End Using
        End Function

        ''' <summary>
        ''' Os rótulos <b>da geração publicada</b> de uma pasta, por
        ''' <c>provider_entry_id</c>.
        '''
        ''' Pasta sem geração publicada devolve vazio — e vazio aqui quer dizer
        ''' "não há rótulo publicado", que é diferente de "não há rótulo".
        '''
        ''' <b>Só o que está PRESENTE na pasta.</b> A encarnação continua no banco
        ''' depois de a mensagem sair, e sem esta condição o rótulo dela voltava —
        ''' uma linha de fila sobre uma mensagem que não está mais lá. É a mesma
        ''' regra que o leitor da fila já aplica ao metadado, e faltava aqui.
        ''' Achado por revisão externa em 31/08/2026.
        '''
        ''' <b>E vem inteiro</b>: rótulo, confiança, ativação e quando. A versão
        ''' anterior devolvia só o rótulo, e então a ativação ficava gravada sem
        ''' ninguém conseguir lê-la — guardar um dado que nenhum consumidor
        ''' alcança é o mesmo que não guardar.
        ''' </summary>
        Public Function Publicados(folderKey As Long) _
                        As IReadOnlyDictionary(Of String, RotuloObservado)

            Dim achados As New Dictionary(Of String, RotuloObservado)(StringComparer.Ordinal)

            Using cmd = _conn.CreateCommand()
                cmd.CommandText =
                    "SELECT i.provider_entry_id, l.label, l.confidence, " &
                    "       l.activation_id, l.observed_at, l.matched_rules " &
                    "FROM folder f " &
                    "JOIN incarnation i ON i.folder_key = f.folder_key " &
                    "JOIN association a ON a.item_key = i.item_key " &
                    "  AND a.folder_key = f.folder_key " &
                    "JOIN label_observation l ON l.incarnation_key = i.incarnation_key " &
                    "  AND l.generation_key = f.published_generation_key " &
                    "WHERE f.folder_key = $f AND f.published_generation_key IS NOT NULL " &
                    "  AND a.presence = 'presente'"
                cmd.Parameters.AddWithValue("$f", folderKey)

                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        achados(rd.GetString(0)) = New RotuloObservado(
                            rd.GetString(1),
                            If(rd.IsDBNull(2), CType(Nothing, Double?), rd.GetDouble(2)),
                            rd.GetString(3),
                            rd.GetString(4),
                            If(rd.IsDBNull(5), Nothing, Desserializar(rd.GetString(5))))
                    End While
                End Using
            End Using

            Return achados
        End Function

        ' ==============================================================

        Private Function IncarnationDe(tx As SqliteTransaction, folderKey As Long,
                                       entryId As String) As Long?
            Using cmd = _conn.CreateCommand()
                cmd.Transaction = tx
                cmd.CommandText =
                    "SELECT incarnation_key FROM incarnation " &
                    "WHERE folder_key = $f AND provider_entry_id = $p"
                cmd.Parameters.AddWithValue("$f", folderKey)
                cmd.Parameters.AddWithValue("$p", entryId)

                Dim valor = cmd.ExecuteScalar()
                If valor Is Nothing OrElse valor Is DBNull.Value Then Return Nothing
                Return Convert.ToInt64(valor, CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Sub Executar(tx As SqliteTransaction, sql As String,
                             ParamArray parametros As (Nome As String, Valor As Object)())
            Using cmd = _conn.CreateCommand()
                cmd.Transaction = tx
                cmd.CommandText = sql
                For Each p In parametros
                    cmd.Parameters.AddWithValue(p.Nome, If(p.Valor, DBNull.Value))
                Next
                cmd.ExecuteNonQuery()
            End Using
        End Sub

    End Class

    ''' <summary>
    ''' Um rótulo como ele foi observado: <b>com a idade e a autorização
    ''' junto</b>.
    '''
    ''' <see cref="Confianca"/> é anulável de propósito. "O modelo disse zero" e
    ''' "o modelo não disse" são coisas diferentes, e colapsá-las faria a tela
    ''' tratar silêncio como certeza mínima — que é uma afirmação.
    ''' </summary>
    Public NotInheritable Class RotuloObservado
        Public ReadOnly Property Rotulo As String
        Public ReadOnly Property Confianca As Double?
        ''' <summary>Sob que autorização esta classificação foi feita.</summary>
        Public ReadOnly Property Ativacao As String
        ''' <summary>Quando, em ISO 8601. Texto, como o resto do cache guarda.</summary>
        Public ReadOnly Property Quando As String

        ''' <summary>
        ''' As regras do dono que esta mensagem satisfez, pelo texto delas.
        '''
        ''' <b><c>Nothing</c> quer dizer "não há resposta"</b> — porque não havia
        ''' regra nenhuma naquela varredura, <i>ou</i> porque havia e o que voltou
        ''' sobre esta mensagem não deu para usar. Vetor vazio quer dizer "havia,
        ''' respondeu, e nenhuma casou".
        '''
        ''' A tela precisa da diferença: a segunda é uma resposta, a primeira é a
        ''' ausência dela. O texto anterior só falava do primeiro caso, e depois
        ''' que a distinção passou a ser por item ele ficou estreito — a tela
        ''' podia ler silêncio de um item como ausência global da pergunta.
        ''' </summary>
        Public ReadOnly Property RegrasCasadas As IReadOnlyList(Of String)

        Friend Sub New(rotulo As String, confianca As Double?,
                       ativacao As String, quando As String,
                       regrasCasadas As IReadOnlyList(Of String))
            Me.Rotulo = If(rotulo, "")
            Me.Confianca = confianca
            Me.Ativacao = If(ativacao, "")
            Me.Quando = If(quando, "")
            Me.RegrasCasadas = regrasCasadas
        End Sub
    End Class

    ''' <summary>
    ''' O que aconteceu numa gravação de lote.
    '''
    ''' <see cref="ForaDaPasta"/> não é erro: é mensagem que saiu entre a
    ''' classificação e a gravação. Mas é <b>contado</b>, porque um lote de
    ''' cinquenta que grava dois precisa dizer isso a quem chamou — senão a
    ''' varredura seguinte acha que aquela pasta já foi classificada.
    ''' </summary>
    Public NotInheritable Class ResultadoDaGravacao
        ''' <summary>
        ''' Falso quando a geração pedida não é a publicada da pasta — e aí
        ''' <b>nada</b> foi gravado.
        ''' </summary>
        Public ReadOnly Property Gravou As Boolean
        Public ReadOnly Property Entraram As Integer
        Public ReadOnly Property ForaDaPasta As Integer

        Private Sub New(gravou As Boolean, entraram As Integer, foraDaPasta As Integer)
            Me.Gravou = gravou
            Me.Entraram = entraram
            Me.ForaDaPasta = foraDaPasta
        End Sub

        Friend Shared Function Nada() As ResultadoDaGravacao
            Return New ResultadoDaGravacao(True, 0, 0)
        End Function

        Friend Shared Function GeracaoErrada() As ResultadoDaGravacao
            Return New ResultadoDaGravacao(False, 0, 0)
        End Function

        Friend Shared Function Feita(entraram As Integer,
                                     foraDaPasta As Integer) As ResultadoDaGravacao
            Return New ResultadoDaGravacao(True, entraram, foraDaPasta)
        End Function
    End Class

End Namespace

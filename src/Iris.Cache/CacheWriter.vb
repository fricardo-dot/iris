Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Cache

    ''' <summary>Uma linha lida do provider, já reduzida a metadado (D1).</summary>
    Public NotInheritable Class StagedRow
        Public Property ProviderEntryId As String
        Public Property SearchKey As String
        Public Property InternetMessageId As String
        Public Property Subject As String
        Public Property SenderName As String
        Public Property ReceivedAt As String
        Public Property LastModifiedAt As String
        Public Property SizeBytes As Long?
        Public Property HasAttachments As Boolean?
        Public Property IsUnread As Boolean?
        Public Property MessageClass As String
    End Class

    Public Enum PublishOutcome
        Publicada
        ''' <summary>A época mudou: o universo em que a varredura correu não existe mais.</summary>
        RecusadaPorEpoca
        ''' <summary>Uma tentativa MAIS NOVA já publicou. Esta é velha chegando tarde.</summary>
        RecusadaPorOrdem
        RecusadaPorEstado
    End Enum

    ''' <summary>
    ''' As três primitivas transacionais, e a razão de cada fronteira.
    '''
    ''' O critério 9 do 2.0 pede que gravar as linhas, avançar o checkpoint e
    ''' publicar sobrevivam a morrer no meio. A resposta não é "tentar de
    ''' novo": é escolher as fronteiras de forma que todos os estados
    ''' intermediários possíveis sejam aceitáveis.
    '''
    ''' <b>Linhas e checkpoint vão na MESMA transação.</b> Separados, morrer
    ''' entre os dois produz um de dois males: checkpoint à frente das linhas
    ''' — e aí a retomada PULA mensagens, perda silenciosa, o pior defeito
    ''' possível num cache de correio; ou linhas à frente do checkpoint — e aí
    ''' a retomada relê, o que só é inofensivo porque a gravação é idempotente.
    ''' Juntos, nenhum dos dois acontece.
    '''
    ''' <b>A publicação é uma LINHA, não um evento.</b> Se publicar fosse
    ''' disparar um evento para a UI, morrer depois do commit e antes do
    ''' evento deixaria a geração no banco e a UI sem saber dela — e nada no
    ''' disco registrando essa dívida. Com <c>publication_log</c> gravado na
    ''' mesma transação da geração, os únicos estados possíveis são: nada, ou
    ''' geração com dívida registrada e consultável.
    '''
    ''' <b>O que está provado, e o que não está.</b> Está provado que a dívida
    ''' persiste ao término abrupto do processo e é consultável depois, e o
    ''' <c>PublicationDrain</c> do 2.2a é o mecanismo que a consome. O que ainda
    ''' não existe é o consumidor da UI: quem implementa
    ''' <c>IPublicationConsumer</c> hoje é só o teste.
    '''
    ''' E quando existir, a entrega será <b>ao menos uma vez</b>: morrer depois
    ''' de a UI agir e antes de <see cref="MarcarDrenada"/> repete a entrega. É
    ''' a escolha certa — repetir é recuperável, perder não —, mas o consumidor
    ''' precisa ser idempotente, e isso é contrato dele, não deste arquivo.
    ''' </summary>
    Public NotInheritable Class CacheWriter

        Private ReadOnly _conn As SqliteConnection

        Public Sub New(db As CacheDatabase)
            _conn = db.Connection
        End Sub

        Public Sub New(conn As SqliteConnection)
            _conn = conn
        End Sub

        ' ==============================================================

        Public Function AbrirTentativa(folderKey As Long, environmentKey As Long,
                                       universo As String, epoca As Long,
                                       numero As Integer,
                                       Optional versaoAlgoritmo As Integer = 1,
                                       Optional corteRetencao As String = Nothing) As Long
            Using tx = Imediata()
                Dim k = Escalar(tx, "INSERT INTO scan_attempt (folder_key, environment_key, " &
                    "universe_fingerprint, retention_cutoff, algorithm_version, reconcile_epoch, " &
                    "attempt_number, stage, rows_read, started_at) " &
                    "VALUES ($f,$e,$u,$c,$v,$p,$n,'aberta',0,$t); SELECT last_insert_rowid()",
                    ("$f", CObj(folderKey)), ("$e", CObj(environmentKey)), ("$u", CObj(universo)),
                    ("$c", CObj(corteRetencao)), ("$v", CObj(versaoAlgoritmo)),
                    ("$p", CObj(epoca)), ("$n", CObj(numero)), ("$t", CObj(Agora())))
                tx.Commit()
                Return Convert.ToInt64(k)
            End Using
        End Function

        ''' <summary>
        ''' Grava uma página E avança o checkpoint, atomicamente.
        ''' Idempotente: reexecutar com as mesmas linhas não duplica nada.
        ''' </summary>
        Public Sub GravarPagina(attemptKey As Long, folderKey As Long, pagina As Integer,
                                linhas As IReadOnlyList(Of StagedRow), cursorDepois As String)
            CrashInjection.Talvez(CrashInjection.AntesDeGravarPagina)

            If CacheWriterDefects.CheckpointAntesDasLinhas Then
                ' DEFEITO deliberado. Ver CacheWriterDefects.
                Using tx0 = Imediata()
                    Executar(tx0, "UPDATE scan_attempt SET cursor = $c, stage = 'varrendo' " &
                        "WHERE attempt_key = $a", ("$a", CObj(attemptKey)), ("$c", CObj(cursorDepois)))
                    tx0.Commit()
                End Using
            End If

            Using tx = Imediata()
                For Each l In linhas
                    ' A pagina SO ENCENA. Nada do acervo — encarnacao,
                    ' metadado, associacao — e tocado aqui.
                    '
                    ' Antes era o contrario, e o defeito era grave: numa
                    ' encarnacao que ja existia, gravar a pagina substituia o
                    ' metadado PUBLICADO e devolvia a associacao para
                    ' 'presente'. Se a tentativa fosse rejeitada depois, nada
                    ' disso era desfeito — uma varredura RECUSADA alterava o
                    ' manifesto que a UI mostra. O teste nao pegava porque so
                    ' usava itens NOVOS, cujas associacoes ainda nao pertencem
                    ' a geracao nenhuma.
                    '
                    ' O unico (attempt_key, provider_entry_id) e o que torna a
                    ' releitura apos crash inofensiva: a mesma mensagem
                    ' encenada duas vezes conta uma vez so.
                    Executar(tx, "INSERT INTO scan_stage (attempt_key, provider_entry_id, " &
                        "page_number, cursor_after, search_key, internet_message_id, subject, " &
                        "sender_name, received_at, last_modified_at, size_bytes, has_attachments, " &
                        "is_unread, message_class) " &
                        "VALUES ($a,$p,$n,$c,$sk,$mid,$s,$sn,$r,$lm,$sz,$ha,$iu,$mc) " &
                        "ON CONFLICT (attempt_key, provider_entry_id) DO NOTHING",
                        ("$a", CObj(attemptKey)), ("$p", CObj(l.ProviderEntryId)),
                        ("$n", CObj(pagina)), ("$c", CObj(cursorDepois)),
                        ("$sk", CObj(l.SearchKey)), ("$mid", CObj(l.InternetMessageId)),
                        ("$s", CObj(l.Subject)), ("$sn", CObj(l.SenderName)),
                        ("$r", CObj(l.ReceivedAt)), ("$lm", CObj(l.LastModifiedAt)),
                        ("$sz", If(l.SizeBytes.HasValue, CObj(l.SizeBytes.Value), Nothing)),
                        ("$ha", If(l.HasAttachments.HasValue, CObj(If(l.HasAttachments.Value, 1, 0)), Nothing)),
                        ("$iu", If(l.IsUnread.HasValue, CObj(If(l.IsUnread.Value, 1, 0)), Nothing)),
                        ("$mc", CObj(l.MessageClass)))
                Next

                ' rows_read e DERIVADO, nao incrementado. Incrementar conta
                ' duas vezes quando a pagina e reexecutada apos crash, e a
                ' contagem inflada faria o S6 rejeitar a varredura inteira por
                ' um sintoma sem relacao nenhuma com a causa.
                Executar(tx, "UPDATE scan_attempt SET cursor = $c, stage = 'varrendo', " &
                    "rows_read = (SELECT COUNT(*) FROM scan_stage WHERE attempt_key = $a) " &
                    "WHERE attempt_key = $a",
                    ("$a", CObj(attemptKey)), ("$c", CObj(cursorDepois)))

                CrashInjection.Talvez(CrashInjection.DentroDaPaginaAntesDoCommit)
                tx.Commit()
            End Using

            CrashInjection.Talvez(CrashInjection.DepoisDoCommitDaPagina)
        End Sub

        ''' <summary>
        ''' Publica: geração + cobertura + cabeça + dívida para a UI, tudo numa
        ''' transação.
        '''
        ''' <paramref name="tipoDeVarredura"/> e <paramref name="alcance"/> sao
        ''' EIXOS DIFERENTES, e o schema os separa de proposito:
        '''
        '''   - <c>generation.coverage_kind</c> diz que TIPO de varredura foi -
        '''     completa ou incremental;
        '''   - <c>coverage_observation.coverage</c> diz QUANTO dela se
        '''     ALCANCOU - completa, parcial ou desconhecida.
        '''
        ''' Uma varredura pode ser de tipo completo e alcance parcial: ela
        ''' percorreu integralmente O CONJUNTO QUE O PROVIDER EXPOS, que nao e a
        ''' pasta inteira. Foi a §19.2, com pastas cheias reportando zero.
        '''
        ''' A observacao de cobertura vai na MESMA TRANSACAO da geracao. Fora
        ''' dela, "publicou" e "registrou o quanto alcancou" seriam dois
        ''' estados que podem divergir - e o divergente seria uma geracao sem
        ''' alcance conhecido, que e pior que nao ter geracao.
        ''' </summary>
        Public Function Publicar(attemptKey As Long, folderKey As Long,
                                 tipoDeVarredura As String,
                                 contagemAntes As Long, contagemDepois As Long,
                                 Optional ByRef geracao As Long = 0,
                                 Optional alcance As String = "desconhecida",
                                 Optional environmentKey As Long = 1) As PublishOutcome
            Using tx = Imediata()
                Dim epocaTentativa = Ler(tx, "SELECT reconcile_epoch FROM scan_attempt WHERE attempt_key=$a",
                                         ("$a", CObj(attemptKey)))
                If epocaTentativa Is Nothing Then Return PublishOutcome.RecusadaPorEstado

                Dim estagio = Convert.ToString(Ler(tx,
                    "SELECT stage FROM scan_attempt WHERE attempt_key=$a", ("$a", CObj(attemptKey))))
                If estagio = "publicada" OrElse estagio = "descartada" Then
                    Return PublishOutcome.RecusadaPorEstado
                End If

                Dim epocaPasta = Convert.ToInt64(Ler(tx,
                    "SELECT reconcile_epoch FROM folder WHERE folder_key=$f", ("$f", CObj(folderKey))))

                ' CAS: a epoca e lida DENTRO da mesma transacao que escreve.
                If Convert.ToInt64(epocaTentativa) <> epocaPasta Then
                    Executar(tx, "UPDATE scan_attempt SET stage='descartada', ended_at=$t, " &
                        "rejection='epoca' WHERE attempt_key=$a",
                        ("$a", CObj(attemptKey)), ("$t", CObj(Agora())))
                    tx.Commit()
                    Return PublishOutcome.RecusadaPorEpoca
                End If

                ' ORDEM: a cabeca avanca por ordem de ABERTURA da tentativa,
                ' nao por ordem de publicacao.
                '
                ' A distincao e o criterio 10 inteiro. generation_key e
                ' atribuido no INSERT, que acontece ao PUBLICAR — entao uma
                ' varredura velha que termina tarde recebe a chave MAIOR, e um
                ' teste de monotonicidade sobre generation_key aprovaria
                ' exatamente o caso que ele deveria barrar.
                '
                ' A politica e "a tentativa aberta por ultimo vence", e ela NAO
                ' e o mesmo que "os dados mais frescos vencem". Tentativas
                ' sobrepostas intercalam leituras: a mais velha pode, em tese,
                ' ter lido parte dos dados depois da mais nova. Ordem de
                ' abertura e aproximacao conservadora — erra para o lado de
                ' descartar trabalho bom, nunca para o lado de deixar a cabeca
                ' recuar.
                '
                ' E nao trava o sistema: a guarda compara so com a cabeca
                ' PUBLICADA, entao uma tentativa nova abandonada nao bloqueia
                ' ninguem. O custo maximo e perder o frescor de uma varredura,
                ' e a proxima recupera. Afirmar frescor de verdade exigiria
                ' serializar as varreduras por pasta ou ter leitura com
                ' snapshot, e nenhuma das duas existe aqui.
                Dim cabeca = Ler(tx, "SELECT g.attempt_key FROM folder f " &
                    "JOIN generation g ON g.generation_key = f.published_generation_key " &
                    "WHERE f.folder_key = $f", ("$f", CObj(folderKey)))
                If cabeca IsNot Nothing AndAlso Convert.ToInt64(cabeca) > attemptKey Then
                    Executar(tx, "UPDATE scan_attempt SET stage='descartada', ended_at=$t, " &
                        "rejection='ordem' WHERE attempt_key=$a",
                        ("$a", CObj(attemptKey)), ("$t", CObj(Agora())))
                    tx.Commit()
                    Return PublishOutcome.RecusadaPorOrdem
                End If

                Dim lidas = Convert.ToInt64(Ler(tx,
                    "SELECT COUNT(*) FROM scan_stage WHERE attempt_key=$a", ("$a", CObj(attemptKey))))
                Dim distintas = Convert.ToInt64(Ler(tx,
                    "SELECT COUNT(DISTINCT provider_entry_id) FROM scan_stage WHERE attempt_key=$a",
                    ("$a", CObj(attemptKey))))
                Dim universo = Convert.ToString(Ler(tx,
                    "SELECT universe_fingerprint FROM scan_attempt WHERE attempt_key=$a",
                    ("$a", CObj(attemptKey))))

                ' A observacao de ALCANCE, na mesma transacao.
                Dim cov = Convert.ToInt64(Escalar(tx,
                    "INSERT INTO coverage_observation (folder_key, environment_key, " &
                    "universe_fingerprint, coverage, source, observed_at) " &
                    "VALUES ($f,$e,$u,$c,'varredura',$t); SELECT last_insert_rowid()",
                    ("$f", CObj(folderKey)), ("$e", CObj(environmentKey)),
                    ("$u", CObj(universo)), ("$c", CObj(alcance)), ("$t", CObj(Agora()))))

                Dim g = Convert.ToInt64(Escalar(tx,
                    "INSERT INTO generation (folder_key, attempt_key, coverage_kind, " &
                    "coverage_key, universe_fingerprint, rows_read, count_before, count_after, " &
                    "distinct_keys, reconcile_epoch, published_at) " &
                    "VALUES ($f,$a,$k,$cov,$u,$r,$b,$d,$q,$p,$t); SELECT last_insert_rowid()",
                    ("$f", CObj(folderKey)), ("$a", CObj(attemptKey)), ("$k", CObj(tipoDeVarredura)),
                    ("$cov", CObj(cov)),
                    ("$u", CObj(universo)), ("$r", CObj(lidas)), ("$b", CObj(contagemAntes)),
                    ("$d", CObj(contagemDepois)), ("$q", CObj(distintas)),
                    ("$p", CObj(epocaPasta)), ("$t", CObj(Agora()))))

                Executar(tx, "UPDATE folder SET published_generation_key = $g WHERE folder_key = $f",
                    ("$g", CObj(g)), ("$f", CObj(folderKey)))

                ' ===== MATERIALIZACAO =====
                '
                ' O acervo so recebe as linhas AQUI, a partir da encenacao, e
                ' dentro desta transacao. Enquanto a tentativa nao publica,
                ' nada do que ela leu e visivel — e uma tentativa rejeitada nao
                ' deixa marca nenhuma no que a UI mostra.
                Materializar(tx, attemptKey, folderKey, g)

                ' ===== NAO VISTOS -> SUSPEITOS =====
                '
                ' O comando MarcarNaoVistosComoSuspeitos era EMITIDO pelo
                ' modelo e nunca EXECUTADO por ninguem: o sink nem tinha a
                ' operacao. Efeito: depois de uma geracao com A e B e outra so
                ' com A, o B continuava 'presente' — e o manifesto o devolvia
                ' como se estivesse la.
                '
                ' So 'presente' vira 'suspeito'. NaoVerificado continua
                ' NaoVerificado (suspeita pressupoe presenca anterior), e
                ' suspeito continua suspeito — geracao que passa NAO promove
                ' suspeita a ausencia, nem por contagem nem por tempo. E o
                ' AplicarGeracao da PresencePolicy, em SQL.
                Executar(tx,
                    "UPDATE association SET presence='suspeito', version = version + 1, " &
                    "       generation_key = $g " &
                    "WHERE folder_key = $f AND presence = 'presente' " &
                    "  AND item_key NOT IN (" &
                    "    SELECT i.item_key FROM incarnation i JOIN scan_stage s " &
                    "    ON s.provider_entry_id = i.provider_entry_id " &
                    "    WHERE i.folder_key = $f AND s.attempt_key = $a)",
                    ("$g", CObj(g)), ("$f", CObj(folderKey)), ("$a", CObj(attemptKey)))

                Executar(tx, "UPDATE association SET generation_key = $g " &
                    "WHERE folder_key = $f AND item_key IN (" &
                    "  SELECT i.item_key FROM incarnation i JOIN scan_stage s " &
                    "  ON s.provider_entry_id = i.provider_entry_id " &
                    "  WHERE i.folder_key = $f AND s.attempt_key = $a)",
                    ("$g", CObj(g)), ("$f", CObj(folderKey)), ("$a", CObj(attemptKey)))

                Executar(tx, "UPDATE scan_attempt SET stage='publicada', ended_at=$t " &
                    "WHERE attempt_key=$a", ("$a", CObj(attemptKey)), ("$t", CObj(Agora())))

                ' A DIVIDA para a UI, na MESMA transacao.
                Executar(tx, "INSERT INTO publication_log (generation_key, emitted_at, " &
                    "delivery_attempts) VALUES ($g,$t,0)", ("$g", CObj(g)), ("$t", CObj(Agora())))

                CrashInjection.Talvez(CrashInjection.DentroDaPublicacaoAntesDoCommit)
                tx.Commit()
                geracao = g
            End Using

            CrashInjection.Talvez(CrashInjection.DepoisDoCommitDaPublicacao)
            Return PublishOutcome.Publicada
        End Function

        ''' <summary>Gerações publicadas que a UI ainda não consumiu.</summary>
        Public Function PublicacoesPendentes() As IReadOnlyList(Of Long)
            Dim r As New List(Of Long)()
            Using cmd = _conn.CreateCommand()
                cmd.CommandText = "SELECT generation_key FROM publication_log " &
                                  "WHERE drained_at IS NULL ORDER BY log_key"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        r.Add(rd.GetInt64(0))
                    End While
                End Using
            End Using
            Return r
        End Function

        ''' <summary>
        ''' Registra que a entrega foi TENTADA e falhou. Sem isto, um consumidor
        ''' que falha sempre trava a fila em silencio.
        ''' </summary>
        Public Sub RegistrarFalhaNaEntrega(geracao As Long, erro As String)
            Using tx = Imediata()
                Executar(tx, "UPDATE publication_log SET delivery_attempts = delivery_attempts + 1, " &
                    "last_error = $e, last_attempt_at = $t WHERE generation_key = $g",
                    ("$g", CObj(geracao)), ("$e", CObj(Recortar(erro))), ("$t", CObj(Agora())))
                tx.Commit()
            End Using
        End Sub

        Public Function TentativasDeEntrega(geracao As Long) As Integer
            Return Convert.ToInt32(Ler(Nothing,
                "SELECT delivery_attempts FROM publication_log WHERE generation_key=$g",
                ("$g", CObj(geracao))))
        End Function

        Public Function UltimoErroDeEntrega(geracao As Long) As String
            Dim v = Ler(Nothing,
                "SELECT last_error FROM publication_log WHERE generation_key=$g",
                ("$g", CObj(geracao)))
            Return If(v Is Nothing, Nothing, Convert.ToString(v))
        End Function

        Public Sub MarcarDrenada(geracao As Long)
            Using tx = Imediata()
                Executar(tx, "UPDATE publication_log SET drained_at = $t WHERE generation_key = $g",
                    ("$g", CObj(geracao)), ("$t", CObj(Agora())))
                tx.Commit()
            End Using
        End Sub

        ''' <summary>Cursor de onde retomar, ou Nothing se não há o que retomar.</summary>
        Public Function CursorDaTentativa(attemptKey As Long) As String
            Dim v = Ler(Nothing, "SELECT cursor FROM scan_attempt WHERE attempt_key=$a",
                        ("$a", CObj(attemptKey)))
            Return If(v Is Nothing, Nothing, Convert.ToString(v))
        End Function

        Public Function LinhasEncenadas(attemptKey As Long) As Integer
            Return Convert.ToInt32(Ler(Nothing,
                "SELECT COUNT(*) FROM scan_stage WHERE attempt_key=$a", ("$a", CObj(attemptKey))))
        End Function

        ''' <summary>
        ''' Marca a tentativa como descartada. Sem isto, toda varredura que não
        ''' publica deixa uma linha <c>aberta</c> ou <c>varrendo</c> no banco,
        ''' e a retomada seguinte encontra lixo que parece trabalho.
        ''' </summary>
        Public Sub Descartar(attemptKey As Long, motivo As String)
            Using tx = Imediata()
                Executar(tx, "UPDATE scan_attempt SET stage='descartada', ended_at=$t, " &
                    "rejection=$r WHERE attempt_key=$a AND stage NOT IN ('publicada','descartada')",
                    ("$a", CObj(attemptKey)), ("$t", CObj(Agora())),
                    ("$r", CObj(Recortar(motivo))))
                tx.Commit()
            End Using
        End Sub

        Public Function EpocaDaPasta(folderKey As Long) As Long
            Return Convert.ToInt64(Ler(Nothing,
                "SELECT reconcile_epoch FROM folder WHERE folder_key=$f", ("$f", CObj(folderKey))))
        End Function

        Private Shared Function Recortar(s As String) As String
            If s Is Nothing Then Return Nothing
            Return If(s.Length <= 500, s, s.Substring(0, 500))
        End Function

        Public Function EstagioDa(attemptKey As Long) As String
            Dim v = Ler(Nothing, "SELECT stage FROM scan_attempt WHERE attempt_key=$a",
                        ("$a", CObj(attemptKey)))
            Return If(v Is Nothing, Nothing, Convert.ToString(v))
        End Function

        Public Function CabecaPublicada(folderKey As Long) As Long?
            Dim v = Ler(Nothing, "SELECT published_generation_key FROM folder WHERE folder_key=$f",
                        ("$f", CObj(folderKey)))
            If v Is Nothing Then Return Nothing
            Return Convert.ToInt64(v)
        End Function

        ''' <summary>Sobe a época: invalida toda varredura em curso nesta pasta.</summary>
        Public Sub InvalidarUniverso(folderKey As Long)
            Using tx = Imediata()
                Executar(tx, "UPDATE folder SET reconcile_epoch = reconcile_epoch + 1 " &
                    "WHERE folder_key = $f", ("$f", CObj(folderKey)))
                tx.Commit()
            End Using
        End Sub

        ' ==============================================================

        ''' <summary>
        ''' Leva a encenação para o acervo, dentro da transação da publicação.
        '''
        ''' Tudo é feito em SQL de conjunto, sem laço por linha, e não é
        ''' otimização: uma varredura da Caixa de Entrada desta caixa tem mil
        ''' linhas, e mil idas ao SQLite dentro de uma transação com o lock de
        ''' escrita segurado é tempo em que ninguém mais publica.
        ''' </summary>
        Private Sub Materializar(tx As SqliteTransaction, attemptKey As Long,
                                 folderKey As Long, geracao As Long)
            ' 1-2. ITEM + ENCARNACAO para cada encenada que ainda nao tem
            ' encarnacao nesta pasta.
            '
            ' Em laco, e nao em SQL de conjunto. A versao de conjunto casava
            ' os itens recem-criados com as linhas encenadas por ROW_NUMBER e
            ' por created_at — e Agora() devolve valor novo a cada chamada,
            ' entao os dois passos nem usavam o mesmo carimbo. Truque frageis
            ' para ganhar idas ao banco dentro de uma transacao que ja e curta
            ' e o tipo de coisa que quebra em silencio quando a ordem muda.
            Dim faltando As New List(Of String)()
            Using cmd = _conn.CreateCommand()
                cmd.CommandText =
                    "SELECT s.provider_entry_id FROM scan_stage s " &
                    "WHERE s.attempt_key = $a AND NOT EXISTS (" &
                    "  SELECT 1 FROM incarnation i WHERE i.folder_key = $f " &
                    "    AND i.provider_entry_id = s.provider_entry_id) " &
                    "ORDER BY s.stage_key"
                cmd.Transaction = tx
                cmd.Parameters.AddWithValue("$a", attemptKey)
                cmd.Parameters.AddWithValue("$f", folderKey)
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        faltando.Add(rd.GetString(0))
                    End While
                End Using
            End Using

            For Each id In faltando
                Dim itemK = Convert.ToInt64(Escalar(tx,
                    "INSERT INTO item (created_at) VALUES ($t); SELECT last_insert_rowid()",
                    ("$t", CObj(Agora()))))
                Executar(tx,
                    "INSERT INTO incarnation (item_key, folder_key, provider_entry_id, " &
                    "  search_key, internet_message_id, first_seen_generation, last_seen_generation) " &
                    "SELECT $i, $f, s.provider_entry_id, s.search_key, s.internet_message_id, $g, $g " &
                    "FROM scan_stage s WHERE s.attempt_key = $a AND s.provider_entry_id = $p",
                    ("$i", CObj(itemK)), ("$f", CObj(folderKey)), ("$g", CObj(geracao)),
                    ("$a", CObj(attemptKey)), ("$p", CObj(id)))
            Next

            ' 3. METADADO: substitui o da encarnacao pelo encenado.
            Executar(tx,
                "DELETE FROM metadata_observation WHERE incarnation_key IN (" &
                "  SELECT i.incarnation_key FROM incarnation i JOIN scan_stage s " &
                "  ON s.provider_entry_id = i.provider_entry_id " &
                "  WHERE i.folder_key = $f AND s.attempt_key = $a)",
                ("$f", CObj(folderKey)), ("$a", CObj(attemptKey)))

            Executar(tx,
                "INSERT INTO metadata_observation (incarnation_key, generation_key, subject, " &
                "  sender_name, received_at, last_modified_at, size_bytes, has_attachments, " &
                "  is_unread, message_class) " &
                "SELECT i.incarnation_key, $g, s.subject, s.sender_name, s.received_at, " &
                "       s.last_modified_at, s.size_bytes, s.has_attachments, s.is_unread, " &
                "       s.message_class " &
                "FROM incarnation i JOIN scan_stage s " &
                "  ON s.provider_entry_id = i.provider_entry_id " &
                "WHERE i.folder_key = $f AND s.attempt_key = $a",
                ("$f", CObj(folderKey)), ("$a", CObj(attemptKey)), ("$g", CObj(geracao)))

            ' 4. last_seen_generation das que ja existiam.
            Executar(tx,
                "UPDATE incarnation SET last_seen_generation = $g " &
                "WHERE folder_key = $f AND provider_entry_id IN (" &
                "  SELECT provider_entry_id FROM scan_stage WHERE attempt_key = $a)",
                ("$f", CObj(folderKey)), ("$a", CObj(attemptKey)), ("$g", CObj(geracao)))

            ' 5. ASSOCIACAO: visto e visto, venha de onde vier — inclusive de
            '    AusenteDaPasta, porque o item voltou.
            Executar(tx,
                "INSERT INTO association (item_key, folder_key, presence, observability, version) " &
                "SELECT i.item_key, $f, 'presente', 'observavel', 0 " &
                "FROM incarnation i JOIN scan_stage s " &
                "  ON s.provider_entry_id = i.provider_entry_id " &
                "WHERE i.folder_key = $f AND s.attempt_key = $a " &
                "ON CONFLICT (item_key, folder_key) DO UPDATE SET " &
                "  presence='presente', observability='observavel', version = version + 1",
                ("$f", CObj(folderKey)), ("$a", CObj(attemptKey)))
        End Sub

        Private Function GarantirEncarnacao(tx As SqliteTransaction, folderKey As Long,
                                            l As StagedRow) As Long
            Dim k = Ler(tx, "SELECT incarnation_key FROM incarnation " &
                "WHERE folder_key=$f AND provider_entry_id=$e",
                ("$f", CObj(folderKey)), ("$e", CObj(l.ProviderEntryId)))
            If k IsNot Nothing Then Return Convert.ToInt64(k)

            Dim itemK = Convert.ToInt64(Escalar(tx,
                "INSERT INTO item (created_at) VALUES ($t); SELECT last_insert_rowid()",
                ("$t", CObj(Agora()))))

            Return Convert.ToInt64(Escalar(tx,
                "INSERT INTO incarnation (item_key, folder_key, provider_entry_id, " &
                "provider_record_key, search_key, internet_message_id) " &
                "VALUES ($i,$f,$e,NULL,$s,$m); SELECT last_insert_rowid()",
                ("$i", CObj(itemK)), ("$f", CObj(folderKey)), ("$e", CObj(l.ProviderEntryId)),
                ("$s", CObj(l.SearchKey)), ("$m", CObj(l.InternetMessageId))))
        End Function

        Private Sub GravarMetadado(tx As SqliteTransaction, incK As Long, l As StagedRow)
            Executar(tx, "DELETE FROM metadata_observation WHERE incarnation_key=$i", ("$i", CObj(incK)))
            Executar(tx, "INSERT INTO metadata_observation (incarnation_key, subject, sender_name, " &
                "received_at, last_modified_at, size_bytes, has_attachments, is_unread, message_class) " &
                "VALUES ($i,$s,$n,$r,$m,$z,$a,$u,$c)",
                ("$i", CObj(incK)), ("$s", CObj(l.Subject)), ("$n", CObj(l.SenderName)),
                ("$r", CObj(l.ReceivedAt)), ("$m", CObj(l.LastModifiedAt)),
                ("$z", If(l.SizeBytes.HasValue, CObj(l.SizeBytes.Value), Nothing)),
                ("$a", If(l.HasAttachments.HasValue, CObj(If(l.HasAttachments.Value, 1, 0)), Nothing)),
                ("$u", If(l.IsUnread.HasValue, CObj(If(l.IsUnread.Value, 1, 0)), Nothing)),
                ("$c", CObj(l.MessageClass)))
        End Sub

        Private Sub GarantirAssociacao(tx As SqliteTransaction, folderKey As Long, incK As Long)
            Dim itemK = Convert.ToInt64(Ler(tx,
                "SELECT item_key FROM incarnation WHERE incarnation_key=$i", ("$i", CObj(incK))))
            Executar(tx, "INSERT INTO association (item_key, folder_key, presence, observability, version) " &
                "VALUES ($i,$f,'presente','observavel',0) " &
                "ON CONFLICT (item_key, folder_key) DO UPDATE SET presence='presente', " &
                "observability='observavel', version = version + 1",
                ("$i", CObj(itemK)), ("$f", CObj(folderKey)))
        End Sub

        ' ============ utilidades =====================================

        Private Function Imediata() As SqliteTransaction
            ' BEGIN IMMEDIATE: pega o lock de escrita JA. Sem isto, duas
            ' publicacoes concorrentes leem a mesma epoca antes de qualquer
            ' uma escrever, e o CAS vira decoracao.
            Return _conn.BeginTransaction(IsolationLevel.Serializable, deferred:=False)
        End Function

        Private Shared Function Agora() As String
            Return DateTime.UtcNow.ToString("o")
        End Function

        Private Shared Sub Aplicar(cmd As SqliteCommand, ps As (String, Object)())
            For Each p In ps
                cmd.Parameters.AddWithValue(p.Item1, If(p.Item2, DBNull.Value))
            Next
        End Sub

        Private Sub Executar(tx As SqliteTransaction, sql As String, ParamArray ps As (String, Object)())
            Using cmd = _conn.CreateCommand()
                cmd.CommandText = sql
                cmd.Transaction = tx
                Aplicar(cmd, ps)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Private Function Escalar(tx As SqliteTransaction, sql As String,
                                 ParamArray ps As (String, Object)()) As Object
            Using cmd = _conn.CreateCommand()
                cmd.CommandText = sql
                cmd.Transaction = tx
                Aplicar(cmd, ps)
                Return cmd.ExecuteScalar()
            End Using
        End Function

        Private Function Ler(tx As SqliteTransaction, sql As String,
                             ParamArray ps As (String, Object)()) As Object
            Using cmd = _conn.CreateCommand()
                cmd.CommandText = sql
                If tx IsNot Nothing Then cmd.Transaction = tx
                Aplicar(cmd, ps)
                Dim v = cmd.ExecuteScalar()
                Return If(v Is Nothing OrElse v Is DBNull.Value, Nothing, v)
            End Using
        End Function

    End Class

End Namespace

Imports System.Collections.Generic
Imports System.Globalization
Imports Iris.Assist
Imports Iris.Cache
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Integration

    ''' <summary>
    ''' O <see cref="IDisclosureJournal"/> sobre o SQLite do cache.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DURÁVEL QUER DIZER DURÁVEL</b>
    '''
    ''' Cada passo é um <c>COMMIT</c> próprio, e não um passo de uma transação
    ''' longa. Uma transação que abrisse na intenção e fechasse na conclusão
    ''' perderia <b>tudo</b> num crash — inclusive a intenção, que é justamente
    ''' o registro que precisa sobreviver.
    '''
    ''' O preço é uma escrita por passo. O que se compra é que morrer no meio
    ''' deixe rastro.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>CADA PASSO DIZ SE PEGOU</b>
    '''
    ''' As guardas moram no <c>WHERE</c>, e o número de linhas alteradas é o que
    ''' volta. Ignorá-lo era o buraco: um <c>Iniciando</c> que não persistisse
    ''' passava em silêncio e quem chamou seguia para o HTTP assim mesmo —
    ''' egress sem registro de voo.
    '''
    ''' A guarda no <c>WHERE</c>, e não num <c>If</c> antes: duas execuções
    ''' podem estar escrevendo, e o banco é quem arbitra.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>SEM CONTEÚDO, E SEM TEXTO DE TERCEIRO</b>
    '''
    ''' Nem trecho, nem assunto, nem nome de rótulo. E nem o corpo do erro do
    ''' provedor: resposta de erro pode <b>ecoar o que foi enviado</b>. Os
    ''' motivos são <b>enums fechados</b> — não há campo por onde texto
    ''' arbitrário entre.
    ''' </summary>
    Public NotInheritable Class SqliteDisclosureJournal
        Implements IDisclosureJournal

        Private ReadOnly _db As CacheDatabase

        Public Sub New(db As CacheDatabase)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _db = db
        End Sub

        ' ==============================================================

        Public Function Intencao(c As DisclosureCapability,
                                 quando As DateTimeOffset) As Boolean _
                                 Implements IDisclosureJournal.Intencao
            If c Is Nothing Then Throw New ArgumentNullException(NameOf(c))
            ' A contagem vem da propria capability. Enquanto vinha de fora, o
            ' diario podia registrar quantidade diferente da autorizada.
            Dim mensagens = c.Itens.Count

            Return Executar(
                "INSERT OR IGNORE INTO disclosure_log (" &
                "  seq, request_id, capability_id, stage, activation_id, " &
                "  activation_version, operation, provider, endpoint, model, " &
                "  payload_hash, payload_bytes, message_count, " &
                "  intended_at, started_at, finished_at, note, gate_reason) " &
                "SELECT COALESCE(MAX(seq), 0) + 1, $r, $c, $s, $ai, $av, $op, $p, $e, $m, " &
                "       $h, $b, $n, $t, NULL, NULL, $nota, $gr FROM disclosure_log",
                ("$r", CObj(c.RequestId.ToString("D"))),
                ("$c", CObj(c.Id.ToString("D"))),
                ("$s", CObj(DisclosureStage.Intencionada.ToString())),
                ("$ai", CObj(c.AtivacaoId)),
                ("$av", CObj(c.AtivacaoVersao)),
                ("$op", CObj(c.Operacao.ToString())),
                ("$p", CObj(c.Destino.Provedor)),
                ("$e", CObj(c.Destino.Endpoint)),
                ("$m", CObj(c.Destino.Modelo)),
                ("$h", CObj(c.Hash)),
                ("$b", CObj(c.Comprimento)),
                ("$n", CObj(mensagens)),
                ("$t", CObj(Instante(quando))),
                ("$nota", CObj(DisclosureNote.Nenhuma.ToString())),
                ("$gr", CObj(DisclosureReason.NaoDecidido.ToString()))) = 1
        End Function

        Public Function Iniciando(requestId As Guid, quando As DateTimeOffset) As Boolean _
                                  Implements IDisclosureJournal.Iniciando
            ' So sai de Intencionada. Reiniciar um envio ja em voo, ou ja
            ' concluido, seria reabrir uma janela que ja fechou.
            Return Executar(
                "UPDATE disclosure_log SET stage = $s, started_at = $t " &
                "WHERE request_id = $r AND stage = $de",
                ("$s", CObj(DisclosureStage.EmVoo.ToString())),
                ("$t", CObj(Instante(quando))),
                ("$r", CObj(requestId.ToString("D"))),
                ("$de", CObj(DisclosureStage.Intencionada.ToString()))) = 1
        End Function

        Public Function Concluir(requestId As Guid, quando As DateTimeOffset) As Boolean _
                                 Implements IDisclosureJournal.Concluir
            Return Terminar(requestId, DisclosureStage.Concluida, quando,
                            DisclosureNote.Nenhuma, DisclosureReason.NaoDecidido,
                            {DisclosureStage.EmVoo})
        End Function

        Public Function Falhar(requestId As Guid, quando As DateTimeOffset,
                               nota As DisclosureNote, podeTerChegado As Boolean,
                               Optional codigoHttp As Integer? = Nothing) As Boolean _
                               Implements IDisclosureJournal.Falhar
            ' Uma vez EM VOO, falhar e SEMPRE ambiguo — mesmo que o chamador
            ' jure que nao chegou. Ele nao pode saber: entre "a conexao caiu" e
            ' "a conexao caiu depois de o servidor ler o corpo" nao ha
            ' diferenca observavel deste lado.
            ' Nota que nao descreve transporte nao entra: "o portao negou" nao
            ' e um jeito de a transmissao falhar, porque ela nem comecaria.
            If Not DisclosureNotes.DeTransporte(nota) Then Return False

            If podeTerChegado Then
                Return Terminar(requestId, DisclosureStage.Ambigua, quando, nota,
                                DisclosureReason.NaoDecidido,
                                {DisclosureStage.Intencionada, DisclosureStage.EmVoo},
                                codigoHttp)
            End If

            ' NaoEnviada so vale a partir de Intencionada, onde a transmissao
            ' comprovadamente nao tinha comecado.
            If Terminar(requestId, DisclosureStage.NaoEnviada, quando, nota,
                        DisclosureReason.NaoDecidido, {DisclosureStage.Intencionada},
                        codigoHttp) Then
                Return True
            End If

            ' Estava EM VOO: o desfecho honesto e ambiguo, nao "nao enviou".
            Return Terminar(requestId, DisclosureStage.Ambigua, quando, nota,
                            DisclosureReason.NaoDecidido, {DisclosureStage.EmVoo},
                            codigoHttp)
        End Function

        Public Function NaoEnviou(requestId As Guid, quando As DateTimeOffset,
                                  nota As DisclosureNote,
                                  Optional motivoDoPortao As DisclosureReason =
                                      DisclosureReason.NaoDecidido) As Boolean _
                                  Implements IDisclosureJournal.NaoEnviou
            If Not DisclosureNotes.AnteriorAoEnvio(nota) Then Return False
            If Not DisclosureNotes.Coerente(nota, motivoDoPortao) Then Return False

            Return Terminar(requestId, DisclosureStage.NaoEnviada, quando, nota,
                            motivoDoPortao, {DisclosureStage.Intencionada})
        End Function

        ' ==============================================================

        ''' <summary>
        ''' A reconciliação da abertura — o quinto passo.
        '''
        ''' O que ficou <b>em voo</b> de uma execução morta vira ambíguo: os
        ''' bytes podem ter chegado, e ninguém vai saber. O que ficou só como
        ''' intenção vira não-enviada, porque ali a transmissão não tinha
        ''' começado.
        ''' </summary>
        Public Function Reconciliar(quando As DateTimeOffset) As Integer _
                                    Implements IDisclosureJournal.Reconciliar
            Dim ambiguas = Executar(
                "UPDATE disclosure_log SET stage = $novo, finished_at = $t, note = $nota " &
                "WHERE stage = $velho",
                ("$novo", CObj(DisclosureStage.Ambigua.ToString())),
                ("$velho", CObj(DisclosureStage.EmVoo.ToString())),
                ("$t", CObj(Instante(quando))),
                ("$nota", CObj(DisclosureNote.ProcessoMorreuEmVoo.ToString())))

            Executar(
                "UPDATE disclosure_log SET stage = $novo, finished_at = $t, note = $nota " &
                "WHERE stage = $velho",
                ("$novo", CObj(DisclosureStage.NaoEnviada.ToString())),
                ("$velho", CObj(DisclosureStage.Intencionada.ToString())),
                ("$t", CObj(Instante(quando))),
                ("$nota", CObj(DisclosureNote.ProcessoMorreuAntesDeTransmitir.ToString())))

            Return ambiguas
        End Function

        ''' <summary>
        ''' Do mais recente para o mais antigo, por <c>intended_at</c> — que é
        ''' <b>imutável</b> — com a sequência de inserção desempatando.
        '''
        ''' Ordenar por um carimbo que cada passo sobrescrevia fazia uma intenção
        ''' abandonada há meses aparecer como atividade recente, logo depois de
        ''' uma reconciliação. E o <c>Guid</c> não serve de desempate: ele é
        ''' aleatório, então a ordem entre dois registros do mesmo instante
        ''' mudava a cada execução.
        ''' </summary>
        Public Function Ler(quantas As Integer) As IReadOnlyList(Of DisclosureEntry) _
                            Implements IDisclosureJournal.Ler
            Dim saida As New List(Of DisclosureEntry)()
            Using cmd = _db.Connection.CreateCommand()
                cmd.CommandText =
                    "SELECT seq, request_id, capability_id, stage, activation_id, " &
                    "       activation_version, operation, provider, endpoint, model, " &
                    "       payload_hash, payload_bytes, message_count, " &
                    "       intended_at, started_at, finished_at, note, gate_reason, " &
                    "       http_status " &
                    "FROM disclosure_log ORDER BY intended_at DESC, seq DESC LIMIT $n"
                cmd.Parameters.AddWithValue("$n", quantas)
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        saida.Add(New DisclosureEntry(
                            r.GetInt64(0), Guid.Parse(r.GetString(1)), Guid.Parse(r.GetString(2)),
                            Ler(Of DisclosureStage)(r.GetString(3)),
                            r.GetString(4), r.GetInt32(5),
                            Ler(Of AssistOperation)(r.GetString(6)),
                            r.GetString(7), r.GetString(8), r.GetString(9),
                            If(r.IsDBNull(10), Nothing, r.GetString(10)),
                            r.GetInt32(11), r.GetInt32(12),
                            Momento(r.GetString(13)),
                            If(r.IsDBNull(14), CType(Nothing, DateTimeOffset?), Momento(r.GetString(14))),
                            If(r.IsDBNull(15), CType(Nothing, DateTimeOffset?), Momento(r.GetString(15))),
                            Ler(Of DisclosureNote)(r.GetString(16)),
                            Ler(Of DisclosureReason)(r.GetString(17)),
                            If(r.IsDBNull(18), CType(Nothing, Integer?), r.GetInt32(18))))
                    End While
                End Using
            End Using
            Return saida
        End Function

        ' ==============================================================

        ''' <summary>
        ''' Termina o registro, se ele estiver num dos estágios permitidos.
        ''' Devolve se a transição <b>aconteceu</b>.
        ''' </summary>
        Private Function Terminar(requestId As Guid, destino As DisclosureStage,
                                  quando As DateTimeOffset, nota As DisclosureNote,
                                  portao As DisclosureReason,
                                  de As DisclosureStage(),
                                  Optional codigoHttp As Integer? = Nothing) As Boolean
            ' Valor inventado - CType(999, DisclosureNote) compila - e dupla
            ' incoerente nao entram. Um registro incoerente e pior que um
            ' registro ausente: ele PARECE resposta.
            If Not DisclosureNotes.Coerente(nota, portao) Then Return False

            ' Codigo fora da faixa vira NULO em vez de abortar a transicao: um
            ' campo de diagnostico nao pode piorar o registro que ele anota.
            ' Ver DisclosureNotes.CodigoDeDiario.
            Dim codigo = DisclosureNotes.CodigoDeDiario(codigoHttp)

            Dim aceitos = String.Join(",", de.Select(Function(e) "'" & e.ToString() & "'"))
            Return Executar(
                "UPDATE disclosure_log SET stage = $s, finished_at = $t, " &
                "  note = $nota, gate_reason = $gr, http_status = $http " &
                "WHERE request_id = $r AND stage IN (" & aceitos & ")",
                ("$s", CObj(destino.ToString())),
                ("$t", CObj(Instante(quando))),
                ("$nota", CObj(nota.ToString())),
                ("$gr", CObj(portao.ToString())),
                ("$http", If(codigo.HasValue, CObj(codigo.Value), CObj(DBNull.Value))),
                ("$r", CObj(requestId.ToString("D")))) = 1
        End Function

        Private Function Executar(sql As String,
                                  ParamArray p As (Nome As String, Valor As Object)()) As Integer
            Using cmd = _db.Connection.CreateCommand()
                cmd.CommandText = sql
                For Each par In p
                    cmd.Parameters.AddWithValue(par.Nome, par.Valor)
                Next
                Return cmd.ExecuteNonQuery()
            End Using
        End Function

        Private Shared Function Ler(Of T)(texto As String) As T
            Return CType([Enum].Parse(GetType(T), texto), T)
        End Function

        Private Shared Function Momento(texto As String) As DateTimeOffset
            Return DateTimeOffset.Parse(texto, CultureInfo.InvariantCulture,
                                        DateTimeStyles.RoundtripKind)
        End Function

        ''' <summary>ISO 8601 com offset, para ordenar como texto e não mentir.</summary>
        Private Shared Function Instante(q As DateTimeOffset) As String
            Return q.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture)
        End Function

    End Class

End Namespace

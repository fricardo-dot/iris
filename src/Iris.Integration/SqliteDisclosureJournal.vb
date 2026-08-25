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
    ''' <b>SEM CONTEÚDO, E SEM CORPO DE RESPOSTA</b>
    '''
    ''' Nem trecho, nem assunto, nem nome de rótulo. E nem o corpo do erro do
    ''' provedor: resposta de erro pode <b>ecoar o que foi enviado</b>, e o
    ''' diário viraria a cópia que ele existe para não criar. O
    ''' <c>reason</c> guarda código, não texto de terceiro.
    ''' </summary>
    Public NotInheritable Class SqliteDisclosureJournal
        Implements IDisclosureJournal

        Private ReadOnly _db As CacheDatabase

        Public Sub New(db As CacheDatabase)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _db = db
        End Sub

        ' ==============================================================

        Public Sub Intencao(c As DisclosureCapability, mensagens As Integer,
                            quando As DateTimeOffset) Implements IDisclosureJournal.Intencao
            If c Is Nothing Then Throw New ArgumentNullException(NameOf(c))

            Executar(
                "INSERT INTO disclosure_log (request_id, capability_id, stage, " &
                "  activation_id, activation_version, operation, provider, endpoint, " &
                "  model, payload_hash, payload_bytes, message_count, at, reason) " &
                "VALUES ($r, $c, $s, $ai, $av, $op, $p, $e, $m, $h, $b, $n, $t, NULL)",
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
                ("$t", CObj(Instante(quando))))
        End Sub

        Public Sub Iniciando(requestId As Guid, quando As DateTimeOffset) _
                             Implements IDisclosureJournal.Iniciando
            ' So sai de Intencionada. Reiniciar um envio ja em voo, ou ja
            ' concluido, seria reabrir uma janela que ja fechou.
            Avancar(requestId, DisclosureStage.EmVoo, quando, Nothing,
                    "stage = '" & DisclosureStage.Intencionada.ToString() & "'")
        End Sub

        Public Sub Concluir(requestId As Guid, quando As DateTimeOffset) _
                            Implements IDisclosureJournal.Concluir
            Avancar(requestId, DisclosureStage.Concluida, quando, Nothing,
                    "stage = '" & DisclosureStage.EmVoo.ToString() & "'")
        End Sub

        Public Sub Falhar(requestId As Guid, quando As DateTimeOffset, motivo As String,
                          podeTerChegado As Boolean) Implements IDisclosureJournal.Falhar
            ' Uma vez EM VOO, falhar e SEMPRE ambiguo — mesmo que o chamador
            ' jure que nao chegou. Ele nao pode saber: entre "a conexao caiu" e
            ' "a conexao caiu depois de o servidor ler o corpo" nao ha
            ' diferenca observavel deste lado.
            '
            ' NaoEnviada so vale a partir de Intencionada, onde a transmissao
            ' comprovadamente nao tinha comecado. E a guarda mora no WHERE, e
            ' nao num If antes: duas execucoes podem estar escrevendo, e o
            ' banco e quem arbitra.
            If podeTerChegado Then
                Avancar(requestId, DisclosureStage.Ambigua, quando, motivo,
                        "stage IN ('" & DisclosureStage.Intencionada.ToString() & "','" &
                        DisclosureStage.EmVoo.ToString() & "')")
            Else
                Avancar(requestId, DisclosureStage.NaoEnviada, quando, motivo,
                        "stage = '" & DisclosureStage.Intencionada.ToString() & "'")
                ' Se estava EM VOO, o UPDATE acima nao pegou nada — e ai o
                ' desfecho honesto e ambiguo, nao "nao enviou".
                Avancar(requestId, DisclosureStage.Ambigua, quando, motivo,
                        "stage = '" & DisclosureStage.EmVoo.ToString() & "'")
            End If
        End Sub

        Public Sub NaoEnviou(requestId As Guid, quando As DateTimeOffset, motivo As String) _
                             Implements IDisclosureJournal.NaoEnviou
            Avancar(requestId, DisclosureStage.NaoEnviada, quando, motivo,
                    "stage = '" & DisclosureStage.Intencionada.ToString() & "'")
        End Sub

        ' ==============================================================

        ''' <summary>
        ''' A reconciliação da abertura — o quinto passo.
        '''
        ''' O que ficou <b>em voo</b> de uma execução morta vira ambíguo: os
        ''' bytes podem ter chegado, e ninguém vai saber. O que ficou só como
        ''' intenção vira não-enviada, porque ali a transmissão não tinha
        ''' começado.
        '''
        ''' Devolve quantas viraram ambíguas, e esse número é para a UI mostrar.
        ''' "Pode ter saído conteúdo desta caixa e não dá para saber" não é
        ''' detalhe de log.
        ''' </summary>
        Public Function Reconciliar(quando As DateTimeOffset) As Integer _
                                    Implements IDisclosureJournal.Reconciliar
            Dim ambiguas = Executar(
                "UPDATE disclosure_log SET stage = $novo, at = $t, " &
                "  reason = COALESCE(reason, 'processo terminou em voo') " &
                "WHERE stage = $velho",
                ("$novo", CObj(DisclosureStage.Ambigua.ToString())),
                ("$velho", CObj(DisclosureStage.EmVoo.ToString())),
                ("$t", CObj(Instante(quando))))

            Executar(
                "UPDATE disclosure_log SET stage = $novo, at = $t, " &
                "  reason = COALESCE(reason, 'processo terminou antes de transmitir') " &
                "WHERE stage = $velho",
                ("$novo", CObj(DisclosureStage.NaoEnviada.ToString())),
                ("$velho", CObj(DisclosureStage.Intencionada.ToString())),
                ("$t", CObj(Instante(quando))))

            Return ambiguas
        End Function

        Public Function Ler(quantas As Integer) As IReadOnlyList(Of DisclosureEntry) _
                            Implements IDisclosureJournal.Ler
            Dim saida As New List(Of DisclosureEntry)()
            Using cmd = _db.Connection.CreateCommand()
                cmd.CommandText =
                    "SELECT request_id, capability_id, stage, activation_id, " &
                    "       activation_version, operation, provider, endpoint, model, " &
                    "       payload_hash, payload_bytes, message_count, at, reason " &
                    "FROM disclosure_log ORDER BY at DESC, request_id DESC LIMIT $n"
                cmd.Parameters.AddWithValue("$n", quantas)
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        saida.Add(New DisclosureEntry(
                            Guid.Parse(r.GetString(0)), Guid.Parse(r.GetString(1)),
                            CType([Enum].Parse(GetType(DisclosureStage), r.GetString(2)),
                                  DisclosureStage),
                            r.GetString(3), r.GetInt32(4),
                            CType([Enum].Parse(GetType(AssistOperation), r.GetString(5)),
                                  AssistOperation),
                            r.GetString(6), r.GetString(7), r.GetString(8),
                            If(r.IsDBNull(9), Nothing, r.GetString(9)),
                            r.GetInt32(10), r.GetInt32(11),
                            DateTimeOffset.Parse(r.GetString(12), CultureInfo.InvariantCulture),
                            If(r.IsDBNull(13), Nothing, r.GetString(13))))
                    End While
                End Using
            End Using
            Return saida
        End Function

        ' ==============================================================

        Private Sub Avancar(requestId As Guid, destino As DisclosureStage,
                            quando As DateTimeOffset, motivo As String, guarda As String)
            Executar(
                "UPDATE disclosure_log SET stage = $s, at = $t, " &
                "  reason = COALESCE($m, reason) " &
                "WHERE request_id = $r AND " & guarda,
                ("$s", CObj(destino.ToString())),
                ("$t", CObj(Instante(quando))),
                ("$m", If(motivo Is Nothing, CObj(DBNull.Value), CObj(motivo))),
                ("$r", CObj(requestId.ToString("D"))))
        End Sub

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

        ''' <summary>ISO 8601 com offset, para ordenar como texto e não mentir.</summary>
        Private Shared Function Instante(q As DateTimeOffset) As String
            Return q.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture)
        End Function

    End Class

End Namespace

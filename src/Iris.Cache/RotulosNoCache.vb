Imports System.Collections.Generic
Imports System.Globalization
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
    ''' <b>GRAVAR É POR LOTE, E O LOTE É ATÔMICO</b>
    '''
    ''' Ou entram todos os rótulos daquele lote, ou nenhum. Meia gravação
    ''' deixaria a pasta com uma parte classificada e outra não, sem nada dizendo
    ''' qual é qual — e a varredura seguinte acharia que aquilo já foi feito.
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
        ''' <returns>Quantos rótulos entraram.</returns>
        Public Function Gravar(folderKey As Long, geracao As Long,
                               ativacao As String,
                               quando As DateTimeOffset,
                               rotulos As IReadOnlyDictionary(Of String, String),
                               confiancas As IReadOnlyDictionary(Of String, Double)) As Integer

            If rotulos Is Nothing OrElse rotulos.Count = 0 Then Return 0

            Dim gravados = 0
            Using tx = _conn.BeginTransaction()
                For Each par In rotulos
                    Dim incarnation = IncarnationDe(tx, folderKey, par.Key)
                    If Not incarnation.HasValue Then Continue For

                    Dim confianca As Double
                    If confiancas Is Nothing OrElse
                       Not confiancas.TryGetValue(par.Key, confianca) Then confianca = 0

                    ' SUBSTITUI, e nao acrescenta: o indice e unico por
                    ' (encarnacao, geracao), e reclassificar a mesma geracao e
                    ' uma correcao, nao um segundo fato.
                    Executar(tx,
                        "INSERT INTO label_observation (incarnation_key, generation_key, " &
                        "  label, confidence, activation_id, observed_at) " &
                        "VALUES ($i,$g,$l,$c,$a,$q) " &
                        "ON CONFLICT (incarnation_key, generation_key) DO UPDATE SET " &
                        "  label = excluded.label, confidence = excluded.confidence, " &
                        "  activation_id = excluded.activation_id, " &
                        "  observed_at = excluded.observed_at",
                        ("$i", CObj(incarnation.Value)), ("$g", CObj(geracao)),
                        ("$l", CObj(par.Value)), ("$c", CObj(confianca)),
                        ("$a", CObj(ativacao)),
                        ("$q", CObj(quando.ToString("o", CultureInfo.InvariantCulture))))
                    gravados += 1
                Next
                tx.Commit()
            End Using

            Return gravados
        End Function

        ''' <summary>
        ''' Os rótulos <b>da geração publicada</b> de uma pasta, por
        ''' <c>provider_entry_id</c>.
        '''
        ''' Pasta sem geração publicada devolve vazio — e vazio aqui quer dizer
        ''' "não há rótulo publicado", que é diferente de "não há rótulo". A
        ''' distinção é a mesma do resto do acervo, e quem mostra a fila precisa
        ''' dela.
        ''' </summary>
        Public Function Publicados(folderKey As Long) _
                        As IReadOnlyDictionary(Of String, String)

            Dim achados As New Dictionary(Of String, String)(StringComparer.Ordinal)

            Using cmd = _conn.CreateCommand()
                cmd.CommandText =
                    "SELECT i.provider_entry_id, l.label " &
                    "FROM folder f " &
                    "JOIN incarnation i ON i.folder_key = f.folder_key " &
                    "JOIN label_observation l ON l.incarnation_key = i.incarnation_key " &
                    "  AND l.generation_key = f.published_generation_key " &
                    "WHERE f.folder_key = $f AND f.published_generation_key IS NOT NULL"
                cmd.Parameters.AddWithValue("$f", folderKey)

                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        achados(rd.GetString(0)) = rd.GetString(1)
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

End Namespace

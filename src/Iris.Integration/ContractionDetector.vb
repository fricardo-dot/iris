Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Cache
Imports Iris.Sync
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Integration

    Public Enum ContractionVerdict
        ''' <summary>Não há com o que comparar — primeira geração da pasta.</summary>
        SemReferencia
        ''' <summary>O alcance não encolheu.</summary>
        Estavel
        ''' <summary>O alcance ENCOLHEU. Alguma coisa mudou no ambiente.</summary>
        Encolheu
    End Enum

    Public NotInheritable Class ContractionReport
        Public ReadOnly Property Verdict As ContractionVerdict
        Public ReadOnly Property AlcanceAntes As Integer
        Public ReadOnly Property AlcanceAgora As Integer
        ''' <summary>Chaves que a geração anterior via e esta não vê.</summary>
        Public ReadOnly Property Sumiram As IReadOnlyList(Of String)
        Public ReadOnly Property Chegaram As Integer

        Friend Sub New(v As ContractionVerdict, antes As Integer, agora As Integer,
                       sumiram As IEnumerable(Of String), chegaram As Integer)
            Verdict = v
            AlcanceAntes = antes
            AlcanceAgora = agora
            Me.Sumiram = If(sumiram, Enumerable.Empty(Of String)()).ToList()
            Me.Chegaram = chegaram
        End Sub

        ''' <summary>
        ''' O que dizer ao usuário. <c>Nothing</c> quando não há o que dizer.
        ''' </summary>
        Public ReadOnly Property Aviso As String
            Get
                If Verdict <> ContractionVerdict.Encolheu Then Return Nothing
                Return $"O Iris passou a enxergar menos nesta pasta: {AlcanceAntes} itens " &
                       $"na varredura anterior, {AlcanceAgora} agora. " &
                       "Alguma coisa mudou no seu Outlook — a janela de sincronização, " &
                       "ou o que o cache local guarda. As mensagens que sumiram da lista " &
                       "não foram necessariamente apagadas."
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Detecta que o alcance do Iris <b>encolheu</b> numa pasta.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO EXISTE, E O QUE ELE NÃO É</b>
    '''
    ''' A §18.4 estabeleceu que a janela de sincronização faz parte da
    ''' identidade do ambiente: mudá-la muda <b>o que existe</b>, e nenhuma
    ''' conclusão de ausência sobrevive à mudança. A §22.4 então mediu que a
    ''' janela <b>não é legível</b> — nem pelo OOM nem pelo registro do perfil.
    ''' E a §22.11 mediu que a contagem do servidor também não vem por
    ''' <c>PropertyAccessor</c>.
    '''
    ''' Sem fonte externa, sobra o que o próprio Iris já guardou: se a
    ''' varredura de hoje enxerga menos do que a de ontem <b>na mesma
    ''' pasta</b>, alguma coisa encolheu.
    '''
    ''' <b>Isto AVISA, nunca AUTORIZA.</b> A distinção não é retórica, e a
    ''' §22.11 registra que eu já escrevi a versão inflada disto uma vez —
    ''' "a cobertura passa a se acumular com o tempo", que era transformar uma
    ''' derrota em recurso. O que se acumula é histórico de alcance observado.
    ''' Cobertura exige referência externa do universo, e essa referência
    ''' continua sem fonte.
    '''
    ''' Os buracos, e são estruturais:
    '''
    '''   - <b>A janela sempre foi pequena.</b> Se o Iris nunca viu o que
    '''     falta, não há contração para detectar — e é o estado desta caixa.
    '''   - <b>Encolhimento compensado</b> por correio novo: a contagem não se
    '''     mexe. Por isso este detector compara <b>conjuntos</b>, não
    '''     contagens — o que fecha esse buraco em particular, e só ele.
    '''
    '''   - Um item que sumiu não se distingue entre excluído, movido e saído
    '''     da janela. O detector diz que sumiu; não diz por quê, e não deve.
    '''
    ''' E o efeito é <b>aviso</b>: o <c>ManifestReader</c> o consulta e a
    ''' contração entra na ressalva do manifesto. Não mexe em época, cobertura
    ''' nem associação — e não precisa: em cached a cobertura já é sempre
    ''' parcial e ausência já é proibida (§23). Não há conclusão a invalidar.
    ''' </summary>
    Public NotInheritable Class ContractionDetector

        Private ReadOnly _conn As SqliteConnection

        ''' <summary>
        ''' A trava do arquivo. A conexão é compartilhada, e uma
        ''' <c>SqliteConnection</c> não tem contrato de uso simultâneo — ver
        ''' <c>CacheWriter._trava</c>, que tem o caso por extenso.
        ''' </summary>
        Private ReadOnly _trava As Object

        Public Sub New(db As CacheDatabase)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _conn = db.Connection
            _trava = db.Trava
        End Sub

        ''' <summary>
        ''' Compara as duas últimas gerações publicadas da pasta.
        ''' </summary>
        Public Function Comparar(folderKey As Long) As ContractionReport
            SyncLock _trava
                Dim geracoes = UltimasDuas(folderKey)
                If geracoes.Count < 2 Then
                    Return New ContractionReport(ContractionVerdict.SemReferencia, 0, 0, Nothing, 0)
                End If

                Dim agora = ChavesVistas(folderKey, geracoes(0))
                Dim antes = ChavesVistas(folderKey, geracoes(1))

                Dim sumiram = antes.Except(agora).ToList()
                Dim chegaram = agora.Except(antes).Count()

                ' Compara CONJUNTOS, e nao contagens. Encolhimento compensado por
                ' correio novo mantem a contagem igual e passaria batido - foi o
                ' buraco que a §22.11 listou.
                Dim v = If(sumiram.Count > 0, ContractionVerdict.Encolheu, ContractionVerdict.Estavel)
                Return New ContractionReport(v, antes.Count, agora.Count, sumiram, chegaram)
            End SyncLock
        End Function

        ' ==============================================================

        Private Function UltimasDuas(folderKey As Long) As List(Of Long)
            Dim r As New List(Of Long)()
            Using cmd = _conn.CreateCommand()
                cmd.CommandText = "SELECT generation_key FROM generation WHERE folder_key = $f " &
                                  "ORDER BY attempt_key DESC LIMIT 2"
                cmd.Parameters.AddWithValue("$f", folderKey)
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        r.Add(rd.GetInt64(0))
                    End While
                End Using
            End Using
            Return r
        End Function

        ''' <summary>
        ''' As chaves que aquela geração VIU — não o estado atual da associação.
        '''
        ''' Vem de <c>scan_stage</c>, pela tentativa que produziu a geração, e
        ''' não de <c>association</c>: a associação carrega o estado
        ''' <b>corrente</b>, e a corrente já foi sobrescrita pela geração
        ''' seguinte. Perguntar à associação o que a geração anterior via daria
        ''' sempre a resposta de agora, e o detector nunca acusaria nada.
        ''' </summary>
        Private Function ChavesVistas(folderKey As Long, geracao As Long) As HashSet(Of String)
            Dim r As New HashSet(Of String)(StringComparer.Ordinal)
            Using cmd = _conn.CreateCommand()
                cmd.CommandText =
                    "SELECT s.provider_entry_id FROM scan_stage s " &
                    "JOIN generation g ON g.attempt_key = s.attempt_key " &
                    "WHERE g.generation_key = $g AND g.folder_key = $f"
                cmd.Parameters.AddWithValue("$g", geracao)
                cmd.Parameters.AddWithValue("$f", folderKey)
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        r.Add(rd.GetString(0))
                    End While
                End Using
            End Using
            Return r
        End Function

    End Class

End Namespace

Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Core
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Cache

    Public NotInheritable Class OpenFailure
        Public ReadOnly Property Reason As String
        Public ReadOnly Property Detail As String
        Friend Sub New(reason As String, detail As String)
            Me.Reason = reason
            Me.Detail = detail
        End Sub
        Public Overrides Function ToString() As String
            Return $"{Reason}: {Detail}"
        End Function
    End Class

    ''' <summary>
    ''' Abre o cache, e RECUSA abrir um banco que não corresponda ao modelo.
    '''
    ''' A diferença entre isto e o <see cref="SchemaGate"/> importa: o gate
    ''' valida o modelo que ESTÁ NO CÓDIGO. Ele prova que o schema que eu
    ''' escrevi respeita os invariantes — não que o arquivo no disco é aquele
    ''' schema. Um banco criado por uma versão antiga, ou editado à mão,
    ''' passa no gate e não corresponde a nada.
    '''
    ''' Por isso a abertura faz as duas coisas:
    '''   1. o gate, sobre o modelo;
    '''   2. a INTROSPECÇÃO, comparando o arquivo real com o modelo.
    '''
    ''' É o mesmo padrão que a Q1 cobrou com a coluna <c>Permission</c> e a
    ''' §16.5 com o <c>Restrict</c>: "não lançou" não é "funciona".
    ''' </summary>
    Public NotInheritable Class CacheDatabase
        Implements IDisposable

        Private ReadOnly _conn As SqliteConnection
        Private _disposed As Boolean

        Public ReadOnly Property Connection As SqliteConnection
            Get
                Return _conn
            End Get
        End Property

        Private Sub New(conn As SqliteConnection)
            _conn = conn
        End Sub

        ''' <summary>
        ''' Abre ou cria. Devolve <c>Nothing</c> em <paramref name="falha"/>
        ''' quando deu certo.
        ''' </summary>
        Public Shared Function Open(caminho As String, schema As CacheSchema,
                                    ByRef falha As OpenFailure) As CacheDatabase
            falha = Nothing

            ' 1. O gate, sobre o MODELO. Se o proprio modelo viola um
            '    invariante, nem faz sentido criar banco.
            Dim violacoes = SchemaGate.Check(schema)
            If violacoes.Count > 0 Then
                falha = New OpenFailure("gate",
                    String.Join(" | ", violacoes.Select(Function(x) x.ToString())))
                Return Nothing
            End If

            Dim conn As SqliteConnection = Nothing
            Try
                conn = New SqliteConnection($"Data Source={caminho}")
                conn.Open()

                ' FK vem DESLIGADA no SQLite. Sem isto, todas as FKs do
                ' schema seriam decorativas.
                Executar(conn, "PRAGMA foreign_keys = ON")

                Dim versao = LerVersao(conn)
                Dim vazio = ContarTabelas(conn) = 0

                If vazio Then
                    Criar(conn, schema)
                ElseIf versao <> SqliteDdl.SchemaVersion Then
                    ' Versao divergente so passa por um caminho CONHECIDO.
                    '
                    ' Migrar sem saber de onde para onde continua sendo pior
                    ' que recusar -- e por isso o que nao esta em
                    ' SqliteDdl.Migracoes continua sendo recusado, com a mesma
                    ' mensagem de antes. O que mudou foi que apagar o arquivo
                    ' deixou de ser barato: o diario do egress mora aqui, e ele
                    ' nao se reconstroi do Outlook como o resto.
                    If Not Migrar(conn, versao, falha) Then
                        conn.Dispose()
                        Return Nothing
                    End If
                End If

                ' 2. INTROSPECCAO: o arquivo REAL corresponde ao modelo?
                Dim diferencas = SchemaIntrospector.Comparar(conn, schema)
                If diferencas.Count > 0 Then
                    falha = New OpenFailure("divergencia", String.Join(" | ", diferencas))
                    conn.Dispose()
                    Return Nothing
                End If

                ' 3. E as FKs estao mesmo LIGADAS?
                If Not ForeignKeysLigadas(conn) Then
                    falha = New OpenFailure("fk", "foreign_keys nao ficou ligada")
                    conn.Dispose()
                    Return Nothing
                End If

                Return New CacheDatabase(conn)

            Catch ex As Exception
                If conn IsNot Nothing Then conn.Dispose()
                falha = New OpenFailure("excecao", ex.Message)
                Return Nothing
            End Try
        End Function

        Private Shared Sub Criar(conn As SqliteConnection, schema As CacheSchema)
            Using tx = conn.BeginTransaction()
                For Each cmd In SqliteDdl.Generate(schema)
                    ' PRAGMA nao vai dentro de transacao em alguns casos;
                    ' journal_mode em particular.
                    If cmd.StartsWith("PRAGMA journal_mode", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                    Executar(conn, cmd, tx)
                Next
                tx.Commit()
            End Using
            Executar(conn, "PRAGMA journal_mode = WAL")
        End Sub

        Private Shared Sub Executar(conn As SqliteConnection, sql As String,
                                    Optional tx As SqliteTransaction = Nothing)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = sql
                If tx IsNot Nothing Then cmd.Transaction = tx
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' Sobe de <paramref name="de"/> até <see cref="SqliteDdl.SchemaVersion"/>
        ''' usando <b>só</b> passos listados. Um degrau faltando recusa tudo.
        ''' </summary>
        ''' <remarks>
        ''' Cada degrau é uma transação própria, com o <c>PRAGMA user_version</c>
        ''' <b>dentro</b> dela: morrer no meio deixa o banco na versão de onde
        ''' ele saiu, e não numa versão que ele não tem a forma de ter.
        '''
        ''' Não confere o resultado: quem confere é a introspecção que roda
        ''' logo depois, em <see cref="Open"/>, igual para banco criado e para
        ''' banco migrado.
        ''' </remarks>
        Private Shared Function Migrar(conn As SqliteConnection, de As Integer,
                                       ByRef falha As OpenFailure) As Boolean
            Dim atual = de

            ' Para tras nao ha caminho: um banco mais NOVO que o programa foi
            ' escrito por uma versao que sabia coisas que esta nao sabe.
            If atual > SqliteDdl.SchemaVersion Then
                falha = New OpenFailure("versao",
                    $"banco na versao {atual}, mais nova que a esperada " &
                    $"{SqliteDdl.SchemaVersion}")
                Return False
            End If

            While atual < SqliteDdl.SchemaVersion
                Dim passo As IReadOnlyList(Of String) = Nothing
                If Not SqliteDdl.Migracoes.TryGetValue(atual, passo) Then
                    falha = New OpenFailure("versao",
                        $"banco na versao {atual}, esperado " &
                        $"{SqliteDdl.SchemaVersion}, e nao ha migracao conhecida")
                    Return False
                End If

                Try
                    Using tx = conn.BeginTransaction()
                        For Each cmd In passo
                            Executar(conn, cmd, tx)
                        Next
                        Executar(conn, $"PRAGMA user_version = {atual + 1}", tx)
                        tx.Commit()
                    End Using
                Catch ex As Exception
                    falha = New OpenFailure("migracao",
                        $"a migracao {atual} -> {atual + 1} falhou: {ex.Message}")
                    Return False
                End Try

                atual += 1
            End While

            Return True
        End Function

        Private Shared Function LerVersao(conn As SqliteConnection) As Integer
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "PRAGMA user_version"
                Return Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Function

        Private Shared Function ContarTabelas(conn As SqliteConnection) As Integer
            Using cmd = conn.CreateCommand()
                cmd.CommandText =
                    "SELECT COUNT(*) FROM sqlite_schema WHERE type='table' AND name NOT LIKE 'sqlite_%'"
                Return Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Function

        Private Shared Function ForeignKeysLigadas(conn As SqliteConnection) As Boolean
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "PRAGMA foreign_keys"
                Return Convert.ToInt32(cmd.ExecuteScalar()) = 1
            End Using
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            _conn?.Dispose()
        End Sub

    End Class

End Namespace

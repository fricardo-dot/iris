Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Core
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Cache

    ''' <summary>
    ''' Lê o schema REAL do arquivo SQLite e compara com o modelo.
    '''
    ''' Existe porque o <see cref="SchemaGate"/> prova o que eu ESCREVI, não
    ''' o que está no disco. Um banco criado por versão antiga, ou editado à
    ''' mão, passa no gate e não corresponde a nada — e a divergência só
    ''' apareceria muito depois, num INSERT que falha ou, pior, num SELECT
    ''' que devolve menos do que deveria.
    ''' </summary>
    Public NotInheritable Class SchemaIntrospector

        Public Shared Function Comparar(conn As SqliteConnection,
                                        esperado As CacheSchema) As IReadOnlyList(Of String)
            Dim diffs As New List(Of String)()

            Dim tabelasReais = LerTabelas(conn)
            Dim esperadas = New HashSet(Of String)(
                esperado.Tables.Select(Function(t) t.Name), StringComparer.OrdinalIgnoreCase)

            For Each t In esperado.Tables
                If Not tabelasReais.Contains(t.Name) Then
                    diffs.Add($"tabela ausente no banco: {t.Name}")
                End If
            Next
            For Each nome In tabelasReais
                If Not esperadas.Contains(nome) Then
                    diffs.Add($"tabela EXTRA no banco: {nome}")
                End If
            Next
            If diffs.Count > 0 Then Return diffs

            For Each t In esperado.Tables
                CompararColunas(conn, t, diffs)
                CompararUnicos(conn, t, diffs)
                CompararFks(conn, t, esperado, diffs)
            Next

            Return diffs
        End Function

        Private Shared Function LerTabelas(conn As SqliteConnection) As HashSet(Of String)
            Dim r As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Using cmd = conn.CreateCommand()
                cmd.CommandText =
                    "SELECT name FROM sqlite_schema WHERE type='table' AND name NOT LIKE 'sqlite_%'"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        r.Add(rd.GetString(0))
                    End While
                End Using
            End Using
            Return r
        End Function

        Private Shared Sub CompararColunas(conn As SqliteConnection, t As SchemaTable,
                                           diffs As List(Of String))
            Dim reais As New Dictionary(Of String, (Tipo As String, NotNull As Boolean, Pk As Boolean))(
                StringComparer.OrdinalIgnoreCase)

            Using cmd = conn.CreateCommand()
                cmd.CommandText = $"PRAGMA table_info({t.Name})"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        reais(rd.GetString(1)) = (rd.GetString(2),
                                                  rd.GetInt32(3) = 1,
                                                  rd.GetInt32(5) > 0)
                    End While
                End Using
            End Using

            For Each c In t.Columns
                If Not reais.ContainsKey(c.Name) Then
                    diffs.Add($"{t.Name}.{c.Name} ausente no banco")
                    Continue For
                End If
                Dim real = reais(c.Name)
                If Not String.Equals(real.Tipo, c.Kind, StringComparison.OrdinalIgnoreCase) Then
                    diffs.Add($"{t.Name}.{c.Name}: tipo {real.Tipo}, esperado {c.Kind}")
                End If
                If c.IsPrimaryKey <> real.Pk Then
                    diffs.Add($"{t.Name}.{c.Name}: chave primaria {real.Pk}, esperado {c.IsPrimaryKey}")
                End If
                ' A PK do SQLite ja e implicitamente NOT NULL numa coluna
                ' INTEGER PRIMARY KEY, entao so cobro NOT NULL fora dela.
                If Not c.IsPrimaryKey AndAlso c.IsRequired <> real.NotNull Then
                    diffs.Add($"{t.Name}.{c.Name}: NOT NULL {real.NotNull}, esperado {c.IsRequired}")
                End If
            Next

            For Each nome In reais.Keys
                If t.Column(nome) Is Nothing Then
                    diffs.Add($"{t.Name}.{nome} EXTRA no banco")
                End If
            Next
        End Sub

        Private Shared Sub CompararUnicos(conn As SqliteConnection, t As SchemaTable,
                                          diffs As List(Of String))
            Dim reais As New List(Of String())()

            Using cmd = conn.CreateCommand()
                cmd.CommandText = $"PRAGMA index_list({t.Name})"
                Dim indices As New List(Of String)()
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        If rd.GetInt32(2) = 1 Then indices.Add(rd.GetString(1))
                    End While
                End Using
                For Each ix In indices
                    reais.Add(ColunasDoIndice(conn, ix))
                Next
            End Using

            For Each esperado In t.UniqueIndexes
                Dim achou = reais.Any(Function(r) r.Length = esperado.Length AndAlso
                                                  r.SequenceEqual(esperado, StringComparer.OrdinalIgnoreCase))
                If Not achou Then
                    diffs.Add($"{t.Name}: falta indice UNICO sobre ({String.Join(", ", esperado)})")
                End If
            Next
        End Sub

        Private Shared Function ColunasDoIndice(conn As SqliteConnection, indice As String) As String()
            Dim cols As New List(Of String)()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = $"PRAGMA index_info({indice})"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        If Not rd.IsDBNull(2) Then cols.Add(rd.GetString(2))
                    End While
                End Using
            End Using
            Return cols.ToArray()
        End Function

        ''' <summary>
        ''' Compara a FK inteira: tabela alvo E COLUNA alvo.
        '''
        ''' A coluna estava de fora, e isso deixou passar um schema quebrado:
        ''' <c>scan_attempt.environment_key</c> apontava para
        ''' <c>environment_profile(environment_profile_key)</c>, coluna que não
        ''' existe. Como a TABELA alvo estava certa, a comparação dizia "sem
        ''' divergência" — e o SQLite aceita o CREATE e só reclama no primeiro
        ''' INSERT, com "foreign key mismatch". Onze testes passaram por cima.
        ''' </summary>
        Private Shared Sub CompararFks(conn As SqliteConnection, t As SchemaTable,
                                       esperado As CacheSchema, diffs As List(Of String))
            Dim reais As New Dictionary(Of String, (Tabela As String, Coluna As String))(
                StringComparer.OrdinalIgnoreCase)

            Using cmd = conn.CreateCommand()
                cmd.CommandText = $"PRAGMA foreign_key_list({t.Name})"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        ' 2 = tabela alvo, 3 = coluna de origem, 4 = coluna alvo
                        Dim destino = If(rd.IsDBNull(4), Nothing, rd.GetString(4))
                        reais(rd.GetString(3)) = (rd.GetString(2), destino)
                    End While
                End Using
            End Using

            For Each c In t.Columns
                If c.References Is Nothing Then Continue For
                If Not reais.ContainsKey(c.Name) Then
                    diffs.Add($"{t.Name}.{c.Name}: FK ausente no banco (esperava -> {c.References})")
                    Continue For
                End If
                Dim real = reais(c.Name)
                If Not String.Equals(real.Tabela, c.References, StringComparison.OrdinalIgnoreCase) Then
                    diffs.Add($"{t.Name}.{c.Name}: FK aponta para {real.Tabela}, esperado {c.References}")
                    Continue For
                End If
                Dim colunaEsperada = SqliteDdl.ChavePrimariaDe(esperado, c.References)
                If real.Coluna IsNot Nothing AndAlso
                   Not String.Equals(real.Coluna, colunaEsperada, StringComparison.OrdinalIgnoreCase) Then
                    diffs.Add($"{t.Name}.{c.Name}: FK aponta para {real.Tabela}.{real.Coluna}, " &
                              $"esperado {c.References}.{colunaEsperada}")
                End If
            Next
        End Sub

    End Class

End Namespace

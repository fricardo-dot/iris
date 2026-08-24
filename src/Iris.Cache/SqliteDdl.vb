Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Iris.Core

Namespace Global.Iris.Cache

    ''' <summary>
    ''' Gera o DDL a partir do <see cref="CacheSchema"/>.
    '''
    ''' O DDL não é escrito à mão em lugar nenhum. Se fosse, o modelo que os
    ''' testes de invariante leem e o banco que o programa abre seriam duas
    ''' coisas que podem divergir — e a Q1 já cobrou esse desenho quando o
    ''' teste sintético verificava um algoritmo diferente do que rodava
    ''' contra o Outlook.
    ''' </summary>
    Public NotInheritable Class SqliteDdl

        ''' <summary>
        ''' Versão do schema gravada em <c>PRAGMA user_version</c>. Abrir um
        ''' banco de versão diferente FALHA FECHADO.
        ''' </summary>
        Public Const SchemaVersion As Integer = 1

        Public Shared Function Generate(schema As CacheSchema) As IReadOnlyList(Of String)
            Dim comandos As New List(Of String)()

            ' FK precisa ser LIGADA por conexão — no SQLite ela é opcional e
            ' vem desligada. Um schema com FK declarada e desligada dá a
            ' impressão de integridade sem tê-la.
            comandos.Add("PRAGMA foreign_keys = ON")
            ' WAL: crash no meio de uma transação não corrompe o banco.
            comandos.Add("PRAGMA journal_mode = WAL")

            For Each t In OrdenarPorDependencia(schema)
                comandos.Add(CriarTabela(t))
            Next

            For Each t In schema.Tables
                Dim n = 0
                For Each ix In t.UniqueIndexes
                    n += 1
                    comandos.Add($"CREATE UNIQUE INDEX ux_{t.Name}_{n} ON {t.Name} ({String.Join(", ", ix)})")
                Next
                n = 0
                For Each ix In t.Indexes
                    n += 1
                    comandos.Add($"CREATE INDEX ix_{t.Name}_{n} ON {t.Name} ({String.Join(", ", ix)})")
                Next
            Next

            comandos.Add($"PRAGMA user_version = {SchemaVersion}")
            Return comandos
        End Function

        ''' <summary>
        ''' Tabelas na ordem em que dá para criar sem FK apontando para o
        ''' vazio. Há ciclo real no modelo — <c>folder</c> aponta para
        ''' <c>generation</c> e vice-versa —, e nesse caso a ordem não
        ''' resolve: o SQLite aceita FK para tabela ainda inexistente desde
        ''' que ela exista na hora do INSERT.
        ''' </summary>
        Private Shared Function OrdenarPorDependencia(schema As CacheSchema) As IReadOnlyList(Of SchemaTable)
            Dim restantes = schema.Tables.ToList()
            Dim ordenadas As New List(Of SchemaTable)()
            Dim jaCriadas As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            While restantes.Count > 0
                Dim prontas = restantes.Where(
                    Function(t) t.Columns.All(
                        Function(c) c.References Is Nothing OrElse
                                    jaCriadas.Contains(c.References) OrElse
                                    String.Equals(c.References, t.Name, StringComparison.OrdinalIgnoreCase))).ToList()

                If prontas.Count = 0 Then
                    ' Ciclo. Cria o resto na ordem em que veio: o SQLite nao
                    ' valida o alvo da FK no CREATE.
                    ordenadas.AddRange(restantes)
                    Exit While
                End If

                For Each t In prontas
                    ordenadas.Add(t)
                    jaCriadas.Add(t.Name)
                    restantes.Remove(t)
                Next
            End While

            Return ordenadas
        End Function

        Private Shared Function CriarTabela(t As SchemaTable) As String
            Dim sb As New StringBuilder()
            sb.Append($"CREATE TABLE {t.Name} (")
            Dim partes As New List(Of String)()

            For Each c In t.Columns
                Dim p As New StringBuilder()
                p.Append($"{c.Name} {c.Kind}")
                If c.IsPrimaryKey Then p.Append(" PRIMARY KEY")
                If c.IsRequired AndAlso Not c.IsPrimaryKey Then p.Append(" NOT NULL")
                If c.Check IsNot Nothing Then p.Append($" CHECK ({c.Check})")
                partes.Add(p.ToString())
            Next

            For Each c In t.Columns
                If c.References Is Nothing Then Continue For
                Dim acao = AcaoSql(c.OnDelete)
                Dim alvo = ChavePrimariaDe(c.References)
                partes.Add($"FOREIGN KEY ({c.Name}) REFERENCES {c.References} ({alvo}){acao}")
            Next

            For Each chk In t.Checks
                partes.Add($"CHECK ({chk})")
            Next

            sb.Append(String.Join(", ", partes))
            sb.Append(")")
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Convenção do modelo: a chave primária de <c>x</c> é <c>x_key</c>.
        ''' </summary>
        Private Shared Function ChavePrimariaDe(tabela As String) As String
            Return tabela & "_key"
        End Function

        Private Shared Function AcaoSql(a As DeleteAction) As String
            Select Case a
                Case DeleteAction.Cascade : Return " ON DELETE CASCADE"
                Case DeleteAction.Restrict : Return " ON DELETE RESTRICT"
                Case DeleteAction.SetNull : Return " ON DELETE SET NULL"
                Case Else : Return ""
            End Select
        End Function

    End Class

End Namespace

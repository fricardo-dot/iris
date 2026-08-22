Imports System.Diagnostics
Imports System.Text

Namespace Checks

    Public Enum CheckStatus
        Pass
        Fail
        Warn
        Info
        Skipped
    End Enum

    Public NotInheritable Class CheckResult
        Public Property Id As String
        Public Property Group As String
        Public Property Title As String
        Public Property Status As CheckStatus
        Public Property Notes As String
        Public Property Elapsed As TimeSpan

        Public ReadOnly Property Marker As String
            Get
                Select Case Status
                    Case CheckStatus.Pass : Return "[ OK ]"
                    Case CheckStatus.Fail : Return "[FALHA]"
                    Case CheckStatus.Warn : Return "[AVISO]"
                    Case CheckStatus.Info : Return "[INFO]"
                    Case Else : Return "[PULADO]"
                End Select
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Executa os critérios de aceitação da Fase 0 e produz o relatório.
    '''
    ''' O relatório é o entregável real do spike: o código aqui é
    ''' descartável, as descobertas não. Elas alimentam o desenho das Fases
    ''' 1 e 2 (identidade estável, estados de conteúdo, reconexão).
    ''' </summary>
    Public NotInheritable Class CheckRunner

        Private ReadOnly _results As New List(Of CheckResult)()

        Public ReadOnly Property Results As IReadOnlyList(Of CheckResult)
            Get
                Return _results
            End Get
        End Property

        Public Async Function RunAsync(id As String,
                                       group As String,
                                       title As String,
                                       body As Func(Of Task(Of (Status As CheckStatus, Notes As String)))) As Task(Of CheckResult)
            Dim sw = Stopwatch.StartNew()
            Dim status As CheckStatus
            Dim notes As String

            Try
                Dim outcome = Await body()
                status = outcome.Status
                notes = outcome.Notes
            Catch ex As Exception
                status = CheckStatus.Fail
                notes = $"{ex.GetType().Name}: {ex.Message}"
            End Try

            sw.Stop()

            Dim result As New CheckResult With {
                .Id = id,
                .Group = group,
                .Title = title,
                .Status = status,
                .Notes = notes,
                .Elapsed = sw.Elapsed
            }
            _results.Add(result)
            Report(result)
            Return result
        End Function

        Public Sub Skip(id As String, group As String, title As String, reason As String)
            Dim result As New CheckResult With {
                .Id = id, .Group = group, .Title = title,
                .Status = CheckStatus.Skipped, .Notes = reason
            }
            _results.Add(result)
            Report(result)
        End Sub

        Private Shared Sub Report(result As CheckResult)
            Dim original = Console.ForegroundColor
            Console.ForegroundColor = ColorFor(result.Status)
            Console.Write($"  {result.Marker,-8}")
            Console.ForegroundColor = original
            Console.Write($" {result.Id}  {result.Title}")
            If result.Elapsed > TimeSpan.Zero Then
                Console.Write($"  ({result.Elapsed.TotalMilliseconds:0} ms)")
            End If
            Console.WriteLine()
            If Not String.IsNullOrWhiteSpace(result.Notes) Then
                Console.WriteLine($"           {result.Notes}")
            End If
        End Sub

        Private Shared Function ColorFor(status As CheckStatus) As ConsoleColor
            Select Case status
                Case CheckStatus.Pass : Return ConsoleColor.Green
                Case CheckStatus.Fail : Return ConsoleColor.Red
                Case CheckStatus.Warn : Return ConsoleColor.Yellow
                Case CheckStatus.Skipped : Return ConsoleColor.DarkGray
                Case Else : Return ConsoleColor.Cyan
            End Select
        End Function

        Public Function Count(status As CheckStatus) As Integer
            Return Enumerable.Count(_results, Function(r) r.Status = status)
        End Function

        Public Function BuildReport(environmentNotes As IEnumerable(Of String)) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("# Iris — Relatório da Fase 0")
            sb.AppendLine()
            sb.AppendLine($"**Execução:** {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            sb.AppendLine($"**Máquina:** {Environment.MachineName}")
            sb.AppendLine($"**SO:** {Environment.OSVersion.VersionString}")
            sb.AppendLine($"**Runtime:** {Runtime.InteropServices.RuntimeInformation.FrameworkDescription}")
            sb.AppendLine($"**Processo:** {Runtime.InteropServices.RuntimeInformation.ProcessArchitecture}")
            sb.AppendLine()

            sb.AppendLine("## Resumo")
            sb.AppendLine()
            sb.AppendLine("| Resultado | Quantidade |")
            sb.AppendLine("|---|---|")
            sb.AppendLine($"| Passou | {Count(CheckStatus.Pass)} |")
            sb.AppendLine($"| Falhou | {Count(CheckStatus.Fail)} |")
            sb.AppendLine($"| Aviso | {Count(CheckStatus.Warn)} |")
            sb.AppendLine($"| Informativo | {Count(CheckStatus.Info)} |")
            sb.AppendLine($"| Pulado | {Count(CheckStatus.Skipped)} |")
            sb.AppendLine()

            sb.AppendLine("## Ambiente")
            sb.AppendLine()
            For Each note In environmentNotes
                sb.AppendLine($"- {note}")
            Next
            sb.AppendLine()

            For Each group In _results.GroupBy(Function(r) r.Group)
                sb.AppendLine($"## {group.Key}")
                sb.AppendLine()
                sb.AppendLine("| | Critério | Resultado | Observação |")
                sb.AppendLine("|---|---|---|---|")
                For Each r In group
                    Dim notes = If(r.Notes, "").Replace("|", "\|").Replace(vbCrLf, " ").Replace(vbLf, " ")
                    sb.AppendLine($"| {r.Id} | {r.Title} | {StatusWord(r.Status)} | {notes} |")
                Next
                sb.AppendLine()
            Next

            Return sb.ToString()
        End Function

        Private Shared Function StatusWord(status As CheckStatus) As String
            Select Case status
                Case CheckStatus.Pass : Return "**passou**"
                Case CheckStatus.Fail : Return "**FALHOU**"
                Case CheckStatus.Warn : Return "aviso"
                Case CheckStatus.Info : Return "info"
                Case Else : Return "pulado"
            End Select
        End Function
    End Class

End Namespace

Imports System.IO
Imports System.Runtime.InteropServices
Imports IrisSpike.Broker
Imports IrisSpike.Checks
Imports IrisSpike.Interop

''' <summary>
''' Spike da Fase 0. Descartável por desenho: o entregável é o relatório,
''' não este código.
'''
''' O grupo A (broker) roda em qualquer máquina. Os grupos B a G exigem o
''' Outlook clássico instalado e EM EXECUÇÃO; sem ele são marcados como
''' pulados, com o motivo, para o relatório não dar falsa impressão de
''' cobertura.
''' </summary>
Module Program

    Private ReadOnly PendingGroups As (Id As String, Title As String)() = {
        ("B", "Leitura — stores, pastas, itens heterogêneos, corpo e anexo"),
        ("C", "Envio — rascunho, Display(), Send() real e entrega confirmada"),
        ("D", "Eventos — ItemAdd/Change/Remove, movimentos, mudanças offline"),
        ("E", "Ciclo de vida — reconexão, Iris antes do Outlook, RPC rejeitado"),
        ("F", "Desempenho — Restrict/Sort, 100 e 1.000 itens, corpo e anexos"),
        ("G", "Ambiente — bitness, COMReference vs pacote, mensagens protegidas")
    }

    Function Main(args As String()) As Integer
        Return MainAsync(args).GetAwaiter().GetResult()
    End Function

    Private Async Function MainAsync(args As String()) As Task(Of Integer)
        Console.OutputEncoding = Text.Encoding.UTF8
        Banner()

        Dim runner As New CheckRunner()
        Dim env = InspectEnvironment()

        Console.WriteLine("Ambiente")
        For Each note In env.Notes
            Console.WriteLine($"  - {note}")
        Next
        Console.WriteLine()

        ' ---- Grupo A: sempre roda -------------------------------------
        Console.WriteLine("A — Broker")
        Dim broker As New OutlookBroker()
        broker.Start()

        Dim brokerChecks As New BrokerChecks(runner, broker)
        Await brokerChecks.RunAsync()

        ' ---- Conexão com o Outlook ------------------------------------
        Console.WriteLine()
        Console.WriteLine("Conexão")

        Await runner.RunAsync(
            "A9", "A — Broker", "Anexar a uma instância do Outlook em execução",
            Async Function()
                Dim state = Await broker.ConnectAsync()
                Select Case state
                    Case SessionState.Connected
                        Return (CheckStatus.Pass, "Anexado ao Outlook em execução.")
                    Case SessionState.Unavailable
                        If Not env.OutlookInstalled Then
                            Return (CheckStatus.Skipped,
                                    "Outlook clássico não instalado nesta máquina.")
                        End If
                        Return (CheckStatus.Skipped,
                                "Outlook instalado, porém não está em execução. " &
                                "Abra o Outlook e rode de novo.")
                    Case Else
                        Return (CheckStatus.Fail, $"Estado inesperado: {state}.")
                End Select
            End Function)

        Dim connected = broker.State = SessionState.Connected

        ' ---- Grupos B a G ---------------------------------------------
        Console.WriteLine()
        Console.WriteLine("B a G — dependem do Outlook")

        Dim reason = If(env.OutlookInstalled,
                        "Outlook instalado mas não está em execução.",
                        "Outlook clássico não instalado nesta máquina.")

        For Each group In PendingGroups
            If connected Then
                runner.Skip(group.Id, $"{group.Id} — dependente do Outlook", group.Title,
                            "Ainda não implementado — próximo incremento do spike.")
            Else
                runner.Skip(group.Id, $"{group.Id} — dependente do Outlook", group.Title, reason)
            End If
        Next

        ' ---- Encerramento ---------------------------------------------
        Console.WriteLine()
        Console.WriteLine("Encerramento")
        Await brokerChecks.RunShutdownCheckAsync()
        broker.Dispose()

        ' ---- Relatório -------------------------------------------------
        Dim reportPath = Path.Combine(ReportDirectory(), "relatorio-fase0.md")
        File.WriteAllText(reportPath, runner.BuildReport(env.Notes), Text.Encoding.UTF8)

        Console.WriteLine()
        Console.WriteLine(New String("-"c, 64))
        Console.WriteLine($"passou {runner.Count(CheckStatus.Pass)}   " &
                          $"falhou {runner.Count(CheckStatus.Fail)}   " &
                          $"aviso {runner.Count(CheckStatus.Warn)}   " &
                          $"pulado {runner.Count(CheckStatus.Skipped)}")
        Console.WriteLine($"relatório: {reportPath}")
        Console.WriteLine()

        Return If(runner.Count(CheckStatus.Fail) > 0, 1, 0)
    End Function

    Private Sub Banner()
        Console.WriteLine("Iris — spike da Fase 0")
        Console.WriteLine(New String("="c, 64))
        Console.WriteLine()
    End Sub

    Private Function ReportDirectory() As String
        ' bin/Debug/net10.0-windows -> sobe até a pasta spike/
        Dim dir As New DirectoryInfo(AppContext.BaseDirectory)
        While dir IsNot Nothing AndAlso Not String.Equals(dir.Name, "spike", StringComparison.OrdinalIgnoreCase)
            dir = dir.Parent
        End While
        Return If(dir?.FullName, AppContext.BaseDirectory)
    End Function

    Private Function InspectEnvironment() As (Notes As List(Of String), OutlookInstalled As Boolean)
        Dim notes As New List(Of String)()

        notes.Add($"Runtime: {RuntimeInformation.FrameworkDescription}")
        notes.Add($"Arquitetura do processo: {RuntimeInformation.ProcessArchitecture} " &
                  $"(SO: {RuntimeInformation.OSArchitecture})")

        ' R12: registrar, não assumir. O Outlook é servidor COM fora do
        ' processo e faz marshaling entre arquiteturas.
        Dim progIdType = Type.GetTypeFromProgID("Outlook.Application")
        Dim installed = progIdType IsNot Nothing
        notes.Add($"ProgID Outlook.Application: {If(installed, "registrado", "AUSENTE")}")

        Dim running = Diagnostics.Process.GetProcessesByName("OUTLOOK")
        Try
            If running.Length > 0 Then
                notes.Add($"OUTLOOK.EXE em execução: {running.Length} processo(s)")
            Else
                notes.Add("OUTLOOK.EXE em execução: nenhum")
            End If
        Finally
            For Each p In running
                p.Dispose()
            Next
        End Try

        If Not installed Then
            notes.Add("Sem Outlook clássico: grupos B a G ficam pulados. " &
                      "O grupo A valida a arquitetura mesmo assim.")
        End If

        Return (notes, installed)
    End Function

End Module

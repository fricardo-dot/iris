Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports IrisSpike.Broker
Imports IrisSpike.Interop
Imports IrisSpike.Model
Imports Outlook = Microsoft.Office.Interop.Outlook

Namespace Checks

    ''' <summary>
    ''' Grupos B, E, F e G da Fase 0 — exigem Outlook clássico em execução.
    '''
    ''' Todo acesso COM passa pelo broker. Nenhuma linha aqui toca um RCW
    ''' fora da thread dele: as lambdas rodam lá dentro e devolvem DTOs.
    ''' </summary>
    Public NotInheritable Class OutlookChecks

        Private ReadOnly _runner As CheckRunner
        Private ReadOnly _broker As OutlookBroker

        Public Sub New(runner As CheckRunner, broker As OutlookBroker)
            _runner = runner
            _broker = broker
        End Sub

        Public Async Function RunAsync() As Task
            Await RunReadChecksAsync()
            Await RunLifecycleChecksAsync()
            Await RunPerformanceChecksAsync()
            Await RunEnvironmentChecksAsync()
        End Function

        ' ===================================================================
        ' B — Leitura
        ' ===================================================================
        Private Async Function RunReadChecksAsync() As Task
            Const Group As String = "B — Leitura"

            Await _runner.RunAsync(
                "B1", Group, "Enumerar stores e seus tipos",
                Async Function()
                    Dim stores = Await _broker.ReadAsync(
                        Function(app, ns)
                            Dim result As New List(Of StoreInfo)()
                            Dim collection = ns.Stores
                            Try
                                For i = 1 To collection.Count
                                    Dim store = TryCast(collection.Item(i), Outlook.Store)
                                    If store Is Nothing Then Continue For
                                    Try
                                        result.Add(New StoreInfo With {
                                            .DisplayName = SafeString(Function() store.DisplayName),
                                            .StoreId = SafeString(Function() store.StoreID),
                                            .FilePath = SafeString(Function() store.FilePath),
                                            .ExchangeStoreType = SafeString(Function() store.ExchangeStoreType.ToString()),
                                            .IsCachedExchange = SafeBool(Function() store.IsCachedExchange)
                                        })
                                    Finally
                                        ComHelpers.Release(store)
                                    End Try
                                Next
                            Finally
                                ComHelpers.Release(collection)
                            End Try
                            Return result
                        End Function)

                    If stores.Count = 0 Then
                        Return (CheckStatus.Fail, "Nenhum store visível.")
                    End If

                    Dim described = String.Join("; ",
                        stores.Select(Function(s) $"{s.DisplayName} [{s.ExchangeStoreType}" &
                                                  If(s.IsCachedExchange, ", cached", "") & "]"))
                    Return (CheckStatus.Pass, $"{stores.Count} store(s): {described}")
                End Function)

            Await _runner.RunAsync(
                "B2", Group, "Abrir a Caixa de Entrada padrão",
                Async Function()
                    Dim info = Await _broker.ReadAsync(
                        Function(app, ns)
                            Dim folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
                            Try
                                Dim items = folder.Items
                                Try
                                    Return New FolderInfo With {
                                        .Name = folder.Name,
                                        .Key = New ItemKey With {
                                            .EntryId = folder.EntryID,
                                            .StoreId = folder.StoreID
                                        },
                                        .DefaultItemType = folder.DefaultItemType.ToString(),
                                        .ItemCount = items.Count,
                                        .UnreadCount = SafeInt(Function() folder.UnReadItemCount)
                                    }
                                Finally
                                    ComHelpers.Release(items)
                                End Try
                            Finally
                                ComHelpers.Release(folder)
                            End Try
                        End Function)

                    Return (CheckStatus.Pass,
                            $"{info.Name}: {info.ItemCount} itens, {info.UnreadCount} não lidos, " &
                            $"tipo padrão {info.DefaultItemType}.")
                End Function)

            Await _runner.RunAsync(
                "B3", Group, "Itens heterogêneos são ignorados corretamente",
                Async Function()
                    ' ESCOPO.md seção 5: uma coleção Items NÃO contém apenas
                    ' MailItem. Convites, relatórios de entrega e itens de
                    ' outros tipos convivem na mesma pasta.
                    Dim survey = Await _broker.ReadAsync(
                        Function(app, ns)
                            Dim folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
                            Try
                                Dim items = folder.Items
                                Try
                                    Dim classes As New Dictionary(Of String, Integer)()
                                    Dim limit = Math.Min(items.Count, 250)
                                    Dim nonMail = 0
                                    For i = 1 To limit
                                        Dim raw = items.Item(i)
                                        Try
                                            Dim mail = TryCast(raw, Outlook.MailItem)
                                            Dim cls = If(mail IsNot Nothing,
                                                         SafeString(Function() mail.MessageClass),
                                                         raw.GetType().Name)
                                            If mail Is Nothing Then nonMail += 1
                                            If String.IsNullOrEmpty(cls) Then cls = "(desconhecida)"
                                            classes(cls) = If(classes.ContainsKey(cls), classes(cls), 0) + 1
                                        Finally
                                            ComHelpers.Release(raw)
                                        End Try
                                    Next
                                    Return (Examined:=limit, NonMail:=nonMail, Classes:=classes)
                                Finally
                                    ComHelpers.Release(items)
                                End Try
                            Finally
                                ComHelpers.Release(folder)
                            End Try
                        End Function)

                    Dim breakdown = String.Join(", ",
                        survey.Classes.OrderByDescending(Function(kv) kv.Value).
                                       Take(6).
                                       Select(Function(kv) $"{kv.Key}×{kv.Value}"))

                    Return (CheckStatus.Info,
                            $"{survey.Examined} itens examinados, {survey.NonMail} não-MailItem. {breakdown}")
                End Function)

            Await _runner.RunAsync(
                "B4", Group, "DTO de mensagem cruza a fronteira sem RCW",
                Async Function()
                    Dim summaries = Await _broker.ReadAsync(
                        Function(app, ns) ReadSummaries(ns, 5, includeBody:=False))

                    If summaries.Count = 0 Then
                        Return (CheckStatus.Skipped, "Caixa de entrada vazia.")
                    End If

                    If ComHelpers.ContainsComReference(summaries) Then
                        Return (CheckStatus.Fail,
                                "RCW vazou junto com o DTO: a fronteira da seção 4 está furada.")
                    End If

                    Dim first = summaries(0)
                    Return (CheckStatus.Pass,
                            $"{summaries.Count} DTOs limpos. Amostra: " &
                            $"""{Truncate(first.Subject, 40)}"" ({first.Content}).")
                End Function)

            Await _runner.RunAsync(
                "B5", Group, "Estado de download do conteúdo (R9)",
                Async Function()
                    Dim states = Await _broker.ReadAsync(
                        Function(app, ns)
                            Dim summaries = ReadSummaries(ns, 100, includeBody:=False)
                            Return summaries.GroupBy(Function(s) s.Content).
                                             ToDictionary(Function(g) g.Key.ToString(),
                                                          Function(g) g.Count())
                        End Function)

                    If states.Count = 0 Then
                        Return (CheckStatus.Skipped, "Caixa de entrada vazia.")
                    End If

                    Dim described = String.Join(", ", states.Select(Function(kv) $"{kv.Key}={kv.Value}"))
                    Dim headerOnly = If(states.ContainsKey(ContentState.MetadataOnly.ToString()),
                                        states(ContentState.MetadataOnly.ToString()), 0)

                    If headerOnly > 0 Then
                        Return (CheckStatus.Warn,
                                $"{described}. Há itens só com cabeçalho: a UI precisa dos estados " &
                                "explícitos do R9, nunca bloquear esperando download.")
                    End If

                    Return (CheckStatus.Pass, described)
                End Function)

            Await _runner.RunAsync(
                "B6", Group, "Ler endereço do remetente (guarda do OOM, R2)",
                Async Function()
                    ' Isolado de propósito: este é um dos campos que a guarda
                    ' de segurança do Object Model protege. Se bloquear, o
                    ' resto da leitura precisa continuar funcionando.
                    Dim outcome = Await _broker.ReadAsync(
                        Function(app, ns)
                            Dim folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
                            Try
                                Dim items = folder.Items
                                Try
                                    If items.Count = 0 Then Return (Ok:=False, Detail:="vazia")
                                    Dim raw = items.Item(1)
                                    Try
                                        Dim mail = TryCast(raw, Outlook.MailItem)
                                        If mail Is Nothing Then Return (Ok:=False, Detail:="primeiro item não é MailItem")
                                        Dim address = mail.SenderEmailAddress
                                        Return (Ok:=True, Detail:=If(String.IsNullOrEmpty(address), "(vazio)", "lido"))
                                    Finally
                                        ComHelpers.Release(raw)
                                    End Try
                                Finally
                                    ComHelpers.Release(items)
                                End Try
                            Finally
                                ComHelpers.Release(folder)
                            End Try
                        End Function)

                    If Not outcome.Ok Then
                        Return (CheckStatus.Skipped, outcome.Detail)
                    End If

                    Return (CheckStatus.Pass,
                            $"SenderEmailAddress {outcome.Detail} sem prompt da guarda.")
                End Function)
        End Function

        ' ===================================================================
        ' E — Ciclo de vida
        ' ===================================================================
        Private Async Function RunLifecycleChecksAsync() As Task
            Const Group As String = "E — Ciclo de vida"

            Await _runner.RunAsync(
                "E1", Group, "Probe confirma sessão viva",
                Async Function()
                    Dim state = Await _broker.ProbeAsync()
                    If state <> SessionState.Connected Then
                        Return (CheckStatus.Fail, $"Probe retornou {state} com o Outlook aberto.")
                    End If
                    Return (CheckStatus.Pass, "Sessão viva e responsiva.")
                End Function)

            Await _runner.RunAsync(
                "E2", Group, "Contadores do message filter (R13)",
                Async Function()
                    Await Task.CompletedTask
                    Dim f = _broker.MessageFilter
                    Dim notes = $"rejeições {f.RejectionsSeen}, retries {f.RetriesIssued}, " &
                                $"canceladas {f.CallsCancelled}."

                    If f.RejectionsSeen = 0 Then
                        Return (CheckStatus.Info,
                                notes & " Outlook nunca ficou ocupado nesta execução — " &
                                "o caminho de retry ficou sem exercício.")
                    End If

                    Return (CheckStatus.Pass, notes)
                End Function)
        End Function

        ' ===================================================================
        ' F — Desempenho
        ' ===================================================================
        Private Async Function RunPerformanceChecksAsync() As Task
            Const Group As String = "F — Desempenho"

            Await _runner.RunAsync(
                "F1", Group, "Restrict + Sort na Caixa de Entrada",
                Async Function()
                    Dim timing = Await _broker.ReadAsync(
                        Function(app, ns)
                            Dim folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
                            Try
                                Dim items = folder.Items
                                Try
                                    Dim sw = Stopwatch.StartNew()
                                    items.Sort("[ReceivedTime]", True)
                                    Dim sortMs = sw.Elapsed.TotalMilliseconds

                                    Dim cutoff = DateTime.Now.AddDays(-30).ToString("g")
                                    sw.Restart()
                                    Dim restricted = items.Restrict($"[ReceivedTime] >= '{cutoff}'")
                                    Dim count = restricted.Count
                                    Dim restrictMs = sw.Elapsed.TotalMilliseconds
                                    ComHelpers.Release(restricted)

                                    Return (Total:=items.Count, Recent:=count,
                                            SortMs:=sortMs, RestrictMs:=restrictMs)
                                Finally
                                    ComHelpers.Release(items)
                                End Try
                            Finally
                                ComHelpers.Release(folder)
                            End Try
                        End Function)

                    Return (CheckStatus.Info,
                            $"{timing.Total} itens. Sort {timing.SortMs:0} ms; " &
                            $"Restrict últimos 30 dias {timing.RestrictMs:0} ms ({timing.Recent} itens).")
                End Function)

            For Each size In {100, 1000}
                Dim pageSize = size
                Await _runner.RunAsync(
                    $"F{If(pageSize = 100, 2, 3)}", Group,
                    $"Ler {pageSize} mensagens só com metadados",
                    Async Function()
                        Dim outcome = Await _broker.ReadAsync(
                            Function(app, ns)
                                Dim sw = Stopwatch.StartNew()
                                Dim summaries = ReadSummaries(ns, pageSize, includeBody:=False)
                                sw.Stop()
                                Return (Count:=summaries.Count, Ms:=sw.Elapsed.TotalMilliseconds)
                            End Function)

                        ' Abaixo de MinSample o custo fixo (abrir pasta,
                        ' obter Items) domina e o ms/item vira ruído. Emitir
                        ' um aviso com base nisso seria pior que não medir:
                        ' produziria um número errado com cara de dado.
                        Const MinSample As Integer = 50

                        If outcome.Count < MinSample Then
                            Return (CheckStatus.Skipped,
                                    $"Amostra insuficiente ({outcome.Count} itens; mínimo {MinSample}). " &
                                    "Repetir quando a caixa tiver histórico — o número atual não " &
                                    "mede desempenho, mede custo fixo.")
                        End If

                        Dim perItem = outcome.Ms / outcome.Count
                        Dim verdict = If(perItem > 5.0, CheckStatus.Warn, CheckStatus.Info)
                        Dim note = $"{outcome.Count} itens em {outcome.Ms:0} ms ({perItem:0.0} ms/item)."
                        If verdict = CheckStatus.Warn Then
                            note &= " Acima de 5 ms/item a leitura direta da Fase 1 fica " &
                                    "desconfortável; reforça a Fase 2."
                        End If
                        Return (verdict, note)
                    End Function)
            Next

            Await _runner.RunAsync(
                "F4", Group, "Custo do corpo separado dos metadados",
                Async Function()
                    Dim outcome = Await _broker.ReadAsync(
                        Function(app, ns)
                            Dim sw = Stopwatch.StartNew()
                            Dim summaries = ReadSummaries(ns, 25, includeBody:=True)
                            sw.Stop()
                            Return (Count:=summaries.Count, Ms:=sw.Elapsed.TotalMilliseconds)
                        End Function)

                    If outcome.Count = 0 Then
                        Return (CheckStatus.Skipped, "Sem itens suficientes.")
                    End If

                    Return (CheckStatus.Info,
                            $"{outcome.Count} itens com corpo em {outcome.Ms:0} ms " &
                            $"({outcome.Ms / outcome.Count:0.0} ms/item).")
                End Function)
        End Function

        ' ===================================================================
        ' G — Ambiente e proteção
        ' ===================================================================
        Private Async Function RunEnvironmentChecksAsync() As Task
            Const Group As String = "G — Ambiente"

            Await _runner.RunAsync(
                "G1", Group, "Versão do Outlook e bitness do processo",
                Async Function()
                    Dim version = Await _broker.ReadAsync(Function(app, ns) app.Version)
                    Dim bitness = OutlookBitness()

                    Return (CheckStatus.Info,
                            $"Outlook {version}, processo {bitness}; Iris " &
                            $"{RuntimeInformation.ProcessArchitecture}. " &
                            If(bitness = "x86" AndAlso RuntimeInformation.ProcessArchitecture = Architecture.X64,
                               "Marshaling entre arquiteturas em uso — R12 confirmado na prática.",
                               "Mesma arquitetura."))
                End Function)

            Await _runner.RunAsync(
                "G2", Group, "Mensagens protegidas e classificadas (R11)",
                Async Function()
                    Dim survey = Await _broker.ReadAsync(
                        Function(app, ns)
                            Dim folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
                            Try
                                Dim items = folder.Items
                                Try
                                    Dim limit = Math.Min(items.Count, 200)
                                    Dim sensitive = 0
                                    Dim restricted = 0
                                    For i = 1 To limit
                                        Dim raw = items.Item(i)
                                        Try
                                            Dim mail = TryCast(raw, Outlook.MailItem)
                                            If mail Is Nothing Then Continue For
                                            If SafeString(Function() mail.Sensitivity.ToString()) <> "olNormal" Then
                                                sensitive += 1
                                            End If
                                            If SafeInt(Function() CInt(mail.Permission)) <> 0 Then
                                                restricted += 1
                                            End If
                                        Finally
                                            ComHelpers.Release(raw)
                                        End Try
                                    Next
                                    Return (Examined:=limit, Sensitive:=sensitive, Restricted:=restricted)
                                Finally
                                    ComHelpers.Release(items)
                                End Try
                            Finally
                                ComHelpers.Release(folder)
                            End Try
                        End Function)

                    Dim note = $"{survey.Examined} examinadas: {survey.Sensitive} com sensitivity " &
                               $"não-normal, {survey.Restricted} com permissão restrita."

                    If survey.Sensitive > 0 OrElse survey.Restricted > 0 Then
                        Return (CheckStatus.Warn,
                                note & " O escopo de pastas da IA precisa excluí-las (R11).")
                    End If

                    Return (CheckStatus.Info, note & " Nenhuma protegida na amostra.")
                End Function)
        End Function

        ' ===================================================================
        ' Auxiliares — todos rodam DENTRO da thread do broker
        ' ===================================================================

        ''' <summary>
        ''' Lê metadados da Caixa de Entrada para DTOs. Referências curtas e
        ''' liberação determinística em cada item (R7).
        ''' </summary>
        Private Shared Function ReadSummaries(ns As Outlook.NameSpace,
                                              count As Integer,
                                              includeBody As Boolean) As List(Of MailSummary)
            Dim result As New List(Of MailSummary)()
            Dim folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
            Try
                Dim items = folder.Items
                Try
                    Dim limit = Math.Min(items.Count, count)
                    For i = 1 To limit
                        Dim raw = items.Item(i)
                        Try
                            Dim mail = TryCast(raw, Outlook.MailItem)
                            If mail Is Nothing Then Continue For
                            result.Add(Summarize(mail, includeBody))
                        Finally
                            ComHelpers.Release(raw)
                        End Try
                    Next
                Finally
                    ComHelpers.Release(items)
                End Try
            Finally
                ComHelpers.Release(folder)
            End Try
            Return result
        End Function

        Private Shared Function Summarize(mail As Outlook.MailItem, includeBody As Boolean) As MailSummary
            Dim state = ContentState.MetadataOnly
            Try
                If mail.DownloadState = Outlook.OlDownloadState.olFullItem Then
                    state = ContentState.BodyAvailable
                End If
            Catch ex As COMException
                state = ContentState.TransientError
            End Try

            If includeBody AndAlso state = ContentState.BodyAvailable Then
                Try
                    Dim body = mail.Body
                    If mail.Attachments.Count > 0 Then state = ContentState.AttachmentsAvailable
                Catch ex As COMException
                    state = ContentState.TransientError
                End Try
            End If

            Return New MailSummary With {
                .Key = New ItemKey With {
                    .EntryId = SafeString(Function() mail.EntryID),
                    .StoreId = ""
                },
                .Subject = SafeString(Function() mail.Subject),
                .SenderName = SafeString(Function() mail.SenderName),
                .SenderAddress = "",
                .ReceivedTime = SafeDate(Function() mail.ReceivedTime),
                .SizeBytes = SafeInt(Function() mail.Size),
                .HasAttachments = SafeInt(Function() mail.Attachments.Count) > 0,
                .IsUnread = SafeBool(Function() mail.UnRead),
                .Content = state,
                .IsProtected = SafeInt(Function() CInt(mail.Permission)) <> 0,
                .MessageClass = SafeString(Function() mail.MessageClass)
            }
        End Function

        ' Propriedades COM lançam por item corrompido, offline ou baixado
        ' parcialmente (seção 5). Um item ruim não pode derrubar a listagem.
        Private Shared Function SafeString(getter As Func(Of String)) As String
            Try
                Return If(getter(), "")
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SafeInt(getter As Func(Of Integer)) As Integer
            Try
                Return getter()
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function SafeBool(getter As Func(Of Boolean)) As Boolean
            Try
                Return getter()
            Catch
                Return False
            End Try
        End Function

        Private Shared Function SafeDate(getter As Func(Of DateTime)) As DateTime?
            Try
                Return getter()
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function Truncate(value As String, max As Integer) As String
            If String.IsNullOrEmpty(value) Then Return "(sem assunto)"
            Return If(value.Length <= max, value, value.Substring(0, max) & "…")
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function IsWow64Process(process As IntPtr, <Out> ByRef wow64 As Boolean) As Boolean
        End Function

        Private Shared Function OutlookBitness() As String
            Dim processes = Process.GetProcessesByName("OUTLOOK")
            Try
                If processes.Length = 0 Then Return "desconhecido"
                Dim wow64 As Boolean
                If Not IsWow64Process(processes(0).Handle, wow64) Then Return "desconhecido"
                Return If(wow64, "x86", "x64")
            Catch
                Return "desconhecido"
            Finally
                For Each p In processes
                    p.Dispose()
                Next
            End Try
        End Function

    End Class

End Namespace

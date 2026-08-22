Imports IrisSpike.Broker
Imports IrisSpike.Interop
Imports Outlook = Microsoft.Office.Interop.Outlook

Namespace Checks

    ''' <summary>
    ''' Grupo C — envio.
    '''
    ''' Regras que valem para o grupo inteiro:
    '''
    '''   • Send() NUNCA é repetido. O message filter tem o retry desligado
    '''     durante a chamada (MutateAsync), porque repetir manda o e-mail
    '''     duas vezes.
    '''   • Send() retornar sem exceção NÃO prova entrega — pode ter apenas
    '''     enfileirado na Caixa de Saída. As observações são separadas.
    '''   • Falha no envio é resultado AMBÍGUO, não fracasso: antes de
    '''     qualquer coisa, procura-se o GUID nas pastas.
    '''   • O envio real só acontece com --send-to informado.
    ''' </summary>
    Public NotInheritable Class SendChecks

        Private Const Group As String = "C — Envio"
        Private Const Marker As String = "[IRIS-SPIKE-C]"

        Private ReadOnly _runner As CheckRunner
        Private ReadOnly _broker As OutlookBroker
        Private ReadOnly _sendTo As String

        Public Sub New(runner As CheckRunner, broker As OutlookBroker, sendTo As String)
            _runner = runner
            _broker = broker
            _sendTo = sendTo
        End Sub

        Public Async Function RunAsync() As Task
            Await C1_DraftAndDisplayAsync()
            Await C3_AmbiguityProcedureAsync()
            Await C2_RealSendAsync()
        End Function

        ' ===================================================================
        ' C1 — rascunho + Display(), sem risco de entrega
        ' ===================================================================
        Private Async Function C1_DraftAndDisplayAsync() As Task
            Await _runner.RunAsync(
                "C1", Group, "Rascunho, Display() e fechar sem enviar",
                Async Function()
                    Dim token = Guid.NewGuid().ToString("N").Substring(0, 12)

                    Dim outcome = Await _broker.MutateAsync(
                        Function(app, ns)
                            Dim item As Outlook.MailItem = Nothing
                            Try
                                item = TryCast(app.CreateItem(Outlook.OlItemType.olMailItem),
                                               Outlook.MailItem)
                                If item Is Nothing Then Return (Ok:=False, EntryId:="", Detail:="CreateItem falhou")

                                item.Subject = $"{Marker} rascunho {token}"
                                item.Body = "Rascunho de teste do spike da Fase 0. Não foi enviado."

                                ' Recipients é objeto COM próprio — nunca
                                ' encadear item.Recipients.Add (R7).
                                Dim recipients As Outlook.Recipients = Nothing
                                Try
                                    recipients = item.Recipients
                                    Dim r As Outlook.Recipient = Nothing
                                    Try
                                        r = recipients.Add(If(String.IsNullOrWhiteSpace(_sendTo),
                                                              "ninguem@exemplo.invalido", _sendTo))
                                    Finally
                                        ComHelpers.Release(r)
                                    End Try
                                Finally
                                    ComHelpers.Release(recipients)
                                End Try

                                item.Save()
                                Dim entryId = item.EntryID

                                ' Abre o inspector e fecha logo em seguida,
                                ' salvando. É o fallback oficial do R2.
                                item.Display(False)
                                item.Close(Outlook.OlInspectorClose.olSave)

                                Return (Ok:=True, EntryId:=entryId, Detail:="")

                            Catch ex As Runtime.InteropServices.COMException
                                ' Se a guarda do OOM bloquear até isto, o
                                ' fallback do R2 não existe e o projeto
                                ' precisa ser repensado.
                                Return (Ok:=False, EntryId:="",
                                        Detail:=$"COMException 0x{ex.HResult:X8}: {ex.Message}")
                            Finally
                                ComHelpers.Release(item)
                            End Try
                        End Function)

                    If Not outcome.Ok Then
                        Return (CheckStatus.Fail,
                                $"Não foi possível criar/exibir o rascunho — {outcome.Detail}. " &
                                "Sem isto, o fallback do R2 (rascunho + envio manual) não existe.")
                    End If

                    ' Confirma que o rascunho ficou onde deveria.
                    Dim location = Await FindByTokenAsync(token)
                    Await DeleteByEntryIdAsync(outcome.EntryId)

                    If location = "" Then
                        Return (CheckStatus.Warn,
                                "Rascunho criado e exibido, mas não localizado em Rascunhos depois.")
                    End If

                    Return (CheckStatus.Pass,
                            $"Rascunho criado, Inspector aberto e fechado, item encontrado em " &
                            $"{location}. Fallback do R2 funciona. Rascunho removido.")
                End Function)
        End Function

        ' ===================================================================
        ' C3 — o procedimento para resultado ambíguo
        ' ===================================================================
        Private Async Function C3_AmbiguityProcedureAsync() As Task
            Await _runner.RunAsync(
                "C3", Group, "Procedimento de ambiguidade localiza (ou não) por GUID",
                Async Function()
                    ' Exercita a busca com um GUID que garantidamente não
                    ' existe. É o caminho que roda quando Send() estoura e
                    ' não se sabe se a mensagem saiu — nessa hora, "tentar de
                    ' novo" é a pior resposta possível.
                    Dim ghost = Guid.NewGuid().ToString("N").Substring(0, 12)
                    Dim location = Await FindByTokenAsync(ghost)

                    If location.StartsWith("ERRO:") Then
                        Return (CheckStatus.Fail,
                                $"A busca não conseguiu executar ({location}). " &
                                "Um procedimento de desambiguação que não roda é pior que nenhum.")
                    End If

                    If location <> "" Then
                        Return (CheckStatus.Fail,
                                $"Busca encontrou um GUID inexistente em {location} — " &
                                "o procedimento de desambiguação não é confiável.")
                    End If

                    Return (CheckStatus.Pass,
                            "Busca varreu Rascunhos, Caixa de Saída, Itens Enviados e Entrada e " &
                            "reportou corretamente 'não encontrado'. É o que roda antes de " &
                            "qualquer decisão após um Send() ambíguo.")
                End Function)
        End Function

        ' ===================================================================
        ' C2 — envio real, só com --send-to
        ' ===================================================================
        Private Async Function C2_RealSendAsync() As Task
            If String.IsNullOrWhiteSpace(_sendTo) Then
                _runner.Skip("C2", Group, "Envio real com entrega confirmada",
                             "Requer --send-to <endereço>. Sem o argumento, NADA é enviado — " &
                             "e a pergunta central do R2 (a política permite Send?) segue aberta.")
                Return
            End If

            Await _runner.RunAsync(
                "C2", Group, $"Envio real para {_sendTo}, com entrega confirmada",
                Async Function()
                    Dim token = Guid.NewGuid().ToString("N").Substring(0, 12)
                    Dim filter = _broker.MessageFilter
                    Dim retriesBefore = filter.RetriesIssued

                    Dim result = Await _broker.MutateAsync(
                        Function(app, ns)
                            Dim item As Outlook.MailItem = Nothing
                            Dim sendStarted = False
                            Try
                                item = TryCast(app.CreateItem(Outlook.OlItemType.olMailItem),
                                               Outlook.MailItem)
                                If item Is Nothing Then Return (Sent:=False, Started:=False, Detail:="CreateItem falhou")

                                item.Subject = $"{Marker} envio {token}"
                                item.Body = "Mensagem automática de teste do spike da Fase 0 do Iris." &
                                            Environment.NewLine & $"Identificador: {token}"

                                Dim recipients As Outlook.Recipients = Nothing
                                Try
                                    recipients = item.Recipients
                                    Dim r As Outlook.Recipient = Nothing
                                    Try
                                        r = recipients.Add(_sendTo)
                                    Finally
                                        ComHelpers.Release(r)
                                    End Try
                                    ' Ignorar este retorno era mandar mensagem
                                    ' para destinatario nao resolvido.
                                    If Not recipients.ResolveAll() Then
                                        Return (Sent:=False, Started:=False,
                                                Detail:=$"ResolveAll falhou para '{_sendTo}'. " &
                                                        "Send() NAO foi chamado.")
                                    End If
                                Finally
                                    ComHelpers.Release(recipients)
                                End Try

                                item.Save()

                                ' Exatamente uma vez. Sem retry (MutateAsync
                                ' desliga o do message filter). Depois disto
                                ' o item NÃO é mais tocado.
                                ' A partir daqui QUALQUER excecao e ambigua,
                                ' nao so COMException: depois de invocar
                                ' Send() nao ha como saber se a mensagem saiu.
                                sendStarted = True
                                item.Send()
                                Return (Sent:=True, Started:=True, Detail:="")

                            Catch ex As Exception
                                Return (Sent:=False, Started:=sendStarted,
                                        Detail:=$"{ex.GetType().Name}: {ex.Message}")
                            Finally
                                ComHelpers.Release(item)
                            End Try
                        End Function)

                    Dim retriesDuring = filter.RetriesIssued - retriesBefore

                    If Not result.Sent Then
                        If Not result.Started Then
                            ' Falhou ANTES de invocar Send(): não há
                            ' ambiguidade nenhuma, nada saiu.
                            Return (CheckStatus.Fail,
                                    $"Preparação falhou e nada foi enviado — {result.Detail}")
                        End If

                        ' AMBÍGUO, não "falhou": a exceção pode ter vindo
                        ' depois de a mensagem sair.
                        Await Task.Delay(4000)
                        Dim whereIsIt = Await FindByTokenAsync(token)
                        Dim verdict = If(whereIsIt = "",
                                         "não localizada em nenhuma pasta",
                                         $"LOCALIZADA em {whereIsIt}")

                        Return (CheckStatus.Warn,
                                $"Send() lançou — {result.Detail}. Estado AMBÍGUO, e nenhum reenvio " &
                                $"foi tentado. Busca por GUID: {verdict}. " &
                                "Se a política bloqueia Send, o fallback do C1 vira o " &
                                "comportamento oficial do produto (R2).")
                    End If

                    ' Send() retornou. Isso ainda não é entrega. Seis segundos
                    ' fixos era curto demais para Exchange em modo cached:
                    ' agora é polling até 90s — e nunca, em hipótese alguma,
                    ' um reenvio.
                    Dim sentCopy = 0, inboxCopy = 0, outboxCopy = 0
                    Dim deadline = DateTime.UtcNow.AddSeconds(90)
                    While DateTime.UtcNow < deadline
                        sentCopy = Await CountByTokenAsync(token, Outlook.OlDefaultFolders.olFolderSentMail)
                        inboxCopy = Await CountByTokenAsync(token, Outlook.OlDefaultFolders.olFolderInbox)
                        outboxCopy = Await CountByTokenAsync(token, Outlook.OlDefaultFolders.olFolderOutbox)
                        If sentCopy > 0 AndAlso inboxCopy > 0 Then Exit While
                        Await Task.Delay(2000)
                    End While

                    Dim note = $"Send() retornou sem exceção; retries emitidos: {retriesDuring}. " &
                               $"Itens Enviados: {sentCopy}; Caixa de Saída: {outboxCopy}; " &
                               $"Entrada: {inboxCopy}."

                    If retriesDuring > 0 Then
                        Return (CheckStatus.Fail,
                                note & " HOUVE RETRY durante um envio — risco de mensagem duplicada.")
                    End If

                    If sentCopy > 1 OrElse inboxCopy > 1 Then
                        Return (CheckStatus.Fail, note & " MAIS DE UMA cópia: envio duplicado.")
                    End If

                    If sentCopy = 1 AndAlso inboxCopy = 1 Then
                        Return (CheckStatus.Pass,
                                note & " Uma cópia enviada e uma entregue: envio programático " &
                                "PERMITIDO pela política. R2 resolvido para envio.")
                    End If

                    If sentCopy = 1 Then
                        Return (CheckStatus.Warn,
                                note & " Saiu, mas a entrega não foi observada na janela de espera. " &
                                "Se o destinatário não é você mesmo, isso é esperado.")
                    End If

                    Return (CheckStatus.Warn,
                            note & " Send() retornou mas nenhuma cópia foi localizada — " &
                            "verificar manualmente antes de concluir qualquer coisa.")
                End Function)
        End Function

        ' ===================================================================
        ' Busca por GUID — o coração da desambiguação
        ' ===================================================================

        Private Shared ReadOnly SearchFolders As (Folder As Outlook.OlDefaultFolders, Name As String)() = {
            (Outlook.OlDefaultFolders.olFolderDrafts, "Rascunhos"),
            (Outlook.OlDefaultFolders.olFolderOutbox, "Caixa de Saída"),
            (Outlook.OlDefaultFolders.olFolderSentMail, "Itens Enviados"),
            (Outlook.OlDefaultFolders.olFolderInbox, "Entrada")
        }

        ''' <summary>Sentinela: a consulta falhou, nao "nada encontrado".</summary>
        Private Const SearchFailed As Integer = -1

        ''' <summary>
        ''' Pasta onde achou, "" se varreu tudo sem achar, ou "ERRO:..."
        ''' se alguma consulta falhou - porque "nao achei" e "nao consegui
        ''' procurar" levam a decisoes opostas.
        ''' </summary>
        Private Async Function FindByTokenAsync(token As String) As Task(Of String)
            Return Await _broker.ReadAsync(
                Function(app, ns)
                    Dim failures As New List(Of String)()
                    For Each target In SearchFolders
                        Dim found = CountInFolder(ns, target.Folder, token)
                        If found = SearchFailed Then
                            failures.Add(target.Name)
                        ElseIf found > 0 Then
                            Return target.Name
                        End If
                    Next
                    If failures.Count > 0 Then
                        Return "ERRO: consulta falhou em " & String.Join(", ", failures)
                    End If
                    Return ""
                End Function)
        End Function

        Private Async Function CountByTokenAsync(token As String,
                                                 folder As Outlook.OlDefaultFolders) As Task(Of Integer)
            Return Await _broker.ReadAsync(Function(app, ns) CountInFolder(ns, folder, token))
        End Function

        Private Shared Function CountInFolder(ns As Outlook.NameSpace,
                                              which As Outlook.OlDefaultFolders,
                                              token As String) As Integer
            Dim folder As Outlook.MAPIFolder = Nothing
            Try
                folder = ns.GetDefaultFolder(which)
                Dim items As Outlook.Items = Nothing
                Try
                    items = folder.Items
                    ' Restrict com o assunto é muito mais rápido que varrer.
                    Dim matches As Outlook.Items = Nothing
                    Try
                        matches = items.Restrict($"@SQL=""urn:schemas:httpmail:subject"" LIKE '%{token}%'")
                        Return matches.Count
                    Finally
                        ComHelpers.Release(matches)
                    End Try
                Finally
                    ComHelpers.Release(items)
                End Try
            Catch
                ' Converter erro de consulta em zero transformava "a busca
                ' falhou" em "nao encontrei" - a mentira mais perigosa
                ' possivel num procedimento de desambiguacao de envio.
                Return SearchFailed
            Finally
                ComHelpers.Release(folder)
            End Try
        End Function

        Private Async Function DeleteByEntryIdAsync(entryId As String) As Task
            Await _broker.MutateAsync(
                Function(app, ns)
                    Try
                        Dim item As Outlook.MailItem = Nothing
                        Try
                            item = TryCast(ns.GetItemFromID(entryId), Outlook.MailItem)
                            If item Is Nothing Then Return False
                            item.Delete()
                            Return True
                        Finally
                            ComHelpers.Release(item)
                        End Try
                    Catch
                        Return False
                    End Try
                End Function)
        End Function

    End Class

End Namespace

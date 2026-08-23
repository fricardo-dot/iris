Imports System.Collections.Generic
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.Core
Imports Iris.Model

''' <summary>
''' Broker de mentira, só para os testes do compositor.
'''
''' Existe porque o de verdade fala com o Outlook: exigir Outlook aberto
''' para verificar a lógica de rascunho tornaria o teste dependente da
''' caixa corporativa do usuário — e, no caso do envio, mandaria mensagem
''' de verdade a cada execução da suíte.
'''
''' Só o grupo de rascunhos é implementado. O resto lança de propósito: se
''' o compositor um dia chamar algo que não é da alçada dele, o teste
''' quebra em vez de passar por sorte.
''' </summary>
Friend NotInheritable Class FakeBroker
    Implements IOutlookBroker

    ''' <summary>Nome de cada operação chamada, na ordem.</summary>
    Friend ReadOnly Chamadas As New List(Of String)()

    ''' <summary>Chave recebida em cada operação de rascunho, na ordem.</summary>
    Friend ReadOnly ChavesRecebidas As New List(Of DraftKey)()

    ''' <summary>Conteúdo de cada UpdateDraft, na ordem.</summary>
    Friend ReadOnly Gravacoes As New List(Of DraftContent)()

    ''' <summary>
    ''' Quando não é Nothing, UpdateDraft fica parado até alguém completar.
    ''' É o que permite escrever "o usuário digitou DURANTE a gravação" sem
    ''' depender de tempo de relógio.
    ''' </summary>
    Friend TravaDoUpdate As TaskCompletionSource(Of Boolean)

    Friend FalhaAoCriar As ErrorKind = ErrorKind.None
    Friend FalhaAoGravar As ErrorKind = ErrorKind.None
    Friend FalhaAoPreparar As ErrorKind = ErrorKind.None
    Friend ResultadoDoEnvio As ErrorKind = ErrorKind.None
    Friend TodosResolvidos As Boolean = True

    ''' <summary>
    ''' Cada gravação devolve uma chave NOVA, como o Outlook faz: o EntryID
    ''' muda a cada Save. É isto que expõe um compositor que guardou a
    ''' chave antiga.
    ''' </summary>
    Private _versao As Integer

    Friend Function ChaveAtual() As DraftKey
        Return New DraftKey(New ItemKey($"draft-{_versao}", "store-1"))
    End Function

    Private Function NovoInfo() As DraftInfo
        _versao += 1
        Return New DraftInfo With {
            .Key = ChaveAtual(),
            .Subject = "",
            .ToLine = "",
            .CcLine = "",
            .UserText = "",
            .QuotedBody = "",
            .Format = BodyFormat.PlainText
        }
    End Function

    ' ---- Rascunhos ------------------------------------------------------

    Public Function CreateDraftAsync(content As DraftContent, cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateDraftAsync

        Chamadas.Add("create")
        If FalhaAoCriar <> ErrorKind.None Then
            Return Task.FromResult(OperationResult(Of DraftInfo).Fail(FalhaAoCriar, "teste"))
        End If
        Return Task.FromResult(OperationResult(Of DraftInfo).Ok(NovoInfo()))
    End Function

    Public Function CreateReplyDraftAsync(item As ItemKey, replyAll As Boolean,
                                          cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateReplyDraftAsync

        Chamadas.Add(If(replyAll, "replyAll", "reply"))
        If FalhaAoCriar <> ErrorKind.None Then
            Return Task.FromResult(OperationResult(Of DraftInfo).Fail(FalhaAoCriar, "teste"))
        End If

        Dim info = NovoInfo()
        info.Subject = "RE: original"
        info.ToLine = "alguem@exemplo.com"
        info.QuotedBody = "<div>----- mensagem original -----</div>"
        info.QuotedPreview = "----- mensagem original -----"
        Return Task.FromResult(OperationResult(Of DraftInfo).Ok(info))
    End Function

    Public Function CreateForwardDraftAsync(item As ItemKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateForwardDraftAsync

        Chamadas.Add("forward")
        Return Task.FromResult(OperationResult(Of DraftInfo).Ok(NovoInfo()))
    End Function

    Public Async Function UpdateDraftAsync(draft As DraftKey, content As DraftContent,
                                           cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.UpdateDraftAsync

        Chamadas.Add("update")
        ChavesRecebidas.Add(draft)
        Gravacoes.Add(content)

        Dim trava = TravaDoUpdate
        If trava IsNot Nothing Then Await trava.Task

        If FalhaAoGravar <> ErrorKind.None Then
            Return OperationResult(Of DraftInfo).Fail(FalhaAoGravar, "teste")
        End If

        Dim info = NovoInfo()
        info.Subject = content.Subject
        info.ToLine = content.ToLine
        info.CcLine = content.CcLine
        info.UserText = content.UserText
        Return OperationResult(Of DraftInfo).Ok(info)
    End Function

    Public Function AddDraftAttachmentAsync(draft As DraftKey, filePath As String,
                                            cancel As CancellationToken) _
        As Task(Of OperationResult(Of AttachmentInfo)) Implements IOutlookBroker.AddDraftAttachmentAsync

        Chamadas.Add("attach")
        ChavesRecebidas.Add(draft)
        Return Task.FromResult(OperationResult(Of AttachmentInfo).Ok(
            New AttachmentInfo With {.FileName = System.IO.Path.GetFileName(filePath)}))
    End Function

    Public Function RemoveDraftAttachmentAsync(draft As DraftKey, attachment As AttachmentKey,
                                               cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.RemoveDraftAttachmentAsync
        Throw New NotSupportedException("O compositor não deveria chamar isto neste marco.")
    End Function

    Public Function PrepareSendAsync(draft As DraftKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of SendPreview)) Implements IOutlookBroker.PrepareSendAsync

        Chamadas.Add("prepare")
        ChavesRecebidas.Add(draft)

        If FalhaAoPreparar <> ErrorKind.None Then
            Return Task.FromResult(OperationResult(Of SendPreview).Fail(FalhaAoPreparar, "teste"))
        End If

        Dim p As New SendPreview With {
            .Draft = draft,
            .SendingAccount = "eu@empresa.com",
            .Subject = "assunto"
        }
        p.Recipients.Add(New RecipientInfo With {
            .DisplayName = "Fulano",
            .Address = If(TodosResolvidos, "fulano@empresa.com", ""),
            .Kind = RecipientKind.To,
            .Resolved = TodosResolvidos
        })
        Return Task.FromResult(OperationResult(Of SendPreview).Ok(p))
    End Function

    Public Function SendDraftAsync(draft As DraftKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.SendDraftAsync

        Chamadas.Add("send")
        ChavesRecebidas.Add(draft)

        If ResultadoDoEnvio <> ErrorKind.None Then
            Return Task.FromResult(OperationResult(Of Boolean).Fail(ResultadoDoEnvio, "teste"))
        End If
        Return Task.FromResult(OperationResult(Of Boolean).Ok(True))
    End Function

    Public Function DeleteDraftAsync(draft As DraftKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.DeleteDraftAsync

        Chamadas.Add("delete")
        ChavesRecebidas.Add(draft)
        Return Task.FromResult(OperationResult(Of Boolean).Ok(True))
    End Function

    ' ---- Fora da alçada do compositor -----------------------------------

    Public ReadOnly Property State As SessionState Implements IOutlookBroker.State
        Get
            Return SessionState.Connected
        End Get
    End Property

    Public Event StateChanged As EventHandler(Of SessionState) Implements IOutlookBroker.StateChanged
    Public Event FolderInvalidated As EventHandler(Of FolderInvalidation) _
        Implements IOutlookBroker.FolderInvalidated

    Private Shared Function ForaDaAlcada(Of T)() As Task(Of T)
        Throw New NotSupportedException("O compositor não deveria chamar isto.")
    End Function

    Public Function ConnectAsync(cancel As CancellationToken) As Task(Of SessionState) _
        Implements IOutlookBroker.ConnectAsync
        Return ForaDaAlcada(Of SessionState)()
    End Function

    Public Function ProbeAsync(cancel As CancellationToken) As Task(Of SessionState) _
        Implements IOutlookBroker.ProbeAsync
        Return ForaDaAlcada(Of SessionState)()
    End Function

    Public Function GetStoresAsync(cancel As CancellationToken) _
        As Task(Of OperationResult(Of IReadOnlyList(Of StoreInfo))) Implements IOutlookBroker.GetStoresAsync
        Return ForaDaAlcada(Of OperationResult(Of IReadOnlyList(Of StoreInfo)))()
    End Function

    Public Function GetFolderChildrenAsync(parent As FolderKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of IReadOnlyList(Of FolderInfo))) Implements IOutlookBroker.GetFolderChildrenAsync
        Return ForaDaAlcada(Of OperationResult(Of IReadOnlyList(Of FolderInfo)))()
    End Function

    Public Function GetMessagePageAsync(query As MessageQuery, offset As Integer, count As Integer,
                                        cancel As CancellationToken) _
        As Task(Of OperationResult(Of MessagePage)) Implements IOutlookBroker.GetMessagePageAsync
        Return ForaDaAlcada(Of OperationResult(Of MessagePage))()
    End Function

    Public Function GetMessageDetailAsync(item As ItemKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of MessageDetail)) Implements IOutlookBroker.GetMessageDetailAsync
        Return ForaDaAlcada(Of OperationResult(Of MessageDetail))()
    End Function

    Public Function SaveAttachmentAsync(attachment As AttachmentKey, destinationPath As String,
                                        overwrite As Boolean, cancel As CancellationToken) _
        As Task(Of OperationResult(Of String)) Implements IOutlookBroker.SaveAttachmentAsync
        Return ForaDaAlcada(Of OperationResult(Of String))()
    End Function

    Public Function MarkReadAsync(item As ItemKey, isRead As Boolean, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.MarkReadAsync
        Return ForaDaAlcada(Of OperationResult(Of Boolean))()
    End Function

    Public Function SubscribeFolderAsync(folder As FolderKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of SubscriptionToken)) Implements IOutlookBroker.SubscribeFolderAsync
        Return ForaDaAlcada(Of OperationResult(Of SubscriptionToken))()
    End Function

    Public Function UnsubscribeFolderAsync(token As SubscriptionToken, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.UnsubscribeFolderAsync
        Return ForaDaAlcada(Of OperationResult(Of Boolean))()
    End Function

    ''' <summary>Cala os avisos de evento nunca disparado.</summary>
    Friend Sub NaoUsado()
        RaiseEvent StateChanged(Me, SessionState.Connected)
        RaiseEvent FolderInvalidated(Me, Nothing)
    End Sub

End Class

''' <summary>Escolhedor de arquivo que não abre diálogo nenhum.</summary>
Friend NotInheritable Class FakePickFile
    Implements Iris.App.IPickFileService

    Friend Escolha As String

    Public Function AskWhichFileToAttach() As String _
        Implements Iris.App.IPickFileService.AskWhichFileToAttach
        Return Escolha
    End Function
End Class

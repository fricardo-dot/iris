Imports System.Collections.Generic
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.Core
Imports Iris.Model

''' <summary>
''' Duplo mínimo para os testes do <c>FolderWatcher</c>.
'''
''' Separado do <c>FakeBroker</c> de propósito. Aquele foi feito para o
''' compositor: estado fixo em Connected e exceção em tudo que não é
''' rascunho. Transformá-lo em simulador universal produziria um objeto
''' grande e permissivo, que aceita qualquer coisa e por isso não prova
''' nada.
'''
''' Aqui só existe o que o watcher usa: assinar, desassinar e disparar
''' invalidação. O resto lança.
''' </summary>
Friend NotInheritable Class WatcherBroker
    Implements IOutlookBroker

    Friend ReadOnly Assinadas As New List(Of FolderKey)()
    Friend ReadOnly Desassinadas As New List(Of Integer)()

    ''' <summary>
    ''' Quando não é Nothing, <c>SubscribeFolderAsync</c> fica parado até
    ''' alguém completar. É o que permite escrever "a assinatura de A
    ''' terminou DEPOIS de o usuário trocar para B" sem depender de tempo.
    ''' </summary>
    Friend TravaDoSubscribe As TaskCompletionSource(Of Boolean)

    Private _proximoId As Integer = 0

    Friend Function UltimoToken() As Integer
        Return _proximoId
    End Function

    Public Async Function SubscribeFolderAsync(folder As FolderKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of SubscriptionToken)) Implements IOutlookBroker.SubscribeFolderAsync

        Assinadas.Add(folder)

        Dim trava = TravaDoSubscribe
        If trava IsNot Nothing Then Await trava.Task

        _proximoId += 1
        Return OperationResult(Of SubscriptionToken).Ok(New SubscriptionToken(_proximoId, folder))
    End Function

    Public Function UnsubscribeFolderAsync(token As SubscriptionToken, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.UnsubscribeFolderAsync

        If token IsNot Nothing Then Desassinadas.Add(token.Id)
        Return Task.FromResult(OperationResult(Of Boolean).Ok(True))
    End Function

    ''' <summary>Simula o Outlook avisando que uma pasta mudou.</summary>
    Friend Sub Invalidar(subscriptionId As Integer, folder As FolderKey)
        RaiseEvent FolderInvalidated(Me, New FolderInvalidation With {
            .SubscriptionId = subscriptionId,
            .Folder = folder,
            .Kind = InvalidationKind.ItemChanged,
            .At = DateTimeOffset.UtcNow
        })
    End Sub

    Public Event FolderInvalidated As EventHandler(Of FolderInvalidation) _
        Implements IOutlookBroker.FolderInvalidated

    Public Event StateChanged As EventHandler(Of SessionState) Implements IOutlookBroker.StateChanged
    Public Event SessionReplaced As EventHandler(Of Long) Implements IOutlookBroker.SessionReplaced

    Public ReadOnly Property State As SessionState Implements IOutlookBroker.State
        Get
            Return SessionState.Connected
        End Get
    End Property

    Public ReadOnly Property SessionEpoch As Long Implements IOutlookBroker.SessionEpoch
        Get
            Return 1
        End Get
    End Property

    Friend Sub NaoUsado()
        RaiseEvent StateChanged(Me, SessionState.Connected)
        RaiseEvent SessionReplaced(Me, 1)
    End Sub

    ' ---- Fora da alçada do watcher --------------------------------------

    Private Shared Function Fora(Of T)() As Task(Of T)
        Throw New NotSupportedException("O watcher não deveria chamar isto.")
    End Function

    Public Function ConnectAsync(cancel As CancellationToken) As Task(Of SessionState) _
        Implements IOutlookBroker.ConnectAsync
        Return Fora(Of SessionState)()
    End Function

    Public Function ProbeAsync(cancel As CancellationToken) As Task(Of SessionState) _
        Implements IOutlookBroker.ProbeAsync
        Return Fora(Of SessionState)()
    End Function

    ''' <summary>
    ''' Fora da alcada, como quase tudo aqui: um duplo que respondesse a
    ''' tudo faria uma chamada indevida passar por sorte em vez de quebrar
    ''' o teste.
    ''' </summary>
    Public Function GetAppointmentsAsync(folder As FolderKey,
                                         de As DateTimeOffset, ate As DateTimeOffset,
                                         cancel As CancellationToken) _
        As Task(Of OperationResult(Of AppointmentWindow)) _
        Implements IAgendaSource.GetAppointmentsAsync
        Throw New NotSupportedException("O watcher não deveria chamar isto.")
    End Function

    Public Function GetStoresAsync(cancel As CancellationToken) _
        As Task(Of OperationResult(Of IReadOnlyList(Of StoreInfo))) Implements IOutlookBroker.GetStoresAsync
        Return Fora(Of OperationResult(Of IReadOnlyList(Of StoreInfo)))()
    End Function

    Public Function GetFolderChildrenAsync(parent As FolderKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of IReadOnlyList(Of FolderInfo))) Implements IOutlookBroker.GetFolderChildrenAsync
        Return Fora(Of OperationResult(Of IReadOnlyList(Of FolderInfo)))()
    End Function

    Public Function GetMessagePageAsync(query As MessageQuery, continuation As String,
                                        targetCount As Integer,
                                        cancel As CancellationToken) _
        As Task(Of OperationResult(Of MessagePage)) Implements IOutlookBroker.GetMessagePageAsync
        Return Fora(Of OperationResult(Of MessagePage))()
    End Function

    Public Function GetMessageDetailAsync(item As ItemKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of MessageDetail)) Implements IOutlookBroker.GetMessageDetailAsync
        Return Fora(Of OperationResult(Of MessageDetail))()
    End Function

    Public Function GetAttachmentPresenceAsync(items As IReadOnlyList(Of ItemKey),
                                               cancel As CancellationToken) _
        As Task(Of OperationResult(Of IReadOnlyList(Of AttachmentPresence))) _
        Implements IOutlookBroker.GetAttachmentPresenceAsync
        Return Fora(Of OperationResult(Of IReadOnlyList(Of AttachmentPresence)))()
    End Function

    Public Function GetMessageSnapshotAsync(item As ItemKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of MessageSnapshot)) _
        Implements IOutlookBroker.GetMessageSnapshotAsync
        Return Fora(Of OperationResult(Of MessageSnapshot))()
    End Function

    Public Function GetSensitivityLabelsAsync(items As IReadOnlyList(Of ItemKey),
                                              cancel As CancellationToken) _
        As Task(Of OperationResult(Of IReadOnlyList(Of LabelReading))) _
        Implements IOutlookBroker.GetSensitivityLabelsAsync
        Return Fora(Of OperationResult(Of IReadOnlyList(Of LabelReading)))()
    End Function

    Public Function ProbeLabelSemanticsAsync(item As ItemKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of NamedPropertyProbe)) _
        Implements IOutlookBroker.ProbeLabelSemanticsAsync
        Return Fora(Of OperationResult(Of NamedPropertyProbe))()
    End Function

    Public Function ProbeLabelColumnAsync(folder As FolderKey, quantas As Integer,
                                          cancel As CancellationToken) _
        As Task(Of OperationResult(Of LabelColumnProbe)) _
        Implements IOutlookBroker.ProbeLabelColumnAsync
        Return Fora(Of OperationResult(Of LabelColumnProbe))()
    End Function

    Public Function SaveAttachmentAsync(attachment As AttachmentKey, destinationPath As String,
                                        overwrite As Boolean, cancel As CancellationToken) _
        As Task(Of OperationResult(Of String)) Implements IOutlookBroker.SaveAttachmentAsync
        Return Fora(Of OperationResult(Of String))()
    End Function

    Public Function MarkReadAsync(item As ItemKey, isRead As Boolean, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.MarkReadAsync
        Return Fora(Of OperationResult(Of Boolean))()
    End Function

    Public Function CreateDraftAsync(content As DraftContent, cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateDraftAsync
        Return Fora(Of OperationResult(Of DraftInfo))()
    End Function

    Public Function CreateReplyDraftAsync(item As ItemKey, replyAll As Boolean,
                                          cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateReplyDraftAsync
        Return Fora(Of OperationResult(Of DraftInfo))()
    End Function

    Public Function CreateForwardDraftAsync(item As ItemKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateForwardDraftAsync
        Return Fora(Of OperationResult(Of DraftInfo))()
    End Function

    Public Function UpdateDraftAsync(draft As DraftKey, content As DraftContent,
                                     cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.UpdateDraftAsync
        Return Fora(Of OperationResult(Of DraftInfo))()
    End Function

    Public Function AddDraftAttachmentAsync(draft As DraftKey, filePath As String,
                                            cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.AddDraftAttachmentAsync
        Return Fora(Of OperationResult(Of DraftInfo))()
    End Function

    Public Function RemoveDraftAttachmentAsync(draft As DraftKey, attachment As AttachmentKey,
                                               cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.RemoveDraftAttachmentAsync
        Return Fora(Of OperationResult(Of DraftInfo))()
    End Function

    Public Function PrepareSendAsync(draft As DraftKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of SendPreview)) Implements IOutlookBroker.PrepareSendAsync
        Return Fora(Of OperationResult(Of SendPreview))()
    End Function

    Public Function SendDraftAsync(draft As DraftKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.SendDraftAsync
        Return Fora(Of OperationResult(Of Boolean))()
    End Function

    Public Function DeleteDraftAsync(draft As DraftKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.DeleteDraftAsync
        Return Fora(Of OperationResult(Of Boolean))()
    End Function

    ''' <summary>
    ''' Escrita no calendário. Fora da alçada por padrão, como o resto: chamada
    ''' que ninguém configurou QUEBRA o teste, em vez de passar por sorte.
    ''' </summary>
    ''' <summary>
    ''' Tarefas. Fora da alçada por padrão, como o resto: chamada que ninguém
    ''' configurou QUEBRA o teste, em vez de passar por sorte.
    ''' </summary>
    Public Function GetTasksAsync(folder As FolderKey, teto As Integer,
                                  cancel As CancellationToken) _
        As Task(Of OperationResult(Of TaskList)) _
        Implements ITarefasBroker.GetTasksAsync
        Throw New NotSupportedException("fora da alcada")
    End Function

    Public Function CreateTaskAsync(folder As FolderKey, rascunho As TaskDraft,
                                    cancel As CancellationToken) _
        As Task(Of OperationResult(Of TaskInfo)) _
        Implements ITarefasBroker.CreateTaskAsync
        Throw New NotSupportedException("fora da alcada")
    End Function

    Public Function CompleteTaskAsync(chave As TaskKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of TaskInfo)) _
        Implements ITarefasBroker.CompleteTaskAsync
        Throw New NotSupportedException("fora da alcada")
    End Function

    Public Function CreateAppointmentAsync(folder As FolderKey, rascunho As AppointmentDraft,
                                           cancel As CancellationToken) _
        As Task(Of OperationResult(Of AppointmentInfo)) _
        Implements IAgendaWriter.CreateAppointmentAsync
        Throw New NotSupportedException("fora da alcada")
    End Function

    Public Function UpdateAppointmentAsync(chave As AppointmentKey, rascunho As AppointmentDraft,
                                           cancel As CancellationToken) _
        As Task(Of OperationResult(Of AppointmentInfo)) _
        Implements IAgendaWriter.UpdateAppointmentAsync
        Throw New NotSupportedException("fora da alcada")
    End Function

    Public Function DeleteAppointmentAsync(chave As AppointmentKey,
                                           cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) _
        Implements IAgendaWriter.DeleteAppointmentAsync
        Throw New NotSupportedException("fora da alcada")
    End Function

End Class

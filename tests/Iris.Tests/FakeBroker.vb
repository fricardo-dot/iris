Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.Core
Imports Iris.Model

''' <summary>Como o "Outlook" de mentira devolve os destinatários.</summary>
Friend Enum ModoDeDestinatario
    ''' <summary>Resolvido, com SMTP de gente.</summary>
    Smtp
    ''' <summary>O Outlook não reconheceu o nome.</summary>
    NaoResolvido
    ''' <summary>
    ''' O caso perigoso: o Outlook diz que RESOLVEU, mas o endereço é
    ''' <c>/O=...</c>. Parece sucesso e não é conferível por ninguém.
    ''' </summary>
    ExchangeLegado
End Enum

''' <summary>
''' Broker de mentira, só para os testes do compositor.
'''
''' Existe porque o de verdade fala com o Outlook: exigir Outlook aberto
''' para verificar a lógica de rascunho tornaria o teste dependente da
''' caixa corporativa do usuário — e, no caso do envio, mandaria mensagem
''' de verdade a cada execução da suíte.
'''
''' Imita três comportamentos do Outlook que o compositor precisa aguentar,
''' e imitá-los é o que dá valor aos testes:
'''
'''   • TODA operação que salva devolve uma chave nova. Inclui anexar.
'''   • A prévia de envio é montada a partir do que foi GRAVADO, não do que
'''     o teste gostaria que estivesse lá. Sem isso, "a confirmação descreve
'''     o que vai sair" seria só um teste de ordem de chamadas.
'''   • Operar com chave vencida devolve NotFound.
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
    ''' Travas para segurar uma operação em voo. É o que permite escrever
    ''' "o usuário digitou DURANTE isto" sem depender de tempo de relógio.
    ''' </summary>
    Friend TravaDoUpdate As TaskCompletionSource(Of Boolean)
    Friend TravaDoAttach As TaskCompletionSource(Of Boolean)
    Friend TravaDoCreate As TaskCompletionSource(Of Boolean)
    Friend TravaDoPrepare As TaskCompletionSource(Of Boolean)
    Friend TravaDoSend As TaskCompletionSource(Of Boolean)

    Friend FalhaAoCriar As ErrorKind = ErrorKind.None
    Friend FalhaAoGravar As ErrorKind = ErrorKind.None
    Friend FalhaAoPreparar As ErrorKind = ErrorKind.None
    Friend ResultadoDoEnvio As ErrorKind = ErrorKind.None
    Friend FalhaAoDescartar As ErrorKind = ErrorKind.None
    Friend Modo As ModoDeDestinatario = ModoDeDestinatario.Smtp

    ''' <summary>Como a leitura da lista de destinatarios se saiu.</summary>
    Friend LeituraDeDestinatarios As PartStatus = PartStatus.Full

    ''' <summary>O que foi de fato enviado, para o teste conferir.</summary>
    Friend Enviado As SendPreview

    ' ---- O "store" ------------------------------------------------------

    Private _versao As Integer
    Private _existe As Boolean
    Private _subject As String = ""
    Private _toLine As String = ""
    Private _ccLine As String = ""
    Private _userText As String = ""
    Private _quoted As String = ""
    Private ReadOnly _anexos As New List(Of AttachmentInfo)()

    Friend Function ChaveAtual() As DraftKey
        Return New DraftKey(New ItemKey($"draft-{_versao}", "store-1"))
    End Function

    ''' <summary>
    ''' Todo Save gira a chave, como o Outlook faz com o EntryID. Quem
    ''' guardou a antiga descobre aqui.
    ''' </summary>
    Private Function Salvar() As DraftInfo
        _versao += 1
        _existe = True

        Dim info As New DraftInfo With {
            .Key = ChaveAtual(),
            .Subject = _subject,
            .ToLine = _toLine,
            .CcLine = _ccLine,
            .UserText = _userText,
            .QuotedBody = _quoted,
            .QuotedPreview = _quoted,
            .Format = BodyFormat.PlainText
        }
        info.Attachments.AddRange(_anexos)
        Return info
    End Function

    Private Function ChaveVale(chave As DraftKey) As Boolean
        Return _existe AndAlso chave IsNot Nothing AndAlso chave.Equals(ChaveAtual())
    End Function

    ' ---- Rascunhos ------------------------------------------------------

    Public Async Function CreateDraftAsync(content As DraftContent, cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateDraftAsync

        Chamadas.Add("create")

        Dim trava = TravaDoCreate
        If trava IsNot Nothing Then Await trava.Task

        If FalhaAoCriar <> ErrorKind.None Then
            Return OperationResult(Of DraftInfo).Fail(FalhaAoCriar, "teste")
        End If
        Return OperationResult(Of DraftInfo).Ok(Salvar())
    End Function

    Public Function CreateReplyDraftAsync(item As ItemKey, replyAll As Boolean,
                                          cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateReplyDraftAsync

        Chamadas.Add(If(replyAll, "replyAll", "reply"))
        If FalhaAoCriar <> ErrorKind.None Then
            Return Task.FromResult(OperationResult(Of DraftInfo).Fail(FalhaAoCriar, "teste"))
        End If

        _subject = "RE: original"
        _toLine = "alguem@exemplo.com"
        _quoted = "----- mensagem original -----"
        Return Task.FromResult(OperationResult(Of DraftInfo).Ok(Salvar()))
    End Function

    Public Function CreateForwardDraftAsync(item As ItemKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.CreateForwardDraftAsync

        Chamadas.Add("forward")
        Return Task.FromResult(OperationResult(Of DraftInfo).Ok(Salvar()))
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
        If Not ChaveVale(draft) Then
            Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "chave vencida")
        End If

        _subject = content.Subject
        _toLine = content.ToLine
        _ccLine = content.CcLine
        _userText = content.UserText
        Return OperationResult(Of DraftInfo).Ok(Salvar())
    End Function

    Public Async Function AddDraftAttachmentAsync(draft As DraftKey, filePath As String,
                                                  cancel As CancellationToken) _
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.AddDraftAttachmentAsync

        Chamadas.Add("attach")
        ChavesRecebidas.Add(draft)

        ' Ponto de espera DENTRO da anexação. Sem ele não dava para provar
        ' que a trava cobre até depois do anexo — só que cobria a descarga.
        Dim trava = TravaDoAttach
        If trava IsNot Nothing Then Await trava.Task

        If Not ChaveVale(draft) Then
            Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "chave vencida")
        End If

        _anexos.Add(New AttachmentInfo With {
            .FileName = System.IO.Path.GetFileName(filePath),
            .SizeBytes = 10
        })

        ' Anexar SALVA. A chave gira aqui também — era exatamente isto que o
        ' duplo antigo não fazia, e por isso o defeito passava batido.
        Dim info = Salvar()

        ' As chaves dos anexos são reconstruídas com o dono NOVO, como o
        ' Descrever de verdade faz. Deixá-las apontando para o dono anterior
        ' modelaria mal a promessa que o comentário da classe faz.
        For i = 0 To _anexos.Count - 1
            _anexos(i).Key = New AttachmentKey(ChaveAtual().Item, i + 1,
                                               _anexos(i).FileName, _anexos(i).SizeBytes)
        Next
        info.Attachments.Clear()
        info.Attachments.AddRange(_anexos)

        Return OperationResult(Of DraftInfo).Ok(info)
    End Function

    Public Function RemoveDraftAttachmentAsync(draft As DraftKey, attachment As AttachmentKey,
                                               cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.RemoveDraftAttachmentAsync
        Throw New NotSupportedException("O compositor não deveria chamar isto neste marco.")
    End Function

    ''' <summary>
    ''' Monta a prévia a partir do que está GRAVADO. Devolver valores fixos
    ''' faria o teste "a confirmação descreve o que vai sair" provar apenas
    ''' que uma chamada veio antes da outra.
    ''' </summary>
    Public Async Function PrepareSendAsync(draft As DraftKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of SendPreview)) Implements IOutlookBroker.PrepareSendAsync

        Chamadas.Add("prepare")
        ChavesRecebidas.Add(draft)

        Dim trava = TravaDoPrepare
        If trava IsNot Nothing Then Await trava.Task

        If FalhaAoPreparar <> ErrorKind.None Then
            Return OperationResult(Of SendPreview).Fail(FalhaAoPreparar, "teste")
        End If
        If Not ChaveVale(draft) Then
            Return OperationResult(Of SendPreview).Fail(ErrorKind.NotFound, "chave vencida")
        End If

        Dim p As New SendPreview With {
            .Draft = draft,
            .SendingAccount = "eu@empresa.com",
            .Subject = _subject
        }

        For Each bruto In _toLine.Split(";"c)
            Dim digitado = bruto.Trim()
            If digitado.Length = 0 Then Continue For

            Select Case Modo
                Case ModoDeDestinatario.NaoResolvido
                    p.Recipients.Add(New RecipientInfo With {
                        .DisplayName = digitado, .Address = "",
                        .Kind = RecipientKind.To, .Resolved = False})

                Case ModoDeDestinatario.ExchangeLegado
                    ' Resolved=True com /O=: o caso que engana.
                    p.Recipients.Add(New RecipientInfo With {
                        .DisplayName = digitado,
                        .Address = "/O=EMPRESA/OU=GRUPO/CN=RECIPIENTS/CN=" & digitado,
                        .Kind = RecipientKind.To, .Resolved = True})

                Case Else
                    p.Recipients.Add(New RecipientInfo With {
                        .DisplayName = digitado, .Address = digitado,
                        .Kind = RecipientKind.To, .Resolved = True})
            End Select
        Next

        p.Attachments.AddRange(_anexos)
        p.RecipientsStatus = LeituraDeDestinatarios
        Return OperationResult(Of SendPreview).Ok(p)
    End Function

    Public Async Function SendDraftAsync(draft As DraftKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.SendDraftAsync

        Chamadas.Add("send")
        ChavesRecebidas.Add(draft)

        ' Fotografa o que saiu ANTES de qualquer espera: é o que o teste usa
        ' para conferir se o enviado bate com o confirmado.
        Enviado = New SendPreview With {.Draft = draft, .Subject = _subject}

        Dim trava = TravaDoSend
        If trava IsNot Nothing Then Await trava.Task

        If ResultadoDoEnvio <> ErrorKind.None Then
            Return OperationResult(Of Boolean).Fail(ResultadoDoEnvio, "teste")
        End If
        If Not ChaveVale(draft) Then
            Return OperationResult(Of Boolean).Fail(ErrorKind.NotFound, "chave vencida")
        End If

        _existe = False
        Return OperationResult(Of Boolean).Ok(True)
    End Function

    Public Function DeleteDraftAsync(draft As DraftKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.DeleteDraftAsync

        Chamadas.Add("delete")
        ChavesRecebidas.Add(draft)

        If FalhaAoDescartar <> ErrorKind.None Then
            Return Task.FromResult(OperationResult(Of Boolean).Fail(FalhaAoDescartar, "teste"))
        End If

        _existe = False
        Return Task.FromResult(OperationResult(Of Boolean).Ok(True))
    End Function

    ' ---- Fora da alçada do compositor -----------------------------------

    Public ReadOnly Property State As SessionState Implements IOutlookBroker.State
        Get
            Return SessionState.Connected
        End Get
    End Property

    Public Event StateChanged As EventHandler(Of SessionState) Implements IOutlookBroker.StateChanged
    Public Event SessionReplaced As EventHandler(Of Long) Implements IOutlookBroker.SessionReplaced

    Private _epoca As Long = 1

    Public ReadOnly Property SessionEpoch As Long Implements IOutlookBroker.SessionEpoch
        Get
            Return _epoca
        End Get
    End Property

    ''' <summary>
    ''' Simula o Outlook morrer e voltar: a sessão é outra, as chaves da
    ''' anterior deixam de valer, e quem depende disso precisa saber.
    ''' </summary>
    Friend Sub SubstituirSessao()
        _epoca += 1
        _existe = False
        RaiseEvent SessionReplaced(Me, _epoca)
    End Sub
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

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

    ' ---- A ARVORE DE PASTAS, e por que ela chegou tarde ------------------
    '
    ' Ate 28/08/2026 este duplo respondia "fora da alcada" para stores e
    ' filhas, e por isso a familia inteira de guardas de TROCA DE SESSAO
    ' DURANTE A EXPANSAO nao tinha como ser testada: sem uma fonte que se
    ' segure no meio, nao da para dizer "a sessao caiu ENQUANTO isto
    ' carregava". Estava escrito no relatorio da Fase 2 como pendencia, com
    ' esta receita exata -- "pede um broker que segure o carregamento de
    ' filhos".
    '
    ' As travas sao TaskCompletionSource pelo mesmo motivo das outras: quem
    ' escreve o teste decide QUANDO a resposta volta, sem relogio.

    ''' <summary>Stores devolvidos por <c>GetStoresAsync</c>.</summary>
    Friend ReadOnly Stores As New List(Of StoreInfo)()

    ''' <summary>Filhas por pasta-mae. O que nao esta aqui volta vazio.</summary>
    Friend ReadOnly Filhas As New Dictionary(Of String, List(Of FolderInfo))()

    Friend TravaDosStores As TaskCompletionSource(Of Boolean)

    ''' <summary>
    ''' Segura a página em voo. É o que permite provar o instante da TROCA de
    ''' pasta: sem ela, o reload da pasta A já terminou quando a B chega, e a
    ''' fila de um do <c>Despachar</c> nunca é exercitada.
    ''' </summary>
    Friend TravaDaPagina As TaskCompletionSource(Of Boolean)
    Friend TravaDasFilhas As TaskCompletionSource(Of Boolean)

    Friend FalhaAoListarStores As ErrorKind = ErrorKind.None
    Friend FalhaAoListarFilhas As ErrorKind = ErrorKind.None

    ''' <summary>Quantas vezes cada pasta-mae foi pedida.</summary>
    Friend ReadOnly PedidosDeFilhas As New List(Of String)()

    ''' <summary>Um store com raiz, do jeito que a arvore espera.</summary>
    Friend Function ComStore(nome As String, id As String) As FakeBroker
        Stores.Add(New StoreInfo With {
            .DisplayName = nome, .StoreId = id,
            .ExchangeStoreType = "olPrimaryExchangeMailbox",
            .IsCachedExchange = True,
            .RootFolder = New FolderKey("raiz", id)})
        Return Me
    End Function

    ''' <summary>Uma pasta filha de <paramref name="mae"/>.</summary>
    Friend Function ComPasta(mae As FolderKey, nome As String, entryId As String,
                             Optional temFilhas As Boolean = False) As FakeBroker
        Dim chave = Trilha(mae)
        If Not Filhas.ContainsKey(chave) Then Filhas(chave) = New List(Of FolderInfo)()
        Filhas(chave).Add(New FolderInfo With {
            .Key = New FolderKey(entryId, mae.StoreId), .Name = nome,
            .ContentKind = FolderContentKind.Mail, .HasChildren = temFilhas})
        Return Me
    End Function

    ''' <summary>
    ''' Marca uma pasta ja registrada como de CALENDARIO.
    '''
    ''' O ContentKind vem do broker e a arvore o repassa ao no. Sem isto
    ''' nao da para exercitar a troca entre pasta de correio e pasta de
    ''' calendario -- que e a transicao onde o acervo e a agenda poderiam
    ''' se sobrepor.
    ''' </summary>
    Friend Sub MarcarComoCalendario(entryId As String)
        For Each par In Filhas
            For Each f In par.Value
                If f.Key.EntryId = entryId Then f.ContentKind = FolderContentKind.Calendar
            Next
        Next
    End Sub

    Private Shared Function Trilha(f As FolderKey) As String
        Return $"{f.StoreId}|{f.EntryId}"
    End Function

    ''' <summary>
    ''' Resposta canônica de <c>GetMessagePageAsync</c>. <c>Nothing</c> mantém o
    ''' padrão "fora da alçada" — só os testes que paginam mexem nisto.
    ''' </summary>
    Friend RespostaDaPagina As OperationResult(Of MessagePage) = Nothing

    Friend FalhaAoCriar As ErrorKind = ErrorKind.None
    Friend FalhaAoGravar As ErrorKind = ErrorKind.None
    Friend FalhaAoPreparar As ErrorKind = ErrorKind.None
    Friend ResultadoDoEnvio As ErrorKind = ErrorKind.None
    Friend FalhaAoDescartar As ErrorKind = ErrorKind.None
    Friend Modo As ModoDeDestinatario = ModoDeDestinatario.Smtp

    ''' <summary>Como a leitura da lista de destinatarios se saiu.</summary>
    Friend LeituraDeDestinatarios As PartStatus = PartStatus.Full

    ''' <summary>
    ''' O que a leitura dos ANEXOS conseguiu. Existe pelo mesmo motivo do de
    ''' cima, e chegou depois: a tela de confirmação olhava só os destinatários,
    ''' e uma lista de anexos incompleta passava.
    ''' </summary>
    Friend LeituraDeAnexos As PartStatus = PartStatus.Full

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
        As Task(Of OperationResult(Of DraftInfo)) Implements IOutlookBroker.RemoveDraftAttachmentAsync

        Chamadas.Add("detach")
        ChavesRecebidas.Add(draft)

        If Not ChaveVale(draft) Then
            Return Task.FromResult(OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "chave vencida"))
        End If

        Dim alvo = _anexos.FirstOrDefault(
            Function(x) x.FileName = attachment.FileName AndAlso x.SizeBytes = attachment.SizeBytes)

        If alvo Is Nothing Then
            Return Task.FromResult(OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "anexo"))
        End If

        _anexos.Remove(alvo)

        ' Remover SALVA: a chave gira, e as chaves dos anexos que sobraram
        ' sao reconstruidas porque o indice dos seguintes mudou.
        Dim info = Salvar()
        For i = 0 To _anexos.Count - 1
            _anexos(i).Key = New AttachmentKey(ChaveAtual().Item, i + 1,
                                               _anexos(i).FileName, _anexos(i).SizeBytes)
        Next
        info.Attachments.Clear()
        info.Attachments.AddRange(_anexos)

        Return Task.FromResult(OperationResult(Of DraftInfo).Ok(info))
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
        p.AttachmentsStatus = LeituraDeAnexos
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

    ''' <summary>
    ''' O que o probe e o connect respondem. <c>Nothing</c> mantem o
    ''' comportamento antigo -- "fora da alcada" -- para nao mudar nenhum
    ''' teste que nunca falou de sessao.
    ''' </summary>
    Friend EstadoDaSessao As SessionState? = Nothing

    ''' <summary>Segura o probe no meio, para fechar a janela durante a abertura.</summary>
    Friend TravaDoProbe As TaskCompletionSource(Of Boolean)

    Public Async Function ConnectAsync(cancel As CancellationToken) As Task(Of SessionState) _
        Implements IOutlookBroker.ConnectAsync
        If Not EstadoDaSessao.HasValue Then Return Await ForaDaAlcada(Of SessionState)()
        Chamadas.Add("Connect")
        Return EstadoDaSessao.Value
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
        Return ForaDaAlcada(Of OperationResult(Of AppointmentWindow))()
    End Function

    Public Async Function ProbeAsync(cancel As CancellationToken) As Task(Of SessionState) _
        Implements IOutlookBroker.ProbeAsync
        If Not EstadoDaSessao.HasValue Then Return Await ForaDaAlcada(Of SessionState)()
        Chamadas.Add("Probe")
        If TravaDoProbe IsNot Nothing Then Await TravaDoProbe.Task
        Return EstadoDaSessao.Value
    End Function

    ''' <summary>
    ''' Sem store configurado continua "fora da alcada", para nao mudar o
    ''' comportamento dos testes que nunca pediram arvore nenhuma.
    ''' </summary>
    Public Async Function GetStoresAsync(cancel As CancellationToken) _
        As Task(Of OperationResult(Of IReadOnlyList(Of StoreInfo))) Implements IOutlookBroker.GetStoresAsync
        Chamadas.Add("GetStores")
        If Stores.Count = 0 AndAlso FalhaAoListarStores = ErrorKind.None Then
            Return Await ForaDaAlcada(Of OperationResult(Of IReadOnlyList(Of StoreInfo)))()
        End If
        If TravaDosStores IsNot Nothing Then Await TravaDosStores.Task
        Chamadas.Add("GetStores-fim")
        If FalhaAoListarStores <> ErrorKind.None Then
            Return OperationResult(Of IReadOnlyList(Of StoreInfo)).Fail(FalhaAoListarStores, "")
        End If
        Return OperationResult(Of IReadOnlyList(Of StoreInfo)).Ok(Stores.ToList())
    End Function

    Public Async Function GetFolderChildrenAsync(parent As FolderKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of IReadOnlyList(Of FolderInfo))) Implements IOutlookBroker.GetFolderChildrenAsync
        If Stores.Count = 0 AndAlso Filhas.Count = 0 AndAlso FalhaAoListarFilhas = ErrorKind.None Then
            Return Await ForaDaAlcada(Of OperationResult(Of IReadOnlyList(Of FolderInfo)))()
        End If
        SyncLock PedidosDeFilhas
            PedidosDeFilhas.Add(Trilha(parent))
        End SyncLock
        If TravaDasFilhas IsNot Nothing Then Await TravaDasFilhas.Task
        If FalhaAoListarFilhas <> ErrorKind.None Then
            Return OperationResult(Of IReadOnlyList(Of FolderInfo)).Fail(FalhaAoListarFilhas, "")
        End If
        Dim lista As List(Of FolderInfo) = Nothing
        If Not Filhas.TryGetValue(Trilha(parent), lista) Then
            Return OperationResult(Of IReadOnlyList(Of FolderInfo)).Ok(New List(Of FolderInfo)())
        End If
        Return OperationResult(Of IReadOnlyList(Of FolderInfo)).Ok(lista.ToList())
    End Function

    Public Function GetMessagePageAsync(query As MessageQuery, continuation As String,
                                        targetCount As Integer,
                                        cancel As CancellationToken) _
        As Task(Of OperationResult(Of MessagePage)) Implements IOutlookBroker.GetMessagePageAsync
        Chamadas.Add("GetMessagePage")

        ' FORA DA ALCADA LANCA AQUI, e nao dentro da Task.
        '
        ' Ao embrulhar tudo no PaginaAsync eu tinha transformado a chamada nao
        ' esperada numa Task com falha -- e um teste que esquecesse o Await
        ' passaria em silencio. A propriedade deste duplo sempre foi a
        ' contraria: chamada fora da alcada QUEBRA o teste.
        If RespostaDaPagina Is Nothing Then Return ForaDaAlcada(Of OperationResult(Of MessagePage))()

        Dim trava = TravaDaPagina
        If trava Is Nothing Then Return Task.FromResult(RespostaDaPagina)
        Return EsperarAPagina(trava)
    End Function

    Private Async Function EsperarAPagina(trava As TaskCompletionSource(Of Boolean)) _
        As Task(Of OperationResult(Of MessagePage))
        Await trava.Task
        Return RespostaDaPagina
    End Function

    ''' <summary>
    ''' Detalhes por chave de item COMPLETA — <c>EntryId</c> e <c>StoreId</c>.
    '''
    ''' Indexar so pelo <c>EntryId</c> deixaria uma leitura do store errado
    ''' passar calada, e a propriedade deste duplo sempre foi a contraria:
    ''' chamada fora da alcada quebra o teste em vez de passar por sorte.
    ''' </summary>
    Friend ReadOnly Detalhes As New Dictionary(Of String, MessageDetail)()

    Friend Sub ComDetalhe(d As MessageDetail)
        Detalhes(Item(d.Key)) = d
    End Sub

    Private Shared Function Item(k As ItemKey) As String
        Return $"{k.StoreId}|{k.EntryId}"
    End Function

    Public Function GetMessageDetailAsync(item As ItemKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of MessageDetail)) Implements IOutlookBroker.GetMessageDetailAsync
        Dim d As MessageDetail = Nothing
        If Not Detalhes.TryGetValue(FakeBroker.Item(item), d) Then
            Return ForaDaAlcada(Of OperationResult(Of MessageDetail))()
        End If
        Chamadas.Add("GetMessageDetail")
        Return Task.FromResult(OperationResult(Of MessageDetail).Ok(d))
    End Function

    ''' <summary>
    ''' As duas leituras que o contexto de produção faz, e só elas.
    '''
    ''' Continuam <b>fora da alçada por padrão</b>: quem não as configurar
    ''' recebe a exceção de sempre. Existem porque provar
    ''' <c>ContextoDoOutlook</c> exige um broker que responda — e provar o
    ''' caminho de produção sobre uma imitação do próprio caminho não prova
    ''' nada.
    ''' </summary>
    Friend Rotulos As Func(Of IReadOnlyList(Of ItemKey),
                              OperationResult(Of IReadOnlyList(Of LabelReading)))
    Friend Anexos As Func(Of IReadOnlyList(Of ItemKey),
                             OperationResult(Of IReadOnlyList(Of AttachmentPresence)))

    Public Function GetAttachmentPresenceAsync(items As IReadOnlyList(Of ItemKey),
                                               cancel As CancellationToken) _
        As Task(Of OperationResult(Of IReadOnlyList(Of AttachmentPresence))) _
        Implements IOutlookBroker.GetAttachmentPresenceAsync
        Chamadas.Add("outlook.getAttachmentPresence")
        If Anexos Is Nothing Then
            Return ForaDaAlcada(Of OperationResult(Of IReadOnlyList(Of AttachmentPresence)))()
        End If
        Return Task.FromResult(Anexos(items))
    End Function

    ''' <summary>
    ''' O instantâneo de cada item. Fora da alçada por padrão, como as outras
    ''' duas leituras que o contexto de produção usa.
    ''' </summary>
    Friend Instantaneos As Func(Of ItemKey, OperationResult(Of MessageSnapshot))

    Public Function GetMessageSnapshotAsync(item As ItemKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of MessageSnapshot)) _
        Implements IOutlookBroker.GetMessageSnapshotAsync
        Chamadas.Add("outlook.getMessageSnapshot")
        If Instantaneos IsNot Nothing Then Return Task.FromResult(Instantaneos(item))
        Return ForaDaAlcada(Of OperationResult(Of MessageSnapshot))()
    End Function

    Public Function GetSensitivityLabelsAsync(items As IReadOnlyList(Of ItemKey),
                                              cancel As CancellationToken) _
        As Task(Of OperationResult(Of IReadOnlyList(Of LabelReading))) _
        Implements IOutlookBroker.GetSensitivityLabelsAsync
        Chamadas.Add("outlook.getSensitivityLabels")
        If Rotulos IsNot Nothing Then Return Task.FromResult(Rotulos(items))
        Return ForaDaAlcada(Of OperationResult(Of IReadOnlyList(Of LabelReading)))()
    End Function

    Public Function ProbeLabelSemanticsAsync(item As ItemKey, cancel As CancellationToken) _
        As Task(Of OperationResult(Of NamedPropertyProbe)) _
        Implements IOutlookBroker.ProbeLabelSemanticsAsync
        Return ForaDaAlcada(Of OperationResult(Of NamedPropertyProbe))()
    End Function

    Public Function ProbeLabelColumnAsync(folder As FolderKey, quantas As Integer,
                                          cancel As CancellationToken) _
        As Task(Of OperationResult(Of LabelColumnProbe)) _
        Implements IOutlookBroker.ProbeLabelColumnAsync
        Return ForaDaAlcada(Of OperationResult(Of LabelColumnProbe))()
    End Function

    ' ---- O LEITOR DE MENSAGEM ------------------------------------------
    '
    ' Mesma historia da arvore: sem uma gravacao de anexo que se segure no
    ' meio, "salvar durante a troca de mensagem" nao tem como ser escrito.

    ''' <summary>Segura a gravacao do anexo no meio.</summary>
    Friend TravaDoAnexo As TaskCompletionSource(Of Boolean)

    ''' <summary>Segura a marcacao de lida no meio.</summary>
    Friend TravaDaLeitura As TaskCompletionSource(Of Boolean)

    ''' <summary>Como a gravacao termina. <c>None</c> e sucesso.</summary>
    Friend FalhaAoSalvarAnexo As ErrorKind = ErrorKind.None

    ''' <summary>Como a marcacao termina. <c>None</c> e sucesso.</summary>
    Friend FalhaAoMarcarLida As ErrorKind = ErrorKind.None

    ''' <summary>Liga os dois metodos acima; sem isto continuam "fora da alcada".</summary>
    Friend LeitorLigado As Boolean

    ''' <summary>
    ''' Ligar o leitor nao pode virar "aceita qualquer coisa". A chave tem de
    ''' ser de um item conhecido, e <c>MarkRead</c> so existe para marcar como
    ''' LIDA — o Iris nunca desmarca. Fora disso continua estourando, que e a
    ''' propriedade que este duplo sempre teve.
    ''' </summary>
    Private Function Conhecido(k As ItemKey) As Boolean
        Return k IsNot Nothing AndAlso Detalhes.ContainsKey(Item(k))
    End Function

    Public Async Function SaveAttachmentAsync(attachment As AttachmentKey, destinationPath As String,
                                              overwrite As Boolean, cancel As CancellationToken) _
        As Task(Of OperationResult(Of String)) Implements IOutlookBroker.SaveAttachmentAsync
        If Not LeitorLigado Then Return Await ForaDaAlcada(Of OperationResult(Of String))()
        If attachment Is Nothing OrElse Not Conhecido(attachment.Owner) Then
            Throw New NotSupportedException("Anexo de um item que este duplo nao conhece.")
        End If
        If String.IsNullOrWhiteSpace(destinationPath) Then
            Throw New NotSupportedException("Gravar anexo sem destino.")
        End If
        Chamadas.Add("SaveAttachment")
        If TravaDoAnexo IsNot Nothing Then Await TravaDoAnexo.Task
        Chamadas.Add("SaveAttachment-fim")
        If FalhaAoSalvarAnexo <> ErrorKind.None Then
            Return OperationResult(Of String).Fail(FalhaAoSalvarAnexo, "")
        End If
        Return OperationResult(Of String).Ok(destinationPath)
    End Function

    Public Async Function MarkReadAsync(item As ItemKey, isRead As Boolean, cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) Implements IOutlookBroker.MarkReadAsync
        If Not LeitorLigado Then Return Await ForaDaAlcada(Of OperationResult(Of Boolean))()
        If Not Conhecido(item) Then
            Throw New NotSupportedException("Marcar um item que este duplo nao conhece.")
        End If
        If Not isRead Then
            Throw New NotSupportedException("O Iris nao desmarca mensagem.")
        End If
        Chamadas.Add("MarkRead")
        If TravaDaLeitura IsNot Nothing Then Await TravaDaLeitura.Task
        ' MARCO DE CONCLUSAO. Sem ele, um teste que cobra "nao reverteu" so
        ' pode esperar tempo, e tempo nao e evidencia de que a operacao
        ' terminou.
        Chamadas.Add("MarkRead-fim")
        If FalhaAoMarcarLida <> ErrorKind.None Then
            Return OperationResult(Of Boolean).Fail(FalhaAoMarcarLida, "")
        End If
        Return OperationResult(Of Boolean).Ok(True)
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

    ''' <summary>
    ''' Escrita no calendário. Fora da alçada por padrão, como o resto: chamada
    ''' que ninguém configurou QUEBRA o teste, em vez de passar por sorte.
    ''' </summary>
    Public Function CreateAppointmentAsync(folder As FolderKey, rascunho As AppointmentDraft,
                                           cancel As CancellationToken) _
        As Task(Of OperationResult(Of AppointmentInfo)) _
        Implements IAgendaWriter.CreateAppointmentAsync
        Chamadas.Add("createAppointment")
        Return ForaDaAlcada(Of OperationResult(Of AppointmentInfo))()
    End Function

    Public Function UpdateAppointmentAsync(chave As AppointmentKey, rascunho As AppointmentDraft,
                                           cancel As CancellationToken) _
        As Task(Of OperationResult(Of AppointmentInfo)) _
        Implements IAgendaWriter.UpdateAppointmentAsync
        Chamadas.Add("updateAppointment")
        Return ForaDaAlcada(Of OperationResult(Of AppointmentInfo))()
    End Function

    Public Function DeleteAppointmentAsync(chave As AppointmentKey,
                                           cancel As CancellationToken) _
        As Task(Of OperationResult(Of Boolean)) _
        Implements IAgendaWriter.DeleteAppointmentAsync
        Chamadas.Add("deleteAppointment")
        Return ForaDaAlcada(Of OperationResult(Of Boolean))()
    End Function

End Class

''' <summary>Gravador de arquivo que não abre diálogo nenhum.</summary>
Friend NotInheritable Class FakeSaveFile
    Implements Iris.App.ISaveFileService

    Friend Escolha As String

    Public Function AskWhereToSave(suggestedName As String) As String _
        Implements Iris.App.ISaveFileService.AskWhereToSave
        Return Escolha
    End Function
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

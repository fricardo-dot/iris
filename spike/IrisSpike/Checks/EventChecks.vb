Imports IrisSpike.Broker
Imports IrisSpike.Interop
Imports Outlook = Microsoft.Office.Interop.Outlook

Namespace Checks

    ''' <summary>
    ''' Grupo D — eventos.
    '''
    ''' A conclusão que este grupo precisa sustentar NÃO é "eventos são
    ''' confiáveis". É:
    '''
    '''   eventos dão baixa latência enquanto a sessão e os sinks estão
    '''   vivos; snapshots e reconciliação detectam tudo que aconteceu fora
    '''   dessa janela.
    '''
    ''' Tudo acontece em pastas dedicadas com itens sintéticos, removidas no
    ''' fim. Nenhuma mensagem real é tocada.
    ''' </summary>
    Public NotInheritable Class EventChecks

        Private Const Group As String = "D — Eventos"
        Private Const TestFolder As String = "Iris Spike"
        Private Const DestFolder As String = "Iris Spike Destino"

        Private ReadOnly _runner As CheckRunner
        Private ReadOnly _broker As OutlookBroker

        Private ReadOnly _log As New List(Of EventRecord)()
        Private ReadOnly _logLock As New Object()
        Private _subscription As FolderSubscription
        Private _destSubscription As FolderSubscription

        Public Sub New(runner As CheckRunner, broker As OutlookBroker)
            _runner = runner
            _broker = broker
        End Sub

        Public Async Function RunAsync() As Task
            ' VB não permite Await dentro de Finally, e a limpeza é
            ' obrigatória mesmo se um critério estourar: sem ela ficam pastas
            ' e assinaturas penduradas na caixa do usuário.
            Dim failure As Exception = Nothing
            Try
                If Not Await D0_StoreAcceptsWritesAsync() Then
                    SkipEverythingDownstream()
                Else
                    Await D1_SubscribeAndSurviveGcAsync()
                    Await D2_EventSemanticsAsync()
                    Await D3_MoveWithinStoreAsync()
                    Await D5_ChangesWhileNotSubscribedAsync()
                    Await D6_BurstAsync()
                    SkipUntestable()
                End If
            Catch ex As Exception
                failure = ex
            End Try

            Await CleanupAsync()

            If failure IsNot Nothing Then
                Throw New InvalidOperationException(
                    "Grupo D interrompido; a limpeza rodou mesmo assim.", failure)
            End If
        End Function

        ' ===================================================================
        ' D0 — pré-condição
        ' ===================================================================

        ''' <summary>
        ''' Sem esta verificação o grupo inteiro produz diagnóstico errado.
        '''
        ''' Na primeira execução real o disco estava a zero byte, o Outlook
        ''' não conseguia escrever no OST, item nenhum era criado — e o D1
        ''' concluiu "o event sink não sobreviveu ao GC", causa que ele nunca
        ''' havia estabelecido. Ausência de evento depois de uma operação que
        ''' NÃO ACONTECEU não é evidência sobre eventos.
        ''' </summary>
        Private Async Function D0_StoreAcceptsWritesAsync() As Task(Of Boolean)
            Dim writable = False

            Await _runner.RunAsync(
                "D0", Group, "Pré-condição: o store aceita escrita",
                Async Function()
                    Dim before = Await CountItemsAsync(TestFolder)

                    Dim entryId As String
                    Try
                        entryId = Await CreateItemAsync(TestFolder, "D0 sonda de escrita")
                    Catch ex As Runtime.InteropServices.COMException
                        Return (CheckStatus.Fail,
                                $"Criação recusada (0x{ex.HResult:X8}): {ex.Message} " &
                                "Todo o grupo D seria inconclusivo.")
                    End Try

                    Dim after = Await CountItemsAsync(TestFolder)

                    If String.IsNullOrEmpty(entryId) Then
                        Return (CheckStatus.Fail,
                                $"Items.Add não devolveu um MailItem utilizável " &
                                $"({before} → {after} itens).")
                    End If

                    If after <= before Then
                        ' Item tem EntryID mas a pasta não cresceu: ele foi
                        ' criado em OUTRO lugar. Descobrir onde é a diferença
                        ' entre um diagnóstico e um chute.
                        Dim landed = Await LocateItemAsync(entryId)
                        Await DeleteByEntryIdAsync(entryId)

                        If landed <> "" AndAlso
                           Not String.Equals(landed, TestFolder, StringComparison.OrdinalIgnoreCase) Then
                            Return (CheckStatus.Fail,
                                    $"O item foi criado, porém em '{landed}', não em '{TestFolder}' " &
                                    $"({before} → {after}). O grupo D precisa criar o item de outra forma.")
                        End If

                        Return (CheckStatus.Fail,
                                $"O item não persistiu ({before} → {after} itens) e não foi " &
                                $"localizado por EntryID. Disco, cota ou store somente leitura.")
                    End If

                    Await DeleteItemAsync(TestFolder, entryId)
                    writable = True
                    Return (CheckStatus.Pass, $"Criação e exclusão persistem ({before} → {after}).")
                End Function)

            Return writable
        End Function

        Private Sub SkipEverythingDownstream()
            Const Reason As String =
                "D0 falhou: o store não aceita escrita. Sem conseguir criar itens, " &
                "a ausência de eventos NÃO diz nada sobre eventos."
            For Each id In {"D1", "D2", "D3", "D5", "D6"}
                _runner.Skip(id, Group, "(depende do D0)", Reason)
            Next
            SkipUntestable()
        End Sub

        ' ===================================================================
        ' D1 — assinatura, sobrevivência ao GC, thread de entrega
        ' ===================================================================
        Private Async Function D1_SubscribeAndSurviveGcAsync() As Task
            Await _runner.RunAsync(
                "D1", Group, "Eventos chegam na thread do broker e sobrevivem ao GC",
                Async Function()
                    Await SubscribeAsync(TestFolder, isDestination:=False)

                    ' Se a coleção Items não tivesse referência forte, o GC
                    ' levaria o event sink junto e os eventos parariam sem
                    ' erro nenhum — o modo de falha do R7.
                    For i = 1 To 3
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    Next

                    ClearLog()
                    Dim entryId = Await CreateItemAsync(TestFolder, "D1 pos-GC")
                    Dim arrived = Await WaitForEventsAsync(1, 10000)

                    If Not arrived Then
                        ' O D0 já garantiu que a escrita persiste, então isto
                        ' É sobre eventos. As hipóteses seguem em aberto e
                        ' não devem ser apresentadas como conclusão.
                        Return (CheckStatus.Fail,
                                "Item criado e persistido, mas nenhum ItemAdd em 10s. " &
                                "Hipóteses a separar: sink coletado pelo GC, EmbedInteropTypes " &
                                "quebrando o event sink COM, ou o Outlook não emitir evento " &
                                "para criação programática.")
                    End If

                    Dim records = Snapshot()
                    Dim strayThreads = records.Where(Function(r) r.ThreadId <> _broker.ThreadId).
                                               Select(Function(r) r.ThreadId).Distinct().ToList()

                    If strayThreads.Count > 0 Then
                        Return (CheckStatus.Fail,
                                $"A LEITURA do COM aconteceu nas threads " &
                                $"{String.Join(",", strayThreads)}, fora da thread do broker " &
                                $"({_broker.ThreadId}). Ver R6.")
                    End If

                    ' Achado da Fase 0: a ENTREGA nao acontece na STA do
                    ' broker. Isso e fato do ambiente, nao defeito nosso; o
                    ' que importa e que a leitura do COM foi remarcada para a
                    ' thread dona dos objetos.
                    Dim delivery = String.Join(", ",
                        records.Select(Function(r) $"{r.DeliveryThreadId}/{r.DeliveryApartment}").Distinct())

                    Return (CheckStatus.Pass,
                            $"{records.Count} evento(s) após 3 ciclos de GC. " &
                            $"ENTREGA em {delivery}; LEITURA do COM remarcada para a thread " &
                            $"{_broker.ThreadId}/{_broker.Apartment} do broker.")
                End Function)
        End Function

        ' ===================================================================
        ' D2 — quantos eventos cada operação gera de fato
        ' ===================================================================
        Private Async Function D2_EventSemanticsAsync() As Task
            Await _runner.RunAsync(
                "D2", Group, "Semântica: eventos por operação",
                Async Function()
                    If _subscription Is Nothing Then
                        Return (CheckStatus.Skipped, "Sem assinatura ativa.")
                    End If

                    ClearLog()
                    Dim entryId = Await CreateItemAsync(TestFolder, "D2 criacao")
                    Await WaitForEventsAsync(1, 6000)
                    Dim onCreate = Snapshot()

                    ClearLog()
                    Await ModifyItemAsync(TestFolder, entryId, "D2 alterado uma vez")
                    Await WaitForEventsAsync(1, 6000)
                    Dim onChange = Snapshot()

                    ClearLog()
                    Await ModifyItemAsync(TestFolder, entryId, "D2 alterado varias", setBody:=True)
                    Await WaitForEventsAsync(1, 6000)
                    Dim onMulti = Snapshot()

                    ClearLog()
                    Await DeleteItemAsync(TestFolder, entryId)
                    Await WaitForEventsAsync(1, 6000)
                    Dim onDelete = Snapshot()

                    ' Não existe resultado "errado" aqui. O objetivo é
                    ' registrar o comportamento real, porque a Fase 2 não
                    ' pode presumir um evento por operação.
                    Return (CheckStatus.Info,
                            $"criar → {Describe(onCreate)}; " &
                            $"alterar 1 prop → {Describe(onChange)}; " &
                            $"alterar 2 props num Save → {Describe(onMulti)}; " &
                            $"excluir → {Describe(onDelete)}.")
                End Function)
        End Function

        ' ===================================================================
        ' D3 — movimento dentro do mesmo store
        ' ===================================================================
        Private Async Function D3_MoveWithinStoreAsync() As Task
            Await _runner.RunAsync(
                "D3", Group, "Movimento no mesmo store: o EntryID muda?",
                Async Function()
                    Await SubscribeAsync(DestFolder, isDestination:=True)

                    Dim before = Await CreateItemAsync(TestFolder, "D3 movimento")
                    If String.IsNullOrEmpty(before) Then
                        Return (CheckStatus.Fail, "Não foi possível criar o item de teste.")
                    End If
                    Await Task.Delay(1500)

                    ClearLog()
                    Dim after = Await _broker.ReadAsync(
                        Function(app, ns)
                            Dim source = EnsureFolder(_broker, TestFolder)
                            Try
                                Dim dest = EnsureFolder(_broker, DestFolder)
                                Try
                                    Dim item = FindItem(source, before)
                                    If item Is Nothing Then Return ""
                                    Try
                                        ' Move devolve um NOVO RCW, com dono
                                        ' próprio e liberação própria.
                                        Dim moved As Outlook.MailItem = Nothing
                                        Try
                                            moved = TryCast(item.Move(dest), Outlook.MailItem)
                                            Return If(moved Is Nothing, "", moved.EntryID)
                                        Finally
                                            ComHelpers.Release(moved)
                                        End Try
                                    Finally
                                        ComHelpers.Release(item)
                                    End Try
                                Finally
                                    ComHelpers.Release(dest)
                                End Try
                            Finally
                                ComHelpers.Release(source)
                            End Try
                        End Function)

                    Await WaitForEventsAsync(2, 10000)
                    Dim records = Snapshot()

                    If String.IsNullOrEmpty(after) Then
                        Return (CheckStatus.Fail, "O item não pôde ser movido.")
                    End If

                    Dim changed = Not String.Equals(before, after, StringComparison.Ordinal)
                    Dim sequence = If(records.Count = 0, "(nenhum)",
                        String.Join(" → ", records.Select(Function(r) $"{r.Kind}@{r.Folder}")))

                    Dim note = $"EntryID {If(changed, "MUDOU", "permaneceu igual")}. Eventos: {sequence}."
                    If changed Then
                        note &= " Confirma a seção 5 do ESCOPO: EntryID não correlaciona item " &
                                "movido, e a Fase 2 precisa de chave interna própria."
                    End If

                    Return (CheckStatus.Info, note)
                End Function)
        End Function

        ' ===================================================================
        ' D5 — o teste que justifica a reconciliação
        ' ===================================================================
        Private Async Function D5_ChangesWhileNotSubscribedAsync() As Task
            Await _runner.RunAsync(
                "D5", Group, "Mudanças com a assinatura cancelada não são recuperadas",
                Async Function()
                    Dim countBefore = Await CountItemsAsync(TestFolder)

                    ' Cancelar a assinatura equivale ao Iris fechado.
                    Await UnsubscribeAsync(destination:=False)

                    Dim created As New List(Of String)()
                    For i = 1 To 3
                        created.Add(Await CreateItemAsync(TestFolder, $"D5 sem assinatura {i}"))
                    Next
                    Await DeleteItemAsync(TestFolder, created(0))

                    Dim countAfter = Await CountItemsAsync(TestFolder)
                    Dim delta = countAfter - countBefore

                    ' Reassina e observa se algo é reproduzido.
                    ClearLog()
                    Await SubscribeAsync(TestFolder, isDestination:=False)
                    Await Task.Delay(3000)
                    Dim replayed = Snapshot()

                    Dim note = $"3 criações e 1 exclusão sem assinatura. " &
                               $"Eventos reproduzidos ao reassinar: {replayed.Count}. " &
                               $"Snapshot viu {delta:+#;-#;0} item(ns)."

                    If delta = 0 Then
                        ' Sem escrita efetiva não há o que concluir — evita o
                        ' falso verde da execução com disco cheio.
                        Return (CheckStatus.Fail,
                                note & " As operações não persistiram; teste inconclusivo.")
                    End If

                    If replayed.Count > 0 Then
                        Return (CheckStatus.Warn,
                                note & " Chegaram eventos após reassinar — investigar.")
                    End If

                    Return (CheckStatus.Pass,
                            note & " Confirma o R8: eventos não recuperam o que passou fora da " &
                            "janela de escuta; só o snapshot enxergou. A Fase 2 precisa de " &
                            "reconciliação com checkpoint, não apenas de eventos.")
                End Function)
        End Function

        ' ===================================================================
        ' D6 — volume
        ' ===================================================================
        Private Async Function D6_BurstAsync() As Task
            Await _runner.RunAsync(
                "D6", Group, "Rajada: 25 criações seguidas",
                Async Function()
                    If _subscription Is Nothing Then
                        Return (CheckStatus.Skipped, "Sem assinatura ativa.")
                    End If

                    Const Burst As Integer = 25
                    ClearLog()

                    Dim createdOk = 0
                    For i = 1 To Burst
                        Dim id = Await CreateItemAsync(TestFolder, $"D6 rajada {i:00}")
                        If Not String.IsNullOrEmpty(id) Then createdOk += 1
                    Next

                    Await WaitForEventsAsync(Burst, 25000)
                    Dim records = Snapshot()
                    Dim adds = Enumerable.Count(records, Function(r) r.Kind = "ItemAdd")

                    Dim note = $"{createdOk}/{Burst} itens criados → {adds} ItemAdd " &
                               $"({records.Count} eventos no total)."

                    If createdOk < Burst Then
                        Return (CheckStatus.Fail,
                                note & " Nem todas as criações persistiram; comparação inválida.")
                    End If

                    If adds < createdOk Then
                        Return (CheckStatus.Warn,
                                note & $" PERDA de {createdOk - adds} evento(s) sob rajada — " &
                                "evidência direta a favor da reconciliação da Fase 2.")
                    End If

                    If adds > createdOk Then
                        Return (CheckStatus.Warn, note & " Eventos DUPLICADOS sob rajada.")
                    End If

                    Return (CheckStatus.Info,
                            note & " Nenhuma perda neste volume. Não generalizar: 25 itens " &
                            "não é uma caixa em sincronização inicial.")
                End Function)
        End Function

        Private Sub SkipUntestable()
            _runner.Skip("D4", Group, "Movimento entre stores",
                         "Só existe um store (Exchange). Sem PST ou caixa secundária não é " &
                         "testável — NÃO inferir comportamento.")
            _runner.Skip("D7", Group, "Reinício do Outlook com assinatura ativa",
                         "Exige fechar e reabrir o Outlook durante a execução.")
        End Sub

        ' ===================================================================
        ' Assinaturas
        ' ===================================================================

        Private Async Function SubscribeAsync(folderName As String, isDestination As Boolean) As Task
            Await _broker.InvokeAsync(
                Sub()
                    ' A subscription vira dona da pasta E da coleção. Liberar
                    ' a pasta aqui desconectaria a fonte de eventos.
                    Dim folder = EnsureFolder(_broker, folderName)
                    Dim subscription As New FolderSubscription(
                        folderName, folder, AddressOf Record, AddressOf _broker.PostToBrokerThread)
                    _broker.TrackSink(subscription)
                    If isDestination Then
                        _destSubscription = subscription
                    Else
                        _subscription = subscription
                    End If
                End Sub)
        End Function

        Private Async Function UnsubscribeAsync(destination As Boolean) As Task
            Await _broker.InvokeAsync(
                Sub()
                    Dim subscription = If(destination, _destSubscription, _subscription)
                    If subscription Is Nothing Then Return
                    _broker.ReleaseSink(subscription)
                    If destination Then
                        _destSubscription = Nothing
                    Else
                        _subscription = Nothing
                    End If
                End Sub)
        End Function

        ' ===================================================================
        ' Log de eventos
        ' ===================================================================

        Private Sub Record(record As EventRecord)
            SyncLock _logLock
                _log.Add(record)
            End SyncLock
        End Sub

        Private Sub ClearLog()
            SyncLock _logLock
                _log.Clear()
            End SyncLock
        End Sub

        Private Function Snapshot() As List(Of EventRecord)
            SyncLock _logLock
                Return New List(Of EventRecord)(_log)
            End SyncLock
        End Function

        Private Async Function WaitForEventsAsync(minCount As Integer, timeoutMs As Integer) As Task(Of Boolean)
            Dim deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs)
            While DateTime.UtcNow < deadline
                SyncLock _logLock
                    If _log.Count >= minCount Then Return True
                End SyncLock
                Await Task.Delay(100)
            End While
            Return False
        End Function

        Private Shared Function Describe(records As List(Of EventRecord)) As String
            If records.Count = 0 Then Return "nenhum evento"
            Return String.Join(" + ",
                records.GroupBy(Function(r) r.Kind).
                        Select(Function(g) $"{g.Count()}×{g.Key}"))
        End Function

        ' ===================================================================
        ' Acesso ao Outlook — tudo roda na thread do broker
        ' ===================================================================

        ''' <summary>
        ''' Obtém, ou cria, uma subpasta da Caixa de Entrada. Devolve um RCW:
        ''' quem chama é dono e precisa liberar (ou entregar a posse a uma
        ''' FolderSubscription).
        ''' </summary>
        Private Shared Function EnsureFolder(broker As OutlookBroker, name As String) As Outlook.MAPIFolder
            broker.AssertOnBrokerThread()
            Return broker.WithNamespace(
                Function(ns)
                    Dim inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
                    Try
                        Dim folders = inbox.Folders
                        Try
                            For i = 1 To folders.Count
                                Dim candidate = folders.Item(i)
                                If String.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase) Then
                                    Return candidate
                                End If
                                ComHelpers.Release(candidate)
                            Next
                            Return folders.Add(name)
                        Finally
                            ComHelpers.Release(folders)
                        End Try
                    Finally
                        ComHelpers.Release(inbox)
                    End Try
                End Function)
        End Function

        Private Shared Function FindItem(folder As Outlook.MAPIFolder, entryId As String) As Outlook.MailItem
            Dim items = folder.Items
            Try
                For i = 1 To items.Count
                    Dim raw = items.Item(i)
                    Dim mail = TryCast(raw, Outlook.MailItem)
                    If mail Is Nothing Then
                        ComHelpers.Release(raw)
                        Continue For
                    End If
                    If String.Equals(mail.EntryID, entryId, StringComparison.Ordinal) Then
                        Return mail
                    End If
                    ComHelpers.Release(raw)
                Next
                Return Nothing
            Finally
                ComHelpers.Release(items)
            End Try
        End Function

        Private Async Function CreateItemAsync(folderName As String, subject As String) As Task(Of String)
            Return Await _broker.ReadAsync(
                Function(app, ns)
                    Dim folder = EnsureFolder(_broker, folderName)
                    Try
                        Dim items = folder.Items
                        Try
                            Dim item As Outlook.MailItem = Nothing
                            Try
                                item = TryCast(items.Add(Outlook.OlItemType.olMailItem), Outlook.MailItem)
                                If item Is Nothing Then Return ""
                                item.Subject = $"[IRIS-SPIKE] {subject}"
                                item.Body = "Item sintético do spike da Fase 0. Pode apagar."
                                item.Save()

                                ' DESCOBERTA DA FASE 0: uma mensagem NÃO
                                ' ENVIADA vai para Rascunhos ao ser salva,
                                ' independentemente da pasta onde Items.Add
                                ' foi chamado. Sem mover, o item existe — só
                                ' que no lugar errado, e a pasta de teste
                                ' continua vazia.
                                Dim moved As Outlook.MailItem = Nothing
                                Try
                                    moved = TryCast(item.Move(folder), Outlook.MailItem)
                                    Return If(moved Is Nothing, "", moved.EntryID)
                                Finally
                                    ComHelpers.Release(moved)
                                End Try
                            Finally
                                ComHelpers.Release(item)
                            End Try
                        Finally
                            ComHelpers.Release(items)
                        End Try
                    Finally
                        ComHelpers.Release(folder)
                    End Try
                End Function)
        End Function

        Private Async Function ModifyItemAsync(folderName As String,
                                               entryId As String,
                                               subject As String,
                                               Optional setBody As Boolean = False) As Task
            Await _broker.ReadAsync(
                Function(app, ns)
                    Dim folder = EnsureFolder(_broker, folderName)
                    Try
                        Dim item = FindItem(folder, entryId)
                        If item Is Nothing Then Return False
                        Try
                            item.Subject = $"[IRIS-SPIKE] {subject}"
                            If setBody Then
                                item.Body = "corpo alterado em " & DateTime.Now.ToString("HH:mm:ss.fff")
                            End If
                            item.Save()
                            Return True
                        Finally
                            ComHelpers.Release(item)
                        End Try
                    Finally
                        ComHelpers.Release(folder)
                    End Try
                End Function)
        End Function

        Private Async Function DeleteItemAsync(folderName As String, entryId As String) As Task
            Await _broker.ReadAsync(
                Function(app, ns)
                    Dim folder = EnsureFolder(_broker, folderName)
                    Try
                        Dim item = FindItem(folder, entryId)
                        If item Is Nothing Then Return False
                        Try
                            item.Delete()
                            Return True
                        Finally
                            ComHelpers.Release(item)
                        End Try
                    Finally
                        ComHelpers.Release(folder)
                    End Try
                End Function)
        End Function


        ''' <summary>
        ''' Onde o item realmente esta, pelo EntryID. Responde "o item sumiu"
        ''' com um lugar, em vez de uma suposicao.
        ''' </summary>
        Private Async Function LocateItemAsync(entryId As String) As Task(Of String)
            Return Await _broker.ReadAsync(
                Function(app, ns)
                    Try
                        Dim item = TryCast(ns.GetItemFromID(entryId), Outlook.MailItem)
                        If item Is Nothing Then Return ""
                        Try
                            Dim parent = TryCast(item.Parent, Outlook.MAPIFolder)
                            If parent Is Nothing Then Return "(sem pasta)"
                            Try
                                Return parent.Name
                            Finally
                                ComHelpers.Release(parent)
                            End Try
                        Finally
                            ComHelpers.Release(item)
                        End Try
                    Catch
                        Return ""
                    End Try
                End Function)
        End Function

        Private Async Function DeleteByEntryIdAsync(entryId As String) As Task
            Await _broker.ReadAsync(
                Function(app, ns)
                    Try
                        Dim item = TryCast(ns.GetItemFromID(entryId), Outlook.MailItem)
                        If item Is Nothing Then Return False
                        Try
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

        Private Async Function CountItemsAsync(folderName As String) As Task(Of Integer)
            Return Await _broker.ReadAsync(
                Function(app, ns)
                    Dim folder = EnsureFolder(_broker, folderName)
                    Try
                        Dim items = folder.Items
                        Try
                            Return items.Count
                        Finally
                            ComHelpers.Release(items)
                        End Try
                    Finally
                        ComHelpers.Release(folder)
                    End Try
                End Function)
        End Function

        ''' <summary>
        ''' Desfaz tudo que o grupo criou. Um spike não pode deixar lixo na
        ''' caixa de correio de ninguém.
        ''' </summary>
        Private Async Function CleanupAsync() As Task
            Await _runner.RunAsync(
                "D8", Group, "Limpeza: assinaturas canceladas e pastas removidas",
                Async Function()
                    Await UnsubscribeAsync(destination:=False)
                    Await UnsubscribeAsync(destination:=True)

                    Dim removed = Await _broker.InvokeAsync(
                        Function()
                            Dim count = 0
                            For Each name In {TestFolder, DestFolder}
                                Try
                                    Dim folder = EnsureFolder(_broker, name)
                                    Try
                                        folder.Delete()
                                        count += 1
                                    Finally
                                        ComHelpers.Release(folder)
                                    End Try
                                Catch
                                    ' Pasta pode não existir.
                                End Try
                            Next
                            Return count
                        End Function)

                    Return (CheckStatus.Pass,
                            $"Assinaturas canceladas; {removed} pasta(s) de teste removida(s).")
                End Function)
        End Function

    End Class

End Namespace

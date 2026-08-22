Imports IrisSpike.Interop
Imports Outlook = Microsoft.Office.Interop.Outlook

Namespace Broker

    ''' <summary>
    ''' Um evento observado, já convertido para dados. O objeto COM que o
    ''' Outlook entrega no callback é lido e liberado na thread do broker —
    ''' nada de COM sai daqui.
    ''' </summary>
    Public NotInheritable Class EventRecord
        Public Property Kind As String
        Public Property At As DateTime
        ''' <summary>Thread onde o COM foi realmente lido (deve ser a do broker).</summary>
        Public Property ThreadId As Integer
        Public Property Apartment As String
        ''' <summary>Thread onde o Outlook ENTREGOU o callback.</summary>
        Public Property DeliveryThreadId As Integer
        Public Property DeliveryApartment As String
        Public Property Folder As String
        Public Property EntryId As String
        Public Property Subject As String
        Public Property MessageClass As String

        Public Overrides Function ToString() As String
            Return $"{Kind}@{Folder} ""{Subject}"""
        End Function
    End Class

    ''' <summary>
    ''' Assinatura de eventos de uma pasta.
    '''
    ''' DESCOBERTA DA FASE 0, que contradiz a premissa original: o Outlook
    ''' NÃO entrega os callbacks na thread STA que assinou. Medido nesta
    ''' máquina, eles chegam numa thread MTA do pool. Consequências:
    '''
    '''   • Ler MailItem.Subject direto no handler é tocar COM fora da thread
    '''     dona do objeto — o R6. Funciona por marshaling implícito, e é
    '''     assim que se constrói travamento sob carga.
    '''   • A leitura é despachada para o dispatcher do broker, onde os
    '''     objetos moram.
    '''   • O despacho é ASSÍNCRONO de propósito. Bloquear o callback à
    '''     espera da STA arriscaria deadlock: se o Outlook dispara o evento
    '''     de dentro de uma chamada que a própria STA está executando, a STA
    '''     só se libera quando o handler retorna, e o handler estaria
    '''     esperando a STA.
    '''
    ''' A referência forte à coleção Items E à pasta pai é obrigatória (R7):
    ''' se o GC as coletar, o event sink morre junto e os eventos param sem
    ''' erro nenhum.
    ''' </summary>
    Public NotInheritable Class FolderSubscription
        Implements IDisposable

        Private _folder As Outlook.MAPIFolder
        Private _items As Outlook.Items
        Private _onAdd As Outlook.ItemsEvents_ItemAddEventHandler
        Private _onChange As Outlook.ItemsEvents_ItemChangeEventHandler
        Private _onRemove As Outlook.ItemsEvents_ItemRemoveEventHandler
        Private _disposed As Boolean

        Public ReadOnly Property FolderName As String

        ''' <summary>
        ''' Precisa ser construída DENTRO da thread STA do broker.
        ''' </summary>
        ''' <param name="folder">
        ''' A assinatura vira DONA da pasta e da coleção; quem chama não deve
        ''' liberar nenhuma das duas.
        ''' </param>
        ''' <param name="postToBroker">
        ''' Enfileira trabalho na thread do broker sem bloquear o chamador.
        ''' </param>
        Public Sub New(folderName As String,
                       folder As Outlook.MAPIFolder,
                       sink As Action(Of EventRecord),
                       postToBroker As Action(Of Action))
            _FolderName = folderName
            _folder = folder
            _items = folder.Items

            _onAdd = Sub(item) Handle("ItemAdd", item, sink, postToBroker)
            _onChange = Sub(item) Handle("ItemChange", item, sink, postToBroker)
            _onRemove = Sub() HandleRemove(sink, postToBroker)

            AddHandler _items.ItemAdd, _onAdd
            AddHandler _items.ItemChange, _onChange
            AddHandler _items.ItemRemove, _onRemove
        End Sub

        ''' <summary>
        ''' Roda na thread de ENTREGA (MTA). Não toca em nenhuma propriedade
        ''' COM: apenas anota onde o callback chegou e devolve o trabalho ao
        ''' broker.
        ''' </summary>
        Private Sub Handle(kind As String,
                           item As Object,
                           sink As Action(Of EventRecord),
                           postToBroker As Action(Of Action))
            Dim deliveryThread = Environment.CurrentManagedThreadId
            Dim deliveryApartment = Threading.Thread.CurrentThread.GetApartmentState().ToString()
            Dim folderName = _FolderName

            postToBroker(
                Sub()
                    ' Aqui já é a STA do broker: ler COM é seguro.
                    Dim record = Describe(kind, folderName, item)
                    record.DeliveryThreadId = deliveryThread
                    record.DeliveryApartment = deliveryApartment
                    sink(record)
                End Sub)
        End Sub

        Private Sub HandleRemove(sink As Action(Of EventRecord),
                                 postToBroker As Action(Of Action))
            Dim deliveryThread = Environment.CurrentManagedThreadId
            Dim deliveryApartment = Threading.Thread.CurrentThread.GetApartmentState().ToString()
            Dim folderName = _FolderName

            postToBroker(
                Sub()
                    sink(New EventRecord With {
                        .Kind = "ItemRemove",
                        .At = DateTime.UtcNow,
                        .ThreadId = Environment.CurrentManagedThreadId,
                        .Apartment = Threading.Thread.CurrentThread.GetApartmentState().ToString(),
                        .DeliveryThreadId = deliveryThread,
                        .DeliveryApartment = deliveryApartment,
                        .Folder = folderName,
                        .EntryId = "",
                        .Subject = "(ItemRemove não entrega o item)",
                        .MessageClass = ""
                    })
                End Sub)
        End Sub

        ''' <summary>
        ''' Extrai os dados e libera o objeto COM. Roda na thread do broker.
        ''' </summary>
        Private Shared Function Describe(kind As String, folderName As String, item As Object) As EventRecord
            Dim record As New EventRecord With {
                .Kind = kind,
                .At = DateTime.UtcNow,
                .ThreadId = Environment.CurrentManagedThreadId,
                .Apartment = Threading.Thread.CurrentThread.GetApartmentState().ToString(),
                .Folder = folderName,
                .EntryId = "",
                .Subject = "",
                .MessageClass = ""
            }

            Try
                Dim mail = TryCast(item, Outlook.MailItem)
                If mail IsNot Nothing Then
                    Try : record.EntryId = mail.EntryID : Catch : End Try
                    Try : record.Subject = mail.Subject : Catch : End Try
                    Try : record.MessageClass = mail.MessageClass : Catch : End Try
                Else
                    record.MessageClass = "(não é MailItem)"
                End If
            Finally
                ComHelpers.Release(item)
            End Try

            Return record
        End Function

        ''' <summary>
        ''' Ordem obrigatória: RemoveHandler ANTES de liberar os RCWs.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True

            If _items Is Nothing Then Return

            Try
                RemoveHandler _items.ItemAdd, _onAdd
                RemoveHandler _items.ItemChange, _onChange
                RemoveHandler _items.ItemRemove, _onRemove
            Catch
                ' Outlook já pode ter ido embora; seguir para a liberação.
            End Try

            _onAdd = Nothing
            _onChange = Nothing
            _onRemove = Nothing

            ' Ordem inversa da aquisição: coleção, depois pasta.
            ComHelpers.Release(_items)
            _items = Nothing
            ComHelpers.Release(_folder)
            _folder = Nothing
        End Sub
    End Class

End Namespace

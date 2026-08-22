Imports IrisSpike.Interop
Imports Outlook = Microsoft.Office.Interop.Outlook

Namespace Broker

    ''' <summary>
    ''' Um evento observado, já convertido para dados. O objeto COM que o
    ''' Outlook entrega no callback é lido e liberado dentro do handler —
    ''' nada de COM sai daqui.
    ''' </summary>
    Public NotInheritable Class EventRecord
        Public Property Kind As String
        Public Property At As DateTime
        Public Property ThreadId As Integer
        Public Property Apartment As String
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
    ''' Guardar apenas o objeto Items numa lista — como a primeira versão do
    ''' broker fazia — mantém a coleção viva, mas torna impossível cancelar:
    ''' sem os delegates originais não há RemoveHandler. Esta classe guarda a
    ''' coleção E os delegates, e desconecta na ordem certa: primeiro os
    ''' handlers, depois o RCW.
    '''
    ''' A referência forte a <c>_items</c> é obrigatória (R7): se o GC
    ''' coletar a coleção, o event sink morre junto e os eventos param sem
    ''' erro nenhum.
    ''' </summary>
    Public NotInheritable Class FolderSubscription
        Implements IDisposable

        ''' <summary>
        ''' A pasta PAI precisa ficar viva junto com a coleção. Liberar o
        ''' MAPIFolder logo após ler folder.Items desconecta o RCW pai, e a
        ''' fonte de eventos morre junto — sem erro, sem aviso: os eventos
        ''' simplesmente nunca chegam. Foi exatamente o que aconteceu na
        ''' primeira execução do grupo D: 0 de 25 eventos.
        ''' </summary>
        Private _folder As Outlook.MAPIFolder

        Private _items As Outlook.Items
        Private _onAdd As Outlook.ItemsEvents_ItemAddEventHandler
        Private _onChange As Outlook.ItemsEvents_ItemChangeEventHandler
        Private _onRemove As Outlook.ItemsEvents_ItemRemoveEventHandler
        Private _disposed As Boolean

        Public ReadOnly Property FolderName As String

        ''' <summary>
        ''' Precisa ser construída DENTRO da thread STA do broker: o sink
        ''' pertence à thread que assina.
        ''' </summary>
        ''' <param name="folder">
        ''' A assinatura passa a ser DONA da pasta e da coleção; quem chama
        ''' não deve liberar nenhuma das duas.
        ''' </param>
        Public Sub New(folderName As String,
                       folder As Outlook.MAPIFolder,
                       sink As Action(Of EventRecord))
            _FolderName = folderName
            _folder = folder
            _items = folder.Items

            _onAdd = Sub(item) sink(Describe("ItemAdd", folderName, item))
            _onChange = Sub(item) sink(Describe("ItemChange", folderName, item))
            _onRemove = Sub() sink(New EventRecord With {
                .Kind = "ItemRemove",
                .At = DateTime.UtcNow,
                .ThreadId = Environment.CurrentManagedThreadId,
                .Apartment = Threading.Thread.CurrentThread.GetApartmentState().ToString(),
                .Folder = folderName,
                .EntryId = "",
                .Subject = "(ItemRemove não entrega o item)",
                .MessageClass = ""
            })

            AddHandler _items.ItemAdd, _onAdd
            AddHandler _items.ItemChange, _onChange
            AddHandler _items.ItemRemove, _onRemove
        End Sub

        ''' <summary>
        ''' Extrai os dados e libera o objeto COM ainda dentro do callback.
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
        ''' Ordem obrigatória: RemoveHandler ANTES de liberar o RCW. Liberar
        ''' primeiro deixaria handlers apontando para um wrapper morto.
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

Imports System.Threading
Imports IrisSpike.Interop
Imports Outlook = Microsoft.Office.Interop.Outlook

Namespace Broker

    ''' <summary>
    ''' Um evento observado, já convertido para dados. O objeto COM entregue
    ''' no callback é lido e liberado na thread do broker — nada de COM sai
    ''' daqui.
    ''' </summary>
    Public NotInheritable Class EventRecord
        Public Property Kind As String
        ''' <summary>Qual assinatura originou. Sem isto, um callback em voo
        ''' de uma assinatura já cancelada é contado como se fosse novo.</summary>
        Public Property SubscriptionId As Integer
        ''' <summary>Ordem de ENTREGA, não de processamento.</summary>
        Public Property DeliverySequence As Long
        Public Property DeliveredAt As DateTime
        Public Property ProcessedAt As DateTime
        ''' <summary>Thread onde o COM foi lido (deve ser a do broker).</summary>
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
    ''' máquina, chegam numa thread MTA do pool. Daí o desenho:
    '''
    '''   • O handler não toca em propriedade COM alguma. Anota onde chegou,
    '''     em que ordem, e devolve o trabalho ao dispatcher do broker.
    '''   • O despacho é ASSÍNCRONO. Bloquear o callback esperando a STA
    '''     arriscaria deadlock quando o Outlook dispara o evento de dentro
    '''     de uma chamada que a própria STA está executando.
    '''   • Se o post não for aceito (broker encerrando), o item é liberado
    '''     ali mesmo — senão o RCW vaza.
    '''   • Trabalho postado por uma assinatura já descartada é jogado fora,
    '''     e o item liberado. Sem isso, evento antigo em voo aparece como
    '''     evento da assinatura nova.
    '''
    ''' RESSALVA HONESTA: o RCW é materializado na MTA e só lido depois na
    ''' STA. Isso não é o mesmo que "o objeto mora na STA" — a leitura ainda
    ''' pode depender de marshaling entre apartments, e o estado lido é o do
    ''' momento do PROCESSAMENTO, não o que causou o evento. Para o produto,
    ''' o certo é tratar evento como aviso de "pasta suja" e reler o estado
    ''' atual, não como transição a aplicar cegamente.
    '''
    ''' A referência forte à coleção Items E à pasta pai é obrigatória (R7):
    ''' se o GC as coletar, o sink morre e os eventos param sem erro nenhum.
    ''' </summary>
    Public NotInheritable Class FolderSubscription
        Implements IDisposable

        Private Shared _nextId As Integer
        Private Shared _nextSequence As Long

        Private _folder As Outlook.MAPIFolder
        Private _items As Outlook.Items
        Private _onAdd As Outlook.ItemsEvents_ItemAddEventHandler
        Private _onChange As Outlook.ItemsEvents_ItemChangeEventHandler
        Private _onRemove As Outlook.ItemsEvents_ItemRemoveEventHandler

        Private _active As Integer = 1
        Private _pending As Integer

        Public ReadOnly Property Id As Integer
        Public ReadOnly Property FolderName As String

        ''' <summary>Callbacks postados e ainda não processados.</summary>
        Public ReadOnly Property Pending As Integer
            Get
                Return Volatile.Read(_pending)
            End Get
        End Property

        Public ReadOnly Property IsActive As Boolean
            Get
                Return Volatile.Read(_active) = 1
            End Get
        End Property

        ''' <summary>Precisa ser construída DENTRO da thread STA do broker.</summary>
        ''' <param name="folder">
        ''' A assinatura vira DONA da pasta e da coleção; quem chama não deve
        ''' liberar nenhuma das duas.
        ''' </param>
        ''' <param name="postToBroker">
        ''' Enfileira trabalho na thread do broker sem bloquear. Retorna
        ''' False se o trabalho não puder ser aceito.
        ''' </param>
        Public Sub New(folderName As String,
                       folder As Outlook.MAPIFolder,
                       sink As Action(Of EventRecord),
                       postToBroker As Func(Of Action, Boolean))
            _Id = Interlocked.Increment(_nextId)
            _FolderName = folderName
            _folder = folder
            _items = folder.Items

            _onAdd = Sub(item) Handle("ItemAdd", item, sink, postToBroker)
            _onChange = Sub(item) Handle("ItemChange", item, sink, postToBroker)
            _onRemove = Sub() Handle("ItemRemove", Nothing, sink, postToBroker)

            AddHandler _items.ItemAdd, _onAdd
            AddHandler _items.ItemChange, _onChange
            AddHandler _items.ItemRemove, _onRemove
        End Sub

        ''' <summary>
        ''' Roda na thread de ENTREGA (MTA). Não lê nada de COM.
        ''' </summary>
        Private Sub Handle(kind As String,
                           item As Object,
                           sink As Action(Of EventRecord),
                           postToBroker As Func(Of Action, Boolean))

            Dim envelope As New EventRecord With {
                .Kind = kind,
                .SubscriptionId = Id,
                .DeliverySequence = Interlocked.Increment(_nextSequence),
                .DeliveredAt = DateTime.UtcNow,
                .DeliveryThreadId = Environment.CurrentManagedThreadId,
                .DeliveryApartment = Thread.CurrentThread.GetApartmentState().ToString(),
                .Folder = FolderName,
                .EntryId = "",
                .Subject = If(kind = "ItemRemove", "(ItemRemove não entrega o item)", ""),
                .MessageClass = ""
            }

            If Not IsActive Then
                ComHelpers.Release(item)
                Return
            End If

            Interlocked.Increment(_pending)

            Dim accepted = postToBroker(
                Sub()
                    Try
                        ' Já é a STA do broker: ler COM é seguro aqui.
                        If Not IsActive Then Return
                        Fill(envelope, item)
                        sink(envelope)
                    Finally
                        ComHelpers.Release(item)
                        Interlocked.Decrement(_pending)
                    End Try
                End Sub)

            If Not accepted Then
                ' Broker encerrando: ninguém mais vai liberar este RCW.
                Interlocked.Decrement(_pending)
                ComHelpers.Release(item)
            End If
        End Sub

        ''' <summary>Lê o item. Roda na thread do broker.</summary>
        Private Shared Sub Fill(record As EventRecord, item As Object)
            record.ProcessedAt = DateTime.UtcNow
            record.ThreadId = Environment.CurrentManagedThreadId
            record.Apartment = Thread.CurrentThread.GetApartmentState().ToString()

            If item Is Nothing Then Return

            Dim mail = TryCast(item, Outlook.MailItem)
            If mail Is Nothing Then
                record.MessageClass = "(não é MailItem)"
                Return
            End If

            Try : record.EntryId = mail.EntryID : Catch : End Try
            Try : record.Subject = mail.Subject : Catch : End Try
            Try : record.MessageClass = mail.MessageClass : Catch : End Try
        End Sub

        ''' <summary>
        ''' Marca inativa ANTES de desconectar, para que trabalho já postado
        ''' seja descartado em vez de contado. Ordem obrigatória:
        ''' RemoveHandler antes de liberar os RCWs.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            If Interlocked.Exchange(_active, 0) = 0 Then Return
            If _items Is Nothing Then Return

            Try
                RemoveHandler _items.ItemAdd, _onAdd
                RemoveHandler _items.ItemChange, _onChange
                RemoveHandler _items.ItemRemove, _onRemove
            Catch
                ' Outlook já pode ter ido embora.
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

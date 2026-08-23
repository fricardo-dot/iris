Imports System.Collections.Generic
Imports System.Threading
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' Assinatura de eventos de uma pasta.
    '''
    ''' O que a Fase 0 mediu e que este arquivo respeita:
    '''
    '''   • Os callbacks NÃO chegam na thread STA que assinou. Chegam numa
    '''     thread MTA do pool. Ler propriedade COM ali é violar a afinidade
    '''     de thread do OOM — funciona por marshaling implícito e trava sob
    '''     carga.
    '''   • O item entregue é liberado IMEDIATAMENTE, sem ser lido. Evento
    '''     aqui é aviso de "pasta suja", não transição a aplicar: o estado
    '''     lido depois pode não ser o que causou o evento, e a ordem dos
    '''     eventos não é causalmente confiável (critério D3).
    '''   • Referência forte à coleção Items E à pasta pai. Se o GC as
    '''     coletar, o sink morre junto e os eventos param sem erro nenhum.
    '''   • RemoveHandler antes de liberar os RCWs, e nunca o contrário.
    ''' </summary>
    Friend NotInheritable Class FolderSubscription
        Implements IDisposable

        Private Shared _nextId As Integer
        Private Shared _nextSequence As Long

        Private _folder As OL.MAPIFolder
        Private _items As OL.Items
        Private _onAdd As OL.ItemsEvents_ItemAddEventHandler
        Private _onChange As OL.ItemsEvents_ItemChangeEventHandler
        Private _onRemove As OL.ItemsEvents_ItemRemoveEventHandler
        Private _active As Integer = 1

        Public ReadOnly Property Id As Integer
        Public ReadOnly Property Key As FolderKey

        ''' <summary>
        ''' Precisa ser construída DENTRO da thread STA do broker: o sink
        ''' pertence à thread que assina. A assinatura vira DONA da pasta e
        ''' da coleção; quem chama não deve liberar nenhuma das duas.
        ''' </summary>
        Public Sub New(key As FolderKey, folder As OL.MAPIFolder,
                       sink As Action(Of FolderInvalidation))
            _Id = Interlocked.Increment(_nextId)
            _Key = key
            _folder = folder
            _items = folder.Items

            _onAdd = Sub(item) Handle(InvalidationKind.ItemAdded, item, sink)
            _onChange = Sub(item) Handle(InvalidationKind.ItemChanged, item, sink)
            _onRemove = Sub() Handle(InvalidationKind.ItemRemoved, Nothing, sink)

            ' Construcao TRANSACIONAL. Se o segundo AddHandler falhasse, o
            ' primeiro ficaria conectado, os RCWs nao seriam liberados, o
            ' objeto nunca chegaria ao dicionario do broker — e, como o
            ' construtor lanca, ninguem chamaria Dispose. Sink pendurado
            ' para sempre.
            Dim conectados As New List(Of String)()
            Try
                AddHandler _items.ItemAdd, _onAdd
                conectados.Add("add")
                AddHandler _items.ItemChange, _onChange
                conectados.Add("change")
                AddHandler _items.ItemRemove, _onRemove
                conectados.Add("remove")
            Catch
                DesconectarParcial(conectados)
                ComHelpers.Release(_items)
                _items = Nothing
                ComHelpers.Release(_folder)
                _folder = Nothing
                Throw
            End Try
        End Sub

        Private Sub DesconectarParcial(conectados As List(Of String))
            For Each qual In conectados
                Try
                    Select Case qual
                        Case "add" : RemoveHandler _items.ItemAdd, _onAdd
                        Case "change" : RemoveHandler _items.ItemChange, _onChange
                        Case "remove" : RemoveHandler _items.ItemRemove, _onRemove
                    End Select
                Catch
                End Try
            Next
        End Sub

        Public ReadOnly Property IsActive As Boolean
            Get
                Return Volatile.Read(_active) = 1
            End Get
        End Property

        ''' <summary>
        ''' Roda na thread de ENTREGA (MTA). Não lê nada do item — só o
        ''' libera e sinaliza que a pasta mudou.
        '''
        ''' É isto que dispensa marshaling, fila e ciclo de vida de trabalho
        ''' postado: não há nada a processar depois.
        ''' </summary>
        Private Sub Handle(kind As InvalidationKind, item As Object,
                           sink As Action(Of FolderInvalidation))
            ComHelpers.Release(item)
            If Not IsActive Then Return

            Try
                sink(New FolderInvalidation With {
                    .Folder = Key,
                    .Kind = kind,
                    .SubscriptionId = Id,
                    .Sequence = Interlocked.Increment(_nextSequence),
                    .At = DateTimeOffset.UtcNow
                })
            Catch
                ' Um callback nunca pode derrubar o Outlook por cima de nós.
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Inativa ANTES de desconectar: um evento em voo deixa de ser
            ' contado como se fosse da assinatura nova.
            If Interlocked.Exchange(_active, 0) = 0 Then Return
            If _items Is Nothing Then Return

            ' Um Try por handler: com os tres juntos, falhar ao remover
            ' ItemAdd deixaria ItemChange e ItemRemove conectados sem sequer
            ' serem tentados.
            Try : RemoveHandler _items.ItemAdd, _onAdd : Catch : End Try
            Try : RemoveHandler _items.ItemChange, _onChange : Catch : End Try
            Try : RemoveHandler _items.ItemRemove, _onRemove : Catch : End Try

            _onAdd = Nothing
            _onChange = Nothing
            _onRemove = Nothing

            ComHelpers.Release(_items)
            _items = Nothing
            ComHelpers.Release(_folder)
            _folder = Nothing
        End Sub

    End Class

End Namespace

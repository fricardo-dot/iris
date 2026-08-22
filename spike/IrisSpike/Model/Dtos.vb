Namespace Model

    ''' <summary>
    ''' R9: o Outlook pode listar um item sem ter o conteúdo localmente. O
    ''' DTO carrega o estado explicitamente para a UI nunca bloquear
    ''' esperando download.
    ''' </summary>
    Public Enum ContentState
        MetadataOnly
        BodyAvailable
        AttachmentsAvailable
        TransientError
    End Enum

    ''' <summary>
    ''' Identidade de um item. EntryId + StoreId é o ponto de partida, e NÃO
    ''' sobrevive a movimentação entre stores (seção 5 do ESCOPO.md) — a
    ''' chave interna estável é entregável da Fase 2. O spike registra os
    ''' pares antes e depois de um move para embasar aquele desenho.
    ''' </summary>
    Public NotInheritable Class ItemKey
        Public Property EntryId As String
        Public Property StoreId As String

        Public Overrides Function ToString() As String
            Dim entry = If(EntryId, "")
            Dim store = If(StoreId, "")
            Return $"{Left(entry, 12)}…/{Left(store, 8)}…"
        End Function
    End Class

    Public NotInheritable Class StoreInfo
        Public Property DisplayName As String
        Public Property StoreId As String
        Public Property ExchangeStoreType As String
        Public Property FilePath As String
        Public Property IsCachedExchange As Boolean
    End Class

    Public NotInheritable Class FolderInfo
        Public Property Name As String
        Public Property Key As ItemKey
        Public Property DefaultItemType As String
        Public Property ItemCount As Integer
        Public Property UnreadCount As Integer
    End Class

    ''' <summary>
    ''' Resumo de mensagem. Deliberadamente só tipos primitivos: é este DTO
    ''' que atravessa a fronteira do broker, e ContainsComReference() é
    ''' rodado contra ele na Fase 0 para provar que nada de COM vaza junto.
    ''' </summary>
    Public NotInheritable Class MailSummary
        Public Property Key As ItemKey
        Public Property Subject As String
        Public Property SenderName As String
        Public Property SenderAddress As String
        Public Property ReceivedTime As DateTime?
        Public Property SizeBytes As Integer
        Public Property HasAttachments As Boolean
        Public Property IsUnread As Boolean
        Public Property Content As ContentState
        ''' <summary>
        ''' Nothing = o corpo nem foi tentado. Distinguir "DownloadState diz
        ''' que está completo" de "o corpo foi realmente lido" é o ponto do
        ''' R9: a primeira afirmação é uma promessa do Outlook, a segunda é
        ''' um fato.
        ''' </summary>
        Public Property BodyLength As Integer?
        ''' <summary>R11: se o item tem sensitivity label / IRM.</summary>
        Public Property IsProtected As Boolean
        ''' <summary>Classe MAPI, para detectar o que não é MailItem.</summary>
        Public Property MessageClass As String
    End Class

End Namespace

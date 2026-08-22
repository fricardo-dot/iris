Namespace Global.Iris.Model

    ''' <summary>
    ''' Estado do conteúdo de uma mensagem (R9).
    '''
    ''' Existe porque o Outlook pode listar um item sem ter o corpo ou os
    ''' anexos localmente, e a chamada pode bloquear enquanto busca. A UI
    ''' precisa saber a diferença em vez de esperar.
    '''
    ''' A Fase 0 mostrou que <c>DownloadState</c> é a PROMESSA do Outlook;
    ''' só a leitura confirma. Por isso <see cref="BodyAvailable"/> significa
    ''' "corpo lido", não "declarado completo".
    ''' </summary>
    Public Enum ContentState
        MetadataOnly
        BodyAvailable
        AttachmentsAvailable
        TransientError
    End Enum

    Public Enum BodyFormat
        Unknown
        PlainText
        Html
        RichText
    End Enum

    ''' <summary>Estado da sessão com o Outlook.</summary>
    Public Enum SessionState
        Disconnected
        Connecting
        Connected
        ''' <summary>Aberto, porém recusando chamadas (R13). Não é "fechado".</summary>
        Busy
        Reconnecting
        ''' <summary>Não está em execução, ou nem instalado.</summary>
        Unavailable
    End Enum

    ''' <summary>
    ''' Classificação de erro. Existe para a UI decidir o que fazer sem
    ''' inspecionar mensagem de exceção nem HRESULT.
    ''' </summary>
    Public Enum ErrorKind
        None
        ''' <summary>Outlook fechado ou não instalado.</summary>
        NotConnected
        ''' <summary>Ocupado agora; tentar de novo mais tarde faz sentido.</summary>
        Busy
        ''' <summary>O item sumiu ou foi movido por baixo (F1-H).</summary>
        NotFound
        ''' <summary>A página pedida é de uma geração vencida (F1-E).</summary>
        Stale
        ''' <summary>Bloqueado por política, IRM ou guarda do Object Model.</summary>
        Denied
        ''' <summary>Conteúdo ainda não baixado.</summary>
        NotDownloaded
        ''' <summary>Cancelado pelo chamador.</summary>
        Cancelled
        ''' <summary>Falhou e não se sabe se o efeito ocorreu. NUNCA repetir.</summary>
        Ambiguous
        ''' <summary>
        ''' Operação ainda não implementada neste marco. Existe para não
        ''' disfarçar ausência de código como erro de execução.
        ''' </summary>
        NotImplemented
        Unexpected
    End Enum

    ''' <summary>Por que uma pasta foi marcada como suja.</summary>
    Public Enum InvalidationKind
        ItemAdded
        ItemChanged
        ItemRemoved
        ''' <summary>Eventos perdidos ou coalescidos: reler tudo.</summary>
        Unknown
    End Enum

    ''' <summary>
    ''' Tipo de destinatario. Enum, e nao String: com "To"/"Cc" soltos, o
    ''' Core e a UI passariam a depender de literais que ninguem valida.
    ''' </summary>
    Public Enum RecipientKind
        [To]
        Cc
        Bcc
        Unknown
    End Enum

    Public Enum MessageSort
        ReceivedDesc
        ReceivedAsc
        SubjectAsc
        SenderAsc
    End Enum

End Namespace

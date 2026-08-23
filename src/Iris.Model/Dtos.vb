Imports System.Collections.Generic

Namespace Global.Iris.Model

    Public NotInheritable Class StoreInfo
        Public Property DisplayName As String = ""
        Public Property StoreId As String = ""
        Public Property ExchangeStoreType As String = ""
        Public Property IsCachedExchange As Boolean
        Public Property RootFolder As FolderKey
    End Class

    Public NotInheritable Class FolderInfo
        Public Property Key As FolderKey
        Public Property Name As String = ""
        Public Property ContentKind As FolderContentKind = FolderContentKind.Unknown
        Public Property ItemCount As Integer
        ''' <summary>
        ''' Eventualmente consistente por desenho: o Outlook atualiza isto
        ''' de forma assíncrona, e o Iris não vai bloquear para conferir.
        ''' </summary>
        Public Property UnreadCount As Integer
        Public Property HasChildren As Boolean

        ''' <summary>
        ''' PR_ATTR_HIDDEN. E a propriedade que o proprio Outlook usa para
        ''' nao mostrar pastas internas como "Conversation Action Settings"
        ''' e "Quick Step Settings".
        '''
        ''' O broker REPORTA; quem decide esconder e a camada de
        ''' apresentacao. Filtrar no broker enterraria uma politica de
        ''' interface na camada de dados.
        ''' </summary>
        Public Property IsHidden As Boolean
    End Class

    Public NotInheritable Class RecipientInfo
        Public Property DisplayName As String = ""
        Public Property Address As String = ""
        ''' <summary>To, Cc ou Bcc.</summary>
        Public Property Kind As RecipientKind = RecipientKind.Unknown
        Public Property Resolved As Boolean
    End Class

    Public NotInheritable Class AttachmentInfo
        Public Property Key As AttachmentKey
        Public Property FileName As String = ""
        Public Property SizeBytes As Integer
        Public Property AttachmentType As String = ""
        ''' <summary>Anexo embutido referenciado pelo corpo HTML por CID.</summary>
        Public Property IsInline As Boolean
        Public Property ContentId As String = ""
    End Class

    ''' <summary>Resumo para a lista. Sem corpo — ver F1-F.</summary>
    Public NotInheritable Class MailSummary
        Public Property Key As ItemKey
        Public Property Subject As String = ""
        Public Property SenderName As String = ""
        ''' <summary>
        ''' DateTimeOffset, nao DateTime: antes de a Fase 2 persistir isto em
        ''' SQLite, o fuso precisa estar no dado. "DateTime sem Kind" e a
        ''' origem classica de mensagem aparecendo com hora errada depois de
        ''' ler do cache.
        ''' </summary>
        Public Property ReceivedTime As DateTimeOffset?
        Public Property SizeBytes As Integer
        Public Property HasAttachments As Boolean
        Public Property IsUnread As Boolean
        Public Property IsProtected As Boolean
        Public Property MessageClass As String = ""
        Public Property Content As ContentState
    End Class

    ''' <summary>
    ''' Cabeçalho, corpo e anexos de uma vez.
    '''
    ''' Juntos de propósito: obtidos em três chamadas separadas, poderiam
    ''' observar estados diferentes de uma mensagem que mudou no meio.
    ''' </summary>
    Public NotInheritable Class MessageDetail
        Public Property Key As ItemKey
        Public Property Subject As String = ""
        Public Property SenderName As String = ""
        Public Property SenderAddress As String = ""
        Public Property ReceivedTime As DateTimeOffset?
        Public Property Recipients As New List(Of RecipientInfo)()
        Public Property Attachments As New List(Of AttachmentInfo)()

        ''' <summary>
        ''' Quão completa veio cada parte. Sem isto, lista vazia era
        ''' ambígua: "não tem" e "não deu para ler" apareciam iguais.
        ''' </summary>
        Public Property RecipientsStatus As PartStatus = PartStatus.Full
        Public Property AttachmentsStatus As PartStatus = PartStatus.Full

        Public Property Content As ContentState
        Public Property Format As BodyFormat
        Public Property HtmlBody As String = ""
        Public Property TextBody As String = ""

        ''' <summary>
        ''' IRM ou rótulo de sensibilidade. Ver R11: item protegido fica fora
        ''' do escopo da IA, e não entra em log.
        ''' </summary>
        Public Property IsProtected As Boolean
        Public Property BodyError As ErrorKind
    End Class

    ''' <summary>
    ''' Consulta de página. A <see cref="Generation"/> é o que impede uma
    ''' resposta lenta da pasta anterior de sobrescrever a tela depois que o
    ''' usuário já trocou de seleção (F1-E).
    ''' </summary>
    Public NotInheritable Class MessageQuery
        Public ReadOnly Property Folder As FolderKey
        Public ReadOnly Property Sort As MessageSort
        Public ReadOnly Property Generation As Long

        Public Sub New(folder As FolderKey, sort As MessageSort, generation As Long)
            Me.Folder = folder
            Me.Sort = sort
            Me.Generation = generation
        End Sub
    End Class

    ''' <summary>
    ''' Uma página de resumos.
    '''
    ''' PAGINAÇÃO VOLÁTIL por desenho (FASE1.md seção 5): offset numa pasta
    ''' viva não é estável — uma mensagem que chega no topo entre duas
    ''' páginas duplica um item e pula outro. Por isso a página devolve a
    ''' geração com que foi lida, e página de geração vencida é DESCARTADA,
    ''' nunca anexada.
    ''' </summary>
    Public NotInheritable Class MessagePage
        Public Property Generation As Long
        Public Property Offset As Integer
        Public Property Items As New List(Of MailSummary)()
        ''' <summary>Total no momento da leitura. Pode já estar desatualizado.</summary>
        Public Property TotalAtRead As Integer
        Public Property HasMore As Boolean

        ''' <summary>
        ''' Posição na coleção BRUTA onde a próxima página deve começar.
        '''
        ''' Existe porque contar DTOs devolvidos estava errado: o broker
        ''' examina N posições e pode devolver menos, já que uma coleção
        ''' Items não contém apenas MailItem e itens corrompidos são
        ''' pulados. Usar Items.Count como próximo offset relia as posições
        ''' puladas, duplicando linhas — e, no limite, travando o avanço.
        ''' </summary>
        Public Property NextOffset As Integer

        ''' <summary>Quantas posições da coleção foram examinadas.</summary>
        Public Property ExaminedCount As Integer

        ''' <summary>
        ''' Quantas foram puladas por não serem mensagem ou por erro. Fica
        ''' visível para "28 de 30" não virar mistério.
        ''' </summary>
        Public Property SkippedCount As Integer
    End Class

    ''' <summary>
    ''' Aviso de que uma pasta mudou. NÃO carrega dados do item de propósito:
    ''' o callback chega numa thread MTA e o estado lido depois pode não ser
    ''' o que causou o evento, então a única resposta correta é reler.
    ''' </summary>
    Public NotInheritable Class FolderInvalidation
        Public Property Folder As FolderKey
        Public Property Kind As InvalidationKind
        Public Property SubscriptionId As Integer
        Public Property Sequence As Long
        Public Property At As DateTimeOffset
    End Class

    ''' <summary>
    ''' O que o usuário edita num rascunho.
    '''
    ''' <see cref="UserText"/> é SÓ o que ele digitou. A citação e a
    ''' assinatura que o Outlook gerou não passam por aqui — elas ficam
    ''' intactas do outro lado, em <see cref="DraftInfo.QuotedBody"/>.
    ''' Trafegar o corpo inteiro como texto editável destruiria a formatação
    ''' corporativa a cada salvamento.
    ''' </summary>
    Public NotInheritable Class DraftContent
        Public Property Subject As String = ""
        Public Property UserText As String = ""
        ''' <summary>Endereços separados por ponto e vírgula, como digitados.</summary>
        Public Property ToLine As String = ""
        Public Property CcLine As String = ""
    End Class

    ''' <summary>
    ''' Um rascunho aberto no compositor.
    '''
    ''' O rascunho é criado e salvo no Outlook ao ABRIR o compositor, não na
    ''' primeira edição: assim ele sobrevive a um fechamento acidental e tem
    ''' chave estável desde o começo.
    ''' </summary>
    Public NotInheritable Class DraftInfo
        Public Property Key As DraftKey
        Public Property Subject As String = ""
        Public Property ToLine As String = ""
        Public Property CcLine As String = ""

        ''' <summary>O que o usuário digitou. Vazio num rascunho recém-criado.</summary>
        Public Property UserText As String = ""

        ''' <summary>
        ''' O corpo que o Outlook gerou — citação da mensagem original e
        ''' assinatura. Preservado INTACTO: o Iris nunca o reescreve, só
        ''' escreve acima dele.
        ''' </summary>
        Public Property QuotedBody As String = ""

        ''' <summary>
        ''' A mesma citação em TEXTO, só para exibir.
        ''' <see cref="QuotedBody"/> é o que volta para o Outlook e por isso
        ''' fica intacto — mas mostrar marcação HTML crua na tela seria pior
        ''' que não mostrar nada.
        ''' </summary>
        Public Property QuotedPreview As String = ""

        Public Property Format As BodyFormat = BodyFormat.PlainText

        Public Property Attachments As New List(Of AttachmentInfo)()

        ''' <summary>
        ''' Conta que o Outlook vai usar para enviar. Aparece na confirmação
        ''' porque delegação e múltiplas contas tornam isso não óbvio (F1-L).
        ''' </summary>
        Public Property SendingAccount As String = ""
    End Class

    ''' <summary>
    ''' O que a confirmação de envio mostra ao usuário. A conta remetente
    ''' entra porque o Outlook pode escolher conta ou delegação de forma não
    ''' óbvia, e enviar pela errada em silêncio é inaceitável (F1-L).
    ''' </summary>
    Public NotInheritable Class SendPreview
        Public Property Draft As DraftKey
        Public Property SendingAccount As String = ""
        Public Property Subject As String = ""
        Public Property Recipients As New List(Of RecipientInfo)()

        ''' <summary>
        ''' O que vai junto. A confirmacao precisa mostrar: mandar o anexo
        ''' errado para fora e tao irreversivel quanto mandar para a pessoa
        ''' errada.
        ''' </summary>
        Public Property Attachments As New List(Of AttachmentInfo)()

        ''' <summary>
        ''' A lista de destinatários veio inteira?
        '''
        ''' É a pergunta mais importante desta tela. Uma lista INCOMPLETA é
        ''' pior que uma vazia: o usuário confere três endereços certos,
        ''' aprova, e a mensagem vai para menos gente do que devia — e o que
        ''' falta é invisível por definição.
        ''' </summary>
        Public Property RecipientsStatus As PartStatus = PartStatus.Full
        Public Property AttachmentsStatus As PartStatus = PartStatus.Full

        ''' <summary>
        ''' Todo destinatario tem endereco SMTP reconhecivel.
        ''' Nao basta o Outlook dizer que resolveu: resolver um nome interno
        ''' para <c>/O=...</c> continua sendo um endereco que o usuario nao
        ''' tem como conferir — e conferir e a unica funcao desta tela.
        ''' </summary>
        Public ReadOnly Property AllResolved As Boolean
            Get
                For Each r In Recipients
                    If Not r.Resolved Then Return False
                Next
                Return True
            End Get
        End Property
    End Class

    ''' <summary>Token lógico de assinatura. Nunca embrulha COM.</summary>
    Public NotInheritable Class SubscriptionToken
        Public ReadOnly Property Id As Integer
        Public ReadOnly Property Folder As FolderKey

        Public Sub New(id As Integer, folder As FolderKey)
            Me.Id = id
            Me.Folder = folder
        End Sub
    End Class

End Namespace

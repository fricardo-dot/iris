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

        ''' <summary>
        ''' NAO existe IsProtected aqui, e a ausencia e deliberada.
        '''
        ''' A Q1 mediu que Permission nao vem por coluna de Table, e o
        ''' caminho rapido de listagem le por Table. Preencher com False
        ''' seria afirmar "nao e protegida" sem ter medido, e um
        ''' consumidor futuro herdaria a mentira. O gate do R11 vive em
        ''' MessageReading, que le o MailItem de verdade.
        ''' </summary>
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
        ''' <summary>
        ''' Sem valor inicial DE PROPÓSITO.
        '''
        ''' Um default "completo" faz qualquer produtor que esqueça de
        ''' preencher declarar completude que não provou — e completude não
        ''' provada é exatamente o que este tipo existe para impedir. Nothing
        ''' é tratado como não confiável por quem consome.
        ''' </summary>
        Public Property RecipientsStatus As PartStatus
        Public Property AttachmentsStatus As PartStatus

        Public Property Content As ContentState
        Public Property Format As BodyFormat
        Public Property HtmlBody As String = ""
        Public Property TextBody As String = ""

        ''' <summary>
        ''' IRM ou rótulo de sensibilidade. Ver R11: item protegido fica fora
        ''' do escopo da IA, e não entra em log.
        ''' </summary>
        ''' <summary>
        ''' Unknown por ser o PRIMEIRO valor do enum, e isso e deliberado: um
        ''' produtor que esqueca de preencher fecha o gate em vez de abrir.
        ''' </summary>
        Public Property Protection As ProtectionState
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
        Public Property Items As New List(Of MailSummary)()

        ''' <summary>
        ''' Continuacao opaca. Nothing significa FIM.
        '''
        ''' Nao existe um HasMore armazenado ao lado: dois campos que
        ''' podem se contradizer sempre acabam se contradizindo.
        ''' </summary>
        Public Property NextCursor As String

        ''' <summary>
        ''' Total no momento em que a travessia comecou. So vem na
        ''' PRIMEIRA pagina.
        '''
        ''' Antes era TotalAtRead, vinha em toda pagina e saia de
        ''' items.Count. O caminho por Table nao tem Count barato, e
        ''' pedir a contagem a cada pagina gastaria uma chamada COM na
        ''' fila unica da STA para reafirmar um numero que ja estava
        ''' desatualizado quando chegou.
        ''' </summary>
        Public Property TotalAtStart As Integer?

        ''' <summary>
        ''' Quantas foram descartadas por nao serem mensagem ou por erro.
        ''' Fica visivel para "28 de 30" nao virar misterio.
        ''' </summary>
        Public Property SkippedCount As Integer

        ''' <summary>
        ''' Quantas linhas vieram da DRENAGEM do grupo empatado, alem do
        ''' alvo pedido.
        '''
        ''' A pagina drenada NAO tem teto: ela vai ate o fim do grupo do
        ''' ultimo instante, senao os empatados que ficaram para tras
        ''' seriam pulados para sempre. Entao pedir 30 pode devolver 45,
        ''' e isso precisa ser explicavel em vez de parecer defeito.
        ''' Medido nesta caixa: o maior empate num mesmo segundo tem 16.
        ''' </summary>
        Public Property DrainedExtra As Integer

        Public ReadOnly Property HasMore As Boolean
            Get
                Return Not String.IsNullOrEmpty(NextCursor)
            End Get
        End Property
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
        ''' A lista de anexos veio inteira? Sem valor inicial de propósito —
        ''' ver o comentário em <see cref="MessageDetail"/>.
        ''' </summary>
        Public Property AttachmentsStatus As PartStatus

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
        Public Property RecipientsStatus As PartStatus
        Public Property AttachmentsStatus As PartStatus

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


    ''' <summary>
    ''' <b>Um compromisso do calendário, já achatado para a UI.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>MEDIDO ANTES DE EXISTIR</b>
    '''
    ''' Os campos são os que <c>tools/medir-grupos.ps1</c> mediu legíveis em
    ''' <b>100 de 100</b> compromissos da caixa real, em 28/08/2026 — e nenhum
    ''' a mais. Um DTO com campo que ninguém mediu é uma promessa que o
    ''' provedor não fez.
    '''
    ''' O custo medido foi <b>30,9 ms por item</b>, quase o dobro dos ~16 ms
    ''' que a Fase 0 mediu por mensagem. Os 434 compromissos desta caixa
    ''' custariam ~13 s numa leitura só, e é por isso que a leitura é sempre
    ''' por <b>janela de datas</b>, nunca "o calendário inteiro".
    ''' </summary>
    Public NotInheritable Class AppointmentInfo
        Public Property Key As ItemKey
        Public Property Subject As String = ""
        Public Property Location As String = ""

        ''' <summary>
        ''' Início e fim <b>com offset</b>.
        '''
        ''' <c>DateTimeOffset</c> e não <c>DateTime</c>, pela mesma razão que
        ''' <c>MailSummary.ReceivedTime</c>: "DateTime sem Kind" é a origem
        ''' clássica de compromisso aparecendo na hora errada. Num calendário
        ''' isso é pior que num e-mail — a hora <b>é</b> o dado.
        ''' </summary>
        Public Property Start As DateTimeOffset
        Public Property [End] As DateTimeOffset

        Public Property AllDayEvent As Boolean
        Public Property Organizer As String = ""
        Public Property BusyStatus As AppointmentBusy
        Public Property ResponseStatus As AppointmentResponse

        ''' <summary>
        ''' Este item é uma ocorrência de uma série.
        '''
        ''' Medido: <b>4 em 100</b> dos compromissos mais recentes desta caixa.
        ''' O número é pequeno e a consequência não: uma leitura que não expanda
        ''' a série mostra a primeira ocorrência e esconde as outras, e o
        ''' usuário não tem como perceber a diferença olhando a tela.
        ''' </summary>
        Public Property IsRecurring As Boolean

        Public Property RecipientCount As Integer
    End Class

    ''' <summary>Como o compromisso ocupa a agenda.</summary>
    Public Enum AppointmentBusy
        Livre
        Provisorio
        Ocupado
        ForaDoEscritorio
        TrabalhandoEmOutroLugar
        Desconhecido
    End Enum

    ''' <summary>
    ''' A resposta a um convite.
    '''
    ''' <c>NaoEhReuniao</c> é distinto de <c>NaoRespondeu</c>: um compromisso
    ''' que o próprio usuário criou não tem resposta pendente, e mostrá-lo como
    ''' "não respondeu" seria inventar uma pendência.
    ''' </summary>
    Public Enum AppointmentResponse
        NaoEhReuniao
        Organizador
        Aceito
        Provisorio
        Recusado
        NaoRespondeu
        Desconhecido
    End Enum

    ''' <summary>
    ''' A janela de datas pedida, e o que a leitura fez com ela.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A JANELA VOLTA JUNTO, E ISSO NÃO É REDUNDÂNCIA</b>
    '''
    ''' Quem pede "os próximos sete dias" e recebe uma lista não sabe se a
    ''' lista está vazia porque não há compromisso ou porque a leitura falhou
    ''' no meio. A janela efetivamente lida vem no mesmo objeto, pelo mesmo
    ''' motivo que a cobertura acompanha o manifesto do acervo: é a §23 aplicada
    ''' a outro grupo.
    ''' </summary>
    Public NotInheritable Class AppointmentWindow
        Public Property De As DateTimeOffset
        Public Property Ate As DateTimeOffset
        Public Property Items As New List(Of AppointmentInfo)()

        ''' <summary>
        ''' Quantas ocorrências vieram da expansão de séries.
        '''
        ''' Fica visível pela mesma razão que <c>MessagePage.SkippedCount</c>:
        ''' "12 compromissos, 5 deles de séries" é diferente de "12
        ''' compromissos", e a diferença muda o que o usuário conclui de uma
        ''' agenda cheia.
        ''' </summary>
        Public Property FromRecurrence As Integer

        ''' <summary>
        ''' Quantos itens a leitura viu e recusou por não serem compromisso ou
        ''' por não dar para ler. <c>Nothing</c> não é zero — zero afirma que
        ''' nada foi recusado.
        ''' </summary>
        Public Property Skipped As Integer?
    End Class

End Namespace

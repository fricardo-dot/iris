Imports System.Collections.Generic
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.Model

Namespace Global.Iris.Core

    ''' <summary>
    ''' Contrato do broker do Outlook.
    '''
    ''' Mora em Iris.Core, e não em Iris.Outlook, para que o núcleo dependa
    ''' do contrato e a implementação dependa do núcleo. A composição dos
    ''' dois acontece só no startup do aplicativo.
    '''
    ''' REGRAS QUE VÊM DA FASE 0 e valem para toda implementação:
    '''
    '''   • Nenhum membro aceita ou devolve tipo do Interop. O spike provou
    '''     que uma assinatura genérica como ReadAsync(Of T) deixa um RCW
    '''     escapar sem o compilador reclamar.
    '''   • Toda mutação roda com o retry do message filter DESLIGADO.
    '''     Criar, Save, Move, Delete e Send não são idempotentes.
    '''   • Cancelar não interrompe uma chamada COM já iniciada. O token
    '''     evita começar trabalho que ainda está na fila e libera o
    '''     chamador de esperar — a operação pode seguir no broker.
    '''   • Falha depois de iniciar uma mutação é <see cref="ErrorKind.Ambiguous"/>,
    '''     nunca um erro comum. Repetir automaticamente é proibido.
    ''' </summary>
    Public Interface IOutlookBroker

        ' ---- Sessão -----------------------------------------------------

        ReadOnly Property State As SessionState

        ''' <summary>Disparado quando o estado da sessão muda.</summary>
        Event StateChanged As EventHandler(Of SessionState)

        ''' <summary>
        ''' Anexa a um Outlook JÁ EM EXECUÇÃO. Nunca inicia o aplicativo:
        ''' uma instância criada por automação não tem perfil interativo.
        ''' "Outlook fechado" é estado previsto, não exceção.
        ''' </summary>
        Function ConnectAsync(cancel As CancellationToken) As Task(Of SessionState)

        ''' <summary>
        ''' Verifica se a sessão ainda responde e reconecta se o Outlook foi
        ''' fechado por baixo. Distingue "morreu" de "está ocupado agora" —
        ''' tratar toda COMException como morte derruba a conexão por uma
        ''' recusa transitória.
        ''' </summary>
        Function ProbeAsync(cancel As CancellationToken) As Task(Of SessionState)

        ' ---- Leitura ----------------------------------------------------

        Function GetStoresAsync(cancel As CancellationToken) _
            As Task(Of OperationResult(Of IReadOnlyList(Of StoreInfo)))

        ''' <summary>
        ''' Filhos de uma pasta, sob demanda. Não existe "carregar a árvore
        ''' inteira": uma caixa com muitas pastas travaria a abertura.
        ''' </summary>
        Function GetFolderChildrenAsync(parent As FolderKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of IReadOnlyList(Of FolderInfo)))

        ''' <summary>
        ''' Uma página de resumos. Restrict e Sort são feitos pelo Outlook,
        ''' nunca em laço nosso — medido na Fase 0: ~16 ms por item, e 770
        ''' itens levaram 12,8 s.
        '''
        ''' A página devolve a geração com que foi lida. Página de geração
        ''' vencida deve ser descartada pelo chamador, nunca anexada.
        ''' </summary>
        Function GetMessagePageAsync(query As MessageQuery, offset As Integer, count As Integer,
                                     cancel As CancellationToken) _
            As Task(Of OperationResult(Of MessagePage))

        ''' <summary>
        ''' Cabeçalho, corpo e anexos numa chamada só: em três chamadas
        ''' separadas, cada uma poderia observar um estado diferente de uma
        ''' mensagem que mudou no meio.
        ''' </summary>
        Function GetMessageDetailAsync(item As ItemKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of MessageDetail))

        ''' <summary>
        ''' Salva um anexo num diretório controlado. NÃO abre o arquivo:
        ''' abrir anexo é executar conteúdo não confiável, e a decisão é da
        ''' UI, com confirmação (F1-J).
        ''' </summary>
        Function SaveAttachmentAsync(attachment As AttachmentKey, destinationPath As String,
                                     overwrite As Boolean, cancel As CancellationToken) _
            As Task(Of OperationResult(Of String))

        ' ---- Mutações ---------------------------------------------------

        ''' <summary>
        ''' Só marca se realmente estiver não lida. Marcar o que já está
        ''' lido gera ItemChange à toa, que invalida a pasta, que recarrega,
        ''' que muda a seleção — o laço do F1-G.
        ''' </summary>
        Function MarkReadAsync(item As ItemKey, isRead As Boolean, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean))

        ' ---- Rascunhos --------------------------------------------------
        ' Criar e enviar são separados de propósito: juntos, impediriam
        ' editar, reabrir, anexar e tratar envio ambíguo sem criar outra
        ' mensagem.

        ''' <summary>
        ''' Mensagem nova. Lembrar: item não enviado nasce em Rascunhos,
        ''' independentemente da pasta — descoberta da Fase 0.
        ''' </summary>
        Function CreateDraftAsync(content As DraftContent, cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo))

        ''' <summary>
        ''' Usa Reply/ReplyAll do próprio Outlook. Responder NÃO é "mensagem
        ''' nova com texto citado": destinatários, assunto e citação têm
        ''' semântica própria que o OOM já implementa.
        ''' </summary>
        Function CreateReplyDraftAsync(item As ItemKey, replyAll As Boolean,
                                       cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo))

        ''' <summary>Usa Forward do próprio Outlook, preservando anexos.</summary>
        Function CreateForwardDraftAsync(item As ItemKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo))

        Function UpdateDraftAsync(draft As DraftKey, content As DraftContent,
                                  cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo))

        Function AddDraftAttachmentAsync(draft As DraftKey, filePath As String,
                                         cancel As CancellationToken) _
            As Task(Of OperationResult(Of AttachmentInfo))

        Function RemoveDraftAttachmentAsync(draft As DraftKey, attachment As AttachmentKey,
                                            cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean))

        ''' <summary>
        ''' O que a confirmação mostra antes de enviar: conta remetente e
        ''' destinatários resolvidos. Destinatário não resolvido bloqueia.
        ''' </summary>
        Function PrepareSendAsync(draft As DraftKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of SendPreview))

        ''' <summary>
        ''' Envia um rascunho já persistido. Send() é chamado UMA vez, com
        ''' retry desligado, e o item não é tocado depois.
        '''
        ''' Retornar sem erro NÃO prova entrega — pode ter apenas enfileirado
        ''' na Caixa de Saída. Falha aqui é <see cref="ErrorKind.Ambiguous"/>:
        ''' a mensagem pode ter saído, e a resposta certa é procurar, não
        ''' reenviar.
        ''' </summary>
        Function SendDraftAsync(draft As DraftKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean))

        Function DeleteDraftAsync(draft As DraftKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean))

        ' ---- Eventos ----------------------------------------------------

        ''' <summary>
        ''' Assina uma pasta. Devolve token lógico — nunca um IDisposable
        ''' embrulhando COM, que poderia escapar da fronteira.
        ''' </summary>
        Function SubscribeFolderAsync(folder As FolderKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of SubscriptionToken))

        Function UnsubscribeFolderAsync(token As SubscriptionToken, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean))

        ''' <summary>
        ''' Uma pasta mudou. É AVISO, não transição: o callback chega numa
        ''' thread MTA e o estado lido depois pode não ser o que causou o
        ''' evento. A resposta correta é reler a página atual.
        ''' </summary>
        Event FolderInvalidated As EventHandler(Of FolderInvalidation)

    End Interface

End Namespace

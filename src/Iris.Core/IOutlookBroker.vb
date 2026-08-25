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

        ''' <summary>
        ''' Identidade da sessão COM atual. Sobe a cada aquisição.
        '''
        ''' Existe porque <see cref="SessionState"/> não consegue dizer
        ''' "conectado, mas é OUTRA sessão" — isso é mudança de identidade,
        ''' não de estado, e tratar como estado produziu uma falha silenciosa
        ''' e permanente: o Outlook morria e voltava dentro da janela do
        ''' watchdog, o broker largava todas as assinaturas e readquiria, o
        ''' estado continuava Connected dos dois lados, nenhum evento era
        ''' emitido, e a lista parava de atualizar até o Iris ser reiniciado.
        '''
        ''' Toda chave — FolderKey, ItemKey, DraftKey, SubscriptionToken —
        ''' pertence a uma época. Chave de época vencida não vale.
        ''' </summary>
        ReadOnly Property SessionEpoch As Long

        ''' <summary>
        ''' A sessão COM foi SUBSTITUÍDA. O argumento é a época nova.
        '''
        ''' Separado de <see cref="StateChanged"/> de propósito: pode disparar
        ''' sem que o estado mude, e é justamente esse o caso que importa.
        '''
        ''' Chega em thread do broker, como todo evento daqui. O assinante
        ''' devolve ao dispatcher dele antes de tocar em qualquer coisa.
        ''' </summary>
        Event SessionReplaced As EventHandler(Of Long)

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
        ''' <param name="continuation">
        ''' Cursor opaco da pagina anterior. Nothing pede a primeira.
        ''' </param>
        ''' <param name="targetCount">
        ''' ALVO, nao teto. A pagina drena o grupo do ultimo instante ate
        ''' o fim, entao pode devolver mais — e e isso que impede pular
        ''' empatado. Ver MessagePage.DrainedExtra.
        ''' </param>
        Function GetMessagePageAsync(query As MessageQuery, continuation As String,
                                     targetCount As Integer,
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
        ''' <summary>
        ''' O rótulo de sensibilidade (Purview) de cada item pedido.
        '''
        ''' <b>Leitura pura</b>, e por isso vai por <c>ReadAsync</c>. Nunca
        ''' escreve propriedade — nem para "testar round-trip".
        '''
        ''' Um item que falha não derruba os outros: cada
        ''' <see cref="LabelReading"/> carrega o próprio desfecho, com a etapa
        ''' e o HRESULT. Um lote que falhasse inteiro por causa de um item
        ''' esconderia qual era.
        '''
        ''' Esta operação <b>não autoriza nada</b>. Ela informa o estado da
        ''' leitura; a decisão de divulgação é da política da Fase 3, e nela
        ''' "não consegui ler" nunca vira "pode".
        ''' </summary>
        ''' <summary>
        ''' Captura a mensagem inteira <b>numa leitura só</b> — assunto,
        ''' remetente, destinatários, corpo e <c>PR_CHANGE_KEY</c>.
        '''
        ''' Cinco chamadas separadas podem observar cinco estados diferentes de
        ''' uma mensagem que mudou no meio, e a <c>ChangeKey</c> serve justamente
        ''' para prender o corpo à versão que o portão classificou. Vinda de
        ''' outra passada, não prende nada.
        '''
        ''' Não torna a leitura atômica — o OOM não oferece isso, e a §29.2 do
        ''' FASE3.md é a resposta a essa falta. O que se ganha é a janela mais
        ''' estreita possível.
        ''' </summary>
        Function GetMessageSnapshotAsync(item As ItemKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of MessageSnapshot))

        Function GetSensitivityLabelsAsync(items As IReadOnlyList(Of ItemKey),
                                           cancel As CancellationToken) _
            As Task(Of OperationResult(Of IReadOnlyList(Of LabelReading)))

        ''' <summary>
        ''' O rótulo é projetável como coluna de <c>Table</c> nesta pasta?
        '''
        ''' Existe porque classificar por item custa uma ida ao COM por
        ''' mensagem — a mesma ordem de grandeza que tornou o cache obrigatório
        ''' na Fase 0 —, e porque a <c>Table</c> é o caminho <b>menos
        ''' invasivo</b> de descobrir isso: projeta coluna sem materializar
        ''' item e sem tocar em corpo.
        '''
        ''' Aceitar a coluna não é entregar o valor. Ver o
        ''' <see cref="LabelColumnProbe"/>.
        ''' </summary>
        ''' <summary>
        ''' O controle negativo da leitura de rótulo: como esta conta responde
        ''' a uma named property que <b>não existe</b>.
        '''
        ''' Existe porque a primeira medição devolveu "vazio" para todos os
        ''' itens, e vazio pode ser "sem rótulo" ou "esta propriedade nunca
        ''' falha". Sem o controle, um portão trataria ruído como decisão.
        '''
        ''' O conjunto de DASLs é <b>fixo no adaptador</b> — nada vem do
        ''' chamador. Named property tem mapeamento próprio no store.
        ''' </summary>
        Function ProbeLabelSemanticsAsync(item As ItemKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of NamedPropertyProbe))

        Function ProbeLabelColumnAsync(folder As FolderKey, quantas As Integer,
                                       cancel As CancellationToken) _
            As Task(Of OperationResult(Of LabelColumnProbe))

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

        ''' <summary>
        ''' Anexa e devolve o rascunho redescrito, nao so o anexo: anexar
        ''' SALVA, e todo Save pode mudar o EntryID. Devolver so o anexo
        ''' deixaria o chamador com a chave velha.
        ''' </summary>
        Function AddDraftAttachmentAsync(draft As DraftKey, filePath As String,
                                         cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo))

        ''' <summary>
        ''' Remove e devolve o rascunho redescrito. Mesmo motivo do anexar:
        ''' remover SALVA, o EntryID pode mudar, e todas as AttachmentKey são
        ''' reconstruídas porque o índice dos anexos seguintes muda ao tirar
        ''' um do meio.
        ''' </summary>
        Function RemoveDraftAttachmentAsync(draft As DraftKey, attachment As AttachmentKey,
                                            cancel As CancellationToken) _
            As Task(Of OperationResult(Of DraftInfo))

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

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
    ''' <summary>
    ''' <b>A porta que a agenda precisa — e só ela.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE UMA INTERFACE DE UM MÉTODO SÓ</b>
    '''
    ''' O <c>AgendaViewModel</c> recebia o <see cref="IOutlookBroker"/> inteiro,
    ''' e com isso o único jeito de testá-lo seria implementar as trinta e
    ''' tantas operações do broker num duplo — ou afrouxar o <c>FakeBroker</c>,
    ''' que responde "fora da alçada" de propósito para que uma chamada indevida
    ''' quebre o teste em vez de passar por sorte.
    '''
    ''' O resultado era previsível e a revisão externa de 28/08/2026 o achou:
    ''' <b>a agenda não tinha teste nenhum</b>, e por isso dois furos de geração
    ''' sobreviveram — o <c>Catch</c> e o <c>Finally</c> não conferiam de quem
    ''' era o voo.
    '''
    ''' Porta estreita não é elegância: é o que decide se existe teste.
    '''
    ''' <see cref="IOutlookBroker"/> herda desta, então o broker real e os
    ''' duplos existentes continuam servindo sem mudar de forma.
    ''' </summary>
    Public Interface IAgendaSource
        Function GetAppointmentsAsync(folder As FolderKey,
                                      de As DateTimeOffset, ate As DateTimeOffset,
                                      cancel As CancellationToken) _
            As Task(Of OperationResult(Of AppointmentWindow))
    End Interface

    ''' <summary>
    ''' <b>A porta estreita de ESCRITA no calendário.</b>
    '''
    ''' Separada da <see cref="IAgendaSource"/> pelo mesmo motivo que ela
    ''' existe: quem só lê não deve ser obrigado a implementar escrita para
    ''' ter um duplo. E aqui a separação vale mais, porque do outro lado
    ''' desta interface há uma operação que <b>apaga</b>.
    ''' </summary>
    Public Interface IAgendaWriter
        ''' <summary>
        ''' Cria um compromisso <b>na pasta indicada</b>.
        '''
        ''' Não há participantes em <see cref="AppointmentDraft"/>, e a
        ''' ausência é a funcionalidade: compromisso com participante é
        ''' reunião, e salvar reunião manda convite por e-mail.
        ''' </summary>
        Function CreateAppointmentAsync(folder As FolderKey, rascunho As AppointmentDraft,
                                        cancel As CancellationToken) _
            As Task(Of OperationResult(Of AppointmentInfo))

        ''' <summary>
        ''' Edita um compromisso. <b>Recusa reunião</b>, porque o
        ''' <c>Save</c> dela manda atualização a quem foi convidado.
        ''' </summary>
        Function UpdateAppointmentAsync(chave As AppointmentKey, rascunho As AppointmentDraft,
                                        cancel As CancellationToken) _
            As Task(Of OperationResult(Of AppointmentInfo))

        ''' <summary>
        ''' Apaga um compromisso. <b>Recusa reunião</b>: apagar reunião manda
        ''' cancelamento, e aí o estrago chega a terceiros.
        ''' </summary>
        Function DeleteAppointmentAsync(chave As AppointmentKey,
                                        cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean))
    End Interface

    ''' <summary>
    ''' <b>A porta estreita das TAREFAS.</b>
    '''
    ''' Leitura e escrita juntas aqui, ao contrário da agenda, e o motivo é
    ''' concreto: não existe consumidor que só leia tarefa. A separação da
    ''' agenda nasceu de uma necessidade real — a faixa da agenda foi entregue
    ''' lendo, um dia antes de existir escrita — e copiar a forma sem a
    ''' necessidade seria cerimônia.
    ''' </summary>
    Public Interface ITarefasBroker
        ''' <summary>
        ''' A pasta padrão de Tarefas.
        '''
        ''' Existe porque a <c>FolderVisibilityPolicy</c> mantém Tarefas fora da
        ''' árvore — a árvore mostra o que se ABRE como lista de mensagens, e
        ''' tarefa não é isso. Sem esta porta, a única forma de chegar na pasta
        ''' seria alargar a política e passar a mostrar na árvore uma pasta que
        ''' a lista não sabe abrir.
        ''' </summary>
        Function GetDefaultTasksFolderAsync(cancel As CancellationToken) _
            As Task(Of OperationResult(Of FolderKey))
        Function GetTasksAsync(folder As FolderKey, teto As Integer,
                               cancel As CancellationToken) _
            As Task(Of OperationResult(Of TaskList))

        ''' <summary>
        ''' Cria uma tarefa. <see cref="TaskDraft"/> não tem responsável, e a
        ''' ausência é a funcionalidade: tarefa atribuída é pedido enviado por
        ''' e-mail.
        ''' </summary>
        Function CreateTaskAsync(folder As FolderKey, rascunho As TaskDraft,
                                 cancel As CancellationToken) _
            As Task(Of OperationResult(Of TaskInfo))

        ''' <summary>
        ''' Conclui uma tarefa. <b>Recusa atribuída</b>: concluir tarefa de
        ''' outro manda atualização por e-mail.
        ''' </summary>
        Function CompleteTaskAsync(chave As TaskKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of TaskInfo))
    End Interface
    ''' <summary>
    ''' <b>Contatos.</b> Porta estreita, como as outras.
    '''
    ''' Não há operação de encaminhar cartão, e a ausência é a
    ''' funcionalidade: <c>ForwardAsVcard</c> devolve um <c>MailItem</c>, e é
    ''' o único caminho de envio que um contato tem. Sem operação na porta,
    ''' não há como chamá-la de fora do módulo.
    '''
    ''' Também não há apagar. Apagar contato não manda e-mail; é que apagar a
    ''' ficha de uma pessoa do catálogo é irreversível de um jeito que criar
    ''' não é, e nada nesta fase precisa disso. O Outlook apaga.
    ''' </summary>
    Public Interface IContatosBroker
        ''' <summary>
        ''' A pasta pessoal de Contatos.
        '''
        ''' Existe pelo mesmo motivo da porta das tarefas: a
        ''' <c>FolderVisibilityPolicy</c> mantém Contatos fora da árvore,
        ''' porque a árvore mostra o que se abre como lista de mensagens.
        ''' </summary>
        Function GetDefaultContactsFolderAsync(cancel As CancellationToken) _
            As Task(Of OperationResult(Of FolderKey))

        ''' <summary>
        ''' Lê a pasta. O resultado carrega <c>ForaDoAlcance</c> <b>sempre</b>,
        ''' inclusive quando deu certo e veio vazio: o GAL está fora de
        ''' escopo, e pasta vazia não é catálogo vazio.
        ''' </summary>
        Function GetContactsAsync(folder As FolderKey, teto As Integer,
                                  cancel As CancellationToken) _
            As Task(Of OperationResult(Of ContactList))

        ''' <summary>
        ''' Cria um contato. <see cref="ContactDraft"/> não tem nota nem
        ''' corpo: o que entra é o que a mensagem já dizia.
        ''' </summary>
        Function CreateContactAsync(folder As FolderKey, rascunho As ContactDraft,
                                    cancel As CancellationToken) _
            As Task(Of OperationResult(Of ContactInfo))
    End Interface

    Public Interface IOutlookBroker
        Inherits IAgendaSource, IAgendaWriter, ITarefasBroker, IContatosBroker


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

        ''' <summary>
        ''' <b>Os endereços pelos quais o dono desta caixa envia.</b>
        '''
        ''' Serve para uma pergunta só: <i>quem escreveu a última mensagem
        ''' desta conversa?</i> — que é o que separa "estou esperando alguém" de
        ''' "alguém está me esperando".
        '''
        ''' Vem em <b>mais de uma forma</b>, e de propósito: numa organização
        ''' Exchange o remetente de uma mensagem interna é um endereço X.500, e
        ''' as internas são justamente as que enchem a fila. Só o SMTP faria as
        ''' mensagens do próprio dono aparecerem como sendo de terceiros.
        '''
        ''' <b>Não sabe alias, caixa compartilhada nem delegação.</b> O que sai
        ''' daqui é um começo para o dono corrigir, e não a verdade — e é por
        ''' isso que o destino é um arquivo editável, e não uma decisão.
        ''' </summary>
        Function GetIdentidadesAsync(cancel As CancellationToken) _
            As Task(Of OperationResult(Of IReadOnlyList(Of String)))

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
        ''' <b>Os compromissos de uma janela de datas.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE JANELA, E NÃO PÁGINA</b>
        '''
        ''' Mensagem se lê por página porque a caixa é uma fila e o usuário
        ''' desce por ela. Calendário não: ninguém pede "os próximos 50
        ''' compromissos", pede "esta semana". E a diferença não é de gosto —
        ''' é que <b>ocorrência de série não existe até ser expandida</b>, e
        ''' expandir sem data-fim é infinito.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>A ARMADILHA DO OOM, E ELA É CLÁSSICA</b>
        '''
        ''' Com <c>IncludeRecurrences = True</c>, a coleção <b>tem</b> de ser
        ''' ordenada por <c>[Start]</c> <b>antes</b> do <c>Restrict</c>. Fora
        ''' dessa ordem o Outlook devolve a expansão errada — e devolve em
        ''' silêncio, sem erro, com uma lista que parece plausível.
        '''
        ''' Isso é responsabilidade da implementação, e está aqui na interface
        ''' porque quem escrever um segundo adaptador um dia precisa saber.
        '''
        ''' Custo medido em 28/08/2026: <b>30,9 ms por item</b>, contra ~16 ms
        ''' por mensagem na Fase 0. Uma janela larga é cara.
        ''' </summary>
        ''' <param name="de">Início da janela, inclusivo.</param>
        ''' <param name="ate">Fim da janela, exclusivo.</param>
        ''' <remarks>
        ''' Declarada em <see cref="IAgendaSource"/>, e não aqui. Ver o motivo lá.
        ''' </remarks>

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

        ''' <summary>
        ''' <b>A borda em lote:</b> o corpo de N mensagens numa visita só ao
        ''' Outlook.
        '''
        ''' A saída tem <b>uma posição por entrada</b>, na mesma ordem, com
        ''' <c>Nothing</c> onde a leitura falhou. Item que falha não derruba o
        ''' lote, e não some da lista — encolhê-la faria quem chamou casar ficha
        ''' com mensagem pelo índice errado.
        '''
        ''' Leitura pura, e portanto retentável: nada aqui escreve.
        ''' </summary>
        Function GetMessageSnapshotsAsync(items As IReadOnlyList(Of ItemKey),
                                          cancel As CancellationToken) _
            As Task(Of OperationResult(Of IReadOnlyList(Of MessageSnapshot)))

        ''' <summary>
        ''' Cada item tem anexo? — <c>Nothing</c> por item que não deu para
        ''' contar.
        '''
        ''' O portão nega mensagem com anexo, e precisa disto para negar pelo
        ''' motivo certo. Ler numa visita separada da do rótulo é deliberado, e
        ''' a corrida que isso abre é fechada em
        ''' <see cref="GetMessageSnapshotAsync"/>, que lê o anexo junto com o
        ''' corpo que vira bytes.
        ''' </summary>
        Function GetAttachmentPresenceAsync(items As IReadOnlyList(Of ItemKey),
                                            cancel As CancellationToken) _
            As Task(Of OperationResult(Of IReadOnlyList(Of AttachmentPresence)))

        Function GetSensitivityLabelsAsync(items As IReadOnlyList(Of ItemKey),
                                           cancel As CancellationToken) _
            As Task(Of OperationResult(Of IReadOnlyList(Of LabelReading)))

        ''' <summary>
        ''' <b>Em que pasta cada mensagem está</b> — perguntado ao Outlook, e não
        ''' declarado por quem chama.
        '''
        ''' O portão autoriza por pasta, e a pasta de cada mensagem vinha do mesmo
        ''' chamador que dizia qual pasta era o pedido: a comparação era entre duas
        ''' cópias da mesma afirmação. Ver <see cref="PastaDoItem"/>.
        '''
        ''' <b>Não lê corpo</b>, de propósito: o portão decide antes de qualquer
        ''' leitura de conteúdo, e juntar as duas seria ler o corpo para descobrir
        ''' se podia lê-lo.
        ''' </summary>
        Function GetItemFoldersAsync(items As IReadOnlyList(Of ItemKey),
                                     cancel As CancellationToken) _
            As Task(Of OperationResult(Of IReadOnlyList(Of PastaDoItem)))

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
        ''' <param name="versaoEsperada">
        ''' A <c>Version</c> da <see cref="SendPreview"/> que o dono aprovou. O
        ''' envio confere se o rascunho ainda é aquele <b>antes</b> de mandar, e
        ''' recusa com <see cref="ErrorKind.Stale"/> se não for — recusa limpa,
        ''' porque acontece antes do <c>Send</c> e portanto nada saiu.
        '''
        ''' <b>Obrigatório, e não opcional.</b> Um parâmetro com valor padrão
        ''' deixaria a chamada desprotegida continuar compilando, que é
        ''' exatamente o estado de onde este parâmetro veio.
        ''' </param>
        Function SendDraftAsync(draft As DraftKey, versaoEsperada As String,
                                cancel As CancellationToken) _
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

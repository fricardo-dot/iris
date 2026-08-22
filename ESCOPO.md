# Iris — Escopo do Projeto

**Status:** aprovado para iniciar a Fase 0
**Data:** 2026-08-22
**Versão:** 3

---

## 1. Visão

Cliente de produtividade pessoal semelhante ao Outlook, escrito em Visual
Basic, com inteligência artificial integrada ao fluxo de trabalho — não como
um chat lateral, mas dentro das ações do dia a dia: ler, triar, responder e
encontrar.

Uso pessoal, máquina única, Windows.

---

## 2. Restrição fundadora

**Microsoft Graph está descartado.** O acesso exige consentimento de
administrador do tenant corporativo, que não está disponível.

Consequência: o Iris não se autentica contra servidor nenhum. Ele usa o
**Outlook clássico já instalado e autenticado** na máquina como camada de
acesso aos dados, via automação COM (Outlook Object Model, daqui em diante
OOM).

Isso elimina de uma vez: OAuth2, registro de aplicativo, tokens, refresh,
IMAP/SMTP, senhas de aplicativo e armazenamento de credenciais de e-mail.

---

## 3. Princípio de desenho

> **O Iris é uma interface inteligente sobre o Outlook, não um substituto
> dele nem um espelho perfeitamente sincronizado da caixa de correio.**

Este princípio existe porque o maior risco do projeto não é o COM ser
incapaz de ler e enviar e-mail — ele é capaz. O risco é construir sobre a
suposição de que o OOM se comporta como uma API moderna: thread-safe,
determinística e com sincronização incremental. Ele não se comporta.

Três consequências que atravessam todo o documento:

- O acesso ao COM vive atrás de um **broker de thread única (STA) com
  message pump**.
- O cache local é um **subsistema**, com fase e critérios próprios, não um
  item de checklist.
- Capacidades que dependem de política corporativa (envio automático, acesso
  a mensagens classificadas) são **opcionais**, com fallback definido.

---

## 4. Arquitetura

```
+----------------------------------------------------------+
|  Iris (VB.NET / WPF, janela própria)                      |
|                                                           |
|  UI (WPF, thread STA de interface)                        |
|     |                                                     |
|     v                                                     |
|  Núcleo (async)  ---->  Cache local (SQLite)              |
|     |        |                                            |
|     |        +------->  API de IA (HTTPS)                 |
|     |                                                     |
|     v                                                     |
|  Broker Outlook  ..... fronteira: só DTOs a partir daqui  |
|  (thread STA própria + message pump + fila)               |
|     |                                                     |
|     v                                                     |
|  Outlook Object Model (COM)                               |
|     |                                                     |
|     v                                                     |
|  Outlook clássico (instalado e EM EXECUÇÃO)               |
+----------------------------------------------------------+
```

**Papéis:**

- **Outlook clássico** — motor de dados e transporte. Envia, recebe e
  sincroniza com o servidor. O Iris nunca fala com o Exchange diretamente.
- **Broker Outlook** — única porta de entrada para o COM. Detalhado abaixo.
- **Núcleo** — assíncrono. Orquestra cache, IA e UI.
- **Cache local (SQLite)** — espelho dos itens para listagem, busca, triagem
  e embeddings.
- **API de IA** — chamadas HTTPS, sempre a partir do núcleo.

### Regra de acesso a dados

> **A UI nunca recebe objetos COM nem chama o OOM diretamente.** Na Fase 1,
> ela recebe DTOs paginados produzidos pelo broker. A partir da Fase 2,
> listagem e busca leem exclusivamente do cache.

Esta formulação substitui o "a UI lê sempre do cache" da versão 2, que
contradizia o faseamento.

### O broker

Thread STA única não basta. A thread precisa também **atender mensagens
COM/Windows**, ou eventos deixam de chegar e chamadas rejeitadas não são
tratadas. O broker precisa ter:

- Thread criada explicitamente como STA
- **Message pump**, provavelmente um `Dispatcher` próprio
- Fila integrada ao `Dispatcher`, **não** um loop bloqueante — um loop que
  não bombeia mensagens impede a entrega de eventos
- `IOleMessageFilter` registrado, para tratar `RPC_E_CALL_REJECTED`
- **Referências fortes às coleções `Items` enquanto seus eventos estiverem
  assinados** — sem isso o coletor de lixo elimina o event sink e os eventos
  "param misteriosamente"
- Cancelamento e encerramento ordenados

**Sobre timeout:** não é possível abortar com segurança uma chamada COM já
em execução. Um timeout limita a espera do *chamador*; a operação pode
continuar rodando no broker. Isso importa muito para operações com efeito
colateral — ver R13.

A thread STA da UI do WPF **não** serve como thread do broker. Os objetos do
Outlook pertencem à STA do broker e só a ela.

**Requisito de operação:** o Outlook precisa estar instalado e em execução.

---

## 5. Módulos

Os quatro pilares do Outlook, com os limites reais de cada um explicitados.

| Módulo | Objeto COM | Limite conhecido |
|---|---|---|
| E-mail | `MailItem` | conversas não são reconstituíveis com perfeição |
| Tarefas | `TaskItem` | não equivale a To Do / Planner / flagged mail |
| Calendário | `AppointmentItem`, `MeetingItem` | recorrência, exceções, fusos, convites |
| Contatos | `ContactItem` | cobre a pasta de contatos, não GAL/LDAP |

**Notas:**

- `ConversationID` e `GetConversation()` dão resultados incompletos entre
  stores, caixas compartilhadas e PSTs. Não prometer thread perfeita.
- Uma coleção `Items` não contém apenas `MailItem`.
- `EntryID` **não** é identificador global permanente: muda quando o item
  troca de store. `EntryID` + `StoreID` é o ponto de partida, mas **não
  resolve movimentos** — em uma mudança de store os dois podem mudar, e o
  cache não consegue provar que é o mesmo item. O Iris precisará de **chave
  interna própria e estável**, definida na Fase 2 com base nas evidências da
  Fase 0.
- Recorrências exigem `IncludeRecurrences`, ordenação correta e intervalo
  limitado — sem isso, coleção praticamente infinita.
- Caixas compartilhadas e delegação introduzem stores e permissões extras.

---

## 6. Recursos de IA

Separados por **custo e implicação de compliance**, não por tema. Esta
divisão é deliberada: os dois grupos têm perfis de risco muito diferentes.

### Grupo A — sob demanda (barato, acionado por você)

1. **Resumir thread ou mensagem** — condensa conversas longas. Maior
   impacto imediato, menor complexidade.
2. **Redigir e responder** — gera rascunho a partir de instrução curta, com
   o contexto da thread. Sempre revisável; **nunca envia sozinho**.

Não depende do cache, mas depende de **montar contexto pelo broker**. Para
threads grandes é obrigatório impor limite de tamanho, truncamento e seleção
de quais mensagens entram.

### Grupo B — processamento em massa (caro, contínuo)

3. **Triagem** — classifica por prioridade e categoria, sinaliza o que exige
   ação. A primeira versão deve ser **lote iniciado explicitamente**, não
   processamento contínuo.
4. **Busca semântica** — encontra por sentido, não por palavra exata. Exige
   embeddings de cada item, armazenados localmente. Depende do cache
   **maduro**, não apenas existente.

**Autorização do Grupo B** (complementa R4): habilitação separada da do
Grupo A, escopo de pastas explícito, limite de orçamento e possibilidade de
interromper no meio.

**Decisão pendente da triagem:** ela grava categorias e flags **no Outlook**
ou somente no cache? Gravar no Outlook transforma uma análise de IA em
mutação de dado corporativo, e exige política de confirmação, idempotência e
reconciliação próprias. Começar somente no cache.

---

## 7. Faseamento

### Fase 0 — Spike técnico (antes de qualquer interface)

Objetivo: responder as perguntas que podem matar o projeto **ou mudar a
arquitetura**, na máquina corporativa real. Console feio, descartável.

Ponto crítico: o spike deve exercitar **a arquitetura proposta**, não apenas
provar que `MailItem` funciona. Se ele chamar COM direto na `Main`, valida o
OOM e não valida o desenho.

**A. Broker**
- [ ] Toda operação numa thread STA dedicada com message pump — não na
      thread principal do console
- [ ] Confirmar que os eventos chegam nessa mesma thread
- [ ] Confirmar que os DTOs voltam ao chamador sem carregar referência COM
      escondida

**B. Leitura**
- [ ] Conectar a uma instância do Outlook em execução
- [ ] Enumerar stores e seus tipos
- [ ] Comportamento em caixa principal, arquivo/PST e caixa compartilhada
- [ ] Encontrar item que não seja `MailItem` numa pasta real e ignorá-lo
      corretamente
- [ ] Ler mensagem, corpo e anexo

**C. Envio**
- [ ] Criar rascunho e exibi-lo com `Display()`
- [ ] **Tentar `MailItem.Send()` de verdade**, com destinatário controlado e
      identificador único no assunto
- [ ] Confirmar as três coisas separadamente: a chamada retornou, foi para a
      Outbox, e **foi entregue** — retorno sem erro não prova entrega
- [ ] Sem retry automático em nenhuma hipótese

**D. Eventos e mudanças**
- [ ] `ItemAdd`, `ItemChange` e `ItemRemove`
- [ ] Mover mensagem dentro do mesmo store
- [ ] Mover mensagem entre stores, se houver mais de um
- [ ] Alterar e excluir com o Iris **fechado**, confirmando que os eventos
      não recuperam essas mudanças

**E. Ciclo de vida e resiliência**
- [ ] Fechar o Iris e confirmar que o Outlook encerra sem processo órfão
- [ ] Fechar e reabrir o Outlook com o Iris rodando: detectar a invalidação
      dos objetos e restabelecer a sessão
- [ ] Iris iniciando antes do Outlook
- [ ] Forçar Outlook ocupado ou com diálogo modal, observar
      `RPC_E_CALL_REJECTED`, validar retry limitado **só** em leitura

**F. Desempenho e estado dos dados**
- [ ] Medir `Restrict`, `Sort` e paginação numa pasta grande real
- [ ] Medir leitura de 100 e de 1.000 mensagens só com metadados
- [ ] Medir corpo e anexos separadamente
- [ ] Testar em modo offline
- [ ] Mensagem com corpo ou anexo ainda não baixado: a propriedade bloqueia,
      falha ou dispara download?

**G. Ambiente e proteção**
- [ ] Bitness do Office instalado; testar `AnyCPU` (sem *Prefer 32-bit*) e a
      arquitetura correspondente — ver R12
- [ ] Comparar referência COM à Object Library com o pacote
      `Microsoft.Office.Interop.Outlook`
- [ ] Mensagens protegidas, em quatro variações: label informativo, label
      com criptografia/IRM, anexo protegido, e leitura de destinatários

**Resultados que mudariam a arquitetura:** broker sem message pump não
receber eventos; impossibilidade de reconectar; chamadas COM bloqueantes
imprevisíveis; política proibindo até `Display()`; incapacidade de ler a
maior parte das mensagens reais por IRM ou download parcial.

Se o envio estiver bloqueado por política, o fallback (rascunho + envio
manual) passa a ser o comportamento oficial do produto.

### Fase 1 — MVP de e-mail

Pastas, lista, leitura, rascunho, resposta, encaminhamento, anexos básicos,
via DTOs paginados do broker. **Sem busca global** — sem cache, a busca pelo
OOM é lenta e inconsistente; ela pertence à Fase 2.

### Fase 2 — Cache e sincronização

O subsistema. Precisa definir, não apenas mencionar:

- Chave interna estável do Iris e como correlacionar item antigo com novo
  após movimento
- Quando aceitar apagar e recriar em vez de preservar identidade
- Como detectar exclusões ocorridas com o Iris fechado
- Qual propriedade serve de checkpoint por pasta
- Reconciliação sem reler a caixa inteira

Mais importação inicial paginada, tombstones, retomada após falha e busca
textual. Só aqui listagem e busca passam a ler exclusivamente do cache.

### Fase 3 — IA sob demanda (Grupo A)

Resumo e redação sobre a mensagem ou thread aberta.

### Fase 4 — Triagem e busca semântica (Grupo B)

Depois de medir qualidade do cache e custo de indexação.

### Fase 5 — Tarefas

Inclui extração de tarefas a partir de e-mails, em duas etapas distintas:
a IA **sugere**, você confirma, o Iris cria o `TaskItem`. Nunca criação
silenciosa em massa.

### Fase 6 — Calendário / Fase 7 — Contatos

Dois marcos distintos, não um.

---

## 8. Fora de escopo

- Múltiplos usuários, perfis ou contas simultâneas
- Instalador e mecanismo de atualização
- Versão web ou mobile
- Substituir o Outlook como cliente de sincronização
- Suporte ao "novo Outlook" (ver R1)
- Reconstrução perfeita de conversas
- Paridade com To Do, Planner ou GAL

---

## 9. Riscos

### R1 — Migração para o "novo Outlook" — *impacto fatal*

O novo cliente não expõe o Object Model via COM. Se a TI forçar a migração e
remover o Outlook clássico, o Iris para por completo.
**Mitigação:** todo acesso COM atrás do broker, com interface única, **para
limitar o impacto de uma futura reimplementação da fonte de dados**. Note
que isso não a torna "substituível": Graph, EWS, IMAP e OOM têm modelos
diferentes de identidade, eventos, conversas e permissões.

### R2 — Guarda de segurança do Object Model

O Outlook pode avisar ou bloquear acesso programático a endereços de
destinatários e ao envio. Depende do antivírus registrado no Windows
Security Center, do Trust Center e de política de grupo. Se a política for
"negar", **não há solução dentro do aplicativo**.
**Mitigação:** validar na Fase 0; rascunho + `Display()` como comportamento
mínimo garantido, envio automático como opcional.

### R3 — Desempenho do COM

Iterar coleções item a item é lento.
**Mitigação:** `Restrict` e `Sort` do próprio Outlook, paginação,
sincronização para o cache. Medir na Fase 0.

### R4 — Privacidade

Recursos de IA enviam conteúdo de e-mails a um serviço externo.
**Mitigação:** enviar só o necessário; Grupo A sempre acionado
explicitamente; Grupo B com habilitação separada, escopo de pastas, limite
de orçamento e interrupção possível. Ver R11.

### R5 — Escopo grande

Quatro módulos e quatro recursos de IA é bastante trabalho.
**Mitigação:** as fases da seção 7; cada fase utilizável sozinha, exceto a
Fase 0, que é investigação descartável e não precisa ser.

### R6 — Threading, STA e reentrância COM

O OOM tem afinidade de thread. Obter um objeto numa thread e usar em outra,
ou despachar chamadas em `Task.Run`, causa falhas erráticas.
**Mitigação:** o broker da seção 4; só DTOs cruzam a fronteira.

### R7 — Referências COM e `OUTLOOK.EXE` órfão

Expressões encadeadas criam wrappers intermediários difíceis de liberar.
Eventos e `For Each` também seguram referências. Sintomas: Outlook não
encerra, processo órfão, itens bloqueados, Iris falando com instância morta.
**Mitigação:** referências curtas e nomeadas, descarte determinístico,
cancelar assinaturas de evento, não guardar objetos COM no modelo da UI,
liberar na ordem inversa. Não usar `FinalReleaseComObject` indiscriminado.
**Tensão deliberada com o broker:** as coleções `Items` com eventos
assinados precisam de referência forte viva. Soltar tudo cedo demais mata os
eventos.

### R8 — Ausência de sincronização incremental confiável

O OOM não tem delta token. Eventos se perdem com Iris fechado, Outlook
reiniciando, sincronização em lote, queda de conexão ou volume alto.
**Mitigação:** eventos para baixa latência **mais** reconciliação periódica
com checkpoints por pasta. O desenho concreto é entregável da Fase 2, com
evidência colhida na Fase 0.

### R9 — Estado offline e download parcial

O Outlook pode listar um item sem ter corpo ou anexos localmente, e a
chamada pode bloquear enquanto busca o conteúdo.
**Mitigação:** estados explícitos no DTO (metadados / corpo disponível /
anexos disponíveis / erro transitório); nunca bloquear a UI esperando
download. A modelagem dos DTOs depende do que a Fase 0 medir.

### R10 — Volume e anexos

Caixas com centenas de milhares de itens, HTML enorme, histórico citado,
assinaturas repetidas, anexos de centenas de MB.
**Mitigação:** limites de tamanho, anexos fora da indexação inicial,
normalização de HTML, remoção de citações e assinaturas, hashing,
processamento sob demanda.

### R11 — DLP, classificação e compliance

"Uso pessoal" não elimina obrigações corporativas. A empresa pode proibir
enviar conteúdo a uma API externa mesmo com acesso legítimo ao e-mail.
Sensitivity labels, IRM, Purview e mensagens criptografadas podem impedir
leitura ou transmissão.
**Mitigação principal: política explícita de permissão e opt-in**, não
tentativa automática de redigir dados. Concretamente:

- Escopo de pastas habilitado explicitamente, nunca por padrão
- Bloquear mensagens protegidas ou classificadas — exige descobrir na Fase 0
  como os labels aparecem nas propriedades MAPI
- Log do que foi enviado à IA registrando **metadados, hash, modelo e
  tamanho**, não o conteúdo — um log com o texto cria mais uma cópia
  sensível
- Confirmação explícita antes de anexos
- Verificar a política corporativa aplicável

Duas mitigações da v2 foram rebaixadas por não serem confiáveis: redigir
dados sensíveis automaticamente não é barreira de compliance, e lista de
domínios não protege contra conteúdo corporativo vindo de remetente externo.

### R12 — Arquitetura e bitness do Office

O Outlook é servidor COM local **fora do processo**, e o COM faz marshaling
entre um Iris x64 e um Outlook x86. Portanto correspondência de bitness
**não é obrigatória** para OOM puro — a v2 era categórica demais. Ela vira
crítica se entrarem componentes nativos no processo, Extended MAPI, certos
add-ins ou Redemption.
**Mitigação:** testar `AnyCPU` (sem *Prefer 32-bit*) e a arquitetura do
Outlook instalado na Fase 0; publicar só a configuração comprovada. Não
transformar "mesmo bitness" em requisito sem o teste demonstrar.

### R13 — Instabilidade do Outlook

O Outlook pode estar iniciando, encerrando, com diálogo modal aberto,
reparando um store ou indisponível por RPC (`RPC_E_CALL_REJECTED`).
**Mitigação:** `IOleMessageFilter`, retry limitado para rejeições RPC,
estados de reconexão no broker. **Nunca** retry de operação não idempotente
como envio. Lembrar que timeout não aborta a chamada COM (seção 4).

### R14 — O cache é uma segunda cópia dos dados corporativos

SQLite, embeddings e logs criam uma cópia do mailbox fora do OST, sem as
mesmas proteções e retenção. **Embeddings também podem revelar conteúdo** e
entram na mesma decisão.
**Mitigação:** ACL restritiva, expiração, botão de apagar o índice, anexos
não armazenados por padrão. Sobre criptografia, é preciso **escolher** — o
SQLite não tem criptografia nativa:

- campos sensíveis via DPAPI, ou
- distribuição compatível com SQLCipher, ou
- BitLocker + ACL, aceitando que o banco não é criptografado pela aplicação

---

## 10. Stack

**Decidido: VB.NET no .NET 10 (LTS), interface em WPF.**

| Opção | COM | HTTPS + JSON | Veredito |
|---|---|---|---|
| VB6 | nativo | sofrível — TLS moderno e JSON são dor | descartado |
| VBA dentro do Outlook | nativo | ruim, e sem janela própria | descartado |
| VB.NET / .NET Framework 4.8 | maduro | JSON e async piores | descartado |
| **VB.NET / .NET 10 LTS** | **funciona** | **excelente** | **escolhido** |

Razões:

- O diferencial do Iris é a camada de IA: HTTP concorrente, JSON, streaming,
  cancelamento e `Async`/`Await` são muito melhores no .NET moderno.
- O .NET 8 sai de suporte em novembro de 2026. Não se inicia projeto nele.
- O COM não fica mais confiável no .NET Framework: STA e liberação de
  referências continuam iguais. O isolamento é **arquitetural** (o broker),
  não uma questão de versão de runtime.
- WPF em VB **é suportado** no .NET moderno (`dotnet new wpf -lang vb`).
  A suposição contrária, presente na versão 1 deste documento, era falsa.
- WPF oferece binding, templates, estilos e virtualização — exatamente o que
  uma UI de cliente de e-mail precisa.

Voltar para o .NET Framework 4.8 só se aparecer dependência concreta que
falhe no teste com .NET moderno.

### Como referenciar o Outlook

Referência **fortemente tipada** à Object Library ou ao pacote
`Microsoft.Office.Interop.Outlook`. Não usar `CreateObject` e late binding
como estratégia principal. Comparar as duas formas na Fase 0 e verificar
qual produz build reprodutível em `net10.0-windows`. Usar *Embed Interop
Types* quando funcionar. Nada disso dispensa ter o Outlook clássico
instalado e registrado.

O núcleo depende da **interface do broker**, nunca dos tipos Interop.

### Armadilhas específicas de VB

Ligar desde o primeiro arquivo — sem isso, default properties e conversões
implícitas escondem chamadas COM:

```vb
Option Strict On
Option Explicit On
Option Infer On
```

- Evitar `For Each` sobre coleções COM com enumerador não controlado
- Evitar expressões encadeadas e `With` longo sobre objetos COM
- Não usar `ByRef` desnecessário em métodos que recebem tipos Interop
- Cuidado com propriedades parametrizadas e default do Outlook, que em VB
  se parecem com acesso comum
- `WithEvents` mantém referência COM viva; desligar explicitamente
- **Nunca** colocar tipo Interop em `Async Function`, view model ou DTO
- Não confiar no GC como mecanismo normal de encerramento

---

## 11. Decisões pendentes

- [x] Confirmar a abordagem COM com Outlook clássico
- [x] Escolher a stack
- [ ] Escolher o provedor de IA e o modelo — atenção: busca semântica exige
      **embeddings**, e nem todo provedor oferece esse endpoint
- [ ] Definir o visual: parecido com o Outlook ou identidade própria
- [ ] Criptografia do cache: DPAPI, SQLCipher ou BitLocker + ACL (R14)
- [ ] Triagem grava no Outlook ou só no cache? (seção 6)
- [ ] Verificar a política corporativa aplicável antes da Fase 3 (R11)

---

## Apêndice — histórico de revisão

**v3 (2026-08-22)** — segunda rodada de revisão externa; aprovado para
iniciar a Fase 0. Correções: resolvida a contradição entre "a UI lê sempre
do cache" e o faseamento, virando a regra de acesso a dados da seção 4;
broker especificado com message pump, `IOleMessageFilter` e referências
vivas para event sinks, mais a ressalva de que timeout não aborta chamada
COM; spike ampliado de 8 para ~30 verificações, cobrindo reconexão,
`ItemRemove` e movimentos, RPC rejeitado, offline e download parcial, e
desempenho medido; R12 deixou de exigir bitness idêntico; R11 rebaixou
redação automática e lista de domínios; R14 passou a exigir escolha de
criptografia; adicionadas as armadilhas de VB e a forma de referenciar o
Interop.

**v2 (2026-08-22)** — primeira revisão externa. Princípio de desenho da
seção 3; broker STA na arquitetura; cache promovido a fase própria; riscos
R6 a R14; IA separada por custo e compliance; Fase 0 criada; stack corrigida
de .NET 8 + WinForms para .NET 10 + WPF, após a premissa sobre WPF em VB se
mostrar falsa.

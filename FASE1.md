# Iris — Plano da Fase 1

**Objetivo:** um cliente de e-mail utilizável ponta a ponta, com janela
própria, lendo e escrevendo pela sessão do Outlook clássico.

**Pré-requisito:** Fase 0 concluída. Ver seção 10 do `ESCOPO.md`.
**Versão:** 7 — marco 1.5 fechado; plano do 1.6 na seção 13.

---

## 1. O que a Fase 1 entrega

- Árvore de pastas da caixa
- Lista de mensagens de uma pasta, com rolagem fluida
- Leitura: cabeçalho, corpo e anexos
- Salvar anexo (abrir é decisão da UI, com confirmação)
- Responder, responder a todos, encaminhar, escrever nova
- Enviar, ou salvar como rascunho
- Estado da conexão com o Outlook, o tempo todo

**Critério de "utilizável":** passar uma manhã lendo e respondendo e-mail no
Iris sem precisar abrir a janela do Outlook.

---

## 2. O que a Fase 1 NÃO entrega

- **Cache local** — é a Fase 2 inteira
- **Busca global** — depende do cache
- **Qualquer recurso de IA** — Fase 3 em diante
- Calendário, contatos, tarefas
- Regras, categorias, sinalizadores
- Formatação rica na composição

---

## 3. Estrutura da solução

```
Iris.sln
├── Iris.Model      DTOs e enums. NENHUMA referência a Interop.
├── Iris.Core       Serviços async E a interface IOutlookBroker.
├── Iris.Outlook    Implementa IOutlookBroker. Interop vive só aqui.
└── Iris.App        WPF. Referencia Core e Model; referencia Outlook
                    APENAS no composition root do startup.
```

`IOutlookBroker` mora em `Iris.Core`, não em `Iris.Outlook`: o núcleo depende
do contrato, e a implementação depende do núcleo. Alguém precisa compor os
dois, e esse alguém é o startup do `Iris.App` — em nenhum ViewModel ou View.

### O grafo AJUDA, não garante

A versão 1 deste plano afirmava que a estrutura "impõe" a fronteira. É forte
demais, e vale registrar por quê:

- **Referências transitivas.** Não ter `ProjectReference` direto não impede
  que tipos fluam pela cadeia. Mitigação: `PrivateAssets="all"` no
  `PackageReference` do Interop.
- **`Object` engole qualquer coisa.** Um DTO com `Property Value As Object`
  carrega um RCW sem nunca conhecer os tipos do Interop. A afirmação de que
  "Model não consegue" era falsa.

### Testes arquiteturais, antes do marco 1.1

Testes que **falham** se:

- `Iris.Model` referenciar `Microsoft.Office.Interop.*`
- Qualquer membro público de Model, Core ou App usar tipo do Interop
- ViewModel ou View referenciar o assembly `Iris.Outlook`
- Um DTO expuser `Object`, `dynamic` ou delegate arbitrário
- O Interop aparecer como dependência transitiva de `Iris.App`

O detector reflexivo de RCW do spike continua, mas como rede secundária nos
testes de integração — nunca como garantia principal.

---

## 4. A API do broker

`ReadAsync` e `MutateAsync` tornam-se **privados**. A superfície pública é um
conjunto de operações nomeadas, nenhuma aceitando ou devolvendo tipo do
Interop, e todas com `CancellationToken`.

### Sessão

```
ConnectAsync / ProbeAsync / DisconnectAsync
```

### Leitura

```
GetStoresAsync()
GetFolderChildrenAsync(parentKey)        ' sob demanda, não a árvore toda
GetMessagePageAsync(query, offset, count)
GetMessageDetailAsync(itemKey)           ' cabeçalho + corpo + anexos juntos
SaveAttachmentAsync(itemKey, attachmentKey, destino, política)
```

`GetMessageDetailAsync` é uma chamada só de propósito: cabeçalho, corpo e
anexos obtidos separadamente podem observar **estados diferentes** se a
mensagem mudar entre as chamadas.

O resultado do corpo carrega `ContentState`, formato, HTML, texto,
`IsProtected` e o tipo de erro — não uma string solta.

`attachmentKey` é um DTO, não um índice: índice é instável se a coleção
mudar. E `SaveAttachmentAsync` recebe a política de sobrescrita
explicitamente, em vez de decidir sozinho.

**Abrir anexo não é operação do broker.** O broker salva num diretório
controlado; a UI decide se pede confirmação e abre. Abrir anexo é executar
conteúdo não confiável.

### Rascunho e envio

Criar e enviar são operações **separadas**. Juntá-las impediria editar,
reabrir, anexar e tratar envio ambíguo sem criar outra mensagem.

```
CreateDraftAsync(novo)
CreateReplyDraftAsync(itemKey, replyAll)   ' usa Reply/ReplyAll do OOM
CreateForwardDraftAsync(itemKey)           ' usa Forward do OOM
UpdateDraftAsync(draftKey, alterações)
AddDraftAttachmentAsync(draftKey, caminho)
RemoveDraftAttachmentAsync(draftKey, attachmentKey)
SendDraftAsync(draftKey)                   ' MutateAsync, sem retry
DeleteDraftAsync(draftKey)
```

Responder e encaminhar **não** são "criar mensagem nova com texto citado":
`Reply`, `ReplyAll` e `Forward` do próprio Outlook preservam destinatários,
assunto, citação e anexos corretamente.

### Eventos

```
SubscribeFolderAsync(folderKey) As Task(Of SubscriptionToken)
UnsubscribeFolderAsync(token)
```

Token lógico, nunca um `IDisposable` embrulhando COM que possa escapar. O
evento público carrega apenas `FolderKey`, geração e tipo de invalidação —
sem dados da mensagem, já que a resposta correta é **reler**.

---

## 5. Paginação é VOLÁTIL, e isso é explícito

`offset`/`count` não dá paginação estável numa pasta viva:

1. Página 1 lê os itens 1–50
2. Chega mensagem nova no topo
3. Página 2 lê 51–100
4. Um item aparece duplicado e outro é pulado

A Fase 1 **aceita** isso, em vez de fingir resolver. A consulta carrega uma
**geração**:

```
GetMessagePageAsync(query{folderKey, sort, generation}, offset, count)
```

A resposta devolve a mesma geração. Quando a pasta muda ou o usuário troca
de seleção:

- a geração é incrementada
- resultados de geração anterior são **descartados**, nunca anexados
- recarrega-se do topo, ou a página atual

Sem isso, a resposta lenta de uma pasta anterior chega depois que o usuário
já clicou em outra, e sobrescreve a tela.

Se, no marco 1.3, paginação volátil e troca rápida de pasta não ficarem
aceitáveis, **a Fase 2 é antecipada**. O 1.3 é um ponto de decisão, não só
uma entrega.

---

## 6. Marcos

### 1.1 — Esqueleto, contratos e conexão ✅

- Quatro projetos e os testes arquiteturais da seção 3
- DTOs e tipos de erro definidos **antes** de migrar o broker:
  `ItemKey`, `FolderKey`, `AttachmentKey`, `DraftKey`, `ContentState`,
  `OperationResult`/`ErrorKind`, gerações
- Broker migrado do spike com a API fechada, **componente a componente**,
  revisando cada um — o spike não vira biblioteca por osmose
- Testes do broker que rodam **sem UI**
- Política de log e redação definida desde o primeiro log
- Janela mostrando estado: conectado / ocupado / fechado / reconectando
- **Reconexão já funcionando aqui**, não no 1.6

*Aceite:* abrir o Iris com o Outlook fechado mostra o estado certo sem
travar; abrir o Outlook depois conecta sozinho; fechar o Iris durante a
conexão não trava nem deixa `OUTLOOK.EXE` órfão; abrir e fechar rápido não
cria dois brokers; exceção na inicialização não deixa thread STA viva.

### 1.2 — Árvore de pastas ✅

Carregamento sob demanda por nível. Definido explicitamente: quais stores
entram, se pastas ocultas e de busca aparecem, e a ordem.

*Aceite:* subpastas carregam sob demanda sem duplicar filhos; criar ou
remover pasta no Outlook aparece após invalidação; contagem de não lidos é
tratada como **eventualmente consistente**, não instantânea.

### 1.3 — Lista de mensagens ✅ *(ponto de decisão)*

- `Restrict` e `Sort` feitos pelo Outlook, nunca em laço nosso
- Paginação volátil da seção 5, com geração
- Virtualização de verdade: recycling ligado, `CanContentScroll=True`,
  nenhum `ScrollViewer` externo quebrando a virtualização
- **Invalidação por pasta suja já aqui** — sem isso a lista fica
  deliberadamente obsoleta até o último marco

*Aceite:* numa pasta de pelo menos 5.000 itens, exibir **30 resumos em até
1 s no p95**, com a UI nunca bloqueada por mais de 100 ms. Medir cold e warm
start separadamente, troca rápida de pasta, cancelamento de página obsoleta,
rolagem contínua por centenas de itens e memória.

*Fixture:* a pasta de 5.000 itens precisa existir antes do aceite. **Não**
gerar 5.000 mensagens na caixa corporativa — usar store ou PST de teste.

### 1.4 — Leitura ✅

- Primeiro incremento: **texto puro**. WebView2 endurecido vem depois
- Estados explícitos do R9: só metadados / corpo disponível / erro
- Lista de anexos, com salvar
- Marcar como lido com a regra do F1-G

*Aceite:* a UI mostra estado de carregamento em até 200 ms e continua
respondendo enquanto o broker trabalha; resultado tardio é descartado se
outra mensagem foi selecionada; falha de HTML cai para texto puro; mensagem
protegida não vaza conteúdo para o log. Exige fixture com mensagem de
download parcial — sem ela, o critério não é verificável.

### 1.5 — Composição e envio ✅

- Responder, responder a todos, encaminhar, nova
- Anexar arquivo, salvar rascunho, enviar
- Confirmação mostrando **conta remetente e destinatários resolvidos**

*Aceite:* numa execução controlada sem falha, chega uma cópia e aparece uma
em Itens Enviados; `Send()` é chamado **uma única vez**; nenhum retry
automático; destinatário não resolvido **bloqueia** o envio; em falha
ambígua a UI mostra estado desconhecido e **não** oferece reenvio cego.

> "Exatamente uma vez" não é prometido como garantia. O spike demonstrou uma
> execução saudável, não semântica *exactly-once*, que o OOM não oferece.

### 1.6 — Consolidação e testes de falha

Não é onde reconexão e invalidação nascem — é onde são exercitadas: Outlook
fechado e reaberto durante o uso, Outlook ocupado, rajada de invalidações,
disco cheio, mensagem removida por baixo.

---

## 7. Regras que vêm da Fase 0

1. **A UI nunca chama COM.**
2. **Toda mutação usa `MutateAsync`**, com retry desligado.
3. **Evento é aviso de pasta suja, não transição.** Reler, nunca aplicar.
4. **Criar item em pasta específica exige `Save()` + `Move()`.**
5. **`EntryID` não sobrevive a movimento.**
6. **Nada bloqueia esperando download** de corpo ou anexo.
7. **Interop embutido**, senão compila e morre em execução.

---

## 8. Riscos da Fase 1

**F1-A — A lista grande.** A 16 ms por item, 5.000 DTOs levariam 80 s. A
paginação é o que torna a fase viável.

**F1-B — HTML hostil.** Ver F1-I, que é a versão detalhada.

**F1-C — Sem cache, a lista é volátil.** Recarregar em vez de remendar.

**F1-D — Escopo inflando.** A Fase 1 termina quando os seis marcos passam.

**F1-E — Resultados assíncronos obsoletos.** O usuário troca de pasta ou
mensagem enquanto uma chamada está em curso; a resposta antiga chega depois
e sobrescreve a seleção nova. Mitigação: geração por pasta e por seleção,
descarte de resultado antigo, e nunca confiar que cancelar interrompe COM.

**F1-F — Fila única do broker e starvation.** Uma leitura lenta de corpo ou
anexo bloqueia paginação e status, porque tudo passa pela mesma STA.
Mitigação: operações curtas, nunca ler corpo durante listagem, prioridades
de fila, coalescer recargas, e mostrar "Outlook ocupado" em vez de enfileirar
trabalho sem limite.

**F1-G — Laço ao marcar como lido.** Abrir mensagem → marcar lido → `Save` →
`ItemChange` → pasta suja → recarregar → seleção muda → marcar de novo.
Mitigação: só mutar se `UnRead = True`; marcar após 1–2 s de visualização,
não ao selecionar; debounce das invalidações; preservar seleção pelo
`ItemKey`; falha ao marcar é secundária e não impede ler.

**F1-H — Concorrência Outlook e Iris.** O usuário move, exclui ou responde
no Outlook enquanto o Iris exibe o item. Mitigação: toda operação pode
retornar `NotFound`/`Stale`; reobter o item pela chave imediatamente antes de
mutar; nunca manter `MailItem` vivo entre operações; após mutação, reler;
mostrar "a mensagem foi movida ou removida" sem quebrar.

**F1-I — HTML e conteúdo remoto.** "Sem script, sem remoto" precisa ser
verificável, porque o HTML pode trazer `iframe`, `object`, `embed`, SVG
ativo, CSS externo, URLs `file:`, `data:` e `javascript:`, pixels de
rastreamento, formulários e redirecionamentos. Ver seção 9.

**F1-J — Anexos perigosos.** Abrir anexo executa conteúdo não confiável.
Mitigação: mostrar nome, tipo e tamanho; confirmação antes de abrir; salvar
em diretório controlado; não sobrescrever em silêncio; nunca abrir
automaticamente; preservar Mark-of-the-Web; limitar tamanho; tratar disco
cheio — que esta máquina já provou não ser hipotético.

**F1-K — Rascunhos.** Autosave gera muitos eventos; `EntryID` muda após
move; o rascunho pode ser editado no Outlook ao mesmo tempo; reabrir exige
reobter pela chave, nunca guardar RCW; fechar o Iris com edição pendente
precisa perguntar salvar ou descartar.

**F1-L — Conta remetente.** O Outlook pode escolher conta ou delegação de
forma não óbvia. O DTO de composição mostra a conta efetiva, e o envio nunca
sai silenciosamente pela conta errada.

**F1-M — Log com conteúdo sensível.** O log não inclui corpo, assunto
completo, endereço completo nem caminho de anexo. Só ID, HRESULT, operação,
duração e hash. Vale desde o primeiro log, não como faxina posterior.

---

## 9. Renderização do corpo: WebView2

**Decidido: WebView2 endurecido, com texto puro primeiro e fallback
permanente para texto.**

Um visualizador próprio parece mais seguro e não é: converter HTML de e-mail
em `FlowDocument` exige parser, sanitizador, CSS parcial, tabelas, imagens
CID, citações e HTML malformado. O resultado tende a ser inseguro **e** feio.

Configuração mínima, toda verificável:

- JavaScript desabilitado; DevTools desabilitado
- Host objects e web messages desabilitados
- Navegação fora do documento inicial bloqueada
- Toda requisição interceptada, bloqueando `http`, `https`, `file`, `ftp` e
  esquemas desconhecidos
- Popups e novas janelas bloqueados
- Link externo só abre após ação explícita **e** confirmação
- HTML sanitizado antes de chegar ao WebView2, como defesa em profundidade
- Nenhuma API .NET injetada na página
- Diretório de perfil e cache dedicado, com política de limpeza definida —
  senão o WebView2 acumula dados corporativos em cache
- Imagens CID resolvidas apenas a partir dos anexos da própria mensagem,
  nunca por caminho escolhido pelo HTML

---

## 10. Antes do marco 1.1

- [ ] Definir os DTOs e tipos de erro (`ItemKey`, `FolderKey`,
      `AttachmentKey`, `DraftKey`, `ContentState`, `ErrorKind`, gerações)
- [ ] Criar os testes arquiteturais anti-Interop e anti-`Object`
- [ ] Fechar a semântica de rascunho, resposta, encaminhamento e envio
      ambíguo
- [ ] Definir a política de log e redação
- [ ] Preparar o fixture de 5.000 itens para o aceite do 1.3
- [ ] Decidir identidade visual: parecido com o Outlook ou própria

O spike fica preservado como está. Cada componente migra revisado, um a um —
ele não vira biblioteca por osmose.

---

## 11. Dívida registrada

O que NÃO foi feito, dito às claras. Nada aqui foi aprovado por
extrapolação nem por semelhança com outra medição.

### Do marco 1.3

- **Fixture de 5.000 itens no OOM: NÃO medido.** A primeira página foi
  validada com 1.033 itens reais, em 391 ms. O critério fala em 5.000, e a
  caixa corporativa não tem esse volume — gerar 5 mil mensagens nela está
  fora de cogitação. O caminho é um PST de teste, com perfil descartável ou
  anexação temporária autorizada, e limpeza garantida.
- **Acesso por índice em offset profundo: NÃO medido.** Só a primeira
  página foi cronometrada. `Items.Item(i)` pode não ser O(1) no OOM, e
  offsets 300, 600 e 900 precisam de medição própria.
- **Virtualização do WPF: MEDIDA e aprovada.** 5.000 DTOs sintéticos numa
  janela real, contando containers realizados: dezenas, não milhares, com
  controle negativo que exige o contador acusar mais de mil quando a
  virtualização é desligada. Isso prova o WPF e não diz nada sobre o custo
  do Object Model — são medições separadas, e confundi-las seria o mesmo
  erro de extrapolar.

### Do marco 1.4

- **WebView2 para corpo HTML: NÃO feito.** O corpo é texto puro. A seção 9
  descreve a configuração endurecida, e ela continua valendo.
- **Endereços Exchange legados: RESOLVIDO no 1.5.** `AddressPolicy`, no
  Core, decide o que é endereço conferível; `/O=...` não é. A leitura tenta
  `ExchangeUser`, `ExchangeDistributionList` e `PR_SMTP_ADDRESS`, e o que
  sobrar sem SMTP **bloqueia o envio** — mesmo que o Outlook diga que
  resolveu. Continua aceitável para EXIBIR na leitura, que é outro
  contexto: ali ninguém vai mandar nada.
- **Corpos grandes: não medidos na UI.** Os 51 ms medidos são de uma
  mensagem comum. Corpos de 100 KB e de 512 KB precisam de medição própria
  antes de o teto ser considerado seguro.
- **Leitura parcial não é sinalizada.** Se destinatários ou anexos falharem
  em parte, a UI não distingue "não tem" de "não deu para ler". Passa no
  1.4; não passa no 1.5, onde responder depende dos destinatários.

### Do marco 1.5

- **Marca de separação em texto puro é VISÍVEL.** Em HTML a marca é um
  comentário e ninguém a vê. Em texto puro não existe marca invisível, e
  numa mensagem nova a linha `----- mensagem original -----` aparece sem
  ter original nenhuma embaixo. O padrão do Outlook é HTML, então este é o
  caminho raro — mas ele existe e está errado.
- **Remover anexo: NÃO implementado.** `RemoveDraftAttachmentAsync` ainda
  devolve `NotImplemented`. Dá para anexar e não dá para desanexar.
- **Rascunho existente não é reaberto.** O compositor só trabalha com
  rascunhos que ele mesmo criou na sessão. Abrir um rascunho antigo da
  pasta Rascunhos cai no palpite conservador do leitor — corpo inteiro
  vira citação — e ainda não foi exercitado.
- **Conta remetente pode não ser determinável.** Quando não há
  `SendUsingAccount` e nenhuma conta entrega no store do rascunho — caixa
  compartilhada, envio delegado, `SentOnBehalfOf` —, a confirmação diz
  "não foi possível identificar a conta" em vez de adivinhar. É honesto e
  **não** bloqueia o envio: bloquear tornaria o Iris inutilizável em
  configurações que ele não consegue inspecionar. Quem trabalha em caixa
  compartilhada precisa saber disso.
- **Envio ambíguo: caminho NÃO exercitado.** O estado terminal está
  testado contra o broker de mentira. Contra o Outlook de verdade não foi
  provocado, porque provocá-lo exigiria fazer um `Send` real falhar no
  meio.
- **Anexos grandes não medidos.** Anexar foi verificado com o diálogo,
  não com arquivo de dezenas de MB, que ocupa a fila única da STA.
- **Anexar pelo diálogo real não foi automatizado.** O `OpenFileDialog` é
  janela do sistema e a automação da verificação não passa por ele. A
  lógica de anexar está coberta por teste contra o duplo, incluindo a
  troca de chave; o diálogo em si depende de conferência manual.
- **O que FOI verificado com Outlook aberto:** criar rascunho, autosave
  gravando, chave relida a cada Save, confirmação com conta remetente e
  destinatário resolvido em SMTP, pergunta de fechamento e descarte. O
  envio em si não foi disparado — a Fase 0 já provou o `Send` no critério
  C2, e repetir aqui mandaria mensagem de verdade sem necessidade.

### Revisão externa do 1.5 — seis passadas

O marco foi revisado externamente até voltar limpo. Não uma vez: seis.

| Passada | Achados | Onde |
|---|---|---|
| 1 | 9 | no código do marco |
| 2 | 7 (3 graves) | nas **correções** da passada 1 |
| 3 | 4 (2 graves) | nas correções da passada 2 |
| 4 | 3 (3 graves) | nas correções da passada 3 |
| 5 | 1 bloqueante + 2 dívidas | nas correções da passada 4 |
| 6 | 0 | aprovado |

Vinte e quatro defeitos ao todo, e **quinze deles foram introduzidos por
mim ao corrigir os anteriores**. Todos na mesma vizinhança: corrida entre
digitação, gravação, conferência, envio e fechamento.

Três lições que vão além deste marco:

**Um teste que passa prova aquele caminho, não a propriedade.** Corrigi a
corrida "digitar DURANTE a gravação", escrevi o teste, ele passou, e
concluí que a família estava coberta. A irmã dela — "digitar DEPOIS da
gravação e ANTES da prévia ficar pronta" — sobreviveu intacta.

**Corrigir corrida cria corrida.** A trava que resolveu a disputa pela
chave cobria só a gravação; como anexar descarrega e DEPOIS anexa, a
disputa migrou para a fresta entre as duas. Depois, com a trava correta,
a geração passou a ser fotografada tarde demais, e quem esperava na fila
comparava com o número errado. Cada correção precisou da sua própria
revisão.

**Teste sem controle negativo é decoração.** O teste de envio duplo
passava verde mesmo sem a trava existir, porque o `AsyncRelayCommand`
barrava a segunda execução por conta própria. A prova só veio quando o
comando passou a permitir concorrência de propósito, deixando a trava ser
a única barreira — e quando desfiz a correção para ver o teste falhar.
Os quatro bloqueios críticos deste marco foram conferidos assim.

### Do fechamento do 1.5

- **Rascunho órfão em caso raro.** Se a criação do rascunho concluir no
  Outlook depois de o compositor ter sido descartado, o item fica na pasta
  Rascunhos, vazio. É consequência de a mutação não ser cancelável, não de
  o estado ressuscitar. Custo: um rascunho vazio que o usuário apaga.
- **Domínio de rótulo único é reprovado.** `fulano@intranet` existe em
  rede interna e não passa em `AddressPolicy.IsUsableSmtp`. Deliberado: a
  regra decide se o usuário consegue CONFERIR o destino, não se o e-mail
  funciona. Errar bloqueando custa um envio pelo Outlook; errar no outro
  sentido custa mensagem na caixa de quem não devia.
- **Encerramento de sessão do Windows não pergunta.** Logoff, desligar e
  reiniciar não passam pela pergunta de fechamento — cancelar o
  desligamento por causa de um rascunho seria arrogante, e um overlay
  assíncrono não tem tempo garantido de resposta. A proteção ali é o
  autosave de 1,5 s; a exposição é o que foi digitado nesse intervalo.

### Geral

- **D4 e D7 da Fase 0** seguem sem teste: movimento entre stores e reinício
  do Outlook com assinatura ativa.
- **Rótulos do Purview** continuam não testados. Bloqueiam a Fase 3.

---

## 12. Decisões do marco 1.5

Fechadas antes de escrever código, para não virarem improviso no meio.

### O rascunho é do Outlook, não nosso

Responder, responder a todos e encaminhar usam `Reply`, `ReplyAll` e
`Forward` do próprio OOM. Reconstruir destinatários, citação e assinatura à
mão seria reimplementar regras que o Outlook já aplica — e errar nelas
significa mandar mensagem para quem não devia.

### Formato: o do rascunho, não o nosso

O compositor edita **texto**, mas a inserção respeita o formato que o
rascunho já tem. Se o Outlook gerou corpo HTML — com citação e assinatura
corporativa —, o texto digitado entra como HTML escapado no topo. Forçar
texto puro destruiria a citação e a assinatura, que é justamente o que
torna uma resposta utilizável no trabalho.

### O rascunho existe cedo

Ele é criado e salvo ao ABRIR o compositor, não na primeira edição. Assim
sobrevive a um fechamento acidental, tem chave estável desde o início, e o
`EntryID` é relido após cada `Save` — porque ele muda em movimentação.

### Autosave com debounce, nunca por caractere

1,5 s de debounce, uma mutação corrente e no máximo um estado pendente. Uma
fila de `Save` produziria versões intermediárias sem valor e ocuparia a fila
única da STA.

### Fechar com alterações pendentes pergunta

Salvar, descartar ou cancelar. Descartar em silêncio é perder trabalho do
usuário.

### Envio

- Destinatário não resolvido **bloqueia** o envio.
- A confirmação mostra conta remetente, destinatários resolvidos em SMTP,
  assunto e anexos.
- `Send()` uma vez, sem retry, com identificador de correlação — o
  procedimento que a Fase 0 provou no critério C2.
- Envio ambíguo **bloqueia novo envio daquele rascunho** até a
  reconciliação. Reenviar no escuro é o único erro irreversível deste
  projeto.

---

## 13. Plano do marco 1.6 — consolidação e testes de falha

Escrito antes do código, para não virar improviso no meio.

### O que este marco NÃO é

Não é onde reconexão e classificação de erro nascem. Elas existem desde o
1.1 e funcionam. Este marco é onde elas são **exercitadas e provadas** — e
onde o que ficou pela metade nos marcos anteriores é fechado.

### O achado que ancora o marco

`Classify` e `ClassifyFailure`, em `OutlookBroker.vb`, têm **zero testes**.

São as regras que decidem se uma falha pode ser repetida. Em particular,
`ClassifyFailure` contém a única defesa contra o pior erro possível deste
projeto: uma falha depois de a mutação começar vira `Ambiguous`, que não é
retentável. O comentário no código conta que já houve o bug oposto — um
`Send` que estourava depois de a mensagem sair virava `NotConnected`, cujo
`IsRetryable` é `True`, e o código convidava a reenviar exatamente no caso
em que reenviar duplica.

Essa regra vive dentro de um arquivo de 852 linhas que só compila em
Windows e só roda com Outlook. Nenhum teste a alcança. É a mesma situação
de `EhSmtp` antes do 1.5, e a solução é a mesma: a regra é lógica pura,
então ela sai para o `Iris.Core` e ganha testes.

### Grupo A — extrair e provar as regras de falha

**A1. `FailureClassification` no `Iris.Core`.**
Move `Classify` (HRESULT → `ErrorKind`) e a regra de ambiguidade. O broker
passa a chamar a política; nada de comportamento muda. Testes cobrindo:
cada HRESULT mapeado, HRESULT desconhecido, e — o principal — que
`isMutation:=True` com efeito iniciado devolve `Ambiguous` para QUALQUER
HRESULT, inclusive os que sozinhos seriam `NotConnected` ou `Busy`.
Controle negativo: leitura com o mesmo HRESULT NÃO vira `Ambiguous`, senão
a regra estaria classificando tudo como ambíguo e passaria de graça.

**A2. `SessionState` a partir do HRESULT.**
A mesma extração para a classificação do probe: quais HRESULTs derrubam a
sessão e forçam reconexão, quais são "ocupado", e a decisão — já tomada e
comentada no código — de que HRESULT não classificado reporta estado
degradado em vez de mentir `Connected`.

### Grupo B — exercitar as falhas na interface

Contra o `FakeBroker`, que já sabe falhar sob comando. Cada um com
controle negativo.

- **B1. Queda durante o uso.** Árvore, lista e leitor são limpos; o
  compositor NÃO é, porque o texto é trabalho do usuário. Já decidido no
  1.5; falta o teste que impede alguém de "simplificar" isso depois.
- **B2. Volta da conexão.** Árvore recarrega, assinatura é restabelecida,
  e a recarga acontece UMA vez — não uma por evento.
- **B3. Item removido por baixo.** `NotFound` tratado ao ler, ao marcar
  como lida, ao responder e ao enviar. Cada um mostra o que aconteceu em
  vez de falhar em silêncio.
- **B4. Rajada de invalidações.** Vinte eventos em sequência produzem uma
  recarga, não vinte. O debounce do `FolderWatcher` já existe; o teste é o
  que impede a regressão.
- **B5. Gravação falhando por disco/permissão.** O texto permanece na
  tela, o status explica, e a próxima tecla tenta de novo. Parcialmente
  coberto no 1.5; falta o caso do disco.

### Grupo C — fechar o que ficou pela metade

- **C1. `RemoveDraftAttachmentAsync`.** Hoje devolve `NotImplemented`: dá
  para anexar e não dá para desanexar. É a lacuna mais visível do 1.5.
- **C2. Marca de separação visível em texto puro.** Numa mensagem nova em
  texto puro, a linha `----- mensagem original -----` aparece sem ter
  original nenhuma embaixo.
- **C3. Sinalização de leitura parcial.** Dívida do 1.4 que o próprio
  documento marcou como "não passa no 1.5". Hoje a UI não distingue "não
  tem destinatário" de "não deu para ler os destinatários".

### Grupo D — o que exige o Outlook de verdade

Estes **não** dependem de código novo, e sim de um roteiro executado com o
Outlook aberto. Envolvem fechar e reabrir o Outlook do usuário, que é o
cliente de e-mail real dele: **pedir autorização antes**, e nunca executar
por conta própria.

- **D1.** Fechar o Outlook com o Iris rodando: estado vira "Outlook
  fechado", operações devolvem `NotConnected`, compositor preservado.
- **D2.** Reabrir: reconexão automática pelo watchdog, sem reiniciar o
  Iris. É o que valida que o RCW morto é descartado e um novo é adquirido.
- **D3.** D4 e D7 da Fase 0, ainda sem teste: movimento entre stores e
  reinício do Outlook com assinatura ativa.

### O que fica FORA do 1.6, e por quê

- **Fixture de 5.000 itens e offsets profundos.** Precisa de PST de teste;
  é trabalho de infraestrutura de medição, não de consolidação.
- **WebView2.** É superfície nova, não consolidação.
- **Reabrir rascunho existente.** É funcionalidade, não robustez.
- **Envio ambíguo contra o Outlook real.** Provocá-lo exigiria fazer um
  `Send` de verdade falhar no meio. O que dá para provar sem isso é a
  regra de classificação — e é exatamente o que o grupo A faz.

### Critério de pronto

1. `Classify` e a regra de ambiguidade fora do broker, com testes e
   controles negativos.
2. Os cinco cenários do grupo B com teste, cada um com controle negativo.
3. Grupo C fechado.
4. Roteiro do grupo D executado com o usuário, com resultado registrado
   aqui — inclusive se algum falhar.
5. Suíte estável em 10 execuções seguidas.
6. Revisão externa até voltar sem bloqueante, incluindo as correções.

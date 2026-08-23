# Iris — Plano da Fase 1

**Objetivo:** um cliente de e-mail utilizável ponta a ponta, com janela
própria, lendo e escrevendo pela sessão do Outlook clássico.

**Pré-requisito:** Fase 0 concluída. Ver seção 10 do `ESCOPO.md`.
**Versão:** 11 — critério de 5.000 itens dispensado; medição substituta na seção 15.

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

### 1.6 — Consolidação e testes de falha ✅

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

- **Fixture de 5.000 itens: critério DISPENSADO pelo usuário** em
  2026-08-23, por não ser viável na caixa corporativa. Ver seção 15: o
  critério não foi cumprido, foi retirado, e há medição substituta.
- **Acesso por índice em offset profundo: MEDIDO** em 2026-08-23, e o
  resultado é bom. Ver seção 15.
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

**Versão 2.** A primeira versão deste plano foi submetida a revisão externa
antes de virar código, e foi reprovada por um bom motivo: ela produziria
muitos testes verdes ao redor de riscos secundários enquanto deixava
descobertos os dois defeitos capazes de duplicar um envio e de silenciar as
atualizações para sempre. Os dois foram encontrados durante a revisão do
PLANO, no código que já estava lá, e confirmados linha a linha.

### Os dois defeitos que reorganizam o marco

**F1-M — Reconexão invisível.** `OnWatchdogTick` só emite evento quando o
estado MUDA:

```vb
If agora <> antes Then SetState(agora)
```

Se o Outlook morre e volta dentro da janela de 15 s do watchdog, `ProbeCore`
detecta a morte, chama `ReleaseSessionCore` — que **limpa `_subscriptions`
inteiro** — e reconecta. O probe devolve `Connected`, igual ao anterior.
Nenhum evento é emitido.

O resultado é uma falha silenciosa e permanente: o broker não tem mais
assinatura nenhuma, o `FolderWatcher` continua guardando um token que já não
existe, a UI não recarrega, e a lista para de se atualizar **até o Iris ser
reiniciado**. O usuário não recebe sinal nenhum.

`SessionState` não consegue representar isto, porque "Connected, mas é outra
sessão" não é um estado — é uma mudança de identidade.

**F1-N — Fase da operação é global.** `_effectStarted` é um campo do broker.
`RunAsync` zera na thread do CHAMADOR antes de postar; a operação marca 1 já
na STA; e o `Catch` lê depois do `Await`. Qualquer operação concorrente —
uma recarga de pasta disparada por evento, por exemplo — zera o campo entre
a falha de um `Send` e a classificação dela.

Consequência: um `Send` que falhou DEPOIS de a mensagem sair pode ser
classificado como `NotConnected`, cujo `IsRetryable` é `True`. É exatamente
o bug que o comentário do próprio `ClassifyFailure` diz ter corrigido — ele
foi corrigido na regra e continuou vivo no estado que a regra lê.

A única defesa contra reenviar uma mensagem que talvez já tenha saído
depende de um campo compartilhado entre operações concorrentes.

### Eixo 1 — época de sessão e reconexão observável

1. **`SessionEpoch`**: um número que sobe a cada aquisição de sessão COM.
   Toda chave — `FolderKey`, `ItemKey`, `DraftKey`, `SubscriptionToken` —
   pertence a uma época.
2. **Evento de sessão substituída**, separado de `StateChanged`. Emitido
   sempre que a época sobe, inclusive no caminho `Connected → Connected`.
3. **A UI reage**: limpa árvore, lista e leitor; relê stores; refaz
   assinatura. O texto do compositor é PRESERVADO — é trabalho do usuário —
   mas a `DraftKey` da época anterior é marcada como desligada, e gravar ou
   enviar por ela fica bloqueado até o usuário resolver. Preservar o texto é
   certo; preservar cegamente a capacidade de gravar por uma chave de outra
   sessão não é.
4. **HRESULT desconhecido não pode significar Busy eterno.** Hoje um código
   não previsto devolve `Busy` e preserva o RCW; se o RCW estiver morto, o
   Iris fica "ocupado" para sempre. Passa a haver limiar: falha desconhecida
   repetida em probes consecutivos força descartar e reanexar; se reanexar
   falhar, `Unavailable`.
5. **HRESULTs de morte a acrescentar**, com a ressalva de que lista nunca é
   prova completa — por isso o limiar acima é que carrega a robustez:
   `RPC_E_SERVER_DIED` (0x80010007), `RPC_E_SERVER_DIED_DNE` (0x80010012),
   `RPC_S_CALL_FAILED` (0x800706BE), `RPC_E_INVALID_OBJECT` (0x80010114).
   `MAPI_E_NETWORK_ERROR` **não** entra: pode ser Outlook vivo com Exchange
   indisponível.

### Eixo 2 — fase local e política pura de falha

1. **A fase da operação sai do broker e vira local à invocação.** É a
   correção de F1-N, e vem antes de qualquer extração: extrair a regra sem
   corrigir o estado que ela lê seria testar a regra certa alimentada por
   dado errado.
2. **`OutlookFailurePolicy` no `Iris.Core`** — nome específico de propósito.
   `AddressPolicy` é política de domínio usada por duas camadas; esta é
   tradução da borda COM, e chamá-la de "política" no mesmo tom seria
   arrumação estética. Assinatura pura:
   `ClassifyFailure(hresult As Integer?, isMutation As Boolean, mutationAttemptStarted As Boolean) As ErrorKind`.
   O broker continua dono de observar a fase e de escrever o log.
3. **Nomear a fase pelo que ela é.** Hoje "efeito iniciado" significa "o
   delegate `work` começou", não "a chamada COM mutante começou". É
   conservador e correto para segurança, mas o nome promete precisão que o
   código não tem: passa a `mutationAttemptStarted`.
4. **Testes**: cada HRESULT mapeado; desconhecido; e o principal — mutação
   com tentativa iniciada devolve `Ambiguous` para QUALQUER HRESULT,
   inclusive os que sozinhos dariam `NotConnected` ou `Busy`. Controle
   negativo: leitura com o mesmo HRESULT NÃO vira `Ambiguous`, senão a regra
   estaria carimbando tudo de ambíguo e passaria de graça.
5. **Teste de concorrência**, que é o que prova F1-N corrigido: duas
   operações sobrepostas não compartilham fase.

### Eixo 3 — reações da UI, de forma determinística

**`ScenarioBroker`, separado do `FakeBroker`.** O duplo atual foi feito para
o compositor: `State` fixo em `Connected` e exceção em tudo mais. Transformá-lo
em simulador universal produziria um objeto enorme e permissivo, que aceita
qualquer coisa e por isso não prova nada.

**O debounce sai do relógio.** O acumulador temporal do `FolderWatcher` vira
uma máquina pura que recebe `now`. Testar 2 s de relógio real com
`DispatcherTimer` transforma a suíte num teste de carga da máquina — e
"estável em 10 execuções" viraria sorte, não prova. A lição do 1.5 vale
aqui: já tive um teste intermitente e ele é pior que teste nenhum.

Os testes que valem, e o que cada um detecta:

| Teste | Mutação que ele pega |
|---|---|
| Eventos em t=0 e t=400 ms não recarregam em t=450 | debounce que conta do PRIMEIRO evento — a implementação que este projeto já teve e corrigiu |
| Rajada contínua abaixo de 450 ms recarrega perto do teto de 2 s | teto que nunca participa |
| Evento da assinatura A pendente, troca para B: B não fica suja | dispatch sem geração |
| `Subscribe(A)` termina depois de `Watch(B)`: A é desassinada | assinatura órfã |
| Segunda rajada gera segunda recarga | debounce que trava depois da primeira |

E as reações de sessão: queda limpa árvore/lista/leitor e preserva o
compositor; substituição de sessão dispara UMA recarga; resultado atrasado
da época anterior é descartado; `NotFound`, `Ambiguous` e falha de
persistência produzem o estado visual certo.

**Controle negativo onde ele detecta alguma coisa**, não como ritual. A
regra mecânica "cada cenário com o seu" produziria asserts duplicados; cada
controle precisa nomear a mutação que pretende pegar.

### Eixo 4 — dívidas funcionais

- **`RemoveDraftAttachmentAsync`.** Hoje devolve `NotImplemented`, e a
  assinatura está errada: devolve `Boolean`, mas remover SALVA, e salvar
  pode trocar a `DraftKey` e reconstruir todas as `AttachmentKey`. Vira
  `OperationResult(Of DraftInfo)`, como `AddDraftAttachmentAsync` — é
  exatamente o mesmo defeito que o 1.5 corrigiu no anexar, sobrevivendo no
  irmão que ainda não tinha sido escrito.
- **Marca de separação visível em texto puro.**
- **Sinalização de leitura parcial.** Fechar o contrato semântico ANTES de
  mexer na tela: destinatários, anexos e corpo falham de forma
  independente, e um `ContentState` único não representa isso. E a decisão
  que importa não é visual — se os destinatários vieram incompletos,
  responder e enviar provavelmente devem ser BLOQUEADOS até uma releitura
  bem-sucedida, não apenas ganhar um ícone.

### Eixo 5 — o Outlook de verdade

Não depende de código novo. Envolve fechar e reabrir o Outlook do usuário,
que é o cliente de e-mail corporativo dele: **pedir autorização antes**,
nunca executar por conta própria.

- Fechar com o Iris rodando: estado, operações e compositor.
- Reabrir: é o teste que prova F1-M corrigido de verdade. O duplo prova o
  protocolo; só o Outlook real prova que o RCW morto foi solto, que o ROT
  entrega a instância nova e que os sinks são recriados.
- D4 e D7 da Fase 0: movimento entre stores e reinício com assinatura ativa.

### O que o duplo prova, e o que não prova

Registrado porque a primeira versão deste plano confundia as duas coisas.

**Prova:** reação ao protocolo de sessão — limpeza, preservação do
compositor, habilitação de comandos, recarga única, descarte de resultado
de época anterior, estado visual de cada `ErrorKind`, e que a UI não repete
mutação.

**Não prova:** que fechar o Outlook gera determinado HRESULT; que o ROT
entrega a instância nova quando se espera; que os RCWs foram liberados; que
sinks COM sobrevivem ou são recriados; que outro perfil se comporta como
modelado; que disco cheio chega como a exceção que o duplo simula.

Por isso o teste de persistência **não** se chama "disco cheio": ele prova
que o compositor recebeu uma falha de gravação e preservou o texto. É o que
ele prova, e é o nome que ele leva.

### Resultado do eixo 5, executado em 2026-08-23

Com autorização do usuário, o Outlook dele foi fechado e reaberto com o
Iris rodando. O log do broker registrou:

```
epoca 1 -> Connected      conexao inicial
state Unavailable         Outlook fechado; watchdog notou em 11 s
epoca 2 -> Connected      sessao NOVA adquirida; reconexao em 33 s
```

**Provado:** o RCW morto é solto, uma instância nova é adquirida, a época
sobe, `SessionReplaced` dispara, e a árvore recarrega com dado fresco —
as contagens de não lidos vieram diferentes das anteriores, ou seja, houve
releitura de verdade e não cache.

**NÃO provado, e é preciso dizer:** o caminho exato do F1-M —
`Connected → Connected` sem transição de estado — não foi reproduzido. O
Outlook desta caixa leva de 30 a 90 s para subir, e o probe roda a cada
15 s, então o watchdog sempre pega o `Unavailable` no meio. O caminho
específico continua provado só pelos testes de ViewModel.

Isso significa que a janela do F1-M é **estreita na prática**: exige o
Outlook voltar entre dois probes. Não muda a decisão de corrigir — falha
silenciosa e permanente merece defesa mesmo quando rara, e a época passou a
ser o sinal confiável independentemente de o estado mudar — mas o registro
tem que dizer o que foi medido, e não o que foi presumido.

**Efeito colateral observado:** encerrar o Outlook com `Stop-Process
-Force` deixou o processo travado no arranque por vários minutos — janela
fantasma de 322×18 px em coordenada negativa, sem registro no ROT. Não é
defeito do Iris; é o Outlook não gostando de morte súbita. Na segunda
execução o `Quit()` funcionou e a reabertura foi limpa. Quem repetir este
roteiro deve usar `Quit()` e ter paciência, não `Stop-Process`.

### Critério de pronto

1. F1-M e F1-N corrigidos, com teste de concorrência para o segundo.
2. Época de sessão observável, e UI reagindo à substituição.
3. `OutlookFailurePolicy` no Core, com testes e controles negativos.
4. Testes determinísticos do debounce, sem relógio real.
5. Eixo 4 fechado, com o contrato de leitura parcial decidido por escrito.
6. Roteiro do eixo 5 executado com o usuário, resultado registrado aqui —
   inclusive o que falhar.
7. Revisão externa até voltar sem bloqueante, incluindo as correções.

### Fora do 1.6

WebView2 e reabrir rascunho existente: são funcionalidade, não robustez.

**A fixture de 5.000 itens fica fora do 1.6 mas passa a BLOQUEAR declarar a
Fase 1 concluída.** É critério de aceite do 1.3 que nunca foi cumprido, e
manter isso como dívida perpétua enquanto se declara a fase consolidada
seria dar por medido o que não foi.

---

## 14. Resultado do marco 1.6

### O que foi feito

**F1-M e F1-N corrigidos.** Os dois defeitos que a revisão do PLANO
encontrou no código existente. A fase da operação virou local à invocação;
a sessão ganhou época observável, com evento próprio.

**`OutlookFailurePolicy` no Core.** A regra que decide se uma falha pode
ser repetida saiu do broker — 852 linhas que só compilam em Windows — e
ganhou 14 testes, incluindo o principal: mutação iniciada é ambígua para
QUALQUER HRESULT, com dois controles negativos.

**`DirtyDebounce` no Core.** O acumulador temporal saiu do relógio. Oito
testes que recebem o instante como parâmetro, cada um nomeando a mutação
que pega.

**Contrato de leitura parcial.** `PartState`/`PartStatus` por componente,
e `ReplyReadiness` decidindo o que fica bloqueado. Destinatários
incompletos bloqueiam responder e o envio; anexos incompletos bloqueiam
encaminhar; corpo incompleto não bloqueia. A contagem é lida duas vezes,
e divergência invalida o snapshot.

**Remover anexo**, com identificação que RECUSA quando há ambiguidade em
vez de escolher — numa operação que apaga, "o mais provável" não serve.

**Marca invisível em texto puro**, com uma propriedade de usuário
identificando rascunho do Iris.

**Restauração de pasta na reconexão**, descendo o caminho e expandindo o
que for preciso. Quem chega no meio de uma carga já em andamento ESPERA
por ela — a primeira versão desistia, e desistir fazia a restauração
concluir "a pasta não existe" quando ela existia e só não tinha terminado
de carregar.

**116 testes** (eram 57 no início do marco).

### Revisão externa — quatro passadas

| Passada | Achados |
|---|---|
| do PLANO | 2 graves, no código que já existia |
| 1 | 3 (2 bloqueantes) |
| 2 | 3 (1 bloqueante) |
| 3 | 4 |
| 4 | 1 bloqueante |
| 5 | 0 — aprovado |

Cinco passadas, treze achados, e o marco fechado sem bloqueante.

O padrão do 1.5 se repetiu: **a maioria dos achados das passadas 1 a 3
foram defeitos que eu introduzi corrigindo os anteriores.** Três exemplos
que valem registrar, porque são a mesma falha de raciocínio em lugares
diferentes:

- Bloqueei gravar quando a sessão troca, e deixei de fora "Salvar e
  fechar" e "Descartar" — que também gravam.
- Corrigi isso, e deixei de fora `ConfirmarEnvioAsync`, que é o caminho
  irreversível.
- Corrigi isso conferindo a sessão na ENTRADA das operações, e a sessão
  pode trocar durante a descarga que vem depois.

Em todos os casos eu tinha afirmado que o problema estava fechado. A
lição não é "revisar mais": é que **"eu bloqueei X" precisa vir com a
lista de todos os caminhos que fazem X**, e eu não a fiz nenhuma das
três vezes.

### O que foi verificado com o Outlook real

Fechar e reabrir o Outlook com o Iris rodando, com autorização do
usuário. O log registrou `epoca 1 → Connected`, `Unavailable` após 11 s,
`epoca 2 → Connected` após 33 s, e a árvore recarregou com contagens
diferentes — releitura, não cache.

**NÃO reproduzido:** o caminho exato do F1-M (`Connected → Connected` sem
transição). O Outlook desta caixa leva de 30 a 90 s para subir e o probe
roda a cada 15 s, então o watchdog sempre pega o `Unavailable` no meio. A
janela do defeito é estreita na prática. A correção continua certa — falha
silenciosa e permanente merece defesa mesmo rara — mas quem provou o
caminho específico foram os testes de ViewModel.

### Dívida que sai deste marco

- **Envio ambíguo contra o Outlook real** continua sem exercício: provocá-lo
  exigiria fazer um `Send` de verdade falhar no meio.
- **Marca `IrisDraft` pode não ser gravável** se `UserProperties` e
  `PropertyAccessor` forem os dois negados por política. A consequência só
  aparece ao REABRIR um rascunho de texto puro numa sessão futura — e
  reabrir rascunho existente não está implementado. Quando estiver, isto
  precisa ser tratado antes.
- **Anexos grandes** continuam sem medição.
- **Handlers de `SessionReplaced` rodam na STA** no caminho do watchdog. O
  `Try` por assinante protege contra exceção, não contra bloqueio: um
  handler que chame o broker e espere trava a STA. O contrato diz que
  handler devolve ao dispatcher dele; nada impõe isso.

### Sobre a fixture de 5.000 itens

Critério dispensado pelo usuário em 2026-08-23. Ver seção 15 para a
decisão, a medição que ficou no lugar, e o que continua sem resposta.

---

## 15. Fixture de 5.000 itens: critério dispensado

**Decisão do usuário, 2026-08-23:** não gerar 5.000 mensagens. O motivo é
bom — a caixa é corporativa, e encher uma pasta com cinco mil mensagens de
teste é intrusivo mesmo com limpeza depois.

O critério **não foi cumprido**. Foi retirado. Esta seção existe para que
ninguém leia "Fase 1 concluída" e presuma que os 5.000 foram medidos.

### O que o critério queria saber

Duas coisas diferentes, que estavam grudadas num número:

1. **A lista aguenta muitos itens sem travar?** — pergunta sobre o WPF.
2. **`Items.Item(i)` degrada em offset profundo?** — pergunta sobre o
   Object Model. Se degradar, "Carregar mais" fica progressivamente mais
   lento e a paginação por índice não escala.

### O que ESTÁ medido

**(1) já estava.** 5.000 DTOs sintéticos numa janela real, contando
containers realizados: dezenas, não milhares, com controle negativo que
exige o contador acusar mais de mil quando a virtualização é desligada.
Isso prova o WPF e não diz nada sobre o custo do OOM — são medições
separadas.

**(2) foi medido agora**, sem criar item nenhum: na Caixa de Entrada real,
com 1.003 itens, cronometrando páginas de 50 em offsets crescentes e
tocando as mesmas propriedades que o DTO da lista usa. Somente leitura.

| offset | exec 1 | exec 2 | exec 3 |
|---|---|---|---|
| 0 | 9,51 | 6,41 | 7,19 |
| 100 | 6,82 | 7,27 | 6,35 |
| 300 | 6,83 | 6,86 | 6,22 |
| 600 | 6,51 | 6,98 | 6,41 |
| 900 | 7,18 | 7,29 | 6,81 |

(ms por item)

**O custo não tem correlação com o offset.** Offset 900 sai igual a offset
0. A dispersão é ruído de cache e sincronização — a primeira execução de
todas deu 13 ms/item em dois offsets, e sumiu nas seguintes, por isso três
execuções e não uma.

**Conclusão:** `Items.Item(i)` é O(1) na prática nesta caixa, em modo
cached. A paginação por índice escala, e "Carregar mais" não fica mais
lento à medida que o usuário desce.

### O que CONTINUA sem resposta

Dizer o que a medição não cobre importa tanto quanto o resultado:

- **Pasta com 5.000 itens no OOM não foi exercitada.** O comportamento foi
  medido até 1.003. Nada garante que `Items.Count` e `Items.Sort` — que
  rodam uma vez por pasta, antes da primeira página — se comportem igual
  numa coleção cinco vezes maior. O `Sort` é o candidato mais provável a
  degradar, e ele não foi cronometrado em separado.
- **Modo online (não cached) não foi medido.** Esta caixa está em cached
  mode. Sem cache local, cada acesso vira ida ao servidor, e o resultado
  acima não se transfere.
- **Outras pastas não foram medidas.** Só a Caixa de Entrada.

### Se um dia isto voltar a importar

O caminho que não toca na caixa corporativa é um **PST local** com itens
gerados, anexado temporariamente a um perfil de teste. Não foi feito, e
não está planejado. Fica registrado como a saída conhecida, não como
pendência.

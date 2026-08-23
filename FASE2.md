# Fase 2 — Cache e sincronização

**Versão:** 6 — Q1 RESPONDIDA. Resultado na seção 9.

A v1 foi reprovada por um bom motivo: ela transformava em **pergunta de
medição** coisas que são **decisões de correção**. Perguntar "qual é a
chave estável?" e esperar que um benchmark responda pode produzir um
desenho logicamente errado, por mais rápido que ele seja.

Esta versão separa as duas coisas: o que já está decidido porque é
invariante, e o que depende de número.

---

## 1. Por que esta fase existe

A Fase 1 mediu o que custa ler do Outlook: **~600 ms por página de 50
itens**, dominados pela obtenção e resumo item a item.

A Fase 0 mediu o que torna isso difícil de resolver:

- **Não há delta token.** O Object Model não tem "o que mudou desde X".
- **Eventos não recuperam o que aconteceu com o Iris fechado** (R8).
- **`EntryID` muda quando o item se move.**

---

## 2. Revisão explícita do ESCOPO

O `ESCOPO.md` diz que, a partir da Fase 2, "listagem e busca leem
exclusivamente do cache". **Esta fase reabre a parte da LISTAGEM**, e o
faz de propósito: se `Folder.GetTable()` tornar a leitura direta rápida o
bastante, ler direto pode continuar sendo a resposta certa para listar.

A BUSCA não é reaberta. Ela continua exigindo cache, e por isso o
subsistema é necessário de qualquer forma — junto com estado local de
triagem, histórico de frescor, e o que a Fase 4 vier a indexar.

**Correção da v1:** eu tinha escrito que a Q1 poderia "encolher a fase
inteira". Não pode. Ela pode encolher a parte "cache como acelerador de
lista". O subsistema de cache e sincronização continua de pé.

---

## 3. Decisões tomadas AGORA, porque são invariantes

Não dependem de medição, e adiá-las seria convidar um desenho errado.

### 3.1 Identidade

**O cache tem chave interna própria, opaca, gerada pelo Iris.** Nenhuma
propriedade MAPI isolada é declarada identidade universal.

Hierarquia, e o papel de cada camada:

| Camada | Papel |
|---|---|
| Chave Iris | Identidade local, definitiva |
| `StoreID` + `EntryID` | **Localizador atual**, não identidade |
| `PR_SEARCH_KEY` | Evidência forte, condicionada ao provider |
| Internet Message-ID | Evidência útil, **nunca sozinha** |
| Tamanho, datas, remetente, assunto | Desambiguação; fracas isoladamente |

`ConversationIndex` **sai** da lista de candidatas a identidade. Serve
para agrupar conversa; cópias o preservam e mensagens relacionadas
compartilham prefixo.

**Em correlação ambígua, NÃO unir.** Apagar e recriar perde continuidade;
unir errado põe o resumo da IA e o estado do usuário na mensagem errada.
Perder continuidade é o dano menor, e é o que se escolhe.

### 3.2 Eventos

**Evento é invalidação, nunca transição a aplicar.** Já vale desde a
Fase 1; passa a valer também para o cache.

### 3.3 Sincronização

- **Incremental é ACELERADOR, nunca prova de ausência.**
- **Reconciliação completa periódica é obrigatória.**
- **Ausência de uma associação item–pasta só é confirmada depois de uma
  varredura completa e BEM-SUCEDIDA daquela pasta.** O cache nunca
  confirma "removido"; confirma, no máximo, "ausente desta pasta".

### 3.4 O que uma ausência prova

**Ausência numa varredura prova, no máximo, que aquela ASSOCIAÇÃO
item–pasta não existia naquele universo, naquele instante.** Não prova que
a mensagem foi excluída.

Produzem observações parecidas: mover, excluir, perder acesso à pasta,
mudar o filtro ou o universo varrido, e provider indisponível.

Consequências no desenho:

- **Presença pertence à relação item–pasta**, não ao item. O estado
  terminal chama-se `Ausente da pasta`, e nunca *mensagem excluída*.
- Um item Iris tem um localizador atual e **uma ou mais associações a
  pastas**. São coisas separadas no modelo, e a consequência aparece em
  3.7 (a geração marca associações), 3.8 (o estado é da associação) e na
  seção 7 (cada linha da lista é uma associação).
- **É resultado aceitável da Q5 concluir "indistinguível pelo OOM".** Se
  for, a política fica conservadora — e o plano não promete uma distinção
  que o Object Model talvez não permita.

### 3.5 Mudança de universo invalida a geração

Store, pasta, filtro, inclusão de ocultos, janela de retenção e provider
fazem parte da **identidade da varredura**. Marcas produzidas sob
universos diferentes não são comparáveis, e comparar é como se apagam
itens que nunca sumiram.

### 3.6 Enumeração concorrente duvidosa NÃO faz sweep

Concluir todas as chamadas COM sem erro é necessário e **pode não ser
suficiente**: a pasta pode ter mudado durante a paginação, e aí a
enumeração não é snapshot de nada.

A geração precisa satisfazer um critério de consistência — que a Q4 vai
calibrar. **Não satisfazendo, ela termina sem confirmar ausência
nenhuma.** O spike calibra o critério; ele não decide se exclusão falsa é
aceitável, porque não é.

### 3.7 Varredura por geração (mark-and-sweep)

**A geração é de uma PASTA, e o que ela marca são ASSOCIAÇÕES
item–pasta** — não itens. Varrer a Caixa de Entrada não diz nada sobre a
presença do mesmo item em Arquivo Morto.

1. Inicia geração N **da pasta P**, sob um universo declarado (ver 3.5).
2. Marca cada **associação (item, P)** vista com N.
3. **Só depois de todas as páginas de P concluírem com sucesso**, e a
   geração satisfazer o critério de consistência de 3.6, as associações
   (item, P) não vistas viram ausentes **daquela pasta**.
4. Cancelamento, erro, universo alterado, ou Outlook indisponível no meio:
   a geração **não confirma ausência nenhuma**.

Um item cuja associação foi removida de TODAS as pastas conhecidas
continua não sendo prova de exclusão: pode estar numa pasta que o Iris
não varre. Ver 3.8.

Sem isto, uma falha no meio apaga metade do cache — o R2-H.

### 3.8 Estado de presença por ASSOCIAÇÃO item–pasta

O estado não é do item. É de cada par (item, pasta), e o mesmo item pode
estar `Presente` numa pasta e `Ausente` de outra ao mesmo tempo — que é o
caso normal, não a exceção.

Quatro estados, não dois:

| Estado | Significa |
|---|---|
| `Presente` | Visto nesta pasta na última geração válida |
| `Não verificado` | Nunca varrido, ou a última geração não foi válida |
| `Suspeito de ausência` | Não visto, mas a geração não permite confirmar |
| `Ausente da pasta` | Não visto numa geração válida e consistente |

"Não verificado" é o que impede o cache de afirmar o que não sabe.

**Nenhum destes estados é "mensagem excluída".** O item que não está em
nenhuma pasta conhecida do Iris ganha, no máximo, um estado derivado —
`Sem localização conhecida` — que a UI apresenta como *"não encontrado"*, e
nunca como *"excluído"*. A diferença importa: excluir é um fato sobre o
mundo, e o Iris não tem como estabelecê-lo.

### 3.9 Lotes interrompíveis

Importação e reconciliação são **sequências de unidades curtas**, nunca
uma operação longa na fila da STA. Ver seção 6.

### 3.10 Armazenamento

**SQLite**, salvo bloqueio concreto de distribuição ou dependência nativa.
Não é benchmark: para busca textual, transação, migração, índice e
recuperação, é a opção natural. LiteDB adiciona dependência com modelo de
consulta menos adequado; arquivo próprio compraria corrupção, locking e
indexação artesanais sem benefício.

**O que precisa de decisão de verdade, e vale mais que o benchmark:**
criptografia em repouso e política de retenção. O cache é cópia local de
correspondência corporativa — é o R14 do escopo aparecendo aqui.

### 3.11 Duas classes de dado, e só uma é reconstruível

A v2 dizia "o cache é sempre reconstruível, além de estado local". Essa
frase escondia duas coisas muito diferentes:

| Classe | Origem | Se o arquivo sumir |
|---|---|---|
| **Derivado** | Outlook | Reconstrói. Custa tempo, não perde nada. |
| **Estado local durável** | O usuário e o Iris | **Perde-se.** Triagem, vínculo com resumo da IA, decisões tomadas. |

O estado local durável precisa de armazenamento logicamente separado, com
backup e migração próprios. Chamar o arquivo inteiro de "sempre
reconstruível" seria falso, e a diferença aparece exatamente no dia em que
o arquivo corromper.

---

## 4. Marco 2.0 — Spike de medição

Formato da Fase 0: código descartável, critério objetivo. Somente leitura,
exceto onde marcado.

### Q1 — `Table` contra iteração, por COLUNA

Não basta "mesmo DTO". Matriz por coluna do `MailSummary`:

| Situação | |
|---|---|
| vem na tabela padrão | |
| vem via `Columns.Add` | |
| não vem | |
| vem com semântica diferente | |
| exige abrir o item | |

Suspeitos principais: `Attachments.Count`, estado de proteção/IRM/Purview,
propriedades calculadas, e valores que possam disparar download.

**Se sete colunas vierem em lote mas proteção e anexos ainda exigirem
abrir cada item, o ganho desaparece** — e é por isso que a matriz importa
mais que o número agregado.

Fases cronometradas em SEPARADO, para não atribuir ao Outlook um custo que
está na conversão: criar a tabela · adicionar colunas · filtrar/ordenar ·
`GetArray` · converter em DTO · fallback por item.

Armadilhas a verificar: tabela padrão tem poucas colunas; nem toda
propriedade vira coluna; propriedades grandes/binárias/multivalor têm
restrição; coluna ausente pode voltar vazia em vez de erro; Jet e DASL
diferem, sobretudo em data; conversão local/UTC; `GetArray` grande
monopoliza a STA; tabela **não** é snapshot estável enquanto a pasta muda;
conteúdo associado/oculto não pode entrar no universo reconciliado.

### Q2 — Evidências de correlação (absorve a antiga Q5)

Não "qual é a chave vencedora", e sim:

> **Quais evidências permitem correlacionar duas manifestações do mesmo
> item, e onde elas erram?**

**Correção da v2:** eu pedia "taxa de falso positivo e falso negativo" e
propunha mover UM item como experimento central. Um item não produz taxa
nenhuma — isso mede sobrevivência de propriedade, não qualidade de
correlação.

O experimento precisa de um **corpus adversarial com oráculo manual** —
uma lista em que eu SEI, de antemão, quais pares são o mesmo item e quais
não são. Casos que o corpus tem de conter:

| Caso | O que ele ameaça |
|---|---|
| A mesma mensagem em duas pastas | falso negativo |
| Enviado e recebido da mesma mensagem | **falso positivo** |
| Rascunho antes e depois de enviar | Message-ID que aparece só no envio |
| Encaminhamento e reenvio | ID novo ou preservado, depende da ferramenta |
| Cópias com Message-ID idêntico | **falso positivo** |
| Message-ID ausente ou vazio | evidência que não existe |
| Pares deliberadamente parecidos, e distintos | **falso positivo** |
| Item movido entre pastas | é o D4 da Fase 0, sem teste desde então |

Para cada evidência (`PR_SEARCH_KEY`, Message-ID, `PR_RECORD_KEY`,
combinações), registrar contagem de acertos e erros **contra o oráculo**,
com os limites explícitos — não uma taxa estatisticamente defensável, que
o tamanho do corpus não sustenta.

**Movimento entre STORES fica como não validado.** Esta máquina tem um
store só, e fabricar conclusão sobre o que não dá para exercitar é pior
que registrar a lacuna.

**Resultado aceitável inclui "não correlacionar automaticamente".**

### Q3 — Incremental como acelerador

`Items.Restrict` com `[LastModificationTime] > X`: funciona, é rápido, e
**onde falha**.

O experimento precisa incluir, porque é onde o desenho quebra:

- vários itens com timestamp idêntico;
- alteração feita DURANTE a leitura;
- item movido para fora da pasta (não aparece na consulta);
- exclusão (não aparece).

**Desenho já decidido**: high-water mark com JANELA DE SOBREPOSIÇÃO —
consultar desde `checkpoint − janela`, reprocessar de forma idempotente, e
**só avançar o checkpoint depois do commit local**. Uma falha parcial
seguida de avanço prematuro perde itens para sempre.

### Q4 — Enumerar chaves, e provar que a varredura foi completa

Custo de obter chave + `LastModificationTime` de todos os itens de uma
pasta, pelo caminho que a Q1 indicar.

**Mais importante que o custo:** testar consistência sob mutação —

- item criado durante a enumeração;
- item removido durante a enumeração;
- item movido entre duas pastas enquanto ambas são varridas;
- falha COM no meio;
- cancelamento no meio;
- Outlook reiniciado no meio.

A pergunta é *"como sei que esta varredura foi completa o bastante para
confirmar AUSÊNCIA DA PASTA?"* — e não "para confirmar remoção", que é uma
afirmação sobre o mundo que o Iris não tem como fazer.

**O que a Q4 calibra, e não decide:** encontrar enumeração afetada por
mutação concorrente é esperado, não é surpresa. As opções de política já
estão na mesa, e o spike escolhe entre elas com dado:

1. descartar a geração e repetir;
2. exigir **duas** observações completas e compatíveis antes de confirmar;
3. manter as associações candidatas em `Suspeito de ausência` até
   verificação individual.

O que **não** está em discussão é aceitar exclusão falsa.

### Q5 — Política de verificação (era "exclusão deixa rastro?")

A pergunta antiga tem resposta pouco útil: esvaziamento e exclusão dura
não deixam tombstone recuperável pelo OOM. A pergunta boa é:

> **Que política de verificação distingue movido, excluído e
> temporariamente invisível sem produzir exclusão falsa?**

**Aviso sobre o experimento:** item esvaziado da lixeira ou apagado com
`Shift+Del` **não volta**. Este teste exige pasta descartável e
consentimento explícito, aceitando a perda do item de teste. Não prometer
restauração.

### Q6 — Identidade de pasta e de store

Pastas movem, são renomeadas, excluídas e recriadas. `FolderKey` hoje é
`EntryID + StoreID`. O que acontece ao mover uma pasta? E ao remontar o
store ou trocar de perfil?

### Q7 — Custo e semântica do conteúdo para busca

Metadado rápido não resolve busca textual se indexar corpo exigir abrir
cada item ou disparar download. **Medir separado**, porque é o que decide
se a busca do cache é viável.

**O spike MEDE o custo e NÃO PERSISTE corpo nem anexo.** Gravar
correspondência corporativa em disco antes de criptografia e retenção
estarem decididas seria criar o R2-I durante o experimento que deveria
informá-lo.

### Q8 — Matriz de providers

Um número da Caixa de Entrada em modo cached não vira garantia geral.
Registrar o que foi medido em: Exchange cached, offline, PST. Caixa
compartilhada só se entrar no produto.

### Q9 — Fronteira de retenção

Todas as pastas e todo o histórico, ou uma janela? **Esta decisão pode
reduzir o problema mais que qualquer otimização**, e é do usuário.

---

## 5. Riscos

Ordenados por gravidade **estimada**. O primeiro é novo, e é o pior.

Duas ressalvas sobre a ordem: **R2-I pode virar bloqueador absoluto** se a
política corporativa exigir, superando qualquer risco operacional; e
**R2-A é categoria ampla**, que engloba várias das linhas abaixo dela — a
tabela mistura risco-raiz com consequência, e isso vai atrapalhar
priorização mais adiante.

| ID | Risco |
|---|---|
| **R2-G** | **Fusão falsa.** Dois itens distintos correlacionados como o mesmo. Resumo da IA, estado local ou ação do usuário vão para a mensagem errada. Pode levar a responder à mensagem errada. Pior que perder continuidade. |
| R2-H | Varredura parcial confirma exclusões em massa. Falha no meio + mark-and-sweep errado apaga dado válido. |
| R2-A | Cache diverge e o usuário confia nele. Mensagem que sumiu continuar aparecendo é pior que lentidão. |
| R2-I | Cache é cópia corporativa desprotegida em disco. É o R14 do escopo, e decide criptografia e o que se indexa. |
| R2-L | Offline confundido com ausência. Tabela vazia por indisponibilidade **nunca** pode confirmar exclusão. |
| R2-B | Importação inicial monopoliza a fila da STA e deixa o Iris inútil enquanto roda. |
| R2-C | Sem correlação boa, todo movimento vira "apagou e criou", perdendo estado local. |
| R2-J | Sincronização inunda o Outlook e deixa os dois lentos. |
| R2-K | Esquema sem migração/transação: atualizar o Iris torna o banco incompatível. |
| R2-M | Pasta ou store muda de identidade: cache órfão ou duplicado. |
| R2-D | Exclusão com o Iris fechado indetectável a custo aceitável. |
| R2-F | Busca do cache diverge da do Outlook sem o usuário entender. |

---

## 6. R2-B: por que não é estrutural

A STA única é estrutural. "Ficar horas sem atender" não é. O erro seria
enfileirar "importe 50 mil itens" como **uma** operação.

Cada unidade: obter página limitada · liberar todos os RCWs · devolver
DTOs · **persistir FORA da STA** · reenfileirar a próxima com prioridade
baixa · permitir leitura e mutação entre páginas · cancelar entre lotes ·
gravar progresso só após commit.

**Não manter `Table`, `Items` ou `MAPIFolder` vivos entre turnos** só para
ganhar desempenho. Reabrir por lote custa milissegundos e compra
isolamento.

Prioridade na fila do broker:

1. mutação explícita do usuário
2. leitura interativa
3. atualização da pasta visível
4. importação e reconciliação de fundo

`GetArray(50000)` é estruturalmente ruim — chamada COM não é abortável.
`GetArray(100–500)` provavelmente é administrável, sujeito à Q1.

---

## 7. Desenho de exclusão invisível

**Condicional à Q1.** Se a leitura direta por `Table` for rápida o
bastante, a lista pode continuar vindo do Outlook, e este desenho vale só
para a busca e para o estado local. A seção 2 reabriu essa decisão de
propósito, e esta seção não pode pressupor o resultado.

Assumindo que a lista venha do cache:

- A lista abre na hora, com o cache.
- **Cada LINHA é uma associação (item, pasta)**, e carrega o
  `LastVerifiedAt` e o estado de presença DAQUELA associação — não do
  item.
- Selecionar a pasta dispara reconciliação prioritária **daquela pasta**.
- Até ela terminar, o cache é **snapshot**, não verdade atual.
- Ausência da pasta só é confirmada após geração completa e consistente.
- Ao abrir uma linha ausente, tentar correlação **controlada** para
  detectar movimento — o item pode ter ido para outra pasta, e aí a
  associação nova é criada em vez de o registro morrer.
- Sem correspondência inequívoca: *"não encontrado no Outlook"*, e **não**
  *"excluído"*. O Iris não sabe se foi excluído.
- **Nunca** atribuir o registro a outro item só por Message-ID.
- Frescor no nível da PASTA — "atualizado em…" — e não um aviso por linha.

---

## 8. Critério de pronto do 2.0

1. Q1, Q3, Q4, Q7 e Q8 com **número**; Q2, Q5, Q6 e Q9 com **conclusão
   semântica** — nem toda pergunta se responde com métrica, e prometer
   isso seria repetir o erro da v1.
2. Cada resposta com a limitação escrita: qual pasta, qual store, cached
   ou online, quantas execuções.
3. Recomendação explícita sobre o tamanho de cada marco seguinte.
4. Revisão externa do RESULTADO, não só do plano.
5. Itens de teste movidos devolvidos ao original. Itens usados no teste de
   exclusão dura: **perda aceita e consentida de antemão**.
6. Corpus adversarial da Q2 com **oráculo escrito antes** de rodar — quais
   pares são o mesmo item, decidido por mim e não pela ferramenta.
7. Critério operacional de invalidação de geração concorrente, escolhido
   entre as três opções da Q4 com dado, não por preferência.
8. **Latência máxima por lote**, e não só tamanho 100–500: quantidade não
   garante tempo limitado, e é tempo que trava a fila da STA.
9. Teste de crash entre o commit no SQLite, o avanço do checkpoint e a
   publicação para a UI — os três passos precisam sobreviver a morrer no
   meio.
10. Prova de que uma reconciliação antiga **não sobrescreve** resultado de
    geração mais nova.

---

## 9. Q1 — RESPONDIDA

Medido em 2026-08-23, na Caixa de Entrada real (1.003 itens, Exchange
cached). Somente leitura. Scripts em `tools/q1-*.ps1`.

### O resultado

| Caminho | ms/item | Caixa inteira | Página de 50 |
|---|---|---|---|
| Iteração (`Items.Item(i)`) — o atual | 11–13 | ~12 s | ~570 ms |
| `Table` + cursor | **0,66** | **666 ms** | **~27 ms** |

**~20x, e constante com a profundidade.** As 21 páginas custaram entre 26 e
29 ms cada; a última página sai pelo mesmo preço da primeira.

### As colunas

Todas as do `MailSummary` vêm em lote, e três delas eram as que mais
importavam:

| Coluna | Como |
|---|---|
| EntryID, Subject, MessageClass, LastModificationTime | tabela padrão |
| SenderName, ReceivedTime, Size, UnRead | `Columns.Add` pelo nome |
| **HasAttachment** | `PR_HASATTACH` (0x0E1B000B) — **sem abrir `Attachments`** |
| **SearchKey** | `PR_SEARCH_KEY` (0x300B0102) |
| **InternetMessageId** | `PR_INTERNET_MESSAGE_ID` (0x1035001E) |

As duas últimas mudam a Q2: as evidências de correlação vêm **de graça,
junto com a listagem**, e não custam uma passada extra.

### O que NÃO vem, e um falso positivo que eu produzi

`Permission` — o `IsProtected` do DTO — **não foi obtido em lote**.

E eu quase registrei que sim: testei com o proptag `0x0E01000B`, que é
`PR_DELETE_AFTER_SUBMIT`, não permissão. A coluna foi **aceita** e devolveu
nulo. É a armadilha exata que a revisão do plano tinha previsto — coluna
ausente volta vazia em vez de dar erro — e ela me pegou na primeira
tentativa.

Pior: **esta caixa não tem mensagem protegida** (0 em 30 itens abertos,
nenhuma classe protegida em 400). Então nem dá para validar a hipótese de
derivar proteção de `MessageClass` (`IPM.Note.rpmsg.Message`,
`IPM.Note.SMIME*`). Fica **NÃO VALIDADO**, e é decisão pendente: derivar de
`MessageClass` e aceitar o risco, ou abrir o item.

### Duas armadilhas que perdem mensagem em silêncio

Achadas medindo, e as duas custariam caro em produção.

**1. O filtro DASL de data é UTC; o `ReceivedTime` da tabela é LOCAL.**

Paginar com a hora local no filtro pulava uma janela do tamanho do offset
do fuso em **cada** fronteira. Resultado: **803 de 1.003 itens**, 20%
perdidos — e a paginação **terminava cedo, parecendo ter acabado**.

Medido isoladamente numa fronteira: string local devolveu 938, string UTC
devolveu 953, e a contagem manual dava 953.

**2. `ReceivedTime` não é ordem total.**

Cinco grupos de itens compartilham o mesmo segundo nesta caixa; um filtro
`<` estrito pularia 6 deles. A saída é `<=` com deduplicação por `EntryID`,
aceitando reler alguns — nesta caixa o custo foi zero releituras, porque
nenhum empate caiu numa fronteira de página, mas o mecanismo precisa
existir.

Com as duas corrigidas: **1.003 de 1.003**.

### O que isto decide

**A listagem NÃO precisa de cache para ser rápida.** 27 ms por página é
instantâneo. A decisão que a seção 2 tinha reaberto está respondida:
**listar continua lendo do Outlook**, por `Table` + cursor.

O cache continua necessário para busca, estado local de triagem, frescor e
o que a Fase 4 indexar. O que ele deixa de ser é **acelerador de lista**.

**Consequência para a Fase 1:** `MessagePaging.ReadPage` usa iteração e
paginação por offset. Trocar por `Table` + cursor é uma melhoria de ~20x
num código que já funciona — decisão de quando fazer, não de se fazer.

### Limitações desta medição

- Uma pasta (Caixa de Entrada), um store, Exchange **cached**, uma máquina.
- Medido por PowerShell, não pelo broker: serve para comparar caminhos, não
  como latência ponta a ponta.
- `GetArray` foi usado com páginas de 50. Arrays grandes monopolizando a
  STA continuam não medidos.
- A tabela devolve o que não é `MailItem` — 6 em 400 eram convite ou
  resposta de reunião. Filtrar é responsabilidade de quem converte.

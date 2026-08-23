# Fase 2 — Cache e sincronização

**Versão:** 13 — Q1 fechada (seção 9). **Q2 RESPONDIDA**: leitura na
seção 10, `Move`/`Copy` na 11, causalidade da PCL na 11.6. Resposta
negativa, com o escopo explicitado.

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

**Esta seção está na 4ª versão.** A primeira foi revisada e tinha três
problemas: uma comparação que não era equivalente, uma paginação que ainda
perdia mensagem, e uma conclusão mais larga que a evidência.

### O ganho, com as comparações separadas

| Comparação | Iteração | `Table` | Ganho |
|---|---|---|---|
| **Mesmo trabalho** — as 7 escalares do `MailSummary` | 7,48 | 1,04 | **7,2x** |
| **DTO completo** | 11–13 | 0,42 | **~18x** |

(ms/item, página de 50)

**Proveniência, porque importa:** a linha do "mesmo trabalho" sai de uma
execução única de `tools/q1-justo.ps1`, com os dois lados medidos lado a
lado. A linha do "DTO completo" **não é comparação direta**: os 11–13 vêm
de `tools/medir-pagina.ps1` e os 0,42 de `q1-justo.ps1`, em execuções
diferentes. O ~18x é aproximação sustentada por duas medições diretas, não
por um experimento pareado.

A primeira versão desta tabela dizia "8 escalares" e incluía
`LastModificationTime`, que **não está no `MailSummary`** — eu tinha metido
uma propriedade alheia e chamado o conjunto de "o DTO atual".

O que torna a iteração cara não são as propriedades escalares — é o que
exige tocar em objeto:

| Extra, na iteração | Custo marginal |
|---|---|
| abrir `Attachments` e ler `Count` | 5,25 ms/item |
| ler `Permission` | 3,96 ms/item |
| **na `Table`: 3 colunas a mais** | **~0** |

Somar os componentes daria 27x, mas essa soma infla — cada medição
carrega overhead próprio. **O número defensável é ~18x**, da medição
direta do DTO completo.

**Travessia completa:** 1.003 itens em **~1,0 s** (960–1167 ms em três
execuções), 22 aberturas de cursor.

Esse número é do algoritmo **correto**. As primeiras medições davam
742–864 ms, mas eram da versão que perdia mensagem: a drenagem do grupo
custa uma leitura a mais por página. **Correção custa ~20%**, e é barato.

Os ~12 s do caminho atual são extrapolação de páginas amostradas, não uma
travessia medida.

### As colunas

Todas as do `MailSummary` vêm em lote, **menos uma**:

| Coluna | Como |
|---|---|
| EntryID, Subject, MessageClass, LastModificationTime | tabela padrão |
| SenderName, ReceivedTime, Size, UnRead | `Columns.Add` pelo nome |
| **HasAttachment** | `PR_HASATTACH` (0x0E1B000B) — sem abrir `Attachments` |
| **SearchKey** | `PR_SEARCH_KEY` (0x300B0102) |
| **InternetMessageId** | `PR_INTERNET_MESSAGE_ID` (0x1035001E) |
| **Permission** | **não foi obtido** pelos candidatos testados |

As duas do meio mudam a Q2: as evidências de correlação vêm **de graça,
junto com a listagem**.

### `Permission`: não foi obtido, e eu quase registrei que sim

Testei com o proptag `0x0E01000B`, que é `PR_DELETE_AFTER_SUBMIT`. A coluna
foi **aceita** e devolveu nulo — e meu script marcava como sucesso qualquer
coluna que não lançasse exceção. Falso positivo produzido por mim, na
primeira tentativa, exatamente na armadilha que a revisão do plano tinha
previsto.

Corrigido: o teste agora procura em 40 itens e exige **ao menos um valor
não nulo**. Com isso, `Permission` deixa de aparecer como disponível.
(Corrigi também um segundo erro do mesmo script: o fallback de
`LastModificationTime` apontava para `datereceived`, e teria declarado
sucesso para a propriedade errada.)

**Não foi obtido pelos candidatos testados** — o que é diferente de "não
existe". Quarenta valores nulos provam que nenhum candidato devolveu valor
nesta pasta, não que não haja propriedade de tabela equivalente.

O que sustenta a conclusão prática é outra coisa: `MailItem.Permission` é
abstração de nível OOM, e `MessageClass` (`IPM.Note.rpmsg.*`,
`IPM.Note.SMIME*`) classifica famílias conhecidas sem ser o mesmo conceito
— IRM, S/MIME e rótulos de sensibilidade são coisas distintas.

E **esta caixa não tem mensagem protegida** — 0 em 30 itens abertos,
nenhuma classe protegida em 400 — então a hipótese do `MessageClass` nem dá
para validar aqui.

**Decisão pendente**, com as opções na mesa: tirar `IsProtected` do resumo e
obtê-lo no detalhe; torná-lo tri-state (`Desconhecido`/`Não`/`Sim`); ou
abrir o item quando `MessageClass` levantar suspeita. O que **não** serve é
`False` significando "não medi".

### Três armadilhas que perdem mensagem em silêncio

**1. O filtro DASL de data é UTC; o `ReceivedTime` da tabela é LOCAL.**

Paginar com a hora local pulava uma janela do tamanho do fuso em **cada**
fronteira: **803 de 1.003**, 20% perdidos — e a paginação **terminava cedo,
parecendo ter acabado**. Isolado numa fronteira: string local devolveu 938,
string UTC devolveu 953, e a contagem manual dava 953.

Regra: converter para UTC e formatar com cultura invariável. Vale para
DASL; filtros Jet com `[Colchetes]` seguem outra regra. E conferir o `Kind`
do `DateTime` que o COM devolve, em vez de confiar no padrão da máquina.

**2. `ReceivedTime` não é ordem total — e `<=` com deduplicação NÃO basta.**

Foi o que eu tinha feito, e está errado. Se o grupo empatado for **maior
que a página**, a consulta seguinte pode devolver os mesmos itens, nenhum
ser novo, e a paginação declarar fim.

A saída tem **duas** partes, e faltar uma já perde mensagem:

1. **DRENAR** o resto do grupo do último instante **no mesmo cursor** —
   sem reabrir, então sem filtro envolvido nessa parte;
2. só então reabrir com `<` **estrito**.

**3. A segunda parte era a que faltava.** Eu reabria com `<=` depois de
drenar, e a consulta seguinte recomeçava no mesmo grupo: nada é novo, e a
paginação declara fim com itens mais antigos por ler. Não aparecia na
caixa real porque aqui o maior empate tem 3 itens.

Uma versão intermediária tinha uma variável `inclusivo` que reabria com
`<=`. Era **código morto** — a primeira fronteira é nula e toda drenagem
bem-sucedida deixa a fronteira estrita — e eu descrevia o algoritmo como
sendo de "três partes" por causa dela. Removida: descrição que não
corresponde ao código é pior que ausência de descrição.

E minha primeira drenagem quebrou de outro jeito ainda: marcava como
vistos os itens de FORA do grupo, e a página seguinte os achava repetidos —
**50 de 1.003**. A drenagem tem de parar no primeiro instante diferente
sem consumi-lo.

**O algoritmo mora em `tools/paginacao.ps1`, e o teste roda ELE.** A versão
anterior do teste avançava com `AddSeconds(-1)`, ou seja `< T`, enquanto o
script real reabria com `<=` — os cenários passavam provando um algoritmo
melhor do que o implementado. Terceira vez neste projeto que um teste
promete mais do que verifica.

Os controles negativos são **executáveis**: os dois defeitos entram por
parâmetro (`-SemDrenagem`, `-FronteiraInclusiva`) e o teste roda as três
variantes lado a lado. Antes eu editava o arquivo à mão e anotava o
resultado — número anotado não é regressão verificável.

E o teste **falha** se qualquer um dos dois deixar de perder item — cada
defeito precisa ser discriminado por si. Um guarda que aceitasse "algum
dos dois" deixaria passar o dia em que um deles parasse de controlar, e
metade do controle negativo viraria decoração sem ninguém notar.

Conferido neutralizando o defeito `-SemDrenagem`: o teste falha nomeando
qual controle parou de discriminar.

| Cenário | Sem drenar | Fronteira inclusiva | Correto |
|---|---|---|---|
| empate de 50 (= página) + antigos | — | perde 200 | OK |
| empate de 100 (2x) + antigos | perde 50 | perde 200 | OK |
| empate de 500 (10x) + antigos | **perde 450** | perde 300 | OK |
| tudo no mesmo segundo | **perde 150 de 200** | — | OK |
| empate no FIM da pasta | perde 50 | — | OK |
| empate de 200, **ordem instável** | perde 150 | perde 100 | OK |

O cenário de ordem instável embaralha as linhas dentro do empate a cada
consulta: o OOM não promete ordem estável ali, e um algoritmo que dependa
disso está errado mesmo passando.

Com as três corrigidas, contra a caixa real: **1.003 de 1.003**.

### O que isto decide, e com que alcance

**No cenário medido — Exchange cached, uma pasta de mil itens — o cache
NÃO é necessário como acelerador de listagem.** 27 ms por página é
instantâneo. `Table` + cursor é o candidato padrão para listar.

**Isto não é decisão universal.** Falta medir: modo online, caixa
compartilhada e delegada, PST, arquivo morto online, pastas de 10 mil e 100
mil, Outlook recém-aberto com cache frio, rede lenta ou store
parcialmente sincronizado, e o caminho dentro da STA com o DTO real. Store
remoto continua exigindo validação e política de fallback.

O cache segue necessário para busca, estado local de triagem, frescor e o
que a Fase 4 indexar. O que ele deixa de ser, **neste cenário**, é
acelerador de lista.

**Consequência para a Fase 1:** `MessagePaging.ReadPage` usa iteração e
paginação por offset. Trocar por `Table` + cursor é ~18x num código que já
funciona — e traz junto as três armadilhas acima, que precisam ir com ele.

### Limitações desta medição

- Uma pasta, um store, Exchange **cached**, uma máquina, uma ordenação.
- Medido por PowerShell, não pelo broker: compara caminhos, não é latência
  ponta a ponta.
- Colunas validadas **individualmente** e depois usadas juntas no cursor;
  o script de matriz sozinho não prova compatibilidade conjunta.
- A contagem de referência do teste de fuso vem da mesma tabela — é prova
  interna da semântica do filtro, não referência independente.
- `$total` é lido antes da travessia: a pasta pode mudar durante.
- A drenagem foi provada contra tabela **sintética**, e a `Table` real
  nunca foi exercitada com grupo empatado maior que a página — esta caixa
  tem no máximo 3 no mesmo segundo. O teste prova que o ALGORITMO está
  correto; que a `Table` do Outlook se comporta como a fonte sintética
  modela continua **não validado**.
- Itens com `ReceivedTime` nulo não foram tratados nem testados. Nesta
  pasta não existem; em Rascunhos podem existir, e um filtro de comparação
  os excluiria da paginação inteira.

---

## 10. Q2 — parte de LEITURA respondida

**Esta seção está na 3ª versão.** Duas rodadas de Codex, e as duas
derrubaram conclusões minhas. O que mudou está em 10.9.

Corpus: **2.281 itens em 127 pastas**, um store. Nada foi criado, movido
ou apagado.

Scripts: `tools/q2-chaves.ps1` (presença, unicidade, matriz de colisão),
`q2-pares.ps1` (cada par de colisão, propriedade a propriedade),
`q2-conferir.ps1` (Table x PropertyAccessor, e deriva), `q2-quase.ps1`
(o que heurística funde).

### 10.1 O oráculo, e o que ele NÃO é

> Message-ID **diferente**, os dois não vazios => **mensagens** diferentes.
>
> Message-ID **igual** => mesma mensagem, e **não autoriza unir**.

A segunda metade não estava na 1ª versão: eu contava "mesmo Message-ID"
como acerto. O par 1 abaixo desmente.

E a primeira metade é mais fraca do que eu escrevi. Eu disse que era "o
RFC 5322, não inferência minha". O RFC identifica uma **mensagem**, não um
**item**, e não fala de continuidade sob movimentação. A leitura honesta
da coluna "errados" da tabela 10.6 é:

> pares unidos pela regra cujos Message-ID não vazios **conflitam**.

É evidência negativa forte para correio transportado normal. Não é ground
truth de identidade de item.

### 10.2 Presença e unicidade, corpus inteiro

| Chave | Presente | Distintos | Grupos repetidos |
|---|---|---|---|
| `PR_RECORD_KEY` | **100%** | **2.281 / 2.281** | **0** |
| `PR_CHANGE_KEY` | **100%** | 2.281 / 2.281 | 0 |
| `EntryID` | 100% | 2.281 / 2.281 | 0 |
| `PR_SEARCH_KEY` | **100%** | 2.278 | 3 |
| Message-ID | 86% | 1.964 | 4 |
| `PR_SOURCE_KEY` | **0%** | — | — |

`PR_SOURCE_KEY` (`0x65E00102`), sugerida pelo Codex por ser a identidade
dos mecanismos de sincronização do Exchange: a coluna é **aceita** pela
`Table` e o valor **volta nulo nos 2.281**; pelo `PropertyAccessor` dá
**erro de leitura**. Não está disponível por nenhum dos dois caminhos.

Message-ID falta em **313** itens, **todos `IPM.Note`**. Compromissos
(`IPM.Appointment`, 490) têm. E **30 dos 68 rascunhos têm Message-ID** —
o que derruba minha explicação da 1ª versão ("rascunho não passou por
transporte, então não tem"). Rascunho **pode** ter, e por que uns têm e
outros não continua **não explicado**.

### 10.3 Os quatro pares, e cada um quebra uma regra diferente

| | par 1 | pares 2 e 3 | par 4 |
|---|---|---|---|
| **o que é** | enviado x recebido (auto-endereçado, artefato meu da Fase 0) | conflito de sincronização: Conflitos x Lixo Eletrônico | consistente com duas entregas do mesmo Message-ID |
| Message-ID | **igual** | **igual** | **igual** |
| `PR_SEARCH_KEY` | **igual** | **igual** | **diferente** |
| `ConversationIndex` | **igual** | **igual** | diferente |
| `PR_CLIENT_SUBMIT_TIME` | **igual** | **igual** | diferente (10 s) |
| `PR_MESSAGE_DELIVERY_TIME` | diferente | **igual** | diferente (9 s) |
| cabeçalhos de transporte | vazio x 4.615 | **iguais** | 10.197 x 10.251 |
| `PR_MESSAGE_FLAGS` | 1 x 34 | 0 x 262146 | iguais |
| `Size` | 8.380 x 33.166 | 127.252 x 127.028 | 70.790 x 70.748 |
| `PR_RECORD_KEY` | **diferente** | **diferente** | **diferente** |

- **Par 1** — dois itens que compartilham **toda** evidência derivada de
  conteúdo. `MsgFlags` 1 = `READ`; 34 = `0x22` = `FROMME|UNMODIFIED`. Um
  foi submetido, o outro **entregue**: 4.615 caracteres de cabeçalho que só
  o transporte escreve.
- **Pares 2 e 3** — eu escrevi que eram "duas manifestações do mesmo item"
  e chamei isso do caso positivo que faltava. **Está errado, e ver 10.3.1.**
- **Par 4** — falso positivo do Message-ID que eu não fabriquei: dois itens
  com o mesmo `Message-ID`, 10 segundos de diferença, cabeçalhos de
  transporte de tamanhos diferentes. **Consistente com** duas entregas
  distintas; provar exigiria ler os cabeçalhos `Received`. `SearchKey` os
  separa.

`SearchKey` **muda de comportamento conforme o alvo**, e na 2ª versão eu
chamei os dois casos de "erro", o que troca o alvo no meio do argumento:

- para identidade de **item**, o par 1 é erro (une dois objetos) e o par 4
  está **certo** (separa dois objetos);
- para identidade de **mensagem**, o par 4 é que é erro.

Em nenhum dos dois alvos ela acerta sempre.

#### 10.3.1 Os pares de conflito NÃO são o mesmo item

Retração. Um item de conflito é um **objeto MAPI novo**: quando a versão
local e a do servidor divergem, o Outlook preserva a perdedora como cópia
em `Problemas de Sincronização\Conflitos`. Os dois coexistem, aparecem
separados, e têm `EntryID`, `RecordKey`, `ChangeKey` e estado próprios.
`RecordKey` diferente aí é **evidência de objeto novo**, não deficiência da
chave.

Verifiquei em vez de aceitar, com `tools/q2-conflito.ps1`, e o vínculo está
**provado nos dois sentidos**: o `PR_CONFLICT_ITEMS` de cada um aponta para
o outro, nos 2 pares. O controle negativo ("algum aponta para si mesmo")
dá não, então o teste discrimina.

Mas o que está provado é **linhagem**, não continuidade:

> O próprio provider registra a relação, e ela é uma **aresta tipada entre
> dois objetos** — `VarianteDeConflitoDe` —, não uma identidade
> compartilhada. Unir os dois numa identidade só seria **política do
> Iris**, e a política correta é não unir: o usuário pode abrir e apagar
> cada um.

**Consequência:** o corpus **continua sem um positivo demonstrado** de
continuidade de item. O único jeito de obter um é a observação
longitudinal antes/depois de um `Move`.

`PR_RESOLVE_METHOD` e `PR_SOURCE_KEY` dão **erro de leitura** também pelo
`PropertyAccessor`, não só nulo pela `Table`.

### 10.4 A resposta da Q2

Minha 2ª versão dizia "duas famílias" e concluia que **não existe
identidade de item obtenível pelo OOM**. As duas coisas estão erradas: a
partição misturava origem com função, e a conclusão não decorre dos dados.

**Cinco papeis, não duas famílias:**

| Papel | Propriedades | Única? | Estável sob `Move`? |
|---|---|---|---|
| Identidade do objeto no store | `PR_RECORD_KEY`, `EntryID` | **sim**, 2.281/2.281 | **NÃO** — medido, §11 |
| Linhagem copiada junto | Message-ID, `PR_SEARCH_KEY` | não | **sim** — medido, §11 |
| ↳ mesma família, **não medida** no `Move` | `ConversationIndex` | não | não medido |
| Versão / causalidade | `PR_CHANGE_KEY`, `PR_PREDECESSOR_CHANGE_LIST` | sim, **e irrelevante** | não, **por desenho** |
| Localização | `StoreID`, pasta pai | — | não, por definição |
| Relação entre objetos | `PR_CONFLICT_ITEMS` | — | — |

`PR_CHANGE_KEY` estava na família errada. Ela é um **token de versão**:
muda quando o objeto muda. Ser única em 2.281 de 2.281 é consequência
disso, e não a torna candidata a identidade — pelo contrário, garante que
não é. O que ela serve é para a **Q3**.

E `PR_SEARCH_KEY` não é "derivada do conteúdo": é uma propriedade de
correlação controlada pelo provider, que costuma ser **copiada junto** com
a mensagem. O valor não precisa ser função do conteúdo — e o par 4, com
`SearchKey` diferente para o mesmo Message-ID, mostra que não é.

**A conclusão defensável:**

> Entre as propriedades **preexistentes** medidas, nenhuma foi ainda
> demonstrada **simultaneamente** única no escopo necessário e estável sob
> `Move`.

*A seção 11 fechou isso: nenhuma é, e agora está medido nas duas
direções.*

Não "não existe identidade obtenível pelo OOM". Há pelo menos duas saídas
que eu não tinha considerado:

1. **Combinação ou estratégia com estado** pode funcionar onde nenhuma
   propriedade isolada funciona.
2. **Propriedade nomeada escrita pelo próprio Iris.** O projeto **já faz
   isso** — o marcador `IrisDraft` da Fase 1. Exige escrita, uma cópia
   duplicaria o valor, e a Fase 1 registra que a política do tenant pode
   impedir a gravação; nada disso a elimina como via.

~~E se a `RecordKey` sobreviver ao `Move`, a condição nem se
materializa.~~ **Ela não sobrevive** (§11.1). A condição se materializou.

De qualquer forma, o teste de `Move` é **o experimento decisivo da Q2** —
com `Copy` como **controle negativo**, porque uma chave que sobrevive ao
`Move` mas também é duplicada numa cópia não serve como identidade.

### 10.5 `PR_RECORD_KEY`: eu a descartei errado

A 1ª versão dizia: "os últimos bytes dela são os do `EntryID`, logo é
outro localizador". **Duas coisas erradas.**

Primeiro, é falso. `EntryID` tem 24 bytes, `RecordKey` tem 46, e em
**2.281 de 2.281** nenhuma contém a outra — nem como sufixo, nem em
qualquer posição. Eu tinha olhado um par e comparado de olho.

Segundo, mesmo que compartilhassem bytes, isso não provaria equivalência
semântica.

**O argumento que eu usei também era fraco.** Eu comparei byte a byte as
18 pastas com 5+ itens, vi que nenhum byte era constante-por-pasta, e
concluí que "não há identificador de pasta dentro da `RecordKey`". O
método não sustenta isso: um identificador de pasta pode estar misturado
ao contador nos mesmos bytes, desalinhado em bits, ou passado por hash —
`H(pasta || item)` faria todos os bytes variarem com a pasta participando
inteira. E foram 18 das 127 pastas, escolhidas por terem 5+ itens.

**O argumento bom veio de outro lugar.** O blob de `PR_CONFLICT_ITEMS`
(10.3.1) revela o formato, porque ele carrega dois pares e a `RecordKey`
carrega um:

```
PR_CONFLICT_ITEMS  cabecalho | GUID+contador da PASTA | GUID+contador da MENSAGEM
PR_RECORD_KEY      cabecalho |                          GUID+contador da MENSAGEM
```

`4 + 16 (GUID do store) + 2 + 16 (GUID de réplica) + 8 (contador) = 46`
bytes, e **nenhum par de pasta**. É argumento de **formato**, verificado
contra um valor em que as duas partes aparecem lado a lado — muito melhor
que a constância de bytes.

Ainda assim, isso **não prova estabilidade**, e não é nem "precondição
estrutural" como eu escrevi: uma chave pode conter a pasta e ainda ser
preservada, e pode não conter pasta nenhuma e o provider substituí-la no
`Move` assim mesmo. Só o teste decide. O que está estabelecido é que **o
motivo pelo qual eu a descartei não existia**.

### 10.6 O que heurística funde

Subconjunto de **1.318 itens nas 4 pastas padrão** — esta tabela é a
única que não foi refeita sobre o corpus completo.

| Regra | Grupos | Pares com Message-ID conflitante | Maior grupo |
|---|---|---|---|
| assunto | 184 | **1.398** | 19 |
| assunto + remetente | 158 | **773** | 19 |
| assunto + remetente + tamanho | 3 | **5** | 3 |
| assunto + remetente + instante | 1 | **1** | 2 |
| os quatro juntos | 0 | 0 | — |

Assunto + remetente + tamanho ainda funde 5 pares conflitantes. A linha que
não erra **não une nada**.

Na 1ª versão eu chamei isso de "recall zero". Está errado: sem positivos
rotulados, recall é **indefinido**, não zero.

Na 2ª versão eu emendei com "e o corpus agora tem positivos — os pares de
conflito". **Também errado**, e a §10.3.1 desmente: os pares de conflito
são objetos distintos com vínculo registrado, não continuidade. O corpus
**continua sem positivo rotulado**, e recall continua indefinido.

### 10.7 Confiabilidade da medição

- **Marshaling.** Conferi `SearchKey`, `RecordKey` e Message-ID de
  `Table.GetArray` contra `PropertyAccessor.GetProperty`: **183 valores,
  0 divergências**, em 4 pastas. `ChangeKey` **não** foi conferida pelos
  dois caminhos. Valores que não voltam como `byte[]` ou `String` são
  contados como anomalia, não como presença — nenhuma ocorreu.
- **Deriva.** `Table` não é snapshot. Duas leituras seguidas de Entrada
  (1003) e Excluídos (141): **0 itens surgiram, 0 sumiram, 0 chaves
  mudaram**. Isso **não** demonstra ausência de deriva durante a varredura
  das 127 pastas, que é muito mais longa.
- `q2-pares.ps1` cobre **1.789** itens, não 2.281: 15 pastas de Contatos
  recusam a coluna `SenderName`. Não afeta os grupos de colisão, que saem
  de `q2-chaves.ps1` sobre os 2.281.

### 10.8 NÃO validado

- ~~**Sobrevivência a `Move`**~~ e ~~**`Copy`**~~ — **feitos**, seção 11.
- **Rascunho antes/depois de enviar**, **encaminhar/reenviar** — exigem
  enviar, proibido fora do C2.
- **Entre stores** — um store nesta máquina.
- **Enviado x recebido em mensagem para terceiros.** O par 1 é
  auto-endereçado. Para mensagem normal a cópia do destinatário não está
  nesta caixa: não há colisão possível **nem informação**. As outras 105
  de Enviados não colidirem **não prova nada**.
- **Por que 30 rascunhos têm Message-ID e 38 não.**

### 10.9 Erros meus nestas duas rodadas

1. **Escopo inventado.** Li 4 pastas e escrevi "a caixa". São 127 pastas e
   2.281 itens — eu tinha visto 58% e relatado 100%.
2. **`$grupos.Count` sem `@()`.** Com um grupo só, o PowerShell devolve o
   tamanho do `GroupInfo`. Eu **documentei** esse erro na 1ª versão e
   deixei ele no código que gerou a tabela publicada.
3. **`q2-par.ps1` selecionava por prefixo de assunto**, achou 13 itens,
   guardou 2 num hashtable e sobrescreveu 11 em silêncio. Os 2 que
   sobraram eram, por acaso, os certos.
4. **Discriminador refutado pelos meus próprios dados.** O comentário dizia
   "enviado tem `SubmitTime` e não tem `DeliveryTime`". As duas
   propriedades estão nos **dois** itens do par 1. Quem discrimina é
   `PR_TRANSPORT_MESSAGE_HEADERS` com `PR_MESSAGE_FLAGS`.

Os três scripts da 1ª rodada foram **removidos**, não corrigidos: um
script com escopo errado convida a ser reusado.

Na 3ª rodada, mais três — e os três são **a mesma falha**: eu comparei
coisas de tipos diferentes e li o resultado como achado.

5. **Interpretação além do dado.** Chamei os itens de conflito de "duas
   manifestações do mesmo item" e de "o caso positivo que faltava". São
   objetos novos com vínculo registrado. Eu tinha inventado o positivo que
   o corpus não tem.
6. **`PT_MV_BINARY` marshalado errado.** `PR_CONFLICT_ITEMS` volta como
   array **de** arrays; meu `Hex()` devolveu a string `"System.Byte[]"`, o
   teste de vínculo deu "não", e aquele "não" **parecia resultado**.
7. **Comparei ID de curto prazo com ID de longo prazo**, e depois comparei
   por igualdade um blob que é composto. Duas vezes o teste disse "não
   há vínculo" quando havia. Só apareceu porque eu fui olhar os bytes.

Um teste que devolve "não" por defeito próprio é pior que um que quebra:
o que quebra eu conserto, o que devolve "não" eu **publico**.

### 10.10 Duas armadilhas de PowerShell, para não repetir

**Concatenação dentro de `@(...)` precisa de parênteses.**

```powershell
@("EntryID", $PT + "0x300B0102")      # 3 elementos: "EntryID", "$PT", "0x.."
@("EntryID", ($PT + "0x300B0102"))    # 2, que era o que eu queria
```

Sem os parênteses o `Columns.Add` recebia a URL do prefixo sozinha, e as
**127 pastas** falhavam com *"Value does not fall within the expected
range"* — dentro de um `catch { }` vazio, que devolveu **corpus: 0 itens**
sem uma linha de erro.

**`$pid` é somente-leitura.** É o PID do processo, e o PowerShell é
case-insensitive: `foreach ($pid in ...)` aborta. É a mesma classe de
eclipse que o `CLAUDE.md` documenta em VB — sete ocorrências lá, agora uma
aqui, em outra linguagem.

---

## 11. Q2 — o experimento decisivo. RESPONDIDA.

Autorizado pelo usuário. `tools/q2-move.ps1`.

**Dois itens**, cada um com ida, volta e um `Copy` como controle negativo:

- **A** — o `[IRIS-SPIKE-C]` **recebido**, artefato meu da Fase 0. Mensagem
  entregue pelo servidor de verdade (4.615 caracteres de cabeçalho de
  transporte), e descartável.
- **B** — a mensagem mais antiga do **Lixo Eletrônico**. O usuário
  autorizou uma mensagem real sem indicar qual; escolhi a pasta onde um erro
  custa menos.

### 11.1 O resultado

Idêntico para A e para B:

| Chave | `Move` | `Copy` |
|---|---|---|
| `EntryID` | **muda** | muda |
| `PR_RECORD_KEY` | **MUDA** | muda |
| `PR_SEARCH_KEY` | **sobrevive** | **DUPLICADA** |
| Message-ID | **sobrevive** | **DUPLICADO** |
| `PR_CHANGE_KEY` | muda | muda |
| `PR_PREDECESSOR_CHANGE_LIST` | muda | muda |

> **`PR_RECORD_KEY` não sobrevive a um `Move`.**

A candidata que a seção 10 tinha recuperado — 100% presente, 2.281 valores
distintos em 2.281 itens, sem par de pasta no formato — **cai aqui**. O
argumento de formato estava certo sobre o formato e **não previu o
comportamento**: o provider aloca uma chave nova no `Move` mesmo sem
precisar codificar a pasta. Era exatamente a ressalva que eu tinha
registrado, e era ela que valia.

### 11.2 A resposta da Q2

> **Entre as propriedades medidas, neste provider e neste store, nenhuma
> é ao mesmo tempo única e estável.**
>
> As **únicas** (`EntryID`, `PR_RECORD_KEY`) mudam no `Move`.
> As **estáveis** (`PR_SEARCH_KEY`, Message-ID) são **duplicadas pelo
> `Copy`** e compartilhadas por enviado e recebido (par 1 da §10.3).
>
> Os dois papéis falham, e falham por motivos **opostos**.

O escopo importa e não é detalhe. Para **rejeitar** uma garantia, n=1
basta: se a `RecordKey` mudou uma vez num `Move` que o produto precisa
suportar, ela já não serve como identidade garantida. O n=2 só reduz a
chance de acidente.

Mas "nenhuma propriedade preexistente" seria afirmação de universo aberto,
e eu não medi todas as propriedades, nem todas as classes, nem
combinações com estado. A §10.4 já tinha a formulação certa — "entre as
propriedades **medidas**" — e a §11 tinha voltado a generalizar.

O controle negativo foi o que deu o "opostos". Sem o `Copy`, "a SearchKey
sobreviveu ao Move" pareceria uma resposta positiva.

**NÃO validado:** `AppointmentItem` e `MeetingItem`, rascunhos, `Move`
entre stores, PST/IMAP/caixa compartilhada, modo online, reconstrução de
OST ou perfil, e as operações que **parecem** `Move` mas podem ser
copy+delete — regras, arquivamento, retenção, envio de rascunho.

### 11.3 A consequência que vai além da Q2

A ida e volta **não restaura** o `EntryID` nem a `RecordKey`: voltar para a
pasta de origem produz **mais uma** chave nova. Então:

> **Não existe igualdade de chave única** que reconheça o item depois de um
> `Move`. Reconhecer exige evidência causal ou comparação com estado.

A 1ª redação dizia "indistinguível de apagado + chegou". É forte demais:
um item movido **deixa rastro** — `SearchKey` e Message-ID preservados,
conteúdo, instantes, e o par sumiu-aqui/apareceu-ali próximo no tempo.
O que não existe é a **igualdade** que dispensaria juntar essas pistas.

Resta um limite que nenhuma propriedade resolve: **`Copy` seguido de
exclusão do original é observacionalmente idêntico a `Move`.** O OOM não
carrega a intenção do usuário. Para o cache isso pode ser aceitável —
sobrando um único descendente, tratá-lo como continuidade dá o mesmo
resultado prático.

Isso atinge a §3.4 ("o que uma ausência prova") e a Q5. Uma verificação
que confie em `EntryID` para dizer "sumiu" vai reportar exclusão toda vez
que o usuário arrastar uma mensagem entre pastas.

E confirma a §3.1 por medição, ampliando — com uma correção de termo:
`EntryID` é **localizador** (abre o objeto); `PR_RECORD_KEY` **não abre
nada**, e é identificador da **encarnação atual** do objeto no provider.
Nenhum dos dois é identidade lógica durável, que era o ponto.

**Para a Q5**, o que este resultado impõe: estado transitório
`ausente, aguardando reconciliação`; janela de graça de várias gerações
consistentes antes de qualquer conclusão; busca de candidatos no store
inteiro, não só na pasta; **nunca unir se o original e o candidato
coexistem**; vários candidatos ⇒ registros separados; e nenhuma ausência
vira "excluído" por decisão do OOM.

### 11.4 O que sobra para o 2.1

Nenhuma dessas saídas é gratuita, e a escolha é do 2.1:

1. **Propriedade nomeada escrita pelo Iris.** Eu escrevi "única e estável
   por construção" e contradisse na linha seguinte: **o `Copy` a duplica**,
   que é o defeito da SearchKey. Ela é identidade de **linhagem**, e só
   vira identidade com protocolo: ao ver o mesmo `IrisID` em dois itens
   coexistentes, **bifurcar** — um mantém, o outro ganha ID novo, com
   procedência. Some-se a isso que a Fase 1 registra que a política do
   tenant pode impedir a gravação (débito do marcador `IrisDraft`), que o
   Iris **não intercepta** cópias feitas pelo Outlook, por regras ou por
   outro cliente, e que escrever em item do usuário para poder listá-lo é
   decisão de produto.
2. **Correlação explícita e reversível.** Eu propus
   `(SearchKey, Message-ID, papel)`, e tem furo: **a cópia preserva os
   três**. O `Copy` da §11.1 prova para os dois primeiros, e
   submetido-x-entregue também é copiado. "Papel" resolve o par 1 e só
   ele. Só fica de pé como algoritmo conservador: as chaves geram
   **candidatos, nunca identidade**; exigir **exatamente um**
   desaparecimento e **um** aparecimento compatível; confirmar
   **não coexistência** numa varredura consistente; união **provisória**,
   com procedência e confiança.
3. **Não correlacionar.** Já estava previsto como resultado aceitável.
4. **Identidade local + grafo de relações.** Cada encarnação ganha chave
   do Iris, e `Move`, `Copy`, conflito e duplicata suspeita viram
   **arestas tipadas** — nunca fusão destrutiva. É o que o
   `PR_CONFLICT_ITEMS` da §10.3.1 já faz nativamente para um caso.
5. ~~**Causalidade via `ChangeKey`/PCL.**~~ Medida na §11.6: **não existe.**

E "reversível" precisa ser **estrutural**. Se lido, triado e ações forem
fisicamente fundidos numa linha só, desfazer depois não recupera a
separação: o banco tem de guardar **duas encarnações** e uma aresta de
canonicalização **removível**.

Em qualquer uma, o invariante da §3.1 — **na dúvida, não unir** — sai
reforçado, porque agora se sabe que não existe chave que dispense o
julgamento.

### 11.5 Limpeza, e o que ela custou

A conta não fechou: esperado `Excluídos +2, Lixo inalterado`; observado
`Excluídos +1, Lixo +1`. Percebi pela contagem, não porque o script
avisou.

Eu escrevi na 1ª redação que "o `Delete()` não tirou as cópias da pasta".
**Não tenho evidência disso.** O balanço também é compatível com um
`Delete()` ter funcionado e a outra cópia ter ficado na temporária, sendo
depois roteada pelo `q2-limpar.ps1` — que decidia o destino **pelo prefixo
do assunto**, e não pelo conjunto exato de chaves capturado no
experimento. **A causa continua não determinada**, e afirmar mecanismo sem
evidência foi o erro; o fato verificado é só o desbalanço.

O defeito de fundo é outro e esse está claro: **o script não conferiu a
pós-condição de nenhuma mutação.** Depois de um `Delete()` não
retentável, ele deveria reler e confirmar que a cópia saiu de onde estava
e chegou onde deveria, **antes** de mover o original ou apagar a pasta. O
`q2-causal.ps1` da §11.6 já faz isso.

Reconciliei pela `SearchKey` dos dois alvos (`tools/q2-achar-copias.ps1`),
que localiza original e cópia de uma vez **porque a cópia herda a
SearchKey** — o próprio achado do experimento serviu para limpar o
experimento. `PR_CREATION_TIME` não serviria: o `Copy` preserva o valor do
original, e eu tentei por aí primeiro.

Estado final:

- **Lixo Eletrônico: 172 itens**, o número de antes.
- **`Iris Q2 (temp)`: removida da raiz** — `Delete()` de pasta é soft, e
  ela está dentro de Itens Excluídos, não eliminada.
- Resíduo: **2 cópias em Itens Excluídos**, artefatos meus. Não as apaguei
  de lá: `Delete()` dentro de Itens Excluídos pode ser **permanente**, e
  exclusão permanente sem consentimento explícito está proibida neste
  projeto.

**Coleção COM obsoleta não serve para decidir apagar pasta.** O
`q2-move.ps1` terminou dizendo "itens restantes: 2" logo depois de
reportar os dois itens de volta na origem; na releitura havia **1**. A
`Items` que ele consultou vinha da referência de pasta segurada desde o
início. Foi sorte o script ter errado para o lado seguro e **mantido** a
pasta.

### 11.6 A PCL carrega ancestralidade através do `Move`? Não.

`tools/q2-causal.ps1`. Na §11.1 eu comparei
`PR_PREDECESSOR_CHANGE_LIST` por **igualdade**, vi que mudava e escrevi
"muda". A pergunta estava errada: a PCL é uma **lista de antecessoras**, e
o que importa é **contência** — a PCL de depois contém a `ChangeKey` de
antes? Se contivesse, haveria continuidade causal sem nenhuma chave igual.

| | `Move` | `Copy` | `Move` de volta |
|---|---|---|---|
| PCL depois contém a `ChangeKey` de antes | **não** | não | **não** |
| PCL cresceu | não | não | não |

A PCL é **substituída, não acumulada**: 21 bytes nas quatro leituras,
contendo um único registro, que é a `ChangeKey` **atual** do próprio item.

```
ChangeKey  ff76b71bb72984498978e8b7ac86c493 0000059b
PCL        14 ff76b71bb72984498978e8b7ac86c493 0000059b
                                    apos o Move:  00002029
```

Só o contador final muda; o GUID é estável. Como o corpus tem 2.281
`ChangeKey` distintas, o GUID é da caixa e o contador é que distingue —
mais uma razão para ela não ser identidade.

**Sai a saída 5 da §11.4.** Não há aresta causal a explorar aqui.

**Limitação, e ela é séria:** meu seletor pegou o **primeiro** item com o
prefixo `[IRIS-SPIKE-C]`, e caiu no artefato
`[IRIS-SPIKE-C] rascunho ...`, **não** no `envio` entregue pelo servidor.
Então este resultado é **n=1 sobre um item de procedência incerta**. O
comportamento da PCL num item entregue pelo transporte **não foi medido**.
O achado de que a PCL é substituída e não acumulada é forte; a
generalização para correio recebido não está feita.

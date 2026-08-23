# Fase 2 — Cache e sincronização

**Versão:** 2 — reescrito após avaliação técnica externa da v1.

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
- **Remoção no cache só é confirmada depois de uma varredura completa e
  BEM-SUCEDIDA.**

### 3.4 Varredura por geração (mark-and-sweep)

1. Inicia geração N.
2. Marca cada item visto com N.
3. **Só depois de todas as páginas concluírem com sucesso**, itens não
   vistos viram ausentes.
4. Cancelamento, erro, ou Outlook indisponível no meio: a geração inteira
   **não pode confirmar exclusão nenhuma**.

Sem isto, uma falha no meio apaga metade do cache — o R2-H.

### 3.5 Estado de presença por item

Quatro estados, não dois:

`Presente` · `Não verificado` · `Suspeito de remoção` · `Remoção confirmada`

"Não verificado" é o que impede o cache de afirmar o que não sabe.

### 3.6 Lotes interrompíveis

Importação e reconciliação são **sequências de unidades curtas**, nunca
uma operação longa na fila da STA. Ver seção 6.

### 3.7 Armazenamento

**SQLite**, salvo bloqueio concreto de distribuição ou dependência nativa.
Não é benchmark: para busca textual, transação, migração, índice e
recuperação, é a opção natural. LiteDB adiciona dependência com modelo de
consulta menos adequado; arquivo próprio compraria corrupção, locking e
indexação artesanais sem benefício.

**O que precisa de decisão de verdade, e vale mais que o benchmark:**
criptografia em repouso e política de retenção. O cache é cópia local de
correspondência corporativa — é o R14 do escopo aparecendo aqui.

### 3.8 O cache é sempre reconstruível

Perder o arquivo não pode custar nada que só exista nele, além de estado
local. Reconstruir a partir do Outlook é sempre possível.

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
> item, e com que taxa de falso positivo e falso negativo?**

Experimento central: mover um item de teste entre pastas, registrando
antes e depois `EntryID`, `PR_SEARCH_KEY`, `PR_INTERNET_MESSAGE_ID`,
`PR_RECORD_KEY`. É o **D4 da Fase 0**, sem teste desde então.

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
confirmar remoções?"*

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

### Q8 — Matriz de providers

Um número da Caixa de Entrada em modo cached não vira garantia geral.
Registrar o que foi medido em: Exchange cached, offline, PST. Caixa
compartilhada só se entrar no produto.

### Q9 — Fronteira de retenção

Todas as pastas e todo o histórico, ou uma janela? **Esta decisão pode
reduzir o problema mais que qualquer otimização**, e é do usuário.

---

## 5. Riscos

Reordenados por gravidade. **O primeiro é novo, e é o pior.**

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

## 7. Desenho de exclusão invisível (já decidido)

- A lista abre na hora, com o cache.
- Cada item carrega `LastVerifiedAt` e estado de presença.
- Selecionar a pasta dispara reconciliação prioritária.
- Até ela terminar, o cache é **snapshot**, não verdade atual.
- Ausência só é confirmada após varredura completa.
- Ao abrir item ausente, tentar correlação **controlada** para detectar
  movimento. Sem correspondência inequívoca: *"não está mais disponível no
  Outlook"*.
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

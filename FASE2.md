# Fase 2 — Cache e sincronização

**Versão:** 1 — plano, escrito antes de qualquer código.

---

## 1. Por que esta fase existe

A Fase 1 mediu o que custa ler do Outlook: **~600 ms por página de 50
itens**, dominados pela obtenção e resumo item a item. Abrir a mesma pasta
duas vezes paga o preço duas vezes. Buscar exige varrer.

A Fase 0 mediu o que torna isso difícil de resolver:

- **Não há delta token.** O Object Model não tem "o que mudou desde X".
- **Eventos não recuperam o que aconteceu com o Iris fechado.** Medido no
  critério R8: assinatura cancelada, 3 criações e 1 exclusão, zero eventos
  ao reassinar.
- **`EntryID` muda quando o item se move.** Medido. Ele não serve como
  identidade do cache.

Ou seja: o cache não pode ser um espelho que se sincroniza sozinho. Ele
precisa de identidade própria, de reconciliação que não releia tudo, e de
uma resposta honesta para exclusões que ninguém viu acontecer.

---

## 2. A pergunta que pode encolher esta fase

**Existe `Folder.GetTable()` no Object Model, e ele lê em LOTE.** A Fase 1
mediu iteração item a item — `Items.Item(i)` seguido de nove leituras de
propriedade. Nunca medi `Table`.

Se `Table` for uma ordem de grandeza mais rápido, duas coisas mudam:

1. Os ~600 ms/página podem virar dezenas de milissegundos, e a pressa pelo
   cache diminui muito.
2. A reconciliação — enumerar as chaves de uma pasta inteira para descobrir
   o que sumiu — pode passar de "inviável" a "trivial".

**Não vou desenhar o cache antes de medir isso.** Desenhar em cima de uma
suposição de custo é como este projeto errou a medição de página duas vezes
seguidas, e ali o custo era só de documentação.

É possível que a conclusão do 2.0 seja **"a Fase 2 deveria ser bem menor do
que o escopo previa"**. Se for, é isso que o documento vai dizer.

---

## 3. Marco 2.0 — Spike de medição

Mesmo formato da Fase 0: código descartável, critérios objetivos, e **as
decisões de desenho ficam bloqueadas até os números existirem**.

Roda contra a caixa real, **somente leitura**, exceto onde marcado.

### Q1 — `Table` contra iteração

Mesmo trabalho, mesmas colunas do `MailSummary`, mesma pasta de 1.003
itens. Compara `Folder.GetTable()` + `Table.GetArray()` com o caminho atual.

**Critério:** número de ms/item dos dois caminhos, e se `Table` entrega
todas as colunas de que o DTO precisa. `Table` não expõe tudo que
`MailItem` expõe — descobrir o que falta é parte do resultado.

### Q2 — Chave estável

Candidatas, e o que precisa ser verificado de cada uma:

| Candidata | Pergunta |
|---|---|
| `PR_INTERNET_MESSAGE_ID` | Está presente? Em rascunho também? Sobrevive a mover? |
| `PR_SEARCH_KEY` | Sobrevive a mover dentro do store? E entre stores? |
| `ConversationIndex` | Serve para agrupar, serve para identificar? |

**Critério:** para cada candidata, presença medida numa amostra real, e
comportamento após mover **um item de teste** entre pastas e entre stores.

**Este é o único ponto do 2.0 que ESCREVE** — move um item. Precisa de item
criado para o teste e autorização do usuário, e o item volta para onde
estava.

### Q3 — Checkpoint incremental

`Items.Restrict` com `[LastModificationTime] > X` funciona? É rápido? A
propriedade é confiável — muda em toda alteração, inclusive marcar como
lida?

**Critério:** tempo da consulta restrita numa pasta de 1.003 itens, e se o
conjunto devolvido bate com as alterações feitas.

### Q4 — Custo de enumerar só as chaves

Quanto custa obter chave + `LastModificationTime` de TODOS os itens de uma
pasta, sem montar DTO? É o custo de uma reconciliação completa, e define se
detectar exclusão precisa ser esperto ou pode ser bruto.

**Critério:** ms para 1.003 itens, pelo caminho mais rápido que o Q1
indicar.

### Q5 — O que um movimento realmente faz

Um item movido entre pastas do mesmo store, e — se houver segundo store —
entre stores. Registrar antes e depois: `EntryID`, `PR_SEARCH_KEY`,
`PR_INTERNET_MESSAGE_ID`.

**É o D4 da Fase 0, que segue sem teste desde então.**

### Q6 — Exclusão deixa rastro?

Três casos, com item de teste:

- excluir normal (vai para Itens Excluídos)
- esvaziar Itens Excluídos
- `Shift+Del` (exclusão dura)

**Critério:** em qual deles o Iris consegue perceber, com o programa
FECHADO durante a exclusão, que o item sumiu — e a que custo.

### Q7 — Onde guardar

Não é medição do Outlook, é decisão de dependência: SQLite (via
`Microsoft.Data.Sqlite`), LiteDB, ou arquivo próprio. Critérios: busca
textual, tamanho em disco para ~50 mil itens, e o que acontece se o arquivo
corromper.

**Critério:** uma escolha, com o motivo escrito e o custo de reverter.

---

## 4. O que só é decidido DEPOIS do 2.0

Estas decisões dependem dos números, e listá-las agora é para que ninguém
as tome antes:

- **Identidade do item no cache** — depende de Q2 e Q5.
- **Se a listagem lê do cache ou continua lendo do Outlook** — depende de
  Q1. Se `Table` for rápido o bastante, ler direto pode continuar sendo a
  resposta certa para listar, e o cache existir só para busca.
- **Estratégia de reconciliação** — depende de Q3 e Q4.
- **O que fazer com exclusão invisível** — depende de Q6. Se não houver
  forma barata de detectar, a resposta honesta pode ser "o cache mostra
  itens que já não existem, e a UI descobre ao abrir" — o que precisa ser
  dito ao usuário, não escondido.

---

## 5. Marcos previstos (esboço, sujeito ao 2.0)

- **2.1** Identidade e armazenamento.
- **2.2** Importação inicial paginada, com retomada após falha. Importar
  50 mil itens não pode exigir que nada dê errado no meio.
- **2.3** Reconciliação incremental.
- **2.4** Exclusões.
- **2.5** Busca textual.
- **2.6** A listagem passa a ler do cache — **se** o 2.0 disser que vale.

---

## 6. Riscos que já dá para nomear

| ID | Risco |
|---|---|
| R2-A | Cache diverge do Outlook e o usuário confia no cache. Um e-mail que sumiu continuar aparecendo é pior que lentidão. |
| R2-B | Importação inicial de caixa grande leva horas e trava a fila da STA, deixando o Iris inútil enquanto roda. |
| R2-C | Não existe chave estável boa, e todo movimento vira "apagou e criou" — perdendo estado local (lido, marcado, o que a IA já resumiu). |
| R2-D | Exclusão com o Iris fechado é indetectável a custo aceitável. |
| R2-E | O arquivo de cache corrompe e o usuário perde estado local. Precisa ser reconstruível a partir do Outlook, sempre. |
| R2-F | Busca do cache diverge da busca do Outlook e o usuário não entende por quê. |

---

## 7. Critério de pronto do 2.0

1. Q1 a Q7 respondidos com número, não com impressão.
2. Cada resposta com a limitação escrita: qual pasta, qual store, cached ou
   online, quantas execuções.
3. Uma recomendação explícita sobre o tamanho da Fase 2 — inclusive a
   possibilidade de recomendar que ela encolha.
4. Revisão externa do RESULTADO, não só do plano.
5. O item de teste movido/excluído devolvido ao estado original, e a caixa
   do usuário sem resíduo.

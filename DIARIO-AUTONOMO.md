# Diário do trabalho autônomo — a partir de 28/08/2026

O usuário autorizou executar as fases restantes sem ele na máquina, validando os
pontos críticos com o Codex e anotando os bloqueios. Este arquivo é o registro
corrido; o relatório final sai dele.

## Limites que eu declarei e não cruzo

1. **Nenhuma cerimônia é executada por mim.** Nem `--autorizar` do ambiente, nem
   alargamento da lista de pastas da IA, nem ACL. Construo e deixo sem apertar.
2. **Nada é enviado por e-mail**, em fase nenhuma.
3. **Nenhum conteúdo real vai para provedor externo.** O que precisar de
   provedor é provado contra `127.0.0.1`.
4. **Mutação no Outlook do usuário só em pasta de teste criada por mim**, nunca
   na Caixa de Entrada e nunca em item dele.

Cruzar qualquer um destes seria decidir no lugar dele enquanto ele não está.

## Ordem de execução escolhida

As dívidas primeiro, porque são pré-condição escrita e porque fecham a Fase 2 de
verdade; depois as medições, que destravam a pré-condição da Fase 4; depois as
fases novas.

| # | Bloco | Estado |
|---|---|---|
| 1 | Nove caminhos de guarda sem teste (4 categorias) | em andamento |
| 2 | Medir efeito da janela de sincronização | a fazer |
| 3 | Medir qualidade e utilidade do acervo parcial | a fazer |
| 4 | Falha rara da suíte: explicar ou encerrar | a fazer |
| 5 | Dívidas: broker/fechamento, auxiliar `Unico` | a fazer |
| 6 | Fase 4 — triagem e busca semântica, fechada | a fazer |
| 7 | Fase 5 — tarefas | a fazer |
| 8 | Fase 6 — calendário | a fazer |
| 9 | Fase 7 — contatos | a fazer |

---

## Registro

### 28/08 — bloco 1: guardas sem teste

**Categoria 1 — descarte do assistente com transmissão em voo.** Fechada. A
condição era um transmissor que ignore o token, e ela existe agora
(`ProvedorControlado.IgnorarCancelamento`). Três testes, controle negativo dos
dois lados. O número mediu diferente do que eu supus: desligar as duas guardas
do `Finally` produz **7** notificações, não 2 — escrever `Ocupado` reavalia a
cadeia inteira de comandos.

**Categoria 2 — troca de sessão durante a expansão da árvore.** Fechada.
`FakeBroker` ganhou stores e filhas configuráveis com travas; nasceu
`ArvoreDePastasTests`. Achado: os `Atual(geracao)` são **correntes de
conferências independentemente suficientes** — remover uma passa, remover todas
falha. Os testes provam a propriedade da cadeia, e está escrito assim.

**Categoria 3 — guardas de descarte da janela principal.** Dois de quatro
caminhos fechados na primeira passada; o terceiro fechou depois da revisão. A razão real de não haver primeiro teste não era a receita:
`MainViewModel` **não podia ser construído** numa suíte, porque o construtor
abria o cache do usuário. O caminho virou parâmetro opcional; produção continua
sem escolher.

- abertura pendente → **coberto**, controle negativo confirmado
- restauração de pasta vencida por sessão nova → **coberto**, controle negativo
  confirmado
- recarga da árvore concluindo depois do fechamento → **cerca de regressão, não
  cobertura**. Medi: removendo o `Folders.Clear()` do `Dispose` o teste continua
  verde, e o estado do broker é idêntico. Quem segura é a cadeia do
  `FolderTreeViewModel`, já coberta em outro arquivo. Vai para o relatório.
- recarga de contas pendente → **não isolável** pelo mesmo motivo: depois do
  `Dispose` a seleção já é `Nothing`, então o `Apontar` não é chamado com ou sem
  guarda. Vai para o relatório.

Suíte: **816**, 0 falhas, 0 pulados.

**Categoria 4 — guardas do leitor de mensagem.** Fechada. Era a única linha que
pedia arquivo novo, e era literal: `MessageDetailViewModel` nunca tinha sido
instanciado por um teste. `LeitorDeMensagemTests` nasceu com cinco testes, dois
deles controles positivos.

- salvar anexo durante troca de mensagem → **coberto**
- salvar anexo concluindo depois do descarte → **coberto** (as duas guardas do
  `SalvarAnexoAsync` desligadas derrubam os dois)
- marcar como lida falhando tarde → **coberto**, e medido nos três estados: cada
  conferência sozinha é suficiente, as duas juntas é o que o teste prova

Erro meu, e é o do `CLAUDE.md` na letra: os locais `leitor` e `linha` eclipsaram
as funções `Leitor()` e `Linha()` do próprio arquivo. A mensagem foi
"tipo não pode ser inferido", que não aponta para nada. Renomeadas para
`AbrirLeitor()` e `LinhaDe()`.

E um segundo, meu e de ferramenta: um `re.sub` com `"\1"` fora de raw string
escreveu o **byte 0x01** dentro do arquivo VB. O compilador disse "Character is
not valid", que dessa vez apontava para o lugar certo.

Suíte: **821**, 0 falhas, 0 pulados. **Bloco 1 encerrado**: das quatro
categorias, três estão cobertas e uma tem dois caminhos que não isolei.

### 28/08 — revisão do bloco 1 com o Codex, e o que ela derrubou

Sete achados. **Dois testes meus passavam com a guarda que eles nomeiam
removida** — exatamente o defeito que este projeto persegue.

1. `Quem_chega_no_meio_da_carga_ESPERA_o_mesmo_voo` só olhava a segunda tarefa
   **depois** de liberar a trava. Com o defeito antigo presente ele ficava verde.
   Passou a cobrar `segundo.IsCompleted = False` **com a trava ainda fechada** —
   é isso que separa "esperou" de "desistiu".
2. `Descartar_duas_vezes_e_inocuo` passava com o `If _descartado Then Return`
   removido, porque a segunda chamada não tem efeito observável de fora.
   **Apagado**, com o motivo no lugar dele.
3. A "cerca" da recarga da árvore não iniciava recarga nenhuma: a trava
   interceptava a leitura de stores do `ConnectionViewModel`, e o `Dispose` caía
   antes de `Folders.ReloadAsync` existir. Reescrito com a receita do Codex —
   abrir a janela inteira, **depois** travar, disparar a recarga por nome e
   guardar a tarefa. Agora o controle negativo dispara, e **o quarto caminho da
   categoria 3 está coberto de verdade**.
4. Dois testes esperavam relógio. `MarkReadAsync` ganhou marco de conclusão
   (`MarkRead-fim`); a restauração de sessão passou a esperar a árvore da E3.
5. "Transmissão em voo" usava `vm.Ocupado` como marco — e `Ocupado` sobe antes
   de o `Task.Run` agendar o provedor. Passou a usar `AoEnviar`, que roda dentro
   da chamada.
6. `FakeBroker`: ligar uma capacidade estava virando "aceita tudo". `Detalhes`
   passou a ser indexado pela chave completa (`StoreId` **e** `EntryId`), e
   salvar anexo / marcar como lida rejeitam item desconhecido, destino vazio e
   `isRead:=False` — o Iris nunca desmarca.
7. Contradição no próprio diário, corrigida.

Codex confirmou o que passou: os **7** avisos, as cadeias `Atual(geracao)`, os
três estados da marcação, e que o `Optional caminhoDoCache` é aceitável porque
só encaminha capacidade que o `AcervoViewModel` já tinha.

Resta **um** caminho não coberto: a recarga de stores pendente. Codex concorda
que não é observável por API pública — só por reflexão ou por um *seam* interno.
Vai para o relatório como decisão, não como esquecimento.

Suíte: **820** (um teste a menos, e o que saiu não provava nada).

### 28/08 — bloco 2: o efeito da janela de sincronização, MEDIDO

Dívida aberta desde a Fase 2. O ESCOPO dizia a saída: *"não é achar a
configuração; é medir o efeito dela"*. Ferramenta nova: `tools/medir-janela.ps1`,
somente leitura, sem abrir corpo nem anexo.

**A medição, 28/08/2026, caixa corporativa real:**

| Pasta | Itens | Mais antigo | Mais novo | Span |
|---|---|---|---|---|
| Caixa de Entrada | 1.098 | 2026-07-28 | 2026-08-28 | 31 d |
| Itens Enviados | 119 | 2026-07-28 | 2026-08-28 | 31 d |
| Itens Excluídos | 98 | 2026-07-28 | 2026-08-27 | 30 d |
| Spam | 22 | 2026-07-28 | 2026-08-28 | 31 d |
| Caixa de Entrada\1. Backup | 37 | 2026-07-28 | 2026-08-28 | 31 d |
| Lixo Eletrônico | 178 | 2026-07-29 | 2026-08-28 | 30 d |
| Rascunhos | 68 | 2026-07-31 | 2026-08-22 | 22 d |
| Problemas de Sincronização | 31 | 2026-08-22 | 2026-08-23 | 1 d |

**Cinco pastas cortam no mesmo dia: 28/07/2026.** Enviados, Excluídos, Spam,
uma subpasta manual de arquivamento e a Caixa de Entrada — usos completamente
diferentes, volumes de 22 a 1.098, e o mesmo primeiro item. Isso não é hábito de
arquivamento. É o horizonte do store.

**O efeito da janela é uma janela deslizante de ~31 dias.** Rascunhos foge
porque rascunho é local; Problemas de Sincronização foge porque a pasta é
recriada.

**O que isto muda, e é material:** "cobertura parcial" deixou de ser ressalva
formal. O acervo do Iris nunca conterá mais de ~1 mês da caixa. Para a Fase 4
isso é decisivo — busca semântica sobre uma janela de 31 dias é um produto
diferente de busca semântica sobre um arquivo histórico.

**O que continua não dito:** quantos itens existem no servidor além do
horizonte. Isso segue inalcançável pelo OOM, e é por isso que o Iris não conclui
ausência.

### 28/08 — bloco 3: qualidade e utilidade do acervo parcial

Pré-condição escrita da Fase 4. Medido sobre as 1.123 linhas da Caixa de Entrada
no cache.

**Completude:** `subject`, `sender_name`, `received_at`, `size_bytes` — **0%
vazios**. `last_modified_at` — **100% vazio**. `internet_message_id` e
`search_key` — **100% vazios**, e isso é *deliberado e documentado*: a Q1 mediu
que nenhum dos dois vem por coluna de `Table`, e preenchê-los custaria abrir o
`MailItem` de cada mensagem.

**Identidade:** 1.123 incarnations, 1.123 items, 1.123 entry_ids distintos. Sem
duplicata.

**Material para triagem e busca:** 338 com anexo, 235 não lidas, 138 remetentes
distintos, 719 assuntos distintos em 1.123 (**64% únicos**). Os cinco assuntos
mais repetidos aparecem 19, 18, 16, 12 e 11 vezes, todos com prefixo `RES:`/`RE:`
— o acervo é **pesado em conversas**, e as conversas são identificáveis pelo
assunto. Um remetente sozinho responde por 252 das 1.123 (22%).

**Defeito encontrado pela medição, e corrigido.** `message_class` tinha **uma**
classe distinta em 1.123 linhas. Número limpo demais para ser medida — e não era:
`OutlookSweepSource.Traduzir` descartava o valor que o broker havia lido da
coluna de `Table` e gravava a constante `"IPM.Note"`.

O efeito era invisível porque o filtro da paginação só deixa passar linha que
*começa* com `IPM.Note`. Mas `IPM.Note` é prefixo, não valor: `IPM.Note.SMIME` e
`IPM.Note.Microsoft.Conversation` passam pelo filtro e viravam a mesma constante.
O cache afirmava uma classificação que ninguém mediu naquela linha — o mesmo erro
que o comentário de `MailSummary` condena para `IsProtected`.

Corrigido, com teste e controle negativo. `DestinoFalso` passou a guardar as
**linhas** e não só as chaves — guardar só a chave é o motivo de a constante ter
sobrevivido à suíte inteira.

E a armadilha do `CLAUDE.md` cobrou pedágio pela terceira vez nesta sessão: o
campo novo não pode se chamar `Linhas`, porque o parâmetro de `GravarPagina` se
chama `linhas` e o eclipsaria dentro do próprio método que precisa escrever nele.

Suíte: **821**.

### 28/08 — blocos 4 e 5: as três dívidas restantes

**Falha rara da suíte — fechada por construção, não por explicação.** Trinta
execuções limpas não provam correção; ausência de sintoma nunca provou nada neste
projeto. O que dá para fazer é transformar a hipótese em **regra imposta**:
`ParalelismoDaSuiteTests` lê os arquivos da própria suíte e exige
`<DoNotParallelize>` em toda classe que toca SQLite. Controle positivo (≥10
classes conferidas, ≥20 arquivos lidos) e controle negativo confirmado — tirando
o atributo de `CacheDatabaseTests`, o teste falha nomeando o arquivo.

O irmão cobra que o assembly **continue** paralelizando o resto: desligar o
paralelismo global satisfaria a regra da maneira preguiçosa e custaria minutos
por execução.

A dívida muda de forma: deixa de ser "explicar a falha" e passa a ser "a causa
provável está fechada por construção; se ela voltar, a hipótese estava errada".
É menos que uma explicação, e está escrito como sendo menos.

**Auxiliar `Unico` — era pior do que a dívida descrevia.** Ele aparecia nas
**duas** posições do construtor de `SchemaTable`, e cinco chamadas estavam na
posição *não* única. Entre elas `Unico("incarnation_key")` em
`metadata_observation` — coluna que se repete de propósito, uma observação por
geração. Quem lesse aquela linha concluiria o contrário do que o esquema faz.
Renomeado para `Indice`. Não pôde ser `Colunas`: o parâmetro se chama `colunas` e
em VB isso colide com a função que o define.

**Coordenação broker/fechamento — corrida real, corrigida.** A liberação roda em
`DispatcherPriority.Send`, que **fura a fila**: uma leitura já enfileirada pela
janela, ainda não executada, corria depois do `ReleaseSessionCore` e tocava RCW
já liberado. As guardas de `_disposed` não cobrem isso — elas garantem que o
*resultado* é ignorado, não que a *chamada* não acontece.

`Shutdown` passou a drenar a fila primeiro, com espera vazia em
`ApplicationIdle` e metade do orçamento de tempo. Se trabalho novo continuar
chegando, desiste e libera assim mesmo: `OUTLOOK.EXE` órfão é pior que RCW tocado
tarde (R7).

**Não verificado contra o Outlook real** — exercitar exige fechar o Outlook do
usuário, e ele não está na máquina. Vai para o relatório como pendência de
medição, não de raciocínio.

Suíte: **823**. **Blocos 1 a 5 encerrados.**

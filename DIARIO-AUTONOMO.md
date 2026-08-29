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
| 1 | Nove caminhos de guarda sem teste (4 categorias) | **fechado**, 8 de 9 cobertos |
| 2 | Medir efeito da janela de sincronização | **fechado** |
| 3 | Medir qualidade e utilidade do acervo parcial | **fechado** |
| 4 | Falha rara da suíte: explicar ou encerrar | **fechado por construção** |
| 5 | Dívidas: broker/fechamento, auxiliar `Unico` | **fechado** |
| 6 | Fase 4 — triagem e busca semântica, fechada | a fazer |
| 7 | Fase 5 — tarefas | a fazer |
| 8 | Fase 6 — calendário | a fazer |
| 9 | Fase 7 — contatos | a fazer |

Esta tabela é atualizada a cada bloco. Ela ficou desatualizada uma vez, e o
Codex pegou — um diário que se contradiz é pior que diário nenhum.

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

### 28/08 — revisão dos blocos 2–5 com o Codex

Sete achados; três de gravidade alta, e o primeiro é o meu.

1. **O dreno do `Shutdown` não fechava a corrida — só estreitava.** Codex
   descreveu a intercalação: uma thread lê `_disposed = False`, é interrompida
   antes de enfileirar, o encerramento inteiro acontece, e só então ela enfileira
   sobre RCWs já liberados. Corrigido de verdade: `InvokeAsync` passou a
   **conferir e enfileirar sob o mesmo bloqueio**, e `Shutdown` fecha esse portão
   antes de drenar. Fechado o portão, a fila é finita e o dreno tem fim
   garantido. Bloqueio curto e sem risco de impasse — protege um `InvokeAsync`,
   que enfileira e volta, nunca um `Invoke`.

2. **Uma exceção no descarte da janela pulava o do broker.** Estavam no mesmo
   `Try` em `Application_Exit`. Separados, cada um com log próprio: o descarte do
   broker é justamente a proteção contra `OUTLOOK.EXE` órfão, e não pode depender
   de o outro ter corrido bem.

3. **A regra do paralelismo tinha cinco furos**, e Codex listou todos: não
   percorria subpastas; comparava texto sensível a maiúsculas; contava ocorrência
   dentro de comentário; creditava `<DoNotParallelize>` de uma classe a todas as
   do arquivo; e o teste irmão procurava a substring `"Parallelize"` — que
   `DoNotParallelize` também contém, então `<Assembly: DoNotParallelize>`
   passaria. Reescrita: por classe, sem comentários, com subpastas, e casando o
   atributo do assembly por expressão ancorada. **Dois controles negativos
   confirmados.**

   E as marcas estavam largas: cobravam `ArchitectureTests`,
   `BindingsDaJanelaTests` e `ContextoDoOutlookTests`, que só citam os tipos por
   reflexão ou leem o fonte como texto. Passaram a casar **construção**, não
   menção. O erro que sobra é o oposto — uma classe que chegue ao banco por um
   auxiliar novo escapa — e quem alarma é o controle positivo.

4. **"Janela deslizante" era forte demais para uma medição só.** Uma execução é
   um retrato: mostra horizonte comum hoje, não que ele anda. Migração, recriação
   do OST, retenção corporativa ou uma operação em massa naquela data produziriam
   o mesmo. O roteiro passou a dizer isso, e a dizer o que separa — rodar de novo
   daqui a alguns dias e ver se o corte anda junto.

5. **Achado novo, da mesma família do `message_class`, e NÃO corrigido.** O
   caminho de paginação transforma ausência e falha de conversão em fato:
   `Size → 0`, `UnRead → False`, anexo → `False`, texto → `""`. É preexistente e
   chega ao cache como se fosse medição. Corrigir exige tornar os campos anuláveis
   e mexer no esquema — decisão de tamanho, não conserto. **Vai para o
   relatório.**

6. **O roteiro pregava R7 no cabeçalho e a violava na travessia:** não liberava
   `$f`, `$ns` nem `$ol`. Corrigido, em ordem inversa à aquisição.

7. Tabela do diário desatualizada e um `<summary>` órfão no `OutlookBroker` — o
   texto da mutação estava empilhado sobre `SemRetryAsync`, documentando
   justamente a função que **não** é mutação. Remanejado para `MutateAsync`.

Suíte: **823**.

### 28/08 — bloco 6: a Fase 4 não vai ser executada, e o motivo é bom

Levei as oito decisões ao Codex antes de escrever qualquer linha. Ele desmontou
o plano, com razão, e o achado mais importante não era sobre a Fase 4.

**O ESCOPO afirmava duas coisas falsas sobre a Fase 2, e elas estavam lá desde
25/08:**

1. *"busca textual"* como entregue. **Não foi.** Não existe consulta textual
   sobre o cache em lugar nenhum: nem esquema, nem serviço, nem tela. O
   `ManifestReader` lê o manifesto de uma pasta; ele não procura nada.
2. *"listagem e busca passam a ler exclusivamente do cache"*. **A listagem não
   lê do cache** — continua lendo ao vivo do Outlook, e isso é deliberado, está
   escrito no `MainViewModel`, e é a §23. O documento afirmava o contrário de
   uma decisão que o projeto tomou de propósito.

A segunda é pior. A primeira é entrega que faltou; a segunda descreve uma
arquitetura que o projeto **decidiu não ter**.

**Sobre as oito decisões, o veredito do Codex:**

| # | Meu voto | O que ele achou |
|---|---|---|
| 1 | cobertura parcial | quase certa — mas a ressalva não pode viver só na UI; tem de ser estrutural, presa a geração/universo/instante |
| 2 | zero egress na v1 | **certa**, e proporcional ao corpus |
| 3 | índice local | incompleta — "local" não resolve: criptografia do cache (R14) segue aberta |
| 4 | cerimônia separada | direção certa, **prova errada**: `127.0.0.1` prova protocolo, não impossibilidade de egress |
| 5 | retenção | **errada** — o Iris não sabe que a mensagem "morreu"; sair da janela, mover e excluir são indistinguíveis |
| 6 | versão do índice | parcial — falta origem, normalização, métrica; e reconstrução tem de trocar ponteiro, não apagar antes |
| 7 | qualidade | **errada como critério** — precisão sozinha premia quem marca quase nada |
| 8 | orçamento | **não decidida** — "teto por execução" não diz qual teto |

**E o que ele disse que eu precisava ouvir:** minhas regras de triagem produzem
*sinais sem direção validada*. `RE:`/`RES:` identifica sintaxe de assunto, não
necessidade de ação. Remetente com 22% do volume pode ser automação. Anexo pode
ser contrato ou assinatura de e-mail. **Nenhuma dessas variáveis tem direção
segura sem rótulos, e só o dono da caixa pode rotular.** Eu posso definir o
protocolo; não posso fabricar o oráculo.

**Decisão: a Fase 4 NÃO é executada.** Chamar isso de Fase 4 seria trocar a
entrega pelo andaime. E a Fase 4 como está escrita provavelmente não faz sentido
agora: 1.123 itens, sem corpo indexável, janela de um mês, sem conjunto de
consultas, e sem nem a busca textual de baseline.

**Correção importante do Codex, que eu tinha errado:** a janela de 31 dias é do
**Outlook**, não do cache do Iris. O cache foi desenhado para acumular — ele pode
crescer para além da janela daqui para frente. A Fase 4 pode ganhar sentido
depois de meses de acúmulo. Não gravar "31 dias" em lugar nenhum.

**O que passa a ser entregue no lugar:** a lacuna real — busca textual sobre o
acervo — e facetas locais nomeadas como **sinais**, não como prioridade, em
tabela derivada e descartável. Nunca em `user_state.triaged`, que é decisão
durável do usuário e o esquema protege de propósito.

### 28/08 — bloco 6 (continuação): a busca ligada à janela

`BuscaNoAcervo` sem tela seria o **sétimo** caso do erro mais comum meu neste
projeto: proteção que existe e não está ligada a nada. Ligada:

- `AcervoViewModel.Procurar` é a porta — a §26.2 e a `ArchitectureTests` proíbem
  a apresentação de instanciar leitor de cache, e quem já tem o banco aberto é o
  acervo.
- `BuscaViewModel` recebe uma **função**, não o banco. Sem isso não haveria
  teste: o construtor abriria SQLite.
- A faixa da busca fica junto do **acervo**, não da lista. A lista lê ao vivo do
  Outlook; a busca lê o cache. Duas fontes diferentes com a mesma cara enganam em
  silêncio.
- `Busca.` entrou nas raízes de `BindingsDaJanelaTests` — sem isso os bindings
  novos não seriam conferidos por ninguém.

Oito testes de tela, com dois controles negativos medidos: fazer a ressalva
aparecer só no resultado vazio derruba um; tratar falha de banco como "não achei"
derruba outro.

O que a tela cobra e que não é óbvio:

- **A ressalva não some quando a busca acha.** Ressalva que só aparece no vazio
  ensina a lê-la como "não achei", e ela diz outra coisa.
- **"Ainda não procurei" ≠ "procurei e não achei".** É o mesmo erro que a lista
  de mensagens já teve, quando "selecione uma pasta" e "esta pasta está vazia"
  eram a mesma frase.
- **Falha de banco não vira "não achei".** É a §23 na forma mais fácil de
  cometer.

E a armadilha do `CLAUDE.md` pela quarta vez: o local `quando` eclipsou a
propriedade `Quando`, e a mensagem foi "String não converte para DateTimeOffset"
numa linha que não tem String nenhuma.

### 28/08 — verificação contra a máquina real

Duas coisas que eu tinha deixado como "não verificado" passaram a ser medidas.

**1. O encerramento novo do broker.** Abri o Iris, ele conectou (época 1, stores
e pastas lidos, sem erro no log), e fechei **pela janela** — nunca por
`Stop-Process`, que pularia justamente o caminho que eu queria exercitar.
Encerrou em **0,5 s**, com `broker.shutdown ok`: sem aviso de fila que não
drenou, sem falha de limpeza. O `OUTLOOK.EXE` do usuário continuou com o **mesmo
PID**. O comentário do código deixou de dizer "não verificado".

Isso prova que o caminho feliz não regrediu e não deixa órfão. **Não** prova a
corrida que o portão fecha — ela é rara por definição, e um encerramento limpo
não a exercita.

**2. A busca, contra o acervo real.** `Iris.CrashHarness buscar <termo>`, modo
novo e somente leitura, rodado sobre uma **cópia** do cache do usuário:

| termo | achados | tempo |
|---|---|---|
| `regulatorio` (sem acento) | **305** | 22,7 ms |
| `brainmetyl` | 26 | 26,3 ms |
| `aquaba tx` (duas palavras) | 13 | 24,1 ms |
| `palavraquenaoexisteemlugarnenhum` | 0 | 24,8 ms |

Sobre **1.127 mensagens**. Três coisas ficam provadas que o teste sintético não
alcançava:

- a normalização funciona no corpus de verdade — `regulatorio` acha
  `Regulatório - Kate`, que é o remetente com 22% do volume;
- a conjunção funciona sobre vocabulário real: `aquaba tx` acha 13 e não 300;
- **22 a 26 ms** é o número que sustenta a decisão de casar em memória. O gatilho
  para trocar por FTS5 deixou de ser uma frase e passou a ter uma medida.

Não consegui verificar a faixa da busca **na tela**: computer-use exige a
aprovação do dono no monitor, e ele não está. O que dá para dizer é que o
aplicativo abriu sem erro de binding no log e que `BindingsDaJanelaTests` passou
a conferir a raiz `Busca.` no XAML de verdade. Ver a faixa desenhada fica para
ele — está no relatório.

### 28/08 — blocos 7, 8 e 9: medir antes de escrever as Fases 5, 6 e 7

As três fases estão no ESCOPO com **uma frase cada**. Planejá-las sem medir
repetiria o erro que a Fase 0 existiu para evitar — a v1 do ESCOPO afirmava que
WPF em VB não era suportado no .NET moderno, e a premissa era falsa.

`tools/medir-grupos.ps1`, somente leitura, contra a caixa real:

| Grupo | Fase | Itens | Custo por item | Propriedades legíveis |
|---|---|---|---|---|
| **Calendário** | 6 | **434** | 30,9 ms | 10 de 10 |
| **Contatos** | 7 | **0** | — | — |
| **Tarefas** | 5 | **3** | 83,3 ms | 9 de 9 |

**Três achados que reescrevem as três fases.**

**Fase 7 (Contatos) não tem consumidor.** A pasta de Contatos está **vazia**. Numa
conta corporativa os contatos vivem no GAL — e o GAL está explicitamente **fora
de escopo** na §8 do ESCOPO ("Paridade com To Do, Planner ou GAL"). Um módulo de
contatos aqui seria uma tela para uma pasta com zero itens.

**Fase 5 (Tarefas) é uma feature de escrita, não de leitura.** Três itens. O
valor não está em ler as três tarefas que existem, está em **criar** tarefas a
partir de e-mails — que é o que o ESCOPO descreve: *a IA sugere, você confirma, o
Iris cria o `TaskItem`*. Escrever na caixa do dono é exatamente o que eu declarei
que não faço com ele ausente.

**Fase 6 (Calendário) é a que tem substância.** 434 itens, todas as dez
propriedades legíveis em 100 de 100, e **30,9 ms por item** — quase o dobro dos
~16 ms que a Fase 0 mediu por mensagem e que tornou o cache obrigatório. Ler o
calendário inteiro uma vez custaria ~13 s. A conclusão da Fase 2 se aplica antes
de a Fase 6 começar.

**4 recorrentes em 100** (os mais recentes por `[Start]`). A primeira versão
amostrava os "30 primeiros" sem ordenar, e `Items.Item(i)` sem `Sort` devolve a
ordem interna do provedor — relatar "0 em 30" a partir daquilo seria dar
aparência de medida a um acaso.

**O que a medição não mediu:** se dá para **criar**. Saber isso exige criar, e
criar é mutação na caixa do dono.

### 28/08 — Fase 6: o calendário, lido e na janela

A única das três fases restantes com substância **e** que é só leitura.

**A leitura.** `CalendarReading` + `CalendarFilter` + `GetAppointmentsAsync`,
por **janela de datas** e não por página: ocorrência de série não existe até ser
expandida, e expandir sem data-fim é infinito.

**O controle negativo é o achado do dia.** A ordem exigida pelo OOM é
`Sort → IncludeRecurrences → Restrict`. Invertendo para `Restrict` antes de
`IncludeRecurrences`, a leitura devolveu **65 compromissos fora da janela
pedida** — ocorrências de *REUNIÃO PERIÓDICA DE P&D* de **janeiro** numa janela
de ±30 dias em torno de agosto.

Eu esperava que a ordem errada *perdesse* ocorrências. Ela faz a expansão
**ignorar o filtro**: uma agenda com sete meses de atraso, sem erro em lugar
nenhum. E dos cinco testes do arquivo, **só um** cai nesse controle — está
escrito assim.

**Cinco testes contra o Outlook real**, somente leitura, mais um puro sobre o
formato do filtro. Nada cria, move, apaga ou responde convite.

**Na janela.** `AgendaViewModel` + faixa, aparecendo só quando a pasta
selecionada é de calendário. E aí apareceu a armadilha de novo: a agenda ocuparia
a **mesma linha 2** do acervo — que é exatamente a configuração que escondeu a
faixa do acervo por dias. Desta vez a exclusão é **declarada**:
`MostrarAcervo = Acervo IsNot Nothing AndAlso Not Agenda.TemPasta`, uma condição
num lugar só, com teste e controle negativo.

**E o teste antigo pegou a mudança sozinho.** `O_acervo_e_a_IA_nao_dividem_a_LINHA`
falhou quando troquei o binding do acervo — ele procura pelo binding exato, e foi
escrito para isso.

**Outro dado medido e descartado, mesmo padrão do `message_class`:**
`FolderNodeViewModel` recebia `ContentKind` no `FolderInfo` e jogava fora. Foi a
agenda que precisou dele — apontá-la para a Caixa de Entrada mostraria
"0 compromissos" sobre uma pasta que não tem compromisso por definição, e zero
por engano é indistinguível de zero por medida.

Suíte: **846**.

### 28/08 — revisão da Fase 6 e da busca: um bloqueador e seis achados

**O bloqueador é o erro mais comum meu, um nível acima do de costume.** A
`FolderVisibilityPolicy` tinha `MailOnly = True`, então a pasta de calendário
**nunca aparecia na árvore**. A leitura funcionava, os cinco testes contra o
Outlook real passavam — e **o usuário não tinha como chegar nela**: os testes
achavam o calendário pelo broker, contornando a árvore. Não era o método sem
chamador; era a funcionalidade inteira sem porta.

`MailOnly` virou `SoOQueAbre`, com uma lista do que o Iris sabe abrir. O nome
antigo tinha virado mentira: a regra sempre foi "não ofereça porta que não
abre", e "só correio" era só uma coincidência enquanto isso era verdade.

**Os outros seis:**

1. **A busca contorna o `PublicationDrain`**, e consultar o estado dele não é
   passar por ele. Fica assim, dito com esse nome: o `AcervoService` é de uma
   pasta, e uma busca entre pastas não cabe nesse formato. Vai para o relatório
   como dívida.
2. **A frase do dreno estava factualmente errada** — dizia que as publicações
   "não foram entregues ao acervo". A publicação já materializou o acervo; o que
   falta é a entrega ao **painel**. A busca podia estar mostrando exatamente a
   geração que dizia não ter chegado.
3. **`CalendarReading` devolvia sucesso truncado** por dois caminhos: exceção no
   `GetNext` virava fim da coleção, e o teto do laço virava sucesso silencioso.
   Agora existe `Truncada` + `MotivoDoCorte`, e a tela diz "LISTA INCOMPLETA".
4. **A isenção de cobertura da agenda não era honesta.** "Lê ao vivo" prova
   frescor contra o OOM local, não cobertura do servidor — em cached, o universo
   local é o mesmo que a §23 diz que não dá para tratar como a caixa. Removida.
5. **`AgendaViewModel` tinha dois furos de geração**: o `Catch` e o `Finally`
   conferiam só o descarte. Uma leitura velha que falhasse apagava a lista boa da
   nova; uma leitura velha que terminasse desligava o indicador da nova. Ambos
   corrigidos, com `AgendaViewModelTests` novo e controle negativo dos dois.
6. **A busca afirmava demais em três lugares**: "nunca foram varridas" virou "não
   têm acervo publicado" (sem geração cabe também a tentativa rejeitada pela S6);
   a contração de alcance era calculada e jogada fora, agora aparece; e cobertura
   `Desconhecida` era tratada como parcial e recebia a causa da parcial.

**E dois testes meus passavam por motivo errado.** O da expansão de séries
cobrava uma desigualdade que vale com ou sem expansão — meu próprio controle
negativo já tinha dito isso e eu não vi. Reescrito para contar a mesma série em
dias diferentes. Aí eu **medi os três estados** e descobri que ele ainda não
guarda a ordem — com a ordem invertida a expansão *acontece*, o que quebra é o
filtro:

| | teste da expansão | teste da janela |
|---|---|---|
| ordem certa | passa | passa |
| `Restrict` antes do `Include` | passa | **falha** (65 fora) |
| `IncludeRecurrences` desligado | **falha** | **falha** (5 fora) |

Um guarda a expansão, o outro guarda a ordem. Nenhum guarda os dois, e está
escrito assim.

O teste da exclusão agenda/acervo só conferia que a propriedade existia.
`Acervo_e_agenda_nunca_aparecem_juntos` exercita a **transição** — correio,
calendário, nada — e cai quando a fórmula quebra, enquanto o antigo continua
verde.

`IAgendaSource` nasceu disso: o ViewModel recebia o broker inteiro, e por isso
não tinha teste nenhum. **Porta estreita não é elegância — é o que decide se
existe teste.**

Suíte: **854**.

### 28/08 — fechamento

`ESCOPO.md` na **v7**: Fase 4 marcada como *não executada por decisão*, com o
motivo; Fase 5 como *medida, não executada*; Fase 6 com a leitura entregue e o
que ficou de fora; Fase 7 adiada, com a ressalva de que "não tem consumidor"
seria forte demais.

`RELATORIO-TRABALHO-AUTONOMO.html` escrito e publicado. A seção que abre não é
o que foi entregue — é **o que eu decidi no lugar dele**, porque é a parte que
ele pode querer desfazer.

**Balanço do dia:** 805 → **860** testes. 19 commits. Cinco passadas de revisão
externa, e todas acharam alguma coisa. Nada foi enviado por e-mail; nenhuma
cerimônia foi executada; nenhuma mutação na caixa dele.

O achado que eu levo: **prova de leitura não é prova de alcance.** Entreguei a
agenda inteira funcionando e inalcançável, com seis testes verdes contra o
Outlook real — porque todos contornavam a árvore.

### 28/08 — cobertura do calendário: a janela é do correio, não do store

Dívida que nasceu de manhã e fechou à tarde.
`tools/medir-cobertura-calendario.ps1`, somente leitura:

| | correio | calendário |
|---|---|---|
| primeiro item | 2026-07-28 | **2024-06-07** |
| último item | 2026-08-28 | 2026-12-15 |
| span | ~31 dias | **921 dias** |

**411 dos 434** compromissos são anteriores ao corte do correio, e a
distribuição por ano é contínua — 53 em 2024, 229 em 2025, 152 em 2026. Sem
penhasco.

**E isso derruba uma afirmação minha da manhã.** O `medir-janela.ps1` dizia, no
cabeçalho, que *"o corte aparece em TODAS as pastas ao mesmo tempo, porque a
janela é do store e não da pasta"*. Era **inferência**, não medida — e está
falsificada: o calendário é do mesmo store e não corta.

A janela alcança o **correio**. Que é, aliás, como o controle deslizante do
Outlook funciona — mas eu não sabia disso por medida, e escrevi como se
soubesse.

**O que muda na tela:** a agenda **não** mostra uma fatia de um mês, e repetir
para ela a ressalva do acervo seria ressalva emprestada. O que continua valendo
é o outro lado: a contagem do servidor segue inalcançável, então ausência
continua proibida — **por falta de prova, e não por janela**.

### 28/08 — valores fabricados: medir antes de migrar

Segunda dívida da lista. A tentação era fazer a migração para campos anuláveis
direto — mas a disciplina deste projeto é medir primeiro, e foi assim que o
`message_class` apareceu.

`tools/medir-nulos-da-table.ps1`, sobre a Caixa de Entrada real: **zero nulos
nas oito colunas, em 1.109 linhas.**

Então a migração seria trabalho grande por risco que não se materializa aqui. O
que **não** podia continuar era o silêncio: `MessagePage.FabricatedCells` conta
as células ausentes que viraram valor, pela mesma regra da varredura — *recusa
declarada é mais forte que recusa silenciosa*.

Seis testes puros, sem Outlook, com controle positivo primeiro (valor presente
não pode contar, senão o número vira ruído que se aprende a ignorar) e controle
negativo: desligando os seis incrementos, cinco dos seis caem.

**E achei outra afirmação falsa minha, de hoje de manhã.** O comentário do
`CalendarFilter` dizia que a extração era necessária porque *"Iris.Outlook não
abre os internos para a suíte"*. **O projeto abre** — eu tinha olhado a lista de
`InternalsVisibleTo` com um comando truncado e concluí do que não vi. A
separação continua certa; o motivo que eu dei não era.

Suíte: **866**.

### 28/08 — a busca passou a PASSAR pelo dreno

Terceira dívida. Era desenho e não conserto, e o desenho foi feito.

`AcervoDeTodasAsPastas` é um segundo `IPublicationConsumer`: guarda o manifesto
de todas as pastas e só o refaz quando o dreno entrega. `ConsumidorComposto` faz
uma entrega alimentar os dois — dois drenos seriam pior, porque cada um marcaria
a geração como entregue por conta própria e o segundo nunca veria o que o
primeiro drenou.

**O teste antigo cristalizava o contorno.** Ele cobrava que a busca *avisasse*
sobre a publicação pendente — e passava porque a busca **via** a geração nova e
ao mesmo tempo dizia que ela não tinha chegado. Agora o teste prova o oposto:
publicou e não drenou ⇒ **a busca não vê**, igual ao painel ao lado. Ficar para
trás junto é honesto; ficar na frente em silêncio não era.

Controle negativo: reintroduzindo a releitura por pergunta, os dois testes caem
com a mensagem exata do defeito.

**E ficou mais rápida:** 22,7 ms → **6,2 ms** sobre as mesmas 1.127 mensagens
reais, mesmos 305 achados. O manifesto passou a ser lido uma vez em vez de a
cada pergunta — o contorno era mais lento *e* mais errado.

Suíte: **868**.

### 28/08 — revisão das três dívidas: um crítico, e ele era o mesmo contorno

Codex achou sete coisas, e a primeira dói: **eu disse ter fechado o contorno do
dreno e tinha deixado ele na abertura.** `AcervoDeTodasAsPastas` lia o manifesto
no construtor, justificando que *"na abertura não há entrega pendente"* — falso,
porque uma queda entre publicar e marcar drenada deixa pendência **persistida**.

E os testes passavam pelo motivo errado: o `Semear` já drenava, então o `Drenar`
seguinte não tinha o que entregar.

**Corrigido:** o construtor não lê nada. Quem enche é o `Receber`, ou um
`Recarregar()` que o dono chame **depois** de drenar. Teste novo que simula a
queda — publica sem drenar, **fecha o banco**, reabre, constrói do zero — e cai
quando devolvo o `Recarregar()` ao construtor.

**Os outros seis:**

1. `Receber` ignora *qual* geração chegou. Com duas pendentes, entregar a
   primeira torna a segunda visível cedo. Escrito no código; consertar pede o
   manifesto de uma geração específica, que o `ManifestReader` não faz.
2. **`FabricatedCells` não era lido por ninguém** — o número morria no DTO. É
   literalmente o erro mais comum meu. A lista passou a mostrá-lo.
3. **O contador subcontava:** eu zerava no início de `Ler`, e o `CursorPaging`
   chama `Ler` **várias vezes** por página. O DTO recebia só o último lote.
   Virou `Zerar()`, chamado uma vez por página — e o teste que eu tinha **não
   pegava**, porque cada teste criava uma fonte nova.
4. O fan-out não é atômico: se o primeiro consumidor conclui e o segundo falha,
   o painel fica à frente da busca até a repetição. Eu escrevi "congelam juntas".
5. A conclusão do calendário estava mais forte que a evidência — os dois
   roteiros nem correlacionam `StoreID`. Reescrita na formulação estreita.
6. Sobrou a frase *"a janela é do store"* na **saída** do `medir-janela`, ainda
   que o cabeçalho já estivesse corrigido.

Suíte: **870**.

### 28/08 — oitava passada: a ressalva afirmava o oposto do código

Codex confirmou que a correção do crítico é real — o construtor não lê mais, e o
`Recarregar()` está dentro do mesmo `Try`, então uma falha no dreno não o
dispara. Sete achados novos.

**O mais afiado:** a ressalva da busca dizia *"a busca já as enxerga; o painel
pode estar atrasado"*. Era verdade enquanto a busca contornava o dreno, e deixou
de ser **no mesmo dia** em que o contorno saiu — e eu não voltei à frase. Ela
passou a afirmar exatamente o oposto do estado. E **o teste passava junto**,
porque só cobrava a presença de "painel do acervo". Agora ele cobra o sentido.

**O contador errou nas duas direções, em dias seguidos.** Primeiro subcontava
(zerado por lote). Eu movi o reset para uma vez por página, e aí **sobrecontava**:
o `CursorPaging` lê um lote inteiro e para na primeira linha de outro instante,
então as linhas de *read-ahead* — que não entram nesta página — já tinham sido
contadas, e seriam contadas de novo na página seguinte.

Cada conserto de um lado abria o outro, porque o número morava entre *quem
converte* e *quem escolhe o que entra*. **Agora ele mora na linha**, e a página
soma o que recebeu. Não há o que errar.

**E o caminho legado ainda fabricava em silêncio** — eu instrumentei só o rápido,
então o zero que a lista mostrava para pastas no caminho lento era um zero
fabricado. Os quatro auxiliares passaram a contar.

Mais três de calibragem: *"temporária"* continuava mais forte que o mecanismo (o
dreno garante possibilidade de convergência, não convergência); sobrevivia
*"congelam juntas"* em três lugares; e o ESCOPO dizia *"921 dias, sem corte"* e
*"derruba"* — a medição não procura cortes mais antigos que o item mais velho que
achou, e os dois roteiros não correlacionam `StoreID`.

`RESULTADO-SUITE.md` remedido: **872** no `118277f` (o `d61037a` que eu tinha
escrito aqui foi invalidado por um `--amend` meu, minutos depois). Os relatórios diziam 870
enquanto o arquivo versionado registrava 805.

Suíte: **872**.

---

## Nona passada de revisão — e o padrão que ela expôs

Sete achados, e **dois altos**. Os dois no mesmo lugar onde eu já tinha
consertado no dia anterior: o contador de fabricação e a ressalva da busca.

**1. A ressalva ainda mentia — agora por generalizar.** Eu tinha trocado *"a
busca já as enxerga; o painel pode estar atrasado"* por *"o retrato anterior —
na busca **e** no painel"*. É verdade quando ninguém drenou, e **falsa na
entrega parcial**: o `ConsumidorComposto` entrega ao painel primeiro e à busca
depois, sem transação, então uma falha entre as duas deixa o painel uma geração
à frente. O ramo travado errava do mesmo jeito, dizendo *"nem a busca nem o
painel"*.

E o meu teste aceitava a generalização, porque só exercitava o caso em que ela
é verdadeira. **Terceira volta da mesma frase, em três eixos diferentes.** O que
mudou desta vez não foi a redação: a ressalva passou a **só afirmar sobre a
busca**, que é o que este objeto controla, e do painel diz apenas o que é certo
— que ele *pode* estar à frente. E a entrega parcial ganhou teste próprio, que é
a dívida "o fan-out não é atômico" saindo do ESCOPO e entrando na suíte.

**2. O caminho legado ainda fabricava calado, por dois buracos.** Eu instrumentei
os três auxiliares e deixei de fora o `ContarAnexos` — que não é auxiliar, e cuja
falha vira `HasAttachments = False`, a **pior** das cinco, porque o usuário lê
como afirmação. E nos auxiliares eu contei só o `Catch`: o getter pode devolver
`Nothing` sem lançar nada, e o caminho rápido já contava esse caso. Duas
instrumentações discordando faz o número depender de *qual caminho a pasta
tomou*.

**3. E o teste do reset provava o teste, não a produção.** Ele fazia
`f.Fabricadas = 0` **com a própria mão** entre as duas linhas — apagar o reset da
produção o deixaria verde. O laço saiu de dentro do `Ler`, que precisa de uma
`Table` do Outlook, e virou `ConverterLinhas`, que recebe o bloco cru. Agora o
teste entrega um `Object(,)` com buracos na primeira linha e nenhum na segunda, e
o **zero da segunda** é o que prova o reset.

O resto foi calibragem que sobreviveu em arquivos que eu não tinha aberto:
*"temporária"* no ESCOPO e no relatório, *"921 dias, sem corte"* e *"é do store"*
em mais quatro lugares — incluindo o `AgendaViewModel` e o cabeçalho do
`medir-janela.ps1` —, dois comentários descrevendo código removido, e o hash que
o `--amend` invalidou.

**Três controles negativos, confirmados desfazendo:** sem o `Fabricadas = 0`,
`Cada_linha_leva_o_SEU_numero` cai; sem a contagem do `Nothing`,
`O_legado_conta_a_excecao_E_o_Nothing_calado` cai; devolvendo *"na busca e no
painel"*, `Entrega_PARCIAL_deixa_o_painel_a_FRENTE` cai.

Suíte: **874**.

---

## Décima passada — a quarta volta da mesma frase, e o padrão nomeado

Quatro achados: um alto, dois médios, um baixo.

**O alto foi a ressalva de novo — a quarta vez.** Na nona eu tinha tirado a
afirmação categórica sobre o **painel** e mantido uma sobre a **busca**. E ela é
falsa pela *outra* dívida, a de o consumidor ignorar qual geração chegou: com 10
e 11 pendentes e o manifesto já apontando para 11, entregar a 10 faz a busca
recarregar a 11; se a entrega da 11 falhar, a busca enxerga exatamente a geração
que a ressalva jurava que ela não via.

**O padrão, agora com nome:** nas quatro voltas eu afirmei *o estado de alguém* —
ora do painel, ora da busca. Este objeto não sabe o estado de ninguém; ele sabe o
estado da **fila**. A versão atual afirma só isso, e no modo certo: havendo
entrega pendente, **nada na tela pode ser *tratado como*** o retrato da última
varredura. Não diz que está atrás, não diz que está à frente.

E o estado virou teste — `Com_DUAS_geracoes_pendentes_a_busca_ve_a_SEGUNDA_cedo`
—, agora com o `AcervoDeTodasAsPastas` **real**. A dívida estava escrita desde a
manhã, com a observação de que *"na prática a janela é curta"*. Curta não é
inexistente, e foi a revisão que ligou a dívida à frase.

**O primeiro médio foi sobre o nome do meu teste.** `Entrega_PARCIAL_deixa_o_
painel_a_FRENTE` usa dois contadores: prova que o fan-out é sequencial e que a
falha do segundo mantém a geração pendente, e **não** prova nada sobre o painel
de produção. O nome prometia a integração e o corpo entregava a unidade.
Renomeado, e o comentário passou a dizer o que ele não prova.

**O segundo foi calibragem sobrevivendo em mais três lugares** — incluindo um que
eu tinha acabado de reescrever: o `AgendaViewModel` concluía pela agenda inteira
o que foi medido só no calendário padrão local, enquanto a UI abre qualquer pasta
classificada como calendário. E o `medir-cobertura-calendario.ps1` ainda dizia
"está provado".

**O baixo:** o cabeçalho do relatório dizia 25 commits e o rodapé 27.

Suíte: **875**.

---

## Décima primeira passada — a que teria chegado à tela dele

Cinco achados. Um alto, dois médios, dois baixos.

**A ressalva errou pela quinta vez, e agora na abertura.** Ela dizia *"ainda não
foram entregues"*, e `Pendentes()` não significa isso: significa
`drained_at IS NULL`. A entrega é **ao menos uma vez**, e o `DrenoAposCrashTests`
— que já existia — diz com todas as letras *"o disco diz que a UI NÃO recebeu, e
ela recebeu"*. A parte modal estava certa; a abertura, não. Agora ela diz o que a
fila sabe: **entrega não confirmada**.

**Mas o achado que importa é o outro.** O comentário da agenda já reconhecia que
a medição de cobertura só alcança o **calendário padrão local**, e a agenda abre
qualquer pasta classificada como calendário. Só que o XAML ao lado continuava
dizendo *"por isso ela não tem ressalva de cobertura"*, e a tela mostrava
`0 compromisso(s)`.

Numa caixa compartilhada, o aplicativo **afirmava ausência** — que é exatamente o
que este projeto proíbe em todo lugar, menos onde ninguém tinha olhado. É o
mesmo defeito do `message_class` constante e do zero fabricado da paginação, pela
terceira vez, e desta vez ele estava na tela. Agora o resumo diz **"nenhum
compromisso LIDO até dd/MM — o que não é o mesmo que não haver"**, e o XAML
explica por quê.

O resto: **falta de controle de causalidade** no meu teste das duas gerações (sem
ele o teste prova o estado final e não a causa — corrigido com um `Assert` de
zero antes de qualquer entrega); o ramo **executável** do
`medir-cobertura-calendario.ps1` concluindo *"MESMO HORIZONTE"* quando não acha
item antigo, o que um calendário só com compromissos recentes produz igual; e a
contagem de commits, que eu tinha *igualado* nos dois lugares e continuava
errada nos dois.

**Dois controles negativos confirmados desfazendo:** devolvendo o
`$"{j.Items.Count} compromisso(s)"` incondicional, o teste da agenda cai;
devolvendo *"ainda não foram entregues"*, o da busca cai.

Suíte: **876**.

---

## Décima segunda passada — o meio-conserto, e a forma comum

**Nenhum achado alto.** Dois médios, três baixos — e o primeiro médio é o
defeito da passada anterior *meio* corrigido.

**Eu qualifiquei o zero e deixei o número positivo.** Ficaram três versões da
mesma história: o comentário da classe prometendo *"quantos compromissos leu"*, o
XAML qualificando só o caso zero, e a tela dizendo `12 compromissos` numa pasta
cuja cobertura ninguém mediu. **É a mesma afirmação que o zero era, só que mais
difícil de notar** — e eu tinha acabado de escrever, no commit anterior, que
afirmar ausência é o que este projeto proíbe em todo lugar. Agora é "lido(s)" nos
dois ramos, e o controle positivo do teste cobra a palavra.

**Os outros três têm uma forma só, e vale nomeá-la:** *tratar "não observei" como
"observei e não há"*.

- O roteiro do calendário dizia `Calendario vazio. Nada a medir.` quando não
  achava item — no mesmo arquivo cujo cabeçalho repete que a contagem do servidor
  é inalcançável pelo OOM.
- `BuscaNoAcervo` com `dreno = Nothing` devolvia `pendentes = 0`, que é a resposta
  de "olhei e a fila está limpa". Agora devolve `-1`, que cai na frase que já
  existia: *"não consegui conferir"*.
- E a minha quinta proibição só **proibia** a frase antiga: passaria com a
  abertura simplesmente apagada. **Proibir é barato; exigir é o que prende.**
  Agora o teste exige `"entrega não confirmada"`.

Suíte: **876** — a única passada que não acrescentou teste nenhum, só reforçou
dois que já existiam.

---

## Décima terceira passada — a que estava na tela principal

Sete achados, **nenhum alto**. Mas o segundo estava na tela que ele abre todo
dia, e é a **terceira instância da mesma família em três passadas**.

**A lista dizia as duas coisas ao mesmo tempo.** O texto do meio era fixo —
*"Esta pasta está vazia"* — e aparecia sempre que a lista convertida ficava sem
linha. Com uma página de `TotalAtStart = 1` e `SkippedCount = 1`, a mesma tela
mostrava *"Esta pasta está vazia"* no meio e *"0 de 1 · 1 item ignorado"* no
rodapé. Uma das duas mentia, e era a que ocupa a tela inteira.

Agora o texto é calculado, com três casos em ordem de força do que se pode
afirmar: a leitura perdeu item ⇒ não se afirma nada sobre a pasta; a pasta
declara N e a leitura trouxe zero ⇒ diz-se o N; a pasta declara zero e nada se
perdeu ⇒ aí sim, vazia.

**O segundo é pior de outro jeito.** A busca dizia *"Nenhuma pasta foi varrida
ainda"*, e o comentário do laço **logo acima** já dizia, desde a manhã, que isso
é mais do que se sabe: sem geração publicada cabe a varredura rejeitada pela S6,
a cancelada e a que falhou. Eu tinha corrigido o comentário e deixado a frase —
**e o meu teste exigia a frase errada**. Teste que exige a formulação errada
transforma o defeito em requisito, e é o pior lugar possível para deixar um.

Os outros cinco são da mesma linhagem, e listá-los junto é o que mostra o
tamanho do padrão:

- o **segundo zero** do `medir-cobertura-calendario.ps1` — item cuja data não foi
  lida entra em `$recusados` e não é classificado, então nem *"nenhum antes do
  horizonte"* se sustenta;
- o `medir-janela.ps1` engolindo falha de `GetFirst`/`GetLast` e depois
  afirmando *"nenhuma pasta com N itens"* — a pasta sumia da tabela sem aparecer
  em lugar nenhum;
- o `Skipped` nulo colapsado em zero no `Descrever`, **no arquivo cujo
  comentário imediatamente acima declara que nulo e zero são coisas
  diferentes**;
- a **quarta** versão documental do "12 compromissos", no comentário do `Resumo`
  e no do DTO, depois de a tela e o XAML já concordarem;
- e um teste meu que exigia `"3 compromisso(s)"` e `"lido(s)"` em asserções
  soltas, e passaria com o total sem qualificação e o "lido(s)" em outra
  cláusula.

**Dois controles negativos confirmados desfazendo:** fazendo o `EmptyMessage`
ignorar o `_skipped`, o teste da lista cai; devolvendo *"Nenhuma pasta foi
varrida"*, o da busca cai.

Suíte: **878**.

---

## Décima quarta passada — o controle negativo que passou

Eu perguntei ao revisor se a caça a "afirmar ausência" estava esgotada, **porque
eu ia parar de procurar**. A resposta foi **não**, com lista: dois no produto e
cinco em roteiros. Foi a pergunta certa a fazer, e a resposta certa a receber.

**O zero que eram dois estados.** `MessagePage.TotalAtStart` é `Integer?` porque
`ContarItens` devolve `Nothing` quando `Items.Count` lança. O `_total` do
ViewModel guardava só o número — então *"a pasta declara zero"* e *"não consegui
contar"* viravam o mesmo zero, e a tela dizia *"Esta pasta está vazia"* nos dois.
E o reload zerava `_skipped` e `_fabricadas` e **não** o total: uma pasta cuja
contagem falhasse declarava o total da pasta anterior. O rodapé agora mostra
`0 de ?`.

**Mas o que importa aqui não é o conserto.** Escrevi o `_totalConhecido`, apaguei
o ramo novo para conferir o controle negativo — **e a suíte inteira continuou
verde**. A correção não tinha teste nenhum. É o bloqueio sem controle negativo
que o `CLAUDE.md` descreve, cometido no mesmo dia em que eu o citei num
relatório. Dois dos quatro testes novos vieram desse susto.

**E a mesma frase, numa terceira superfície.** O `ManifestReader` ainda dizia
*"Esta pasta ainda não foi varrida"* — com **três testes** prendendo a
formulação —, e ela chega à faixa visível do acervo. Eu tinha corrigido a
superfície da busca na passada anterior e **não procurei as irmãs dela**, que é
literalmente o que o `CLAUDE.md` manda fazer: *ao corrigir uma corrida, procure
as irmãs antes de declarar a família coberta*.

Entrou também a **prova de alcance** que faltava: os dois testes do
`EmptyMessage` passariam com o texto literal de volta no XAML, e o comentário
deles anunciava um controle negativo que não existia. Agora um teste lê o XAML.

O resto foi inventário: o `medir-janela` engolia falha de tipo, de leitura e de
ramo da árvore (o `$semDatas` cobria só `GetFirst`/`GetLast`), e cinco outros
roteiros concluíam ausência sobre o que não leram — `medir-grupos`,
`inventario-pastas`, `preparar-ativacao`, `conferir-mensagem` e o histórico
`q1-nulos-empates`.

**Seis controles negativos confirmados desfazendo** — e um sétimo que passou, o
que originou dois testes.

Suíte: **882**.

---

## Décima quinta passada — o teste que passou pelo motivo errado, de novo

Nenhum achado alto. Quatro médios.

**O nome mudava e os números ficavam.** O `Despachar` tem fila de um: com uma
operação em voo, o pedido novo vira `_pending` e a chamada **volta na hora**. O
`ShowFolderAsync` terminava com o nome da pasta B na tela e as mensagens, o total
e o descarte de A — e o `_totalConhecido = False` só chegava quando o pendente
começasse. Se a operação de A travasse, durava para sempre. A limpeza passou a
ser síncrona, no instante da troca.

**E o meu primeiro teste disso passou com a correção desfeita.** Usei um
*reload* como operação em voo — e o reload já limpa a tela no começo dele, então
o estado estaria limpo por outro motivo. Tem de ser `LoadMore`. É a **segunda
passada seguida** em que um controle negativo meu passa quando devia falhar: da
primeira vez o conserto não tinha teste, desta vez o teste não tocava o defeito.
O `FakeBroker` ganhou `TravaDaPagina` para tornar a fila de um observável.

**Nos roteiros, o achado com dente foi o `conferir-mensagem`:** falha ao ler os
anexos virava *"não tem anexo"* e o roteiro chegava a imprimir **"O IRIS
ACEITA"** — enquanto a produção recusa expressamente quando não sabe.
Diagnóstico que contradiz a produção é pior que diagnóstico nenhum.

O resto: o `q1-nulos-empates` tinha qualificado só a conclusão dos nulos, e as
outras duas usam o mesmo corpus incompleto; o `inventario-pastas` não contava
corte por profundidade, falha de store nem `PR_ATTR_HIDDEN` ilegível, e a seção
dos artefatos imprimia "nenhum" sem consultar os contadores; o `medir-grupos`
traduzia qualquer exceção de `GetDefaultFolder` como "a pasta não existe"; o
`medir-janela` concluía "SEM horizonte comum" mesmo com leituras falhas; o
`q8-caca-contagem` dizia "não encontrada" sobre uma árvore cortada na
profundidade 6; o `q8-janela` dizia "nenhum" sobre chaves lidas com
`SilentlyContinue`; e o comentário do `ManifestReader` ainda dizia que nenhuma
UI consome a ressalva — o `AcervoViewModel` consome e o XAML exibe.

Suíte: **883**.

---

## Décima sexta passada — o defeito que eu introduzi consertando

Nenhum achado alto, e **o único médio fui eu**, na passada anterior: o
`$script:cego` do `inventario-pastas` era somado **depois** da seção que o
consulta, então "PASTAS DO USUARIO" lia `$null` e imprimia "nenhuma" mesmo com
falha de leitura. Consertar o inventário de cegueiras e deixar a primeira
consulta cega.

**E o meu controle negativo passou de novo — terceira vez em três passadas.** A
duração da última página também vazava para a pasta nova; mas no meu teste a
página de A custava zero milissegundo, então `0 ms` valia com e sem a correção.
Agora a página de A é segurada de propósito, o teste exige que ela custe tempo,
e só então exige que B não a herde.

O padrão das três: **eu escrevo a asserção olhando para o código que acabei de
escrever, e não para o estado que ela precisa distinguir.** Uma asserção só vale
se existir um mundo em que ela falha.

O resto: o `$semStore` descrevia como "não consegui abrir a raiz" um `catch` que
pega a recursão inteira; o `FakeBroker` tinha perdido o fail-fast síncrono do
"fora da alçada" — virava `Task` com falha, e um teste sem `Await` passaria
calado; e os quatro roteiros que restavam da lista — `q8-caca-total`,
`q8-caca-contagem`, `q2-pares` e `q2-quase` — afirmavam ausência sobre árvores
cortadas na profundidade e ramos que falharam em silêncio.

Suíte: **883**.

---

## Décima sétima passada — o fundo da família

**Nenhum achado alto nem médio — pela primeira vez.** Três baixos, todos a mesma
família, todos em roteiros históricos que eu ainda não tinha aberto:

- `q2-achar-copias` prometia *"TODAS as manifestações"* e terminava com *"TOTAL
  de cópias deixadas pelo experimento"* — que podia ser zero com ramos não lidos;
- `q1-protecao` concluía *"nenhuma mensagem protegida nesta amostra"* e depois
  extrapolava para a caixa, **sem contar as leituras de `Permission` que
  falharam**: uma amostra em que todas falhassem daria a mesma conclusão;
- `q2-chaves` prometia *"TODAS as pastas"* e *"a árvore inteira é percorrida"*,
  cortando em silêncio na profundidade 12. Os zeros da matriz apareciam sob
  promessa de cobertura completa, que é a forma mais cara de afirmar ausência.

**E entrou a regressão que o revisor pediu:** o contrato do `FakeBroker` de
**lançar na hora** quando chamam a página fora da alçada. Eu quase o perdi ao
acrescentar a `TravaDaPagina` — o embrulho `Async` transformava a exceção em
`Task` com falha, e um teste sem `Await` passaria calado. Controle negativo
confirmado devolvendo o embrulho.

Suíte: **884**.

---

## Varredura própria — o Codex sem cota, e a irmã que eu não tinha procurado

O revisor externo bateu no limite de uso da conta (volta em 29/08, 01:44). Em vez
de parar, varri sozinho as áreas que ele ainda não tinha aberto, com a lente das
últimas passadas: `Catch` que engole e devolve um padrão que a tela lê como fato.

**Achei uma, e é da família inteira.** O `CalendarReading` tem *exatamente* os
mesmos auxiliares da paginação legada — `Texto` devolve `""` na exceção e no
`Nothing`, `Booleano` devolve `False`, `Contar` devolve `0` — e **nenhum
contava**. Na tela isso vira `AllDayEvent = False`, `IsRecurring = False`, "sem
participantes", assunto e organizador vazios. E o `StoreDe` devolvia `""` calado:
`StoreID` vazio é chave que nunca casa, com sintoma longe daqui — é o `EntryID`
fabricado da paginação outra vez.

Eu instrumentei a listagem na manhã do mesmo dia e **não procurei a irmã dela**,
que é literalmente a regra do `CLAUDE.md` que eu tinha citado num relatório horas
antes.

`AppointmentWindow` ganhou `FabricatedCells`; os auxiliares contam por `ByRef`,
com sufixo `-DoCompromisso` porque `Friend` em `Module` vale para o assembly e
`Texto`/`Booleano` colidiriam com quatro homônimos; e o número **sobe para a
tela** no resumo da agenda.

O resto da varredura deu limpo: `Duravel` (grava ou não grava), o marcador do
rascunho e o `AnexosPresentes` já falham fechado ou preservam `Nothing`; as
frases de ausência que sobraram em `src/` são sobre o esquema do próprio banco e
sobre o arquivo de ativação, que são coisas que o código acabou de ler.

Dois controles negativos confirmados desfazendo.

Suíte: **886**.

---

## Décima oitava passada — o achado mais grave do dia inteiro

A cota voltou, e a passada trouxe **um alto**. É sobre conteúdo saindo da
máquina dele.

**O aviso de egresso ambíguo podia sumir para sempre.** Quando uma execução
morre no meio de um envio à IA, aquele envio fica *ambíguo*: os bytes podem ter
chegado, e ninguém sabe. A abertura seguinte deve avisar — e a conta era de
**quantas transitaram naquela chamada**. Bastava a segunda escrita falhar, ou o
processo morrer entre as duas, para as ambíguas ficarem gravadas e a abertura
falhar; na abertura **seguinte** a transição não pegava mais nada, a conta dava
zero, o aviso ficava vazio, e o egresso religava.

A conta passou a ser do **estado**: quantas *estão* ambíguas, de qualquer
execução. O preço está no ESCOPO: o aviso **fica**, porque não existe
reconhecimento.

**E mais um controle negativo meu que passou** — o quarto. Eu tinha escrito que
o controle do teste da queda era *tirar a transação*. Tirei, e ele continuou
verde: com a conta por estado, a linha é achada na reabertura de qualquer jeito.
Quem segura o aviso é a contagem por estado; a transação é **guarda não
observável pela API pública**, e ficou dita com esse nome.

**Três outras com dente:**

- **O segredo saía pela fresta do fechamento com espaço.** O
  `HtmlInterpretavel` conta `</script` — sem o `>` — então `</script >` conta
  como fechamento e o HTML passa. Mas o padrão que *remove o bloco* exigia
  `</script>` exato. O bloco sobrevivia, a limpeza de tags comia só as tags, e o
  conteúdo do script ia para o provedor **como se fosse a mensagem**.
- **O anexo errado podia ser gravado com o nome certo.** A guarda de identidade
  lê nome e tamanho dos dois lados com os auxiliares tolerantes: se as duas
  leituras falhassem, `""/0` casava com `""/0` e a guarda passava. É o dano
  exato que ela existe para impedir, chegando por dentro dela.
- **Um ramo inteiro da árvore sumia.** Falha ao abrir `Folders` virava
  `HasChildren = False`: a pasta perdia o triângulo de expandir, o correio
  existia sem ter como ser alcançado, e nada dizia por quê.

E o revisor disse, com razão, que **a minha varredura própria foi rasa**: eu
tinha concluído que só sobravam frases de ausência sobre o esquema do banco e o
arquivo de ativação, e faltavam a árvore de pastas, os detalhes, os anexos e a
reconciliação.

Nos roteiros, o pior foi o `q2-chaves`: `Select-Object -ExpandProperty`
**descarta `$null` do pipeline**, então um grupo com um Message-ID presente e
outro ausente caía em "igual" em vez de "falta" — na matriz que decide o desenho
da correlação.

Cinco controles negativos confirmados.

Suíte: **893**.

---

## Décima nona passada — eu declarei impossível uma coisa que dava

**O achado que mais ensina:** eu tinha classificado a transação da reconciliação
como *guarda não observável pela API pública*, porque o meu controle negativo
passou. A revisão mostrou a observação que eu não tinha visto: logo depois da
queda injetada, **antes** de qualquer nova reconciliação, basta `Ler`. Com
transação houve *rollback* e a primeira continua `EmVoo`; sem ela, já está
`Ambigua` com a segunda ainda `Intencionada` — metade do evento gravada.

**Declarar "não dá para testar" transforma uma lacuna em decisão permanente.**
Corrigido no teste, no ESCOPO e no comentário.

**E dois consertos meus estavam pela metade:**

- A guarda de identidade do anexo fechava quando a leitura de *agora* falhava e
  continuava cega para a leitura *de antes*, gravada na chave. Uma chave montada
  com `""/0` por falha casava com qualquer anexo que hoje leia `""/0` — e
  `"x.dat"/0` casava com uma chave em que só o nome tinha sido lido.
  `AttachmentKey.IdentidadeConhecida`, e a indexação parou de fabricar calada.
- O sanitizador: eu tinha trocado `</\1>` por `</\1\s*>`, que fecha o espaço e
  **não** fecha a família. A contagem aceita qualquer coisa que comece com
  `</script`, e o parser HTML também trata `</script x>` e `</script/>` como
  fechamento. **Consertar o caso citado e deixar os irmãos é o erro que este
  projeto já cometeu quatro vezes** — e eu o cometi enquanto o citava.

O resto: o `ItemCountConhecido` morria na fronteira do `FolderNodeViewModel`,
então não protegia o próximo consumidor; o `q2-chaves` somava duas causas
diferentes num contador só e explicava a causa errada; pegar o store estava fora
do `try` em dois roteiros; e o `<summary>` novo do `CrashInjection` tinha
engolido a documentação da constante do dreno.

Suíte: **896**.

---

## Vigésima passada — o mesmo defeito numa operação destrutiva

**Perguntei explicitamente pelos irmãos, e o revisor achou o pior.**

Eu tinha acabado de consertar a guarda de identidade do anexo na *leitura*. O
caminho dos **rascunhos** tem a mesma guarda decidindo um `Delete()` — e o
`MesmoArquivo` lia os dois lados com os auxiliares tolerantes, contra uma chave
montada com os mesmos. Se as leituras falhassem nos dois momentos, `""/0` casava
com `""/0`, a comparação dizia "é este", e **o anexo errado era apagado**.

**E o sanitizador estava na terceira versão do mesmo conserto** e ainda não
fechava a família. A contagem procurava a substring `</script`, então um
fechamento **sem o terminador** contava: o balanço fechava, o HTML passava, o
padrão de bloco não removia nada, e sobrava `SEGREDO</script` no texto que vai
para o provedor. Agora os dois lados usam o mesmo critério — o do removedor.

**Duas asserções minhas passavam pelo motivo errado:** usavam uma chave já
inconclusiva, então a função retornava antes de olhar os argumentos. Apagar as
duas guardas da leitura atual as deixaria verdes.

**E uma guarda que eu tinha acabado de escrever saiu.** O `MesmoArquivo`
repetia o teste de confiança, e o controle negativo não derrubou nada — porque o
`MesmaIdentidade` já o fazia. *Guarda duplicada é guarda que ninguém prova:* a
cópia sai, e o que fica é a que tem teste.

A igualdade e o hash do `AttachmentKey` passaram a incluir o
`IdentidadeConhecida`: duas chaves com os mesmos campos e confianças diferentes
não são a mesma chave.

Suíte: **899**.

---

## Vigésima primeira passada — contar era a abordagem errada

Três médios, e o primeiro é a **quarta** versão do mesmo conserto — com um caso
que **eu** abri.

**O sanitizador.** Contar aberturas e fechamentos no HTML bruto aceita
fechamento falso vindo de comentário ou de atributo:
`<!-- </script> --><script>SEGREDO` equilibra e passa. E a minha correção
anterior — contar só fechamento *terminado* — fez
`<!-- </script> --><script>SEGREDO</script` passar a ser aceito, porque o
fechamento de dentro do comentário equilibrava a abertura real. **A contagem
antiga recusava esse.**

Consertar contagem com contagem estava sempre a um contraexemplo de distância.
A pergunta certa não é *"está balanceado"*, é ***"sobrou alguma coisa que eu não
soube remover"*** — e é isso que o código faz agora: tira comentário, tira
bloco, e recusa se ainda restar `<script` em qualquer forma.

**A remoção validava um objeto e apagava outro.** As duas passadas — achar quem
é, e só então apagar — guardavam o **índice** e soltavam o objeto. Se a coleção
mudasse entre elas, o índice apontava para outro anexo: o defeito que as duas
passadas existem para impedir, sobrevivendo dentro delas.

**Identidade fabricada ainda deixava enviar.** O anexo com identidade não
conferida contava como *obtido*, a lista fechava como **completa**, e é a
completude que a tela consulta antes de encaminhar e de enviar. Agora não conta
— e a confirmação de envio passou a conferir o `AttachmentsStatus`, que o
`PrepareSend` já entregava e ela ignorava.

**E o roteiro novo mentia sobre si mesmo:** dizia atualizar três documentos e não
toca no diário; um argumento não era usado; duas substituições podiam trocar zero
ocorrências em silêncio; e o `git` rodava sem `check=True`, então uma falha
gravaria contagem vazia como se fosse medida.

Suíte: **903**.

---

## Vigésima segunda passada — o dano pelo outro lado

O revisor achou o mesmo estrago **invertido**. Com abertura *e* fechamento
falsos em atributos — `<p title="<script>">VISIVEL</p><p title="</script>">` —
o removedor come o `VISIVEL` que está entre os dois, não sobra nada, a
verificação aceita, e o que vai para o provedor **perdeu texto que o usuário
vê**. Eu tinha consertado o vazamento e aberto um sumiço: o mesmo erro de sinal
trocado.

Entrou um `MarcadorDentroDeAtributo` — varredura de estado, e não expressão
regular, porque é exatamente a noção que expressão regular não tem: *estou
dentro de aspas?* Ele recusa HTML legítimo, e esse é o preço, declarado.

**E o comentário tinha ficado na contagem** que eu abandonei para
`script`/`style`: `--> ... <!--` fecha na conta, o removedor não acha par, e o
texto que o navegador trata como comentário aberto sai como mensagem. Agora é a
mesma regra: se sobrou, recusa.

**A armadilha do `CLAUDE.md` ao vivo.** Extraí a regra "a identidade foi lida"
para uma função `IdentidadeLida`, e o local chamado `identidadeLida` a eclipsou
— VB é case-insensitive. O compilador disse *"o tipo não pode ser inferido"*,
que é exatamente o sintoma que a tabela do `CLAUDE.md` lista doze vezes.

O resto: a remoção ganhou comparação de contagem entre as duas passadas, com o
resíduo declarado; o aviso de leitura parcial dizia "Responder e encaminhar
ficam bloqueados" mesmo quando só encaminhar estava — e isso ficou mais comum
**por causa do meu conserto anterior**; o roteiro de evidência passou a montar
as duas edições em memória antes de gravar qualquer uma, e trocou `assert` por
`SystemExit`, que não some sob `python -O`; e o endereço de um teste novo tinha
saído corrompido pela minha própria substituição de aspas.

Suíte: **907**.

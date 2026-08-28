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

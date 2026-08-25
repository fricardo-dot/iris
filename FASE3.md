# Fase 3 — IA sob demanda

> Plano, segunda versão. A primeira foi escrita antes de qualquer código
> justamente para ser derrubada barato, e foi: o Codex achou sete bloqueadores
> estruturais. Estão incorporados aqui, com os nomes dele.

## 28. O que esta fase é, e o que ela não pode ser

O ESCOPO descreve a Fase 3 em uma linha: *"Resumo e redação sobre a mensagem ou
thread aberta."* A linha é curta e o assunto não é.

O que muda aqui, e não mudou em nenhuma fase anterior: **conteúdo da caixa
corporativa passa a poder sair da máquina**. Todas as fases até agora leram,
guardaram e mostraram — sempre dentro do computador. Um resumo por API externa é
a primeira vez que texto de e-mail de trabalho atravessa a fronteira.

Isso é o risco **R11**, e ele já prescreve a mitigação: *política explícita de
permissão e opt-in, não tentativa automática de redigir dados*.

### 28.1 O bloqueio herdado da Fase 0

A §10 do ESCOPO, sob *"O que continua NÃO validado"*:

> **Rótulos de sensibilidade do Purview** (`MSIP_Labels` via `PropertyAccessor`).
> **Obrigatório antes da Fase 3.**

A fase começa por uma medição, não por código de produto. Enquanto o Iris não
souber ler o rótulo, ele não sabe o que está proibido de mandar — e nesse estado
a única política defensável é mandar nada.

### 28.2 O que eu decido, e o que continua sendo do usuário

O usuário pediu execução independente e não está no computador. Duas coisas
dependem dele e não podem ser supridas por mim:

1. **A política corporativa aplicável.** Não é inferível de arquivo nenhum nesta
   máquina.
2. **O provedor e a credencial.** A escolha é dele; chave de API não é coisa que
   eu configure.

Então a transmissão fica **fechada**. Mas — e esta é a primeira correção do
Codex — **não vou chamar isso de "Fase 3 concluída"**. A analogia com a §23 é
fraca: lá o modo degradado ainda entrega um acervo útil; aqui transmissão zero
significa que o resultado central da fase **não funciona em produção**.

A formulação honesta, e a que vale para o relatório:

> **Implementação e provas locais concluídas. Ativação operacional e aceitação
> contra provedor real bloqueadas por decisão externa.**

Pelo mesmo motivo, **não escolho provedor**. Escrever um adaptador para uma API
específica seria antecipar uma decisão que este plano diz ser dele. O que fica
pronto é a **porta** e o servidor HTTP **local falso** que a exercita; o
provedor HTTP concreto é escrito como **referência não configurável**, e **não
conta como integração validada**.

**Nada de e-mail é enviado nesta fase**, como em todas as anteriores.

### 28.3 A cerimônia de ativação

Ligar a transmissão, quando ele quiser, exige um ato explícito e versionado que
declare: a autoridade/política que permite; provedor, endpoint, modelo e região;
a política de retenção e treinamento aceita; as pastas e rótulos permitidos; a
credencial no armazenamento apropriado; e um teste controlado escolhido por ele.
Só depois disso a fase passa a **operacionalmente aceita**.

---

## 29. Os sete bloqueadores, e como o desenho responde

O Codex mostrou que um portão pode **parecer** fail-closed e ser contornável sem
nunca devolver "permitido" formalmente. Estes são os caminhos, e a resposta de
cada um.

### 29.1 A unidade de decisão não é "um pedaço de conteúdo"

Uma thread tem mensagens com **rótulos diferentes**, pastas diferentes e
resultados de leitura diferentes. Decidir sobre "o conteúdo" apaga isso.

**Regra:** cada mensagem carrega identidade, pasta, rótulo e resultado de
leitura próprios. E na primeira versão, **um membro que não seja comprovadamente
permitido nega a thread inteira** — resumo parcial é fácil demais de confundir
com resumo completo.

### 29.2 Exigir o veredito no construtor não basta

Um veredito pode ser reutilizado para outro texto, emitido antes de o texto
mudar, ou associado ao item certo e a uma serialização diferente.

**Regra:** a autorização é uma **capability opaca**, emitida só pela política,
vinculada ao **hash dos bytes exatos** a transmitir, ao conjunto e versão dos
itens, à pasta, à operação, ao provedor/modelo, à versão da política e da
autorização — com validade curta e **consumo único**.

E "recalcular o hash antes de enviar" **não basta**: ainda permite serializar
para autorizar, serializar de novo, conferir, e mandar uma terceira
representação — escaping, ordem de JSON e normalização divergem entre as etapas.
A regra é mais forte:

> O pipeline materializa **um envelope imutável de bytes, uma única vez**. A
> política autoriza o hash **daqueles** bytes, e o transmissor manda
> **exatamente aquele buffer**.

A capability guarda hash, comprimento e identidade do envelope. **Não** guarda
texto remontável — o conteúdo continua só em memória.

### 29.3 Conteúdo citado destrói a segurança "por mensagem"

Mensagem sem rótulo pode citar integralmente uma classificada. Remover citação
não é barreira: heurística falha por idioma, cliente, HTML e edição manual.

**Regra:** esta fase adota a premissa de que **o rótulo do item governa o corpo
inteiro daquele item**, e a registra como **premissa corporativa a confirmar** —
não como fato. É mais uma razão para a transmissão nascer fechada: a premissa
não foi confirmada por ninguém.

### 29.4 O pipeline de conteúdo precisa ser especificado, não improvisado

Assunto, remetentes e destinatários **são conteúdo**. Corpo tem HTML malformado,
texto oculto, CSS, comentários, entidades, RTF, `cid:`, data URIs, imagens
inline e OLE. O plano especifica: qual campo entra, como é normalizado, que
limite em **bytes UTF-8 da requisição final**, truncamento **determinístico por
fronteira de mensagem** e **visível**, e **nenhuma busca de recurso remoto**.
Anexo inline é anexo, não "corpo grátis" — e anexo está fora desta fase.

### 29.5 O e-mail é dado hostil, e a resposta do modelo também

A resposta **nunca** vira comando: não escolhe endpoint, modelo ou header; não
pede nova chamada; não aciona COM; não cria destinatário nem anexo; não envia;
não abre URL; não renderiza HTML ou Markdown ativo; não sobrescreve o rascunho
sem revisão.

O provedor **não tem tools nem function calling** nesta fase. Instrução do
sistema, instrução do usuário e conteúdo do e-mail vão em **campos estruturais
separados** — concatenar tudo num prompt único torna a injeção mais fácil.

Mas campo separado **não impede** injeção: o modelo ainda pode obedecer ao
conteúdo do e-mail. A barreira real é a de cima — **saída tratada como texto
passivo, sem tools e sem efeito**. Campo estruturado reduz ambiguidade; não é
defesa.

Redigir insere no composer, que é **mutação local**: preserva o texto anterior e
oferece desfazer.

### 29.6 O diário precisa de protocolo de crash

Registrar depois da chamada perde justamente os casos que importam.

**Cinco passos:** intenção durável **antes** da tentativa → hash dos bytes
exatos → início da transmissão → conclusão ou falha → **`Ambiguous`** se o
processo morrer ou a conexão cair depois de o envio poder ter começado. É a
mesma disciplina do `ErrorKind.Ambiguous` que o CLAUDE.md já impõe às mutações.

Decisão **negada** não registra hash como se algo tivesse saído.

### 29.7 Concorrência e obsolescência

Enquanto a IA responde, o usuário troca de mensagem, edita o rascunho, pede
outra geração ou fecha a janela — e o item pode mudar no Outlook. **Resposta
velha não aparece em contexto novo nem sobrescreve edição posterior.** Cada
operação tem `RequestId` e geração, e o resultado só é publicado depois de
conferir o contexto.

---

## 30. A fronteira HTTP

Quando existir provedor: redirects **desativados**; endpoint fixo pela
configuração autorizada e **nunca vindo do prompt**; só HTTPS; credencial nunca
em query string nem em log; **nenhum retry automático depois de a transmissão
começar** — é a regra "leitura tem retry, mutação não" do CLAUDE.md aplicada ao
egress; timeout e cancelamento **não** significam que o provedor não recebeu, e
viram `Ambiguous`; limites de request e response; corpo de status **não
registrado**, porque pode ecoar conteúdo; exceções e telemetria sanitizadas.

E a regra arquitetural, corrigida: um teste de IL que só provasse que o
`Iris.App` não instancia provedor **não prova que outro assembly não abre rede**.
A regra é **todo egress de IA vive num único assembly de infraestrutura**, com
teste arquitetural sobre as referências de rede e **controle positivo**.

---

## 31. Arquitetura

```
Iris.Model
    ↑
Iris.Core
    ↑
Iris.Assist          contratos, política, montagem, orquestração
    ↑                (sem COM, sem WPF, sem SQLite, sem HTTP)
├─ Iris.Integration.Assist.Http     o UNICO com egress
├─ Iris.Integration.Assist          diário sobre o Iris.Cache
└─ extração via broker (sem tipo COM cruzando)
    ↑
Iris.App             composition root + UI
```

Distinção que o Codex cobrou e que eu não tinha: **`IAssistantProvider`** é a
porta externa — "chamar o modelo". O **serviço de aplicação** que aplica
política, monta contexto, registra intenção e publica resultado é outra coisa.
Um nome só para os dois faria alguém confundir chamar o modelo com executar a
operação segura.

---

## 32. Marcos

| | O quê |
|---|---|
| **3.0** | A medição do Purview, **pelo broker real**, somente leitura |
| **3.1** | `LabelReading` e `DisclosurePolicy` — o portão, falha fechada |
| **3.2** | Montagem de contexto e a **capability vinculada aos bytes** |
| **3.3** | O diário durável, com os cinco passos e o `Ambiguous` |
| **3.4** | A porta, o servidor falso local, e o provedor de referência |
| **3.5** | A UI: motivo em português, resumo, rascunho com desfazer |
| **3.6** | Provas adversariais ponta a ponta |
| **3.7** | Relatório — e a ativação real **fica pendente** |

### 3.0 — a medição, e por onde ela passa

**Pelo broker/STA real, não por um script à parte.** Medir por outro caminho
mede o Outlook, não o caminho do produto — o erro metodológico que o próprio
ESCOPO adverte.

| | Pergunta |
|---|---|
| **P1** | `MSIP_Labels` existe nesta conta? Sob qual DASL? |
| **P2** | Dá para ler sem prompt do Object Model guard? |
| **P3** | Em que fração das mensagens aparece? |
| **P4** | O `Sensitivity` clássico concorda com o rótulo moderno? |
| **P5** | Existe mensagem cujo corpo o OOM recusa (IRM)? Como recusa? |
| **P6** | O rótulo vem por `Table` (barato) ou exige abrir o item? |
| **P7** | Como o `PropertyAccessor` reage a propriedade **ausente**? |
| **P8** | Pode haver **múltiplos** rótulos? Qual prevalece? Como distinguir removido, obsoleto e desconhecido? |
| **P9** | O valor informa **proteção/criptografia** separada da classificação? |
| **P10** | Que **evidência de versão** existe (`EntryID`, `LastModificationTime`, `PR_CHANGE_KEY`, `PR_RECORD_KEY`), e que mudanças ela detecta? O que fica **sem garantia atômica**? |
| **P11** | Enviados, rascunhos, compartilhadas e outros stores expõem igual? |
| **P12** | Anexos e mensagens embutidas têm classificação própria? |
| **P13** | Ler corpo ou rótulo **dispara download** ou altera estado local? |
| **P14** | *(não é pergunta de medição — ver abaixo)* |
| **P15** | Aparecem propriedades alternativas quando `MSIP_Labels` não aparece? |

| **P16** | Qual a **autoridade** de `MSIP_Labels` nesta conta? Um remetente externo pode fornecer ou falsificar o mesmo header? |

#### A P14 saiu da medição, e é de propósito

Nenhuma leitura do Outlook responde se *"sem rótulo"* é corporativamente
permitido. O resultado do 3.0 para ela **já é conhecido antes de medir**:

> Semântica corporativa de ausência: **desconhecida**; não autoriza transmissão.

Prometer que o experimento responde isso seria fingir que uma pergunta de
política é uma pergunta técnica. Ela fica bloqueada pela cerimônia da §28.3.

#### A P16, e por que ela derruba o uso *positivo* do rótulo

O DASL usual aponta para uma named property do namespace de **cabeçalhos de
internet**:

```
http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/MSIP_Labels
```

Cabeçalho recebido pode ter origem fora do mecanismo corporativo. Mesmo lendo
`MSIP_Labels` com perfeição, isso **não prova** que ninguém consegue apresentar
um valor falso de classificação baixa.

Daí a assimetria, que vale até a política corporativa dizer o contrário:

- **sinal restritivo serve para NEGAR**;
- **ausência, rótulo baixo ou valor em allowlist não servem, sozinhos, para
  PERMITIR**.

#### O tipo de leitura, e por que "ausente versus exceção" não basta

A **P7** não é sozinha a mais importante — a **P10** é estrutural do mesmo
jeito, porque é ela que amarra a decisão ao que de fato foi transmitido. O tipo
precisa representar, no mínimo: **ausente comprovado**, **lido**, **leitura
negada**, **item protegido/IRM**, **erro transitório**, **item indisponível ou
parcialmente baixado**, **valor vazio**, **valor malformado ou desconhecido**,
**múltiplos/conflito**, **item mudou durante a leitura**.

A medição registra **contagens e valores pseudonimizados**. Nome completo de
rótulo e amostra de corpo não vão para o relatório se a contagem basta.

### 3.0 — como a medição é executada, sem o usuário na máquina

A caixa é corporativa e viva, e o usuário não está aqui para fechar um diálogo
modal. A ordem é do menos invasivo para o mais:

**A. Inspeção passiva.** Numa amostra pequena, registrar `EntryID`,
`LastModificationTime`, `UnRead`, classe e tamanho — um antes/depois observável.

**B. Piloto por `Table`, 20 linhas.** Adicionar a coluna DASL e ler. Evita
materializar 20 RCWs de `MailItem`, responde cedo se a propriedade é projetável,
e mostra como coluna ausente é representada. **Sem tocar** em `Body`,
`HTMLBody`, `RTFBody`, `Attachments.Item`, `GetConversation()` ou propriedade de
IRM.

**C. Confirmação por item**, em poucos casos escolhidos do piloto: presente,
aparentemente ausente, vazio, valor diferente, item normal. Via
`PropertyAccessor` pelo broker, comparando com o que a `Table` deu.

**D. Expansão adaptativa** — 20 → 100 → até 400 por pasta —, e só se não houve
prompt, bloqueio, mutação observável nem latência anormal. **Ao primeiro indício
adverso, para.**

**E. Protegidos ficam de fora.** Nada de corpo protegido com o usuário ausente:
uma chamada dessas pode abrir diálogo de autenticação ou de direitos. Para
mensagens com indício de IRM, só metadado já demonstrado seguro; P5 e P13 saem
como **"não medido por restrição operacional"**, e o estado é tratado como
proibido. Isso **é** resultado válido — o spike não tem obrigação de provocar
todo comportamento possível.

**O ovo e a galinha da P13.** Não há como provar de antemão que uma leitura
nunca hidrata conteúdo. O que dá para fazer é reduzir exposição, e registrar
honestamente o que se sabe no fim:

> Nenhuma alteração foi observada nesta amostra. **Ausência de download ou de
> efeito colateral não foi provada.**

#### Amostragem

Adaptativa e **estratificada**, não "as N mais recentes": recentes e antigas
dentro da janela acessível, remetente interno e externo, lidas e não lidas,
tamanhos diferentes, com e sem anexo (por metadado).

Para prevalência (P3), 384–400 observações dão margem aproximada de ±5 pontos a
95%; 100 dão quase ±10. Teto inicial de 400 **por pasta**, não por caixa.

A **P8 é diferente**: amostra nenhuma prova que múltiplo rótulo não existe. Zero
casos em 400 vira

> Nenhum observado; limite superior aproximado de prevalência ≈ 0,75% a 95% —
> **não** impossibilidade.

A **P11** não é respondível só pela Caixa de Entrada. Inbox primeiro; depois, se
o piloto for seguro, amostras pequenas em Enviados e Rascunhos. Compartilhadas e
outros stores ficam **"não disponíveis/não medidos"** enquanto a conta tiver um
store só.

#### Armadilhas do `PropertyAccessor`, listadas antes de tropeçar

- Propriedade ausente costuma vir como `MAPI_E_NOT_FOUND` (`0x8004010F`).
  **Nem toda `COMException` é ausência.**
- `GetProperty` e coluna de `Table` podem representar ausência de formas
  **diferentes**. Comparar tipo real, HRESULT e valor — não string.
- `GetProperties` em lote devolve erro **por elemento**; o adaptador tem de
  preservar resultado por propriedade.
- O valor pode conter **vários GUIDs** e registros históricos, inclusive
  `Enabled=False`. Não modelar como um par nome/valor.
- `Name` **não é identidade estável**; usar o GUID.
- Malformado, GUID inválido, duplicidade conflitante ou campo obrigatório
  faltando vira `MalformedOrUnknown`, **nunca "sem rótulo"**.
- A propriedade pode existir com **string vazia** — distinto de ausente.
- `Table.Columns.Add` pode **aceitar a coluna e entregar erro nas linhas**.
  Armadilha já vista nesta base com outras propriedades.
- A propriedade pode **não aparecer justamente em mensagem criptografada**.
- **Nunca** `SetProperty`, nem para testar round-trip.
- Não gerar nomes de propriedade candidatos em massa: named properties têm
  mapeamento próprio no store.
- Preservar o HRESULT **e a etapa**: obter o `PropertyAccessor`, resolver a
  propriedade e ler o valor são falhas diferentes.
- **R7 do CLAUDE.md**: nada de `mail.PropertyAccessor.GetProperty(...)`
  encadeado.

---

## 34. Resultados do 3.0 — medido em 25/08/2026

Contra a caixa corporativa real, pelo broker/STA, somente leitura. Reproduzível:
`dotnet test --filter FullyQualifiedName~PurviewMedicao`.

### 34.1 O achado que quase virou conclusão errada

A primeira rodada devolveu `Blank` para **120 de 120** itens. Lido sem cuidado,
isso é *"ninguém nesta caixa tem rótulo"* — e a fase inteira teria sido
desenhada em cima disso.

O controle negativo da P7 — ler, **no mesmo item e pelo mesmo caminho**, uma
propriedade que comprovadamente não existe — mostrou outra coisa:

| DASL | O que aconteceu |
|---|---|
| `MSIP_Labels` | lançou `MAPI_E_NOT_FOUND` (`0x8004010F`) |
| `msip_labels` | devolveu `String` de **0 caracteres** |
| nome inventado | lançou `MAPI_E_NOT_FOUND` |

**Nome de named property é sensível a maiúsculas**, e a versão minúscula é
**outra propriedade**, que existe vazia neste store. Eu estava lendo a
propriedade errada, e ela respondia "vazio" com a mesma cara de "sem rótulo".

E o nome inventado ter lançado prova uma segunda coisa: **`GetProperty` não cria
mapeamento de named property**. O `msip_labels` minúsculo já existia aqui.

> Sem o controle, a medição teria produzido um número plausível, redondo e
> falso — a partir de um artefato do meu próprio código.

### 34.2 As respostas

Números do recorte medido: **400 mensagens mais recentes** da Entrada, 86 dos
Enviados, 68 dos Rascunhos. Recorte por data, **não amostra aleatória** — a
distinção está nos próprios números abaixo, e não é preciosismo: intervalo de
confiança pressupõe representatividade que este recorte não tem.

| | Pergunta | Resposta medida |
|---|---|---|
| **P1** | A propriedade existe? | **Sim**, sob `.../{00020386-…}/MSIP_Labels` — com a caixa alta exata |
| **P2** | Lê sem prompt? | **Sim**. Nenhum diálogo em ~1.100 leituras |
| **P3** | Rotulados no recorte | Entrada **6/400** · Enviados **4/86** · Rascunhos **0/68** |
| **P4** | O `Sensitivity` clássico concorda? | **Não medido por desenho** — ele não responde por rótulo moderno, então concordância não seria evidência |
| **P5** | Corpo recusado por IRM? | **Não medido por restrição operacional** |
| **P6** | Vem por `Table`? | **NÃO.** `Columns.Add` recusa o DASL com `E_INVALIDARG` |
| **P7** | Como reage a ausente? | **Lança `MAPI_E_NOT_FOUND`** — é isso que torna `Absent` distinguível de `Blank` |
| **P8** | Múltiplos rótulos? | **Nenhum** em 554. Se fosse amostra aleatória o teto seria ≈ 0,54 %; por ser recorte, vale menos — e **nunca** impossibilidade |
| **P9** | Proteção separada da classificação? | **O formato separa.** Campos observados: `Enabled`, `Name`, `SetDate`, `Method`, `SiteId` e **`ContentBits`**. A *semântica* de `ContentBits` não foi validada |
| **P10** | Evidência de versão | `PR_CHANGE_KEY` em **20 de 20**, com `EntryID` e `LastModificationTime` |
| **P11** | Outras pastas | Entrada, Enviados e Rascunhos expõem **igual**. Outros stores e caixas compartilhadas: **não disponíveis** nesta conta |
| **P12** | Anexos têm classificação própria? | **Não medido** — anexo está fora da fase |
| **P13** | Ler altera estado? | Nada observado; **ausência de efeito não foi provada** |
| **P16** | Autoridade | Continua sendo cabeçalho. **Ler bem não prova autoridade** |

Estabilidade: os mesmos 20 itens, relidos — **0 mudaram de desfecho**. Estável
aqui não é garantia; é ausência de instabilidade observada.

### 34.3 Quatro coisas que mudam o desenho

**Rótulos existem nesta conta, e são poucos.** 10 rotulados em 554, com **três
GUIDs distintos**. Não é uma caixa sem classificação — é uma caixa em que a
classificação é rara e, por isso, fácil de ignorar por acidente. Um portão que
liberasse "o caso comum" liberaria a esmagadora maioria e vazaria justamente o
resto.

E raro por mensagem não é raro por **thread**: se a raridade fosse independente
entre mensagens — e não é —, uma thread de 30 já teria chance apreciável de
conter uma rotulada. É mais um argumento para a regra da §29.1: **um membro não
comprovadamente permitido nega a thread inteira**.

**`ContentBits` existe, e é só isso que está medido.** O formato observado
traz `ContentBits` **separado** do GUID. A consequência que se sustenta é
estreita: **o GUID sozinho não descreve o registro inteiro**, então o portão não
pode decidir olhando só ele.

O que eu tinha escrito — *"dois itens com o mesmo rótulo podem ter proteção
diferente"* — vai além da evidência. Eu observei um campo, não observei valores
divergentes para o mesmo GUID, e não demonstrei que o campo reflete a proteção
corrente, que é autêntico, que sua ausência significa "sem proteção", que não
está obsoleto, nem que cobre toda forma de IRM. **A P16 vale para ele também**:
vem no mesmo cabeçalho possivelmente não autoritativo.

Para o portão isso quer dizer: bit restritivo reconhecido pode contribuir para
**negar**; valor ausente, inválido ou desconhecido **não prova ausência de
proteção**; e `ContentBits=0` não autoriza sozinho.

**O caminho barato por `Table` não existe — nesta conta, com este DASL.** Custa
**~16 a 18 ms por item**, a mesma ordem que tornou o cache obrigatório na Fase 0.
Numa pasta de 1.000 itens isso é ~18 segundos; nos 17.728 do servidor, ~5
minutos.

A conclusão de escopo é: **a Fase 3 classifica sob demanda, no item aberto, e
não implementa varredura classificatória de fundo.** Note o que essa frase *não*
diz — que nenhum caminho barato exista no OOM. Não tentei `Items.Restrict` nem
outras formas do DASL, e não vou tentar: mesmo que funcionassem, classificação
de fundo **não poderia autorizar transmissão depois**, porque o rótulo
envelhece. O item teria de ser relido sob demanda de qualquer jeito.

**`Absent` e `Blank` são estados diferentes, e a maioria é `Blank`.** Dos 400 da
Entrada: 150 `Absent`, 244 `Blank`, 6 `Present`. Só significa *"a propriedade
existe com string vazia em 61 % deste recorte"* — **não** significa "sem
rótulo", "rótulo removido" nem "seguro". Pode ser pipeline do Exchange, add-in
inicializando a propriedade, cabeçalho normalizado, ou origem das mensagens.

### 34.4 Erros meus que o Codex pegou nesta rodada

Nenhum deles apareceria nos testes verdes, e os quatro liberariam mais do que
deviam:

**`E_INVALIDARG` virava `Absent`.** A minha própria medição mostrava que esse
HRESULT é o que a `Table` devolve para *"não aceito este DASL"* — ou seja, eu
tinha a prova de que ele significa "operação recusada" e o estava lendo como "o
item não tem rótulo". Agora é `Unsupported`. Ausência é `MAPI_E_NOT_FOUND`, e
só ela.

**O parser aceitava valor meio corrompido.** Pares não reconhecidos eram
ignorados em silêncio, então um GUID bom junto de um inválido saía `Present`.
E o **mesmo** GUID com `Enabled=True` e `Enabled=False` também saía `Present`,
porque o conjunto ordenado terminava com um ativo só — o comentário prometia
detectar conflito e o código detectava apenas *mais de um GUID*. Os dois estão
cobertos em `RotuloParserTests`, que existe separado porque **esta caixa não tem
nenhum caso difícil**: um parser cuja única prova é a caixa real é um parser
cujos ramos perigosos nunca rodam.

**Havia uma propriedade `Conclusiva`.** Ela juntava `Present`, `Absent` e
`Blank` — exatamente os três cuja política **difere** — sob um booleano que
nenhum `If leitura.Conclusiva Then` iria interpretar com o cuidado do
comentário. Removida. Quem decide trata cada membro do enum explicitamente.

**Eu inverti a direção do offset.** Escrevi que remoção repete e inserção pula;
é o contrário, e disso eu tinha derivado uma regra de teste que também estava
errada. A aritmética saiu do adaptador para o `OffsetPaging`, com
`OffsetPagingTests` demonstrando cada mutação sobre uma coleção controlada —
inclusive que **reordenar sem mudar o conjunto também repete**, que era a
segunda metade da minha afirmação falsa. Está na §27.8 da Fase 2, junto do
motivo.

### 34.5 O que continua NÃO respondido

- **P5 e P13**, por restrição operacional: abrir corpo protegido pode disparar
  diálogo de direitos, e o usuário não está na máquina. O estado fica tratado
  como **proibido**, que é o desfecho seguro.
- **P14**, por não ser pergunta técnica: a semântica corporativa de ausência é
  **desconhecida** e não autoriza transmissão.
- **P16**, por não ser respondível por leitura: só a política corporativa diz se
  o rótulo tem autoridade.
- **P12**, porque anexo está fora desta fase.
- **Efeito colateral de leitura**: não observado ≠ inexistente.

---

## 35. Marco 3.1 — o portão

`Iris.Assist`, `net10.0`, sem COM, sem WPF, sem SQLite e sem HTTP. O TFM é
quem garante — a mesma disciplina do `Iris.Sync` na Fase 2.

### 35.1 Permissão é conjunção fechada de provas positivas

Nunca *"não achei motivo suficiente para negar"*. A diferença não é
estilística: um portão escrito como lista de negativas libera todo caso que
ninguém pensou em proibir, e o caso que ninguém pensou é exatamente o que vaza.

Cada prova é uma pergunta cuja resposta tem de ser **sim**:

| Prova | Nega com |
|---|---|
| Existe autorização | `SemAtivacao` |
| Ela está completa | `AtivacaoIncompleta` |
| Está vigente | `AtivacaoForaDeVigencia` |
| O endpoint é HTTPS | `EndpointInseguro` |
| A operação está listada | `OperacaoNaoAutorizada` |
| Provedor e modelo são os autorizados | `ProvedorNaoAutorizado` |
| A pasta está listada | `PastaNaoAutorizada` |
| Há mensagem no pedido | `PedidoVazio` |
| **Por mensagem:** está na pasta autorizada | `MensagemDeOutraPasta` |
| **Por mensagem:** não tem anexo | `AnexoForaDeEscopo` |
| **Por mensagem:** o desfecho da leitura está listado | `LeituraNaoAceita` |
| **Por mensagem:** há evidência de versão | `SemEvidenciaDeVersao` |
| **Por registro ativo:** o GUID está listado | `RotuloNaoPermitido` |
| **Por registro ativo:** o `ContentBits` veio e é legível | `ContentBitsDesconhecido` |
| **Por registro ativo:** o `ContentBits` está listado | `ContentBitsNaoAceito` |

O `ActivationRecord` lista tudo **pelo nome**: as operações, as pastas, os
GUIDs, os `LabelReadingKind` aceitos e os valores de `ContentBits`. Não existe
"aceita os seguros" nem "aceita os conclusivos" — um estado que ninguém listou
nega, **inclusive um que ainda não existe**.

### 35.2 A ordem das provas também é decisão

As provas sobre a autorização vêm **antes** das provas por mensagem, e não é só
lógica. O 3.0 mediu ~17 ms por item para classificar; classificar uma thread
inteira para depois descobrir que a IA está desligada seria pagar meio segundo
de fila da STA para nada. Há teste cobrando essa ordem.

### 35.3 Um membro que não passa nega a thread inteira

Não a mensagem — a **thread**. Resumo parcial é fácil demais de confundir com
resumo completo, e o usuário não tem como saber que faltou pedaço.

E isso importa mais nesta caixa do que pareceria: rótulo é raro **por
mensagem** — 10 em 554 — e justamente por isso deixa de ser raro **por thread**.
Uma thread de trinta com vinte e nove permitidas e uma não é uma thread que não
passa.

### 35.4 O que a produção tem

`ActivationRecord.DaProducao` é `Nothing`. Em produção o portão nega **tudo**,
sempre, com `SemAtivacao`, e a explicação que chega ao usuário é em português:

> A IA externa não está habilitada. Nada deste computador é enviado para fora
> enquanto você não autorizar, com a política da empresa e um provedor à sua
> escolha.

Isso é a §28.2, não uma pendência de implementação.

### 35.5 As duas propriedades que a suíte cobra

**Exaustividade.** `TODO_desfecho_nao_listado_NEGA` percorre
`Enum.GetValues(GetType(LabelReadingKind))`. Acrescentar um estado e esquecer de
reler o portão **quebra a suíte** em vez de abrir uma porta que ninguém sabe que
abriu. E o contraponto `TODO_desfecho_LISTADO_passa` impede que isso seja
satisfeito por um portão que ignore a lista e negue sempre.

**Monotonicidade de segurança.** Acrescentar mensagem, tirar item da lista de
permitidos ou trocar um `ContentBits` conhecido por outro **nunca** transforma
um "não" em "sim". Três testes, um por eixo.

**E o controle positivo**, que é o que separa decisão de defeito: sem
`Autorizada_com_TUDO_no_lugar_PERMITE`, um portão que negasse por bug passaria
em todos os outros testes do arquivo, e "nega tudo" deixaria de ser uma escolha
para virar uma quebra que ninguém veria.

### 35.6 O que o portão deliberadamente NÃO faz

- **Não infere severidade pelo GUID.** GUID desconhecido nega; ninguém tenta
  adivinhar se "parece um rótulo baixo".
- **Não trata `ContentBits` ausente como zero.** "Não vi proteção" e "comprovei
  que não há proteção" são coisas diferentes, e só a segunda autorizaria.
- **Não usa o rótulo como autoridade positiva.** A P16 vale: `MSIP_Labels` mora
  no namespace de cabeçalhos de internet, então lê-lo com perfeição não prova
  que ninguém apresenta uma classificação baixa falsa. Rótulo entra como mais
  uma condição; o que autoriza é a autorização.
- **Não confere registro desligado.** Rótulo já removido não vale, então o GUID
  dele não precisa estar na lista — senão uma mensagem que um dia teve rótulo
  seria recusada para sempre, o que não é o que a autorização diz.

---

## 36. Marcos 3.2 e 3.3

### 36.1 O envelope — um buffer só

`AssistEnvelope` é o **corpo final da requisição**, materializado uma vez.
`Bytes()` devolve cópia; `Integro()` confere; a serialização mora num lugar só,
chamada tanto para medir quanto para produzir.

Truncamento é **por fronteira de mensagem** e vai declarado no próprio payload
(`conteudoOmitido`, `mensagensOmitidas`). O limite é em **bytes UTF-8 do pedido
final**, medido a cada mensagem acrescentada — estimar e somar daria um número
próximo e errado, porque escaping depende do conteúdo. E se nem o envelope
vazio couber, **recusa**: um envelope maior que o teto é um envelope que o
provedor rejeita *depois* de o conteúdo ter saído da máquina.

Assunto, remetente e destinatários entram como **conteúdo**. O `EntryID` não
entra — o provedor não precisa dele.

### 36.2 A capability — presa a bytes, itens e versões

Não descreve conteúdo: descreve o hash, o comprimento, o destino, a operação, a
ativação (id **e** versão), os itens **e as versões deles**, e um prazo que é o
menor entre dois minutos e o fim da ativação.

O que o Codex derrubou aqui, em três rodadas:

- **O "sim" não estava preso ao que o tornou um sim.** O cofre pedia
  `decisao.Permitido` e depois aceitava, em parâmetros separados, qualquer
  ativação, destino e envelope. `DisclosureGrant` fechou isso.
- **O hash não prova proveniência.** O `EntryID` não entra nos bytes, então dois
  envelopes com o mesmo texto e itens diferentes têm o **mesmo hash**. Há teste
  que constrói as gêmeas e mostra.
- **O grant prendia o rótulo à versão e o corpo a nada.** O rótulo do item X é
  lido em `CK-1`, o portão aprova, o corpo muda, o corpo é extraído em `CK-2`, e
  o envelope continuava dizendo apenas "item X". Agora `MessagePart` carrega a
  `ChangeKey` da mesma leitura, e o par `(item, versão)` é comparado.
- **A operação do grant não estava presa à operação serializada.** Um grant para
  `Resumir` emitia sobre um envelope montado como `Redigir`.

E **truncado ou com corpo incompleto não vira autorização**: uma thread que não
coube não vira uma thread menor.

### 36.3 O pipeline de conteúdo

`MessagePart` *afirmava* que o corpo já era texto seguro, e nada cobrava —
qualquer chamador passava HTML, `cid:` ou corpo pela metade. Escaping de JSON
impede quebrar a estrutura do documento; não impede o conteúdo ser o que não
devia.

`ContentPipeline` converte HTML em texto com comentário, `script` e `style`
saindo **antes** das tags — se saíssem depois, o conteúdo deles viraria texto
visível, e é ali que mora o que o usuário nunca viu na tela. Referência embutida
é procurada **no cru e no decodificado**, em todo campo: `<img src="cid&#58;x">`
não contém `cid:` no cru e o provedor lê a entidade do mesmo jeito.

**HTML mal formado recusa.** Expressão regular não é parser: `<script>SEGREDO`
sem fechar não casa com o padrão de bloco, a tag some junto com as outras, e
`SEGREDO` vira texto legítimo. Enquanto não houver parser de verdade, o mínimo
honesto é recusar o que não dá para interpretar.

Unicode é **preservado sem normalização** — normalizar mudaria o que o usuário
escreveu. O que sai são caracteres de controle e marcadores de direção, que não
são texto.

### 36.4 O diário, e o protocolo de crash

Um diário escrito no fim registra os envios que terminaram e perde justamente os
que importam: se o processo morre durante a transmissão não há linha nenhuma, e
o registro passa a afirmar **por omissão** que nada saiu.

Cinco passos: intenção durável **antes** da tentativa, com o hash dos bytes →
início do voo → conclusão ou falha → reconciliação na abertura seguinte.

| Onde morreu | Vira | Por quê |
|---|---|---|
| Depois da intenção | `NaoEnviada` | a transmissão não tinha começado, e isso se sabe |
| **Em voo** | **`Ambigua`** | os bytes podem ter chegado, e ninguém vai saber |

`Ambigua` **nunca** volta a ser `NaoEnviada`, nem com uma chamada explícita
dizendo que não chegou. E falhar em voo é ambíguo mesmo quando o chamador jura
que não chegou: entre *"a conexão caiu"* e *"a conexão caiu depois de o servidor
ler o corpo"* não há diferença observável deste lado.

**O diário nunca guarda conteúdo** — nem trecho, nem assunto, nem nome de
rótulo, nem corpo de resposta de erro, que pode **ecoar o que foi enviado**. O
teste planta uma isca e varre o arquivo do banco inteiro, byte a byte, porque
procurar coluna por coluna provaria só as colunas de que eu lembrei.

**Toda transição diz se pegou.** Os passos devolvem `Boolean`. Um `Iniciando`
que não persistisse — pedido inexistente, estado errado, corrida — passava em
silêncio, e quem chamou seguia para o HTTP assim mesmo: **egress sem registro de
voo**, que é o buraco que o diário existe para não ter. Quem transmite só toca
na rede depois de `Iniciando` devolver `True`.

**Os motivos são enums fechados.** Enquanto o campo era `String`, "nunca guarda
conteúdo" era convenção: qualquer adaptador podia passar a mensagem de uma
exceção ou o corpo de erro do provedor. Não existe mais campo por onde texto
arbitrário entre; a tradução para português mora na apresentação.

**Três carimbos, não um.** `intended_at` é imutável e é por ele que a ordem
histórica se guia; `started_at` e `finished_at` guardam o resto. Com um carimbo
só, sobrescrito a cada passo, uma intenção abandonada há meses aparecia como
atividade recente logo depois de uma reconciliação. A sequência de inserção
desempata — o `Guid` é aleatório, e uma lista que muda de ordem sozinha é uma
lista em que ninguém confia.

**A reconciliação não é trabalho da UI.** É recuperação de segurança, e roda na
composição, antes de o assistente ficar apto a transmitir: se falhar ou não
terminar, o egress fica fechado. A ligação é do 3.5, e até lá isso é **pendência
declarada**, não esquecimento.

**E a prova de crash é com processo morto de verdade.** Fechar o `Using` não é
morrer: dá ao SQLite a chance de descarregar tudo com ordem, que é exatamente o
que um crash não dá. O harness da Fase 2 ganhou um modo `diario` e é morto com
`TerminateProcess` depois de gravar o passo.

---

## 37. Marco 3.4 — a porta, o transporte e a ordem

### 37.1 O recorte, e por que ele é este

O Codex recomendou **cortar o "provedor de referência"** se ele significasse o
contrato de alguma API externa: autenticação, formato de request e response,
streaming, códigos, identificadores. Nada disso é inferível daqui, e escrever
seria **inventar requisito**.

O que ficou:

- **`IAssistantProvider`** — a porta externa. Recebe `Byte()`, e só. Se
  recebesse um DTO, haveria uma segunda serialização, e o que foi autorizado
  deixaria de ser o que sai.
- **`AssistTransmitter`** — o serviço de aplicação, que é outra coisa. A
  separação existe para ninguém confundir *chamar o modelo* com *executar a
  operação segura*.
- **`HttpAssistantProvider`** — transporte genérico, no único assembly com
  egress. Não sabe o formato de ninguém.
- **`AssistenteIndisponivel`** — o provedor da produção. Recusa **por decisão**,
  e existe em vez de `Nothing` porque `Nothing` vira
  `NullReferenceException` em algum caminho esquecido, e "explodiu" e "recusou"
  não são a mesma coisa para quem lê depois.

**Pendente, declarado:** o adaptador de fornecedor, a autenticação dele, o
formato dos bytes que ele aceita, a semântica de streaming, os códigos de erro e
os limites reais.

### 37.2 A ordem, e o que ela custa

```
portão → capability sobre AQUELES bytes → intenção durável → consumo
       → provedor pronto? → voo durável → rede → desfecho
```

Cada passo que falha para tudo, e o diário fica dizendo **onde** parou.

**O custo de marcar o voo antes**, declarado: uma falha que comprovadamente não
enviou nada seria contada como ambígua. A alternativa — marcar depois do
primeiro byte — exigiria confiar no transporte para dizer quando o byte saiu.
Quem erra nessa direção **esconde egress**; quem erra nesta **conta a mais**.

O que dá para tirar desse custo sem trapacear é `IAssistantProvider.Pronto()`:
endereço não-HTTPS, credencial ausente e provedor nenhum são recusas que se
**sabem** antes de qualquer byte, e por isso são perguntadas antes do voo.

### 37.3 As regras da §30, provadas contra um servidor de verdade

| Regra | Como está provada |
|---|---|
| Redirect **não** é seguido | 302 vira recusa, e o servidor de destino **não recebe nada** |
| Só HTTPS | a superfície **pública** é incapaz de aceitar HTTP: o parâmetro de loopback vive num construtor `Friend`. Antes era público com padrão `False`, o que prova o padrão e não impede a produção de passar `True` |
| Sem credencial, nem começa | e sem tocar na rede |
| Timeout **não** é "não chegou" | o servidor já tinha o corpo, e o teste confere que teve |
| Cancelamento idem | |
| O corpo do erro **não** atravessa | só o código HTTP; o corpo pode **ecoar o que foi enviado** |
| Resposta maior que o teto | **não** vira sucesso truncado: tem estado próprio e sem texto. Com contraponto — no teto exato, passa |
| **Nenhum retry** | nem com `503`, que é o código que mais convida a repetir |
| Credencial no cabeçalho | nunca em query string, que aparece em log e proxy |
| Credencial lida **na hora** | o teste troca o valor entre chamadas e a segunda usa o novo |

Os testes de tempo levaram um **aquecimento** antes de medir: sob a suíte
inteira em paralelo, o custo de subir a conexão passava do limite que o teste
dava, e ele media a corrida entre o cliente desistir e o TCP se estabelecer — não
a propriedade.

### 37.4 Depois do voo, a UI e o diário não podem discordar

Duas coisas que o Codex pegou, e as duas eram a mesma:

**Promessa quebrada do provedor.** Se `Pronto()` diz sim e `Enviar` depois diz
`NaoComecou`, o diário fecha em `Ambigua` — corretamente — e o transmissor
devolvia `NaoComecou`. A tela diria "não saiu" enquanto o registro dizia "pode
ter saído". `Pronto()` é otimização **antes** do voo, não palavra final depois.

**Transição terminal ignorada.** `Concluir` e `Falhar` devolvem `Boolean`, e o
transmissor publicava sucesso sem conferir. HTTP respondendo com o diário sem
fechar deixaria o registro `EmVoo` para sempre, e a reconciliação seguinte o
marcaria ambíguo — depois de a tela ter dito sucesso.

Agora **todo insucesso depois do voo é ambíguo**, e falha de persistência vira
`SemDiario`, nunca sucesso limpo. Há teste com um diário que recusa as
transições finais de propósito.

### 37.5 Egress mora num assembly só

A regra do plano v1 era um teste sobre o IL do `Iris.App` provando que ele não
instancia provedor de rede. O Codex derrubou: isso prova que **uma** camada não
chama, e não que nenhuma outra abre rede.

A regra é sobre **capacidade**, e a asserção é **exatamente um**: nenhum
assembly de produção além dele referencia biblioteca de rede, e ninguém depende
dele — depender é ganhar a capacidade de segunda mão. "Exatamente um" em vez de
"nenhum além do esperado" porque zero significaria que a busca procura a coisa
errada, e passaria em qualquer base.

E os assemblies são **descobertos**, não listados. Era uma lista escrita à mão, e
lista prova o que está nela: um projeto novo passaria calado — que é exatamente o
caso em que a regra importa. O padrão de busca também estava errado
(`Iris.*.dll` não casa com `Iris.dll`, que é o assembly do `Iris.App`), então a
camada que compõe tudo ficava de fora do teste que dizia cobri-la.

O que isso **não** prova: ausência de qualquer egress concebível — socket cru,
processo externo, um COM que busque URL. Essas são outras portas; o que está
fechado é a que a Fase 3 abre.

### 37.6 Um recorte de prova que ficou declarado

O portão exige **HTTPS**, e um servidor HTTPS local exige certificado que o
cliente aceite. Há dois caminhos, e é **escolha**, não impossibilidade:

- **certificado local confiado pelo sistema**, com infraestrutura de teste
  dedicada — custa montagem e manutenção, e é o caminho legítimo;
- **desativar a validação de certificado no código de produção** — barato, e um
  buraco **maior** que o que ajudaria a testar. *"Aceite qualquer certificado"* é
  pior que *"aceite http em loopback"*.

A primeira versão deste parágrafo dizia que o segundo caminho seria *exigido*.
Não é, e o Codex apontou. O recorte fica pelo custo do primeiro, e não por
impossibilidade — e afirmar impossibilidade onde há custo é a mesma família de
erro que esta fase inteira tenta não cometer.

Então as provas ficam separadas: o transporte contra servidor de verdade, e a
ordem com provedor falso. **O que não está provado é os dois juntos** — HTTP
real dentro da ordem inteira. Isso pertence à aceitação contra provedor real, e
está declarado em vez de simulado.

---

## 38. Marco 3.5 — a UI

### 38.1 O que o usuário vê hoje

Uma frase dizendo que a IA externa **não está habilitada**, e por quê. Não é
"recurso em construção" e também não é "só falta ligar": implementação e provas
**locais** estão concluídas, e continuam faltando duas coisas — a cerimônia de
ativação da §28.3, que é decisão do usuário, e o **adaptador do provedor
externo**, que é código e só pode ser escrito depois dela.

Se algum envio ficou sem desfecho conhecido numa execução anterior, isso aparece
**junto**: *"pode ter saído conteúdo, e não dá para saber"*. Um número desses não
vive só no banco.

### 38.2 A reconciliação roda na composição, e é pré-condição

O que ficou *em voo* numa execução que morreu vira ambíguo na abertura seguinte.
Isso é **recuperação de segurança**, não um número para mostrar: roda no
`MainViewModel`, antes de a IA ficar apta a transmitir, e **se falhar o egress
fica fechado** — ativação válida não basta.

Sem cache aberto não há diário, e sem diário a IA fica desligada: transmitir sem
poder registrar seria pior que não transmitir. O `DiarioAusente` existe para isso
em vez de `Nothing`, porque `Nothing` vira `NullReferenceException` em algum
caminho esquecido, e "explodiu" e "recusou por decisão" não são a mesma coisa
para quem lê depois.

### 38.3 Resposta velha não aparece em contexto novo

O modo de falha mais fácil de escrever e mais difícil de perceber: o usuário pede
o resumo da mensagem A, troca para a B enquanto a IA pensa, e a resposta de A
volta e é exibida. **Um resumo errado com cara de certo é pior que resumo
nenhum**, porque ninguém desconfia.

Cada pedido carrega uma **geração**; trocar de mensagem incrementa; um resultado
só é publicado se a geração dele ainda for a corrente. Com contraponto: sem
troca, a resposta aparece — senão um ViewModel que descartasse tudo passaria.

### 38.4 A resposta é dado, e a barreira é estrutural

O texto do modelo vem de um lugar que **leu o e-mail**, que por sua vez veio de
fora. Ele atravessa passivo até a tela, e a tela o mostra num `TextBlock` — não
num controle que interprete Markdown, HTML ou link.

Há teste que lê o XAML e confere que o elemento que recebe o binding é mesmo um
`TextBlock`. A barreira da §29.5 não é uma instrução ao modelo; é onde o texto
para.

### 38.5 Os motivos em português

Todo `DisclosureReason` tem frase própria, e há teste percorrendo o enum inteiro
para garantir que nenhum vaza como nome de código para a tela.

A tradução mora no ViewModel e **não** no diário — lá o motivo é enum fechado,
justamente para não haver campo por onde texto arbitrário entre.

### 38.6 A ação existe na janela, e o fluxo está ligado

Quatro coisas que o Codex pegou, e as três primeiras tinham o mesmo feitio: a
proteção existia no ViewModel e **não estava ligada a nada**.

**Não havia botão.** `Pedir` exigia três parâmetros e não estava ligado a
comando nenhum — nem uma ativação futura tornaria a funcionalidade utilizável.
Agora há `Resumir`, `Redigir resposta`, `Desfazer` e `Cancelar`, **visíveis e
desabilitados** quando a IA está desligada: um botão que some esconderia a
funcionalidade *e* o motivo dela estar desligada, e o motivo é o que o usuário
precisa ler no lugar onde procuraria a ação.

De onde a operação tira o que precisa é uma **porta** (`IAssistContext`), porque
classificar exige ir ao COM e montar exige o corpo — nada disso pode viver no
ViewModel sem arrastar o Outlook para dentro da tela.

**`Trocou()` não estava ligado à seleção.** Os testes chamavam o método direto e
na aplicação ninguém chamava — a proteção contra resposta obsoleta existia
isoladamente. Agora troca de mensagem, troca de pasta e sessão nova invalidam a
geração.

**O aviso da reconciliação podia ficar invisível.** A visibilidade olhava só o
`Aviso`, e com ativação válida ele fica vazio — então uma reconciliação que achou
envios ambíguos sumiria justamente no caso em que tem algo grave a contar.
Virou `TemAlgoADizer`, composto.

**Redigir com desfazer.** Escrever por cima do que o usuário digitou é mutação
local, e mutação local sem volta é a que ele descobre tarde demais. O texto
anterior fica guardado. Com contraponto: redação que **não veio** não mexe no
rascunho — senão a IA falhando apagaria o texto dele.

### 38.7 A segunda passada do Codex: quatro correções

**O contexto de verdade estava desligado.** A produção usava
`ContextoIndisponivel`, e ligar o contexto real era "pendência declarada, só faz
sentido depois de haver provedor". Não fazia sentido: ler a mensagem, classificar
cada membro e montar o envelope são requisitos **do Iris**, e independem de qual
API vai receber os bytes. Deixá-los para depois faria a frase "implementação e
provas locais concluídas" ser falsa — mesmo depois da cerimônia de ativação ainda
faltaria o caminho central até o Outlook.

Agora `ContextoDoOutlook` é o contexto de produção. A leitura é
`GetMessageSnapshotAsync`, **uma ida ao COM só**: cinco chamadas separadas podem
observar cinco estados de uma mensagem que mudou no meio, e a `ChangeKey` serve
justamente para prender o corpo à versão que o portão classificou — vinda de
outra passada, não prende nada. O que continua fechado é a transmissão: o destino
vem do provedor, e o provedor da produção é o `AssistenteIndisponivel`, que não
tem destino nenhum. O portão recusa **antes** de qualquer leitura.

**`Trocou()` invalidava sem reavaliar.** A geração era incrementada e os comandos
continuavam refletindo o contexto anterior — pasta nova pode ter outra
autorização, e o botão seguiria habilitado (ou desabilitado) pelo motivo errado.
Agora `Trocou()` chama `Avaliar()`.

**Redigir sobrescrevia edição concorrente.** A corrida gêmea da §38.3, do outro
lado: o usuário pede a redação, continua digitando enquanto a IA pensa, e o que
ele escreveu some. Pior que só sumir — o `Desfazer` devolveria o texto de
**antes do pedido**, e não a edição dele, então ele perderia o que escreveu por
duas vias. Agora o texto de partida é comparado no retorno e a redação **não é
aplicada** se ele mudou.

O resultado **fica na tela**. Descartá-lo resolveria o mesmo problema jogando
fora trabalho já feito, e que já saiu daqui: o conteúdo já foi ao provedor, e
apagar a resposta não desfaz divulgação nenhuma. Com controle negativo: sem
edição concorrente a redação **tem** de entrar, senão um comando que nunca
escrevesse passaria nos dois testes.

**Os botões viviam dentro do `Border` do aviso.** Eles apareciam exatamente
quando havia algo errado a dizer, e sumiam quando a IA estava funcionando —
o oposto do necessário, e o inverso do que a §38.6 acabara de justificar. A faixa
tem hoje três linhas próprias: botões (sempre), aviso, resultado.

Isso só foi pego porque o Codex leu o XAML. O teste de renderização não podia
pegar: ele montava uma **imitação** da faixa em código, então media uma árvore
que ninguém usava. Por isso a faixa virou um `UserControl` próprio
(`Views/FaixaDaIa.xaml`), que o teste agora instancia de verdade — e os testes de
binding leem o arquivo dela, não a janela.

### 38.8 A terceira passada: seis defeitos, e o que eles tinham em comum

Cinco dos seis eram a mesma forma de erro — **uma afirmação que ninguém tinha
verificado**, escrita num lugar onde outra coisa decidia por ela.

**O caminho de produção afirmava que não havia anexo.** A classificação montava
toda `MessageClassification` com `temAnexo:=False` fixo. O portão nega mensagem
com anexo, e para negar ele depende de alguém lhe dizer se tem — e quem lhe dizia
não tinha olhado. Pior: o parâmetro era `Optional temAnexo As Boolean = False`,
ou seja, o chamador que **não sabia** afirmava "não tem" sem escrever nada. O
padrão foi removido: agora é obrigatório, e quem constrói uma classificação tem
de declarar o que sabe.

A leitura é real, e falha **fechada**: se a contagem não for possível — guarda do
Object Model, item de classe inesperada, erro de COM — o valor é `Nothing`, e
`Nothing` conta como **tem**. É a mesma disciplina do 3.0, onde ler
`E_INVALIDARG` como "não tem rótulo" transformou ignorância em prova de ausência.

E há uma segunda barreira, porque a primeira não fecha a corrida: o portão
classifica numa visita ao COM e o corpo é lido em outra, então um anexo
acrescentado entre as duas passaria. `MessageSnapshot` traz o anexo lido **na
mesma visita que o corpo**, e o `ContentPipeline` recusa. A verificação que
importa está presa aos bytes que sairiam.

**Trocar de mensagem durante um pedido travava o assistente para sempre.** O
`Finally` devolvia `Ocupado = False` só se a geração ainda fosse a mesma — e
trocar de mensagem é exatamente o que muda a geração. A operação antiga terminava
sem devolver o estado: todos os botões desabilitados, nada na tela dizendo por
quê, e nenhuma saída a não ser fechar o Iris.

Era o irmão da §38.3 e o mais fácil de não notar, porque o teste de obsolescência
olhava o `Resultado` — e o resultado estava certo. Agora quem devolve o estado é
o **dono do voo**, identificado pelo `CancellationTokenSource`: a geração decide
se o resultado vale, e não se o estado é meu para limpar.

**Um portão só, para duas operações.** A disponibilidade era calculada com
`AssistOperation.Resumir` e usada para habilitar os dois botões. Como a ativação
lista as operações uma a uma, uma autorização só para resumo habilitava
visualmente a redação — negada depois, com o motivo aparecendo tarde e longe de
onde o usuário clicou — e uma autorização só para redação a deixava inalcançável.
Agora são dois preflights. Quando só uma passa, a faixa **diz qual não passa e
por quê**: botão desabilitado sem motivo ao lado é a forma mais silenciosa de
esconder uma recusa.

**A guarda do rascunho comparava só o texto.** A correção da passada anterior
prendia a resposta ao texto de partida, e não ao rascunho. Fechar o compositor e
abrir outro durante a espera dá um rascunho **diferente** que pode ter o mesmo
texto — e o caso comum é o pior: os dois vazios. A redação de uma mensagem
entraria na outra. Agora a comparação inclui a **sessão** do compositor, que é o
contador que ele já mantinha para largar continuações em voo quando o rascunho
acaba; reaproveitá-lo evita duas noções de "rascunho novo" que um dia
discordariam.

Junto com isso, `PodeRedigir` exigia só que o adaptador existisse — e em produção
ele existe sempre. O botão ficava habilitado com o compositor fechado e durante a
confirmação de envio, quando os campos estão travados justamente para que ninguém
mexa no que o usuário já aprovou. Agora exige `PodeEditar`, e abrir ou fechar o
compositor reavalia.

**E um RCW sem dono.** `TryCast(ns.GetItemFromID(...), MailItem)` perde a
referência quando o item existe e **não** é `MailItem`: o `TryCast` devolve
`Nothing`, a variável fica nula, e o `Finally` não tem o que liberar. É a R7 pela
sexta vez nesta base, sempre no mesmo formato — a expressão encadeada que parece
uma linha só. O objeto agora é adquirido numa variável `Object` própria.

**O que faltava era o controle negativo, não a correção.** Duas das correções da
passada anterior não tinham prova de estarem ligadas: voltar a produção para
`ContextoIndisponivel`, ou tirar o `Avaliar()` de `Trocou()`, deixaria a suíte
inteira verde. Hoje as sete correções foram confirmadas **desfazendo cada uma** e
vendo o teste correspondente falhar, com o controle da mesma família continuando
a passar.

### 38.9 A quarta passada: a recusa estava no botão, não na execução

Dois defeitos, e os dois eram a **metade que faltava** de correções da passada
anterior.

**O portão foi separado por operação só na tela.** `PodePedir` significa "pode
resumir", e era ele que guardava a execução inteira. Com ativação só para
redigir, o botão habilitava e clicar nele não fazia nada — a funcionalidade
existia e era inalcançável, que é o defeito da §38.6 chegando por outro caminho.
E a exigência de rascunho editável vivia só no `CanExecute`: uma chamada direta a
`Redigir()` atravessava e **transmitia** conteúdo sem haver lugar válido para
aplicar a resposta.

Botão desabilitado é conveniência. A recusa mora na execução, e agora cada
operação tem a sua — operação fora da lista recusa, como o resto da §29.

**O desfazer não tinha identidade.** A guarda nova protegia a ida da redação e
deixava a volta aberta: depois de redigir em A, fechar A e abrir B mantinha o
botão habilitado, e clicar nele escrevia o texto antigo de A dentro de B —
apagando o que houvesse lá, numa mensagem que a IA nunca tocou. O mesmo botão
escrevia num rascunho travado durante a confirmação de envio, quando os campos
estão bloqueados justamente para que a confirmação não vire mentira.

`PodeDesfazer` são hoje quatro condições, e todas são a mesma pergunta: *o que eu
desfaria ainda é o que eu fiz?* Há o que desfazer; o rascunho aceita escrita
agora; é o mesmo rascunho; e o texto ainda é o que a IA escreveu. A última fecha o
desfazer assim que o usuário digita por cima da redação — restaurar ali apagaria a
edição dele para desfazer algo que ele já desfez à mão. E a recusa **explica**:
um botão que o usuário clicou e que não faz nada não se distingue de um quebrado.

### 38.10 A quinta passada: o estado estava certo e invisível

Um defeito, e do tipo que só aparece quando alguém lê o binding em vez do
`if`. `PodeDesfazer` recusava corretamente depois de o usuário digitar por cima da
redação — e ninguém avisava o WPF. O `RelayCommand` não se reconsulta sozinho, e o
único assinante do `Composer.PropertyChanged` era o `MainViewModel`, que escutava
`PodeEditar`, `IsOpen` e `State`. `UserText` notificava desde sempre, e não havia
quem ouvisse.

O resultado: o botão "Desfazer" ficava **habilitado mostrando um estado que já não
existia**, até alguma outra mudança passar por perto. Clicar recusava com
segurança, e a promessa da §38.6 — ação desabilitada quando indisponível — estava
quebrada.

`IRascunho` ganhou um evento `Mudou`, que o adaptador do compositor levanta para
`UserText`, `PodeEditar`, `IsOpen` e `State`; o assistente escuta e reconsulta os
comandos.

A notificação atravessa **três saltos** — `Composer.UserText` →
`RascunhoDoCompositor.Mudou` → `AssistenteViewModel` → botão — e a primeira
tentativa de prova cobria só os dois últimos: todos os testes injetavam um
rascunho de mentira, e apagar a assinatura de `PropertyChanged` dentro do
adaptador de produção deixaria a suíte verde com a ligação quebrada. O
`RascunhoDoCompositorTests` monta o compositor de verdade e o adaptador de
produção, e fecha o primeiro salto — com controle que prova que o compositor
notificou, para distinguir qual dos dois lados quebrou.

**E a prova precisou de dois níveis.** Um teste que pergunta `CanExecute` depois
de editar passa mesmo sem existir notificação nenhuma — é o falso positivo de
binding silencioso na sua forma mais pura. Então há um teste que observa o
`CanExecuteChanged`, e outro que monta a **faixa de verdade** com o **ViewModel de
verdade** e lê o `Button.IsEnabled` do botão real.

O segundo custou infraestrutura: ler `IsEnabled` constrói o `InputManager`, que
exige STA, e o `Await` precisa voltar para a mesma thread — senão o
`CanExecuteChanged` é levantado numa thread do pool e o `Button`, que é
`DispatcherObject`, recusa a visita. Daí o ajudante `NaSTA`, com `Dispatcher`
próprio. É o preço de provar o que o usuário vê em vez do que o objeto responde.

### 38.11 E a faixa foi medida

`FaixaDaIaRenderizaTests` faz `Measure`/`Arrange` fora do vídeo, como o
equivalente da Fase 2. Binding correto não detecta faixa com altura zero nem
controle colapsado — e aqui há um risco a mais: **o aviso e o resultado ocupam a
mesma `Grid.Row`**. Um teste mostra que, com os dois visíveis, um cobre o outro;
o ViewModel garante que isso não aconteça, e o teste é o que acusa se um dia
acontecer.

A montagem precisa de **duas passadas** de layout: a primeira mede antes de os
bindings de `Visibility` terem sido aplicados, e o `Grid` sai com altura de quem
não tem nada a mostrar.

### 38.12 O que ficou de fora, por recorte

Configuração de credencial, escolha de provedor ou modelo, e qualquer UX moldada
pelas capacidades de um fornecedor específico. Sem provedor escolhido, isso seria
inventar requisito.

---

## 39. Marco 3.6 — o adversário, na cadeia inteira

### 39.1 Até onde esta prova vai, e até onde não vai

Do **contrato do broker** até a **porta do provedor**: `IOutlookBroker` →
`ContextoDoOutlook` → `DisclosurePolicy` → `CapabilityLedger` →
`AssistTransmitter` → provedor → diário SQLite de verdade, com o
`AssistenteViewModel` na ponta.

Não vai do COM ao socket, e dizer que vai seria mentira. O Outlook real tem
provas próprias no 3.0; o transporte real tem as dele contra `HttpListener`.
Juntar os dois aqui exigiria furar a exigência de HTTPS do portão, e esse furo
seria maior que o buraco que ele fecharia.

### 39.2 O que os testes unitários não pegam

Cada camada já tinha prova própria — 37 testes de envelope e capability, 44 do
portão, 19 de ordem e diário. O que faltava é o que mora **entre** elas, e que
nenhum teste unitário vê, porque em cada um deles a camada de baixo é fabricada
pelo próprio teste:

- a classificação diz "sem anexo" e a leitura do corpo diz outra coisa;
- a thread chega pela metade porque a leitura de um membro falhou;
- a seleção do usuário se move entre classificar e montar;
- a versão da mensagem muda entre uma coisa e outra.

### 39.3 Injeção de prompt: a fronteira honesta

O que se prova é **estrutural**, e é o que dá para provar:

- o endereço chamado é o da **ativação**, e não um que apareceu dentro de um
  e-mail;
- a instrução de sistema é a do Iris, fixa no código;
- a instrução do usuário é a que o **botão** manda;
- a carga viaja no campo de conteúdo, e **não quebra o JSON**.

Esta última tem teste próprio, com aspas, chaves e barras montadas para fechar o
campo `corpo` e abrir um `instrucaoDoSistema` do remetente. A conferência é feita
com `JsonDocument` **sobre os bytes que saíram** — não sobre o objeto que os
produziu — e vale para assunto, remetente e destinatário também: os quatro campos
vêm do e-mail, e uma barreira que olhasse só o corpo deixaria três portas
abertas.

O que **não** se prova, e não se finge provar: que o modelo não obedeça à frase.
Campos separados reduzem ambiguidade e **não são barreira** contra injeção de
prompt. Uma barreira de mentira é pior que nenhuma, porque alguém confia nela.

### 39.4 HTML hostil tem três desfechos, e agrupá-los esconderia o contrato

| Corpo | Desfecho |
|---|---|
| HTML bem-formado com `<script>` | transmite **só o texto visível** |
| corpo só de script | para: não sobrou texto |
| script ou comentário desbalanceado | para: não dá para interpretar |

Com controle: HTML **válido e inofensivo** atravessa. Sem ele, um pipeline que
recusasse todo HTML passaria nos dois últimos e deixaria a IA inútil para a maior
parte do correio real.

### 39.5 O que o controle negativo revelou, e nenhum teste unitário revelaria

Duas barreiras diferentes, e cada uma faz metade do trabalho:

- o **pipeline** impede que conteúdo recusado — anexo, referência embutida, HTML
  ilegível — entre nos bytes: ele descarta a mensagem;
- a **cobertura do grant** impede que o envelope resultante, vazio ou parcial,
  seja enviado.

O que só a prova integrada mostra é que a segunda é indispensável para a
primeira. Descartar a mensagem numa seleção de um item só produz um envelope
**vazio**, e envelope vazio é válido por decisão declarada no 3.2 — o
`EnvelopeBuilder` não conhece o conjunto aprovado, e não teria como julgar. Quem
recusa é o cofre.

Os controles negativos, medidos:

- desligando só a verificação de anexo do pipeline, cai um teste adversarial;
- desligando só a cobertura no `Emitir`, **nenhum dos 25 adversariais** cai —
  porque o consumo da capability reconfere a proveniência por conta própria. (Os
  testes unitários de `Emitir` caem, como devem.)
- desligando só a reconferência do consumo, **nenhum** cai — pela razão
  simétrica;
- desligando **as duas**, caem sete — inclusive os três de conteúdo, que
  passariam a mandar envelope vazio.

As duas são independentemente suficientes, e é por isso que nenhuma sozinha
aparece nos controles. Redundância aqui é escolha, não descuido: uma delas roda
na emissão e a outra no consumo, e entre os dois momentos o envelope passa por
código que pode mudar.

#### O controle negativo quase custou uma guarda

O roteiro que desliga e restaura fotografava o arquivo **antes de cada edição**.
Com duas edições no mesmo arquivo, a segunda foto já era da versão quebrada, e a
restauração devolveu o `Capability.vb` **sem a chamada a `Cobre`** — que foi
commitada assim. A suíte continuou verde porque os 642 tinham sido medidos antes.

Quem pegou foi o Codex, lendo o código: *"o teste unitário
`Sim_para_uns_itens_NAO_emite_para_outros` necessariamente deve falhar com esse
HEAD"*. Estava certo.

O roteiro passou a fotografar cada arquivo **uma vez**, e a conferir o SHA-256
depois de restaurar. A lição é a da própria fase: ferramenta de verificação
também precisa de verificação, e a que edita código de produção precisa mais que
as outras.

O encadeamento omissão → cobertura já estava escrito no `ContextoDoOutlook`; o
que não existia era a prova de que ele é o que segura. Nenhum teste de camada
poderia mostrar: cada um fabrica a camada de baixo, e por isso nunca vê o
envelope vazio chegar ao cofre.

### 39.6 A resposta vazia ganhou contrato — e espaço em branco é vazio

`Respondeu` com texto vazio fechava o diário como sucesso e não deixava nada na
tela: nem resultado, nem aviso. A operação **sumia**, e o usuário não teria como
distinguir "o provedor não tinha o que dizer" de "o botão não funciona".

Não é ambíguo — o conteúdo saiu e a resposta chegou —, então o diário fecha como
concluída e a faixa diz exatamente isso.

A primeira versão do contrato tinha um buraco: `TemResultado` era
`Resultado.Length > 0`, então três espaços ou uma quebra de linha escapavam do
aviso, deixavam a faixa visualmente vazia e — pior — eram **aplicados por cima do
rascunho do usuário** na redação. Trocar o texto dele por espaços é perda de
trabalho com cara de sucesso. Hoje é `IsNullOrWhiteSpace`.

### 39.7 A resposta hostil tem dois consumidores

Resumir e redigir. O primeiro é provado até o `TextBlock` da faixa **real**,
instanciada com o ViewModel real; o segundo até o rascunho editável. Nos dois, a
resposta chega **literal**: nada abre URL, nada interpreta markdown, nada envia
e-mail.

### 39.8 Os controles, um por mecanismo

Um controle global não basta — ele prova que o equipamento funciona, e não que
cada barreira é a que está barrando. Há um por mecanismo: HTML válido, anexo
ausente dos dois lados, mesma `ChangeKey` dos dois lados, thread inteira, seleção
estável, resposta normal, e a cadeia inteira funcionando.

Além disso, todo teste de "zero chamadas" confere antes que o preflight
**passou** — `CanExecute` habilitado — e que o broker chegou até onde deveria. Sem
isso, uma recusa acidental mais cedo faria a prova passar dizendo outra coisa.

### 39.9 A suíte deixou de brigar consigo mesma

`SqliteConnection.ClearAllPools()` é **global**: ele derruba a conexão de
qualquer classe que esteja abrindo banco no mesmo instante. Dez classes de teste
o chamavam — todas para poder apagar o arquivo temporário no Windows — enquanto o
MSTest rodava métodos em paralelo. Foi a causa provável do
`Cannot access a disposed object: SQLitePCL.sqlite3` que apareceu três vezes hoje,
sempre na primeira execução depois de uma compilação.

As dez classes passaram a `<DoNotParallelize>`. A suíte foi de 38 s para 58 s, e
o preço é justo: uma suíte que falha uma vez em vinte não prova o que diz provar,
e uma falha intermitente treina quem a vê a reexecutar em vez de olhar.

Junto disso, os testes que varrem variantes abriam um banco por variante e
limpavam só a última pasta — as outras ficavam no `%TEMP%` do usuário para
sempre. Agora cada abertura registra a sua, e a limpeza apaga todas.

---

## 33. O que esta fase NÃO faz

- Não envia e-mail.
- Não manda nada para fora — a transmissão nasce fechada e depende da §28.3.
- Não escolhe provedor.
- Não redige dados sensíveis automaticamente: o ESCOPO rebaixou isso por não ser
  barreira de compliance, e fingir que é seria pior que não ter.
- Não trata anexos.
- Não faz triagem nem busca semântica (Fase 4).

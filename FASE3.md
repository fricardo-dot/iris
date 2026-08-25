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

## 33. O que esta fase NÃO faz

- Não envia e-mail.
- Não manda nada para fora — a transmissão nasce fechada e depende da §28.3.
- Não escolhe provedor.
- Não redige dados sensíveis automaticamente: o ESCOPO rebaixou isso por não ser
  barreira de compliance, e fingir que é seria pior que não ter.
- Não trata anexos.
- Não faz triagem nem busca semântica (Fase 4).

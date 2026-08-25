# Fase 3 — IA sob demanda

> Plano. Escrito antes de qualquer código, para ser derrubado antes de
> qualquer código.

## 28. O que esta fase é, e o que ela não pode ser

O ESCOPO.md descreve a Fase 3 em uma linha: *"Resumo e redação sobre a mensagem
ou thread aberta."* A linha é curta e o assunto não é.

O que muda aqui, e não mudou em nenhuma fase anterior: **conteúdo da caixa
corporativa do usuário passa a poder sair da máquina**. Todas as fases até agora
leram, guardaram e mostraram — sempre dentro do computador dele. Um resumo feito
por API externa é a primeira vez que texto de e-mail de trabalho atravessa a
fronteira.

Isso não é detalhe de implementação. É o risco **R11** do ESCOPO, e ele é
explícito sobre a mitigação:

> Mitigação principal: **política explícita de permissão e opt-in**, não
> tentativa automática de redigir dados.

E lista, item por item: escopo de pastas habilitado explicitamente e nunca por
padrão; bloquear mensagens protegidas ou classificadas; log de metadados e nunca
de conteúdo; confirmação explícita antes de anexos; verificar a política
corporativa aplicável.

### 28.1 O bloqueio herdado da Fase 0

A §10 do ESCOPO registra, sob *"O que continua NÃO validado"*:

> **Rótulos de sensibilidade do Purview** (`MSIP_Labels` via `PropertyAccessor`).
> `MailItem.Sensitivity` é a propriedade clássica e não responde por rótulos
> modernos. **Obrigatório antes da Fase 3.**

Ou seja: a fase começa por uma medição, não por código de produto. Enquanto o
Iris não souber ler o rótulo de uma mensagem, ele não tem como saber o que está
proibido de mandar para fora — e nesse estado a única política defensável é
mandar nada.

### 28.2 A decisão que eu tomo, e o usuário revisa depois

O usuário pediu execução independente e não está no computador. Duas coisas
dependem dele e não podem ser supridas por mim:

1. **A política corporativa aplicável.** Não dá para inferir de arquivo nenhum
   nesta máquina se a empresa permite mandar corpo de e-mail para API externa.
2. **A escolha do provedor e a credencial.** Chave de API é dele, e não é coisa
   que eu configure.

Então a fase é construída **inteira** e entregue com a transmissão **fechada** —
o mesmo desenho da §23, que ele já aceitou: o mecanismo existe, é testado, e a
política de produção autoriza **zero**. Ligar exige um ato explícito dele,
registrado, com data e texto.

Não é o mesmo que "não fiz". O portão, o diário, a porta, os provedores, a UI e
a prova de que nada escapa por fora ficam prontos e verificáveis. O que não
acontece é conteúdo dele sair desta máquina por decisão minha.

**Nada de e-mail é enviado nesta fase**, como em todas as anteriores.

---

## 29. Marcos

### 3.0 — A medição do Purview (desbloqueia a fase)

Contra a caixa real, somente leitura. Perguntas, cada uma com resposta medida:

| | Pergunta |
|---|---|
| **P1** | `MSIP_Labels` existe nesta conta? Sob qual DASL? |
| **P2** | Dá para ler sem prompt do Object Model guard? |
| **P3** | Em que fração das mensagens aparece, e com que valores? |
| **P4** | O `Sensitivity` clássico concorda com o rótulo moderno? |
| **P5** | Existe mensagem cujo corpo o OOM recusa (IRM)? Como recusa? |
| **P6** | O rótulo vem por `Table` (barato) ou exige abrir o item? |
| **P7** | Como `PropertyAccessor` reage a propriedade **ausente**? |

A **P7 é a mais importante da lista**, e é por um motivo que já custou caro nesta
base: se "não tem rótulo" e "não consegui ler o rótulo" chegarem ao código como
o mesmo valor, o portão vai liberar mensagem classificada toda vez que a leitura
falhar. As duas têm de ser distinguíveis no tipo, não por convenção.

Saída: `tools/medir-purview.ps1`-equivalente em teste de integração, mais a
resposta escrita aqui.

### 3.1 — O portão (`DisclosurePolicy`)

Decide se um pedaço de conteúdo pode sair da máquina. Falha **fechada**, e em
graus, como o `EnvironmentPolicy` da Fase 2.

Nega, e diz por quê, quando:
- o rótulo não pôde ser lido — **ilegível é proibido**, não "sem rótulo";
- há rótulo e ele não está numa allowlist explícita;
- a pasta não foi habilitada explicitamente;
- não há autorização registrada do usuário;
- é anexo (fora de escopo nesta fase, por inteiro).

Prova obrigatória: **a configuração de produção autoriza zero**, com teste, como
a matriz da §22.

### 3.2 — A porta (`Iris.Assist`)

Projeto novo, `net10.0`, sem COM e sem WPF — como o `Iris.Sync`. Define
`IAssistant` e os DTOs. Nenhum provedor concreto mora aqui.

Um `AssistantRequest` **só pode ser construído a partir de conteúdo já liberado
pelo portão** — não por disciplina de quem chama, mas porque o construtor exige
o veredito. Se der para montar um pedido sem passar pelo portão, o portão é
decoração.

### 3.3 — O diário (`DisclosureLog`)

Tabela nova no cache: quando, qual item, veredito, hash SHA-256 do que saiu,
tamanho em bytes, modelo, desfecho. **Nunca o conteúdo** — um log com o texto
cria mais uma cópia sensível, e é o R11 dizendo isso.

Controle negativo obrigatório: plantar uma isca no corpo, rodar o caminho
inteiro, e varrer o banco atrás da isca. Achou, reprova.

### 3.4 — Provedores

- `AssistenteFalso`, determinístico, para teste.
- `AssistenteIndisponivel` — **o padrão de produção**. Recusa e diz por quê.
- `AssistenteHttp` para a Messages API da Anthropic, escrito e testado contra um
  servidor HTTP **local** falso. Sem chave, sem endereço de produção, desligado.

Prova: teste sobre o IL do `Iris.App` de que nenhum provedor de rede é
instanciado sem autorização — o mesmo instrumento da §26.2.

### 3.5 — A UI

Painel sobre a mensagem aberta:
- quando o portão nega, **o motivo aparece em português**, não um código;
- quando libera, o resumo aparece com o modelo e o tamanho enviados à vista;
- redigir produz **rascunho no composer**, nunca envio.

### 3.6 — Ponta a ponta, e o relatório

---

## 30. O que esta fase NÃO faz

- Não envia e-mail.
- Não manda nada para fora sem ato explícito do usuário, registrado.
- Não redige dados sensíveis automaticamente — o ESCOPO rebaixou isso por não
  ser barreira de compliance, e fingir que é seria pior que não ter.
- Não trata anexos.
- Não faz triagem nem busca semântica (Fase 4).

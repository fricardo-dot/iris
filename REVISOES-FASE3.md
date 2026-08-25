# Revisões da Fase 3 — passadas, achados e vereditos

O relatório afirma que os sete marcos foram aprovados pelo Codex e que não ficou
divergência em aberto. Os vereditos são trocas de conversa, não artefatos do
repositório — então esta é a transcrição resumida deles, feita por mim.

**Isto não é prova independente.** É o meu registro do que foi dito. O que o
repositório sustenta sozinho são os commits: cada correção tem o seu, com a
mensagem descrevendo o defeito. As colunas "achados" abaixo batem com os commits
citados, e é por aí que dá para conferir.

## Contagem

| Marco | Passadas | Achados numerados | Commits |
|---|---|---|---|
| Plano da fase | 3 revisões | — | — |
| 3.0 | rejeitado e corrigido | 4 (§34.4) | `95c19f1`, `fee83d1`, `e853b9b`, `9e11c20` |
| 3.1 | rejeitado e corrigido | ver commits | `95c790f`, `9c4c4f6`, `86f1cdf` |
| 3.2 | rejeitado e corrigido | ver commits | `277fd67`, `7d64a39`, `1a9fd34`, `b1c30de`, `b2fdcdf` |
| 3.3 | rejeitado e corrigido | ver commits | `04d2dea`, `c786070`, `be73453` |
| 3.4 | rejeitado e corrigido | ver commits | `b60b896`, `e6a8b62`, `24054af`, `8e3120e` |
| **3.5** | **7** | **14** | `61a8123`, `9455e7f`, `6f8bc7b`, `8167425`, `a25a6c0`, `84119b4`, `3c95c9f` |
| **3.6** | **4** (1 de plano + 3 de revisão) | **8** | `95416d7`, `b377b78`, `7c39cb5` |
| Relatório | 1 | 9 | este |

Os achados do 3.0 estão enumerados na §34.4 do `FASE3.md`. Para 3.1–3.4 o
enumerado é a mensagem de cada commit de correção — não recontei um a um, e por
isso o relatório só afirma o total onde ele foi contado: **27 achados nas onze
passadas do 3.5 e do 3.6**, mais 4 do 3.0.

## Marco 3.5 — sete passadas

| # | Achados | Veredito |
|---|---|---|
| 1 | — | REJEITADO |
| 2 | 5 — não havia botão; `Trocou()` não ligado à seleção; aviso da reconciliação podia ficar invisível; redigir sem desfazer; faltava o teste offscreen | REJEITADO |
| 3 | 4 — contexto de produção era imitação; `Trocou()` não reavaliava; redigir sobrescrevia edição concorrente; botões dentro do `Border` do aviso | REJEITADO |
| 4 | 6 — `temAnexo:=False` fixo; `Ocupado` travava para sempre; rascunho sem identidade nem `PodeEditar`; portão calculado só para `Resumir`; RCW sem dono no `TryCast`; faltavam controles negativos da produção | REJEITADO |
| 5 | 2 — execução guardada por `PodePedir`; `Desfazer` sem identidade de sessão | REJEITADO |
| 6 | 1 — `PodeDesfazer` recusava e ninguém avisava o WPF | REJEITADO |
| 7 | 1 — o elo `ComposerViewModel` → `RascunhoDoCompositor` sem prova | REJEITADO |
| 8 | — | **APROVADO** |

*(A passada 1 e a 2 aconteceram na sessão anterior; a contagem "sete passadas" do
relatório conta as rodadas de revisão, e a aprovação é a oitava troca.)*

## Marco 3.6 — quatro passadas

| # | Achados | Veredito |
|---|---|---|
| plano | 6 ajustes exigidos antes de escrever | APROVADO (plano) |
| 1 | 4 — resposta só de espaço; seleção móvel provava outra coisa; reconciliação que falhou sem prova ponta a ponta; `ClearAllPools` global e pastas temporárias abandonadas | REJEITADO |
| 2 | 4 — guarda `g.Cobre` desligada no commit pelo próprio roteiro de controle; décima classe ainda paralela; item não pedido parava no portão e não na cobertura; contagem 21 vs 25 | REJEITADO |
| 3 | — | **APROVADO** |

## Relatório — uma passada

Nove correções, todas aceitas e aplicadas: evidência da suíte não versionada;
contagem de achados incompatível com o histórico; "mecanismo inteiro" forte
demais; "0 bytes enviados" sem escopo; a limitação de **uma mensagem, não a
thread**, ausente; pendências incompletas; ressalva do teste de egress; a
inevitabilidade alegada do recorte de HTTPS; e os vereditos não auditáveis — que
é o que este arquivo passou a registrar.

## Divergências sem consenso

**Nenhuma.** Duas foram resolvidas por argumento dele, e as registro porque não
foram eu cedendo por cansaço:

- **Envelope vazio continua válido.** Eu perguntei se recusar no
  `EnvelopeBuilder` não seria mais seguro; ele respondeu que o builder não
  conhece o conjunto aprovado, que a cobertura exata pertence ao cofre, e que uma
  recusa ali seria defesa especial e incompleta.
- **Provar composição lendo o código-fonte é suficiente aqui.** Eu submeti o
  desconforto explicitamente; ele aceitou, porque a propriedade é estática, está
  acompanhada de testes executáveis do contexto real, e construir um
  `MainViewModel` exigiria Outlook, cache e Dispatcher.

E uma em que **eu cedi, porque ele estava certo**: eu escrevi que provar HTTPS
local "exigiria" um desvio de validação de certificado no código de produção. Não
exigiria — um certificado local confiado pelo sistema também resolveria. O
recorte continua, mas como decisão de custo e risco, e não como impossibilidade.

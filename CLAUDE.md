# Iris — notas para quem mexe no código

## VB é case-insensitive, e isso já custou sete bugs neste projeto

Um identificador local — variável, parâmetro, campo — **eclipsa** qualquer
membro ou tipo de mesmo nome, ignorando maiúsculas. O compilador não avisa
que houve sombra: ele só reclama muito depois, com uma mensagem que não
tem nada a ver com o problema.

Casos reais desta base, todos custando tempo de depuração:

| Nome local | O que eclipsou | Erro que apareceu |
|---|---|---|
| `path` | `System.IO.Path` | membro não encontrado |
| `POINT` | `System.Windows.Point` | conversão inválida |
| `RECT` | `System.Windows.Rect` | conversão inválida |
| `texto` | função `Texto()` do módulo | "lambda não pode ser convertida para Integer" |
| `legado` | função `Legado()` do teste | "tipo não pode ser inferido" |
| `lista` | função `Lista()` do teste | "tipo não pode ser inferido" |
| `voltarAEditar` | método `VoltarAEditar()` | "expressão não é um método" |

**Regra prática:** antes de nomear um local, procure o nome no arquivo
ignorando maiúsculas. Se já existe como método, tipo ou propriedade,
escolha outro. Nomes genéricos de domínio — `texto`, `lista`, `item`,
`conta`, `chave` — são os que mais colidem, justamente porque são bons
nomes para as duas coisas.

Quando a mensagem do compilador não fizer sentido nenhuma no ponto em que
aparece, suspeite disto antes de suspeitar do resto.

## RCW: nunca encadeie expressões COM

Cada acesso que **devolve outro objeto COM** pode criar um RCW
intermediário que ninguém libera. Propriedade escalar — `Count`, `Subject`,
`EntryID` — não cria RCW próprio; o perigo é a coleção ou o objeto no meio
da cadeia.

```vb
' NÃO
Dim n = pasta.Folders.Count

' SIM
Dim filhas As OL.Folders = Nothing
Try
    filhas = pasta.Folders
    Dim n = filhas.Count
Finally
    ComHelpers.Release(filhas)
End Try
```

Liberar em ordem inversa à aquisição. Isto é a regra R7 do ESCOPO.md, e
já foi violada quatro vezes — sempre em código que "só lia uma contagem".

## Leitura tem retry, mutação não

Regra sem exceção. `ReadAsync` pode repetir; `MutateAsync` não pode, porque
criar, `Save`, `Move`, `Delete` e `Send` não são idempotentes. Falha depois
de a mutação começar vira `ErrorKind.Ambiguous`, que **nunca** é retentável.

Operação que não é leitura pura nem deixa efeito no mundo — hoje só
`PrepareSend` — vai por `SemRetryAsync`. Não abra exceção dentro de
`ReadAsync`: uma regra com exceção escondida no código, enquanto o
contrato a declara absoluta, é o mesmo que não ter regra.

## Toda operação que salva devolve a identidade nova

O `EntryID` **pode** mudar num `Save` — não é garantido que mude, e é
justamente por isso que o código não pode apostar em nenhuma das duas
hipóteses. A regra é sempre relê-lo depois de qualquer operação capaz de
mudar a identidade, e devolver o item redescrito em vez de só o resultado
da ação. Isto já foi esquecido em `AddAttachment`, e o sintoma apareceu
longe: `NotFound` no envio seguinte.

## Verificação: teste verde não é prova

Um teste que passa prova que **aquele caminho** vale, não que a
propriedade vale. Ao corrigir uma corrida, procure as irmãs dela antes de
declarar a família coberta — neste projeto, a corrida "digitar durante a
gravação" foi corrigida e a "digitar durante a conferência" sobreviveu.

Todo bloqueio precisa de controle negativo: sem ele, um compositor que
simplesmente nunca envia passa em todos os testes de "não envia errado".
Quando o controle negativo for barato, confirme desfazendo a correção e
vendo o teste falhar.

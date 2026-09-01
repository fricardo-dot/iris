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
| `reconciliacao` | propriedade `Reconciliacao` | `NullReferenceException` na primeira leitura, longe daqui |
| `corpo` | função `Corpo()` do teste | "argumento não especificado para o parâmetro" |
| `versao` | função `Versao()` | "tipo não pode ser inferido" |
| `pasta` | função `Pasta()` | "tipo não pode ser inferido" |
| `destino` | função `Destino()` | "tipo não pode ser inferido" |
| `Contatos` | módulo `Contatos` do modelo | "ForaDoAlcance não é membro de ObservableCollection" |
| `anexado` | função `Anexado()` | "tipo não pode ser inferido" |
| `confianca` | função `Confianca()` | "Double não pode ser indexado" |
| `Count` (propriedade) | extensão LINQ `Count(Of T)` | "Integer não pode ser indexado" |
| `lote` | função `Lote()` do teste | "tipo não pode ser inferido a partir de expressão contendo `lote`" |
| `chave` | função `Chave()` do teste | "argumento não especificado para o parâmetro `id`" |
| `item` | função `Item()` do fake | "argumento não especificado para o parâmetro `k`" |

O `reconciliacao` é o pior da lista e merece nome: num construtor,
`Reconciliacao = reconciliacao` atribui o **parâmetro a ele mesmo**, a
propriedade fica `Nothing`, e o compilador não diz nada. Dentro de um
construtor, escreva `Me.Reconciliacao = reconciliacao`.

**Regra prática:** antes de nomear um local, procure o nome no arquivo
ignorando maiúsculas. Se já existe como método, tipo ou propriedade,
escolha outro. Nomes genéricos de domínio — `texto`, `lista`, `item`,
`conta`, `chave` — são os que mais colidem, justamente porque são bons
nomes para as duas coisas.

Quando a mensagem do compilador não fizer sentido nenhuma no ponto em que
aparece, suspeite disto antes de suspeitar do resto.

### Palavra reservada como nome de membro

Parente próximo, e o sintoma é igual: a mensagem aponta para longe.

| Nome | Onde | Erro que apareceu |
|---|---|---|
| `Partial` | membro de enum | "tipo não pode ser inferido", no arquivo todo |
| `Protected` | membro de enum | "Enum must end with a matching End Enum" |

Foram renomeados para `Incomplete` e `Restricted`. Colchetes (`[Protected]`)
resolvem o compilador e não resolvem a leitura — quem lê depois tropeça no
mesmo lugar.

### Comentário no meio de uma continuação implícita

Em VB, três coisas continuam a expressão sozinhas quando estão no fim da
linha — a **vírgula** de um inicializador, o **`&`** de uma concatenação e o
**`.`** de uma cadeia de métodos. Uma linha de **comentário** no meio quebra
essa continuação, e o erro sai **nas linhas seguintes**, longe da causa.

| Onde | O que continua | Erro que aparece |
|---|---|---|
| `{ Col("a"), Col("b") }` | vírgula | `')' expected`, `Syntax error` |
| `"SELECT …" & <br> "  WHERE …"` | `&` | `Expression expected` |
| `lista. <br> OrderBy(…). <br> ThenBy(…)` | `.` | `Identifier expected` |

Sete ocorrências até 01/09/2026 — três num dia só, nos três arquivos onde o
esquema do cache é declarado; depois no SQL da herança de rótulos e nas duas
cadeias de ordenação. O conserto é sempre o mesmo: **o comentário vai para
antes da instrução inteira**, nunca para dentro dela. Quando ele explica um
trecho do meio, diga qual — *"as duas últimas condições…"* — em vez de
colá-lo lá.

### PowerShell tem a mesma doença

O `tools/` é PowerShell, e lá também: `$pid` é somente-leitura (é o PID do
processo) e o PowerShell também é case-insensitive, então
`foreach ($pid in ...)` aborta a execução.

## Comentário de produção envelhece

Um XML-doc que diz **o que a produção monta hoje** vira mentira no dia em que
a composição muda — e ninguém o relê ao mexer no compositor. Aconteceu quatro
vezes de uma vez: quatro comentários afirmando que a produção usava
`ActivationRecord.DaProducao = Nothing` e um provedor indisponível, meses
depois de o carregador de ativação e o adaptador de OpenRouter terem entrado.

O estrago não é o comentário errado: é ele **convidar a remoção do código
certo**. Quem lê "a produção não usa isto" e encontra o carregador conclui que
o carregador é resto.

**Regra prática:** comentário sobre produção aponta para *quem compõe*
(`MontarAssistente`, `ActivationLoader`), e não para *qual implementação ele
escolheu numa revisão anterior*. Se precisar nomear a implementação, diga a
data e o motivo — assim o próximo sabe que é um retrato, e não uma regra.

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

# Resultado da suíte

Este arquivo existe porque os relatórios afirmam "N testes, 0 falhas, 0 pulados",
e `TestResults/` é ignorado pelo git. Sem isto, os dois zeros seriam uma
afirmação minha que ninguém consegue conferir pelo repositório.

Reproduza com:

```
dotnet test Iris.slnx
```

## Medição corrente — depois da vigésima sétima revisão

| | |
|---|---|
| **Commit** | `1214fa9` — a árvore da solução .NET que foi medida |
| **Data** | 29 de agosto de 2026, de manhã |
| **SDK** | .NET 10.0.301 |
| **Alvo** | `net10.0-windows` |
| **Máquina** | Windows 11 Pro 10.0.26200 |

```
Passed!  - Failed:     0, Passed:   946, Skipped:     0, Total:   946, Duration: 1 m 7 s - Iris.Tests.dll (net10.0)
```

É este o número que `RELATORIO-TRABALHO-AUTONOMO.html` cita. A execução foi a
última antes do commit `1214fa9`, que contém **exatamente** a árvore medida —
`src/`, `tests/` e `tools/` entraram nele; os documentos vieram depois, em
commit separado.

Esta frase já errou duas vezes, nas duas direções: uma dizendo "só documentos"
quando um roteiro executável tinha mudado, e a correção seguinte repetindo a
omissão. Ela agora diz **o que o commit contém**, e não o que eu lembro de ter
tocado.

**A conta não fecha por soma, e é de propósito.** A nona passada acrescentou
três testes e **apagou um** — o `O_acumulador_da_linha_zera_entre_linhas` fazia
os dois resets com a própria mão e passaria com a correção desfeita. A décima
acrescentou um, que mede a dívida das duas gerações pendentes com o acervo real.
A décima primeira acrescentou o da agenda com zero compromissos, a décima
segunda não acrescentou teste nenhum — reforçou dois que já existiam —, e a
décima terceira acrescentou dois: a lista vazia com item ignorado, e o controle
positivo da pasta legitimamente vazia. A décima quarta acrescentou quatro — a
contagem que falhou, o total herdado da pasta anterior, os recusados
desconhecidos da agenda, e a prova de alcance do XAML. A décima quinta
acrescentou um: a troca de pasta durante uma carga. A décima sexta não
acrescentou teste — reforçou o mesmo, que passava com a correção desfeita. A
décima sétima acrescentou um: o contrato do duplo de lançar na hora. A
varredura própria acrescentou dois, do calendário que fabricava calado. A
décima oitava acrescentou sete: dois da reconciliação do diário de divulgação,
quatro do fechamento de bloco com espaço, e um da identidade do anexo. A
décima nona acrescentou três casos aos irmãos do fechamento.
872 → 874 → 875 → 876 → 878 → 882 → 883 → 884 → 886 → 893 → 896.

### Medição depois da vigésima sexta revisão

| | |
|---|---|
| **Commit** | `349056f` |
| **Data** | 29 de agosto de 2026, de manhã |

```
Passed!  - Failed:     0, Passed:   940, Skipped:     0, Total:   940, Duration: 1 m 9 s - Iris.Tests.dll (net10.0)
```

### Medição depois da vigésima quinta revisão

| | |
|---|---|
| **Commit** | `2f3edcb` |
| **Data** | 29 de agosto de 2026, de manhã |

```
Passed!  - Failed:     0, Passed:   933, Skipped:     0, Total:   933, Duration: 1 m 11 s - Iris.Tests.dll (net10.0)
```

### Medição depois da vigésima quarta revisão

| | |
|---|---|
| **Commit** | `c615e9c` |
| **Data** | 29 de agosto de 2026, de manhã |

```
Passed!  - Failed:     0, Passed:   927, Skipped:     0, Total:   927, Duration: 1 m 7 s - Iris.Tests.dll (net10.0)
```

### Medição depois da vigésima terceira revisão

| | |
|---|---|
| **Commit** | `4592cfc` |
| **Data** | 29 de agosto de 2026, de manhã |

```
Passed!  - Failed:     0, Passed:   918, Skipped:     0, Total:   918, Duration: 1 m 18 s - Iris.Tests.dll (net10.0)
```

### Medição depois do autômato testado

| | |
|---|---|
| **Commit** | `1dfd782` |
| **Data** | 29 de agosto de 2026, de manhã |

```
Passed!  - Failed:     0, Passed:   913, Skipped:     0, Total:   913, Duration: 1 m 7 s - Iris.Tests.dll (net10.0)
```

### Medição da vigésima segunda revisão

| | |
|---|---|
| **Commit** | `423906a` |
| **Data** | 29 de agosto de 2026, de manhã |

```
Passed!  - Failed:     0, Passed:   907, Skipped:     0, Total:   907, Duration: 1 m 8 s - Iris.Tests.dll (net10.0)
```

### Medição da vigésima primeira revisão

| | |
|---|---|
| **Commit** | `92259de` |
| **Data** | 29 de agosto de 2026, de manhã |

```
Passed!  - Failed:     0, Passed:   903, Skipped:     0, Total:   903, Duration: 1 m 7 s - Iris.Tests.dll (net10.0)
```

### Medição da vigésima revisão

| | |
|---|---|
| **Commit** | `8162120` |
| **Data** | 29 de agosto de 2026, de manhã |

```
Passed!  - Failed:     0, Passed:   899, Skipped:     0, Total:   899, Duration: 1 m 8 s - Iris.Tests.dll (net10.0)
```

### Medição da décima nona revisão

| | |
|---|---|
| **Commit** | `46b9d04` |
| **Data** | 29 de agosto de 2026, de manhã |

```
Passed!  - Failed:     0, Passed:   896, Skipped:     0, Total:   896, Duration: 1 m 16 s - Iris.Tests.dll (net10.0)
```

### Medição da décima oitava revisão

| | |
|---|---|
| **Commit** | `9b18122` |
| **Data** | 29 de agosto de 2026, de manhã |

```
Passed!  - Failed:     0, Passed:   893, Skipped:     0, Total:   893, Duration: 1 m 13 s - Iris.Tests.dll (net10.0)
```

### Medição da varredura própria

| | |
|---|---|
| **Commit** | `e574506` |
| **Data** | 28 de agosto de 2026, à noite |

```
Passed!  - Failed:     0, Passed:   886, Skipped:     0, Total:   886, Duration: 1 m 17 s - Iris.Tests.dll (net10.0)
```

### Medição da décima sétima revisão

| | |
|---|---|
| **Commit** | `ec575bd` |
| **Data** | 28 de agosto de 2026, à noite |

```
Passed!  - Failed:     0, Passed:   884, Skipped:     0, Total:   884, Duration: 1 m 8 s - Iris.Tests.dll (net10.0)
```

### Medição da décima sexta revisão

| | |
|---|---|
| **Commit** | `f61119c` |
| **Data** | 28 de agosto de 2026, à noite |

```
Passed!  - Failed:     0, Passed:   883, Skipped:     0, Total:   883, Duration: 1 m 11 s - Iris.Tests.dll (net10.0)
```

### Medição da décima quinta revisão

| | |
|---|---|
| **Commit** | `a011ae3` |
| **Data** | 28 de agosto de 2026, à noite |

```
Passed!  - Failed:     0, Passed:   883, Skipped:     0, Total:   883, Duration: 1 m 7 s - Iris.Tests.dll (net10.0)
```

### Medição da décima quarta revisão

| | |
|---|---|
| **Commit** | `e3b44c6` |
| **Data** | 28 de agosto de 2026, à noite |

```
Passed!  - Failed:     0, Passed:   882, Skipped:     0, Total:   882, Duration: 1 m 13 s - Iris.Tests.dll (net10.0)
```

### Medição da décima terceira revisão

| | |
|---|---|
| **Commit** | `783b8e2` |
| **Data** | 28 de agosto de 2026, à noite |

```
Passed!  - Failed:     0, Passed:   878, Skipped:     0, Total:   878, Duration: 1 m 18 s - Iris.Tests.dll (net10.0)
```

### Medição da décima segunda revisão

| | |
|---|---|
| **Commit** | `23df5ec` |
| **Data** | 28 de agosto de 2026, à noite |

```
Passed!  - Failed:     0, Passed:   876, Skipped:     0, Total:   876, Duration: 1 m 9 s - Iris.Tests.dll (net10.0)
```

### Medição da décima primeira revisão

| | |
|---|---|
| **Commit** | `c96be34` |
| **Data** | 28 de agosto de 2026, à noite |

```
Passed!  - Failed:     0, Passed:   876, Skipped:     0, Total:   876, Duration: 1 m 9 s - Iris.Tests.dll (net10.0)
```

### Medição da décima revisão

| | |
|---|---|
| **Commit** | `602d8e8` |
| **Data** | 28 de agosto de 2026, à noite |

```
Passed!  - Failed:     0, Passed:   875, Skipped:     0, Total:   875, Duration: 1 m 6 s - Iris.Tests.dll (net10.0)
```

### Medição da nona revisão

| | |
|---|---|
| **Commit** | `ff37f0d` |
| **Data** | 28 de agosto de 2026, à noite |

```
Passed!  - Failed:     0, Passed:   874, Skipped:     0, Total:   874, Duration: 1 m 12 s - Iris.Tests.dll (net10.0)
```

### Medição anterior — fim do dia autônomo

| | |
|---|---|
| **Commit** | `118277f` |
| **Data** | 28 de agosto de 2026, à tarde |

```
Passed!  - Failed:     0, Passed:   872, Skipped:     0, Total:   872, Duration: 1 m 4 s - Iris.Tests.dll (net10.0)
```

Fica registrada porque a revisão externa apontou que os relatórios diziam 870
enquanto este arquivo registrava 805 — a afirmação existia sem a evidência
versionada que o próprio projeto exige.

## Medição do fechamento da Fase 2

| | |
|---|---|
| **Commit** | `ef7096a` |
| **Data** | 28 de agosto de 2026, de manhã |
| **SDK** | .NET 10.0.301 |
| **Alvo** | `net10.0-windows` |
| **Máquina** | Windows 11 Pro 10.0.26200 |

```
Test run for C:\Users\Ricardo\Documents\Iris\tests\Iris.Tests\bin\Debug\net10.0-windows\Iris.Tests.dll (.NETCoreApp,Version=v10.0)
A total of 1 test files matched the specified pattern.

Passed!  - Failed:     0, Passed:   805, Skipped:     0, Total:   805, Duration: 1 m 1 s - Iris.Tests.dll (net10.0)
```

É o número que `RELATORIO-FASE2-FECHAMENTO.html` cita. Foi medido **depois** de
todas as correções das seis passadas de revisão daquele bloco, e antes das
correções de redação em documentos — que não são código executável.

## Medição anterior — fechamento da Fase 3

| | |
|---|---|
| **Commit** | `33ecc84` |
| **Data** | 25 de agosto de 2026 |

```
Passed!  - Failed:     0, Passed:   642, Skipped:     0, Total:   642, Duration: 59 s - Iris.Tests.dll (net10.0)
```

É o número de `RELATORIO-FASE3.html`. Fica registrado porque aquele relatório
continua citando o resultado da árvore dele, e substituir a medição sem deixar
rastro faria um documento correto para a sua data parecer errado.

O último commit do **ciclo técnico original** é o `433c244`. Depois dele entraram
documentação, estes arquivos de evidência, correções de redação em comentários de
teste e mudanças na **ferramenta de verificação** `tools/controle-negativo.py` —
que é executável, e não documentação.

Esta medição é do `33ecc84`, e **não** do `433c244`: um resultado de suíte vale
para a árvore em que foi medido, e não para a que se gostaria que ele valesse. Ela
continua valendo enquanto o **código executável da solução .NET** não mudar. A
ferramenta de verificação não faz parte dele; e o que houve em `tests/` depois de
`433c244` foi alteração de **comentário** em dois arquivos, no `33ecc84`, que é
justamente a árvore medida. Se mudar código executável, este arquivo tem de ser
remedido.

## O que este número não diz

**Não diz que a suíte é determinística.** Três vezes ao longo de 25/08 um teste
falhou com `Cannot access a disposed object: SQLitePCL.sqlite3` e não reproduziu
na execução seguinte — sempre na primeira execução depois de uma compilação. A
causa provável é `SqliteConnection.ClearAllPools()`, que é global, derrubando a
conexão de uma classe vizinha rodando em paralelo. As dez classes que tocam
SQLite passaram a `<DoNotParallelize>`, e a duração subiu de 38 s para ~58 s.
Depois disso não houve nova ocorrência em mais de dez execuções — mas **a causa
continua hipótese**, e uma das três ocorrências não chegou a ser identificada por
nome.

**Não diz que as guardas estão ligadas.** Teste verde não é prova. O que sustenta
essa parte é `tools/controle-negativo.py`, que desliga **as guardas que ele
enumera** — cobertura e proveniência da capability, anexo no pipeline, espaço em
branco no resultado, e a corrida de seleção.

Ele **não** exige que toda mutação fique vermelha: cada cenário declara o
desfecho esperado e o **conjunto exato** de testes que devem cair. As duas
redundâncias isoladas — cobertura sozinha, proveniência sozinha — esperam
**verde**, porque a outra guarda sozinha já segura, e estão ali para documentar
que são independentemente suficientes; a dupla e os demais cenários esperam
**vermelho**, com os testes nomeados. Qualquer divergência, nos dois sentidos,
faz o roteiro sair com código 1.

Não é varredura de todas as guardas do produto: é uma lista, e ela está no
arquivo.

**Não diz que a IA funciona contra um provedor real.** Nenhum teste chama um
provedor externo de IA: o HTTP do assistente usa só `127.0.0.1`, contra um
`HttpListener`, e a ordem e o diário são provados com provedor falso. A suíte
tem testes que falam com o **Outlook real**, e o comportamento de rede dele não
é controlado aqui.

Isso continua verdade depois da ativação de 27/08. O egress real aconteceu
**fora da suíte**, por uso do aplicativo, e a evidência dele mora no diário de
divulgação dentro do cache do usuário, que não está versionado. A suíte prova a
semântica do caminho; ela não prova que o caminho foi percorrido naquela data.
Quem quiser conferir o egress tem de olhar o `disclosure_log` da própria
máquina, e não este repositório.

**Não diz nada sobre os números da varredura real.** As 1.123 guardadas e as 12
recusadas vêm do mesmo cache não versionado. Mesma distinção: o repositório
sustenta a semântica, a máquina sustenta a medição.

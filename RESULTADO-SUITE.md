# Resultado da suíte

Este arquivo existe porque os relatórios afirmam "N testes, 0 falhas, 0 pulados",
e `TestResults/` é ignorado pelo git. Sem isto, os dois zeros seriam uma
afirmação minha que ninguém consegue conferir pelo repositório.

Reproduza com:

```
dotnet test Iris.slnx
```

## Medição corrente — depois da nona revisão

| | |
|---|---|
| **Commit** | `ff37f0d` — a árvore da solução .NET que foi medida |
| **Data** | 28 de agosto de 2026, à noite |
| **SDK** | .NET 10.0.301 |
| **Alvo** | `net10.0-windows` |
| **Máquina** | Windows 11 Pro 10.0.26200 |

```
Passed!  - Failed:     0, Passed:   874, Skipped:     0, Total:   874, Duration: 1 m 12 s - Iris.Tests.dll (net10.0)
```

É este o número que `RELATORIO-TRABALHO-AUTONOMO.html` cita. A execução foi a
última antes do commit, e o único arquivo tocado depois dela foi este diário —
que não é código executável.

**São 874 e não 875** porque a nona passada acrescentou três testes e **apagou
um**: o `O_acumulador_da_linha_zera_entre_linhas` fazia os dois resets com a
própria mão e passaria com a correção desfeita. Trocá-lo por um teste que dirige
a produção é o motivo de a conta não fechar por soma.

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

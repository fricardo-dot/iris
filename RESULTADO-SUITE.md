# Resultado da suíte — Fase 3

Este arquivo existe porque o relatório afirma "642 testes, 0 falhas, 0 pulados",
e `TestResults/` é ignorado pelo git. Sem isto, os dois zeros seriam uma
afirmação minha que ninguém consegue conferir pelo repositório.

Reproduza com:

```
dotnet test Iris.slnx
```

## Medição

| | |
|---|---|
| **Commit** | `33ecc84` — o commit de encerramento da fase |
| **Data** | 25 de agosto de 2026 |
| **SDK** | .NET 10.0.301 |
| **Alvo** | `net10.0-windows` |
| **Máquina** | Windows 11 Pro 10.0.26200 |

```
Test run for C:\Users\Ricardo\Documents\Iris\tests\Iris.Tests\bin\Debug\net10.0-windows\Iris.Tests.dll (.NETCoreApp,Version=v10.0)
A total of 1 test files matched the specified pattern.

Passed!  - Failed:     0, Passed:   642, Skipped:     0, Total:   642, Duration: 59 s - Iris.Tests.dll (net10.0)
```

O último commit **técnico** da fase é o `433c244`; daí em diante entraram só
documentação, este arquivo, o `REVISOES-FASE3.md` e correções de redação em
comentários de teste. Esta medição é do `33ecc84`, e **não** do `433c244` — a
distinção importa porque um resultado de suíte vale para a árvore em que foi
medido, e não para a que se gostaria que ele valesse.

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
branco no resultado, e a corrida de seleção — e confere que o teste de cada uma
**falha**. Não é varredura de todas as guardas do produto: é uma lista, e ela
está no arquivo.

**Não diz que a IA funciona contra um provedor real.** Nenhum teste chama um
provedor externo de IA: o HTTP do assistente usa só `127.0.0.1`, contra um
`HttpListener`, e a ordem e o diário são provados com provedor falso. A suíte
tem testes que falam com o **Outlook real**, e o comportamento de rede dele não
é controlado aqui.

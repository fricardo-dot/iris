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
| **Commit** | `433c244` |
| **Data** | 25 de agosto de 2026 |
| **SDK** | .NET 10.0.301 |
| **Alvo** | `net10.0-windows` |
| **Máquina** | Windows 11 Pro 10.0.26200 |

```
Test run for C:\Users\Ricardo\Documents\Iris\tests\Iris.Tests\bin\Debug\net10.0-windows\Iris.Tests.dll (.NETCoreApp,Version=v10.0)
A total of 1 test files matched the specified pattern.

Passed!  - Failed:     0, Passed:   642, Skipped:     0, Total:   642, Duration: 58 s - Iris.Tests.dll (net10.0)
```

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
essa parte é `tools/controle-negativo.py`, que desliga cada guarda e confere que
o teste dela **falha**.

**Não diz que a IA funciona contra um provedor real.** Nenhum teste toca a
internet. O transporte é provado contra um `HttpListener` em `127.0.0.1`, e a
ordem e o diário com provedor falso.

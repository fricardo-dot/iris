# -*- coding: utf-8 -*-
"""Atualiza os NUMEROS de RESULTADO-SUITE.md e RELATORIO-TRABALHO-AUTONOMO.html.

NAO toca no DIARIO-AUTONOMO.md: o diario e prosa por dia, escrita a mao. A
primeira versao deste cabecalho dizia "os tres documentos", e a revisao externa
pegou -- roteiro que descreve errado o que faz e a mesma familia dos zeros
fabricados, num lugar diferente.

Existe porque eu ja errei essa atualizacao duas vezes a mao: uma vez o script
abortou no meio e nao gravou (o relatorio ficou dizendo 17 no cabecalho e
dezoito no rodape), e outra eu somei o commit de cabeca em vez de medir.
Aqui a contagem vem do git, TODA substituicao e conferida, e as gravacoes so
acontecem depois de as duas edicoes terem sido montadas em memoria.
"""
import io, os, subprocess, sys
os.chdir(r"C:\Users\Ricardo\Documents\Iris")

HASH_NOVO, ANT_HASH = sys.argv[1], sys.argv[2]
TESTES_NOVO, ANT_TESTES = sys.argv[3], sys.argv[4]
DUR = sys.argv[5]
ORD_NOVO, ORD_ANT = sys.argv[6], sys.argv[7]      # "vigésima", "décima nona"
PASSADA_NOVA, PASSADA_ANT = sys.argv[8], sys.argv[9]   # "20", "19"
DATA = sys.argv[10]

# check=True: sem ele, um git que falha devolveria contagem VAZIA e ela seria
# gravada como se fosse medida.
n = subprocess.run(["git", "rev-list", "--count", "874f223.." + HASH_NOVO],
                   capture_output=True, text=True, check=True).stdout.strip()
assert n.isdigit() and n != "0", f"contagem de commits invalida: {n!r}"

def edita(p, pares):
    s = io.open(p, encoding="utf-8", newline="").read()
    for a, b in pares:
        assert s.count(a) == 1, (p, s.count(a), a[:70])
        s = s.replace(a, b, 1)
    io.open(p, "w", encoding="utf-8", newline="").write(s)

edita("RESULTADO-SUITE.md", [
    (f"## Medição corrente — depois da {ORD_ANT} revisão",
     f"## Medição corrente — depois da {ORD_NOVO} revisão"),
    (f"| **Commit** | `{ANT_HASH}` — a árvore da solução .NET que foi medida |",
     f"| **Commit** | `{HASH_NOVO}` — a árvore da solução .NET que foi medida |"),
    (f"última antes do commit `{ANT_HASH}`, que contém",
     f"última antes do commit `{HASH_NOVO}`, que contém"),
])

s = io.open("RESULTADO-SUITE.md", encoding="utf-8", newline="").read()
nl = "\r\n" if "\r\n" in s else "\n"
linha_ant = f"Passed!  - Failed:     0, Passed:   {ANT_TESTES}, Skipped:     0, Total:   {ANT_TESTES},"
i = s.index(linha_ant)
fim = s.index(nl, i)
antiga = s[i:fim]
nova = (f"Passed!  - Failed:     0, Passed:   {TESTES_NOVO}, Skipped:     0, "
        f"Total:   {TESTES_NOVO}, Duration: {DUR} - Iris.Tests.dll (net10.0)")
s = s[:i] + nova + s[fim:]

marca = "### Medição da"
j = s.index(marca)
bloco = nl.join([
    f"### Medição da {ORD_ANT} revisão",
    "",
    "| | |",
    "|---|---|",
    f"| **Commit** | `{ANT_HASH}` |",
    f"| **Data** | {DATA} |",
    "",
    "```",
    antiga,
    "```",
    "",
    "",
])
s = s[:j] + bloco + s[j:]
io.open("RESULTADO-SUITE.md", "w", encoding="utf-8", newline="").write(s)

edita("RELATORIO-TRABALHO-AUTONOMO.html", [
    (f"<span><b>{ANT_TESTES}</b> testes · 0 falhas · 0 pulados</span>",
     f"<span><b>{TESTES_NOVO}</b> testes · 0 falhas · 0 pulados</span>"),
    (f"commits até o <code>{ANT_HASH}</code></span>",
     f"commits até o <code>{HASH_NOVO}</code></span>"),
    (f"<span><b>{PASSADA_ANT}</b> passadas de revisão externa</span>",
     f"<span><b>{PASSADA_NOVA}</b> passadas de revisão externa</span>"),
    (f"<span><code>874f223</code> → <code>{ANT_HASH}</code></span>",
     f"<span><code>874f223</code> → <code>{HASH_NOVO}</code></span>"),
    (f'<div><span class="v">{ANT_TESTES}</span><span class="k">testes ao terminar</span></div>',
     f'<div><span class="v">{TESTES_NOVO}</span><span class="k">testes ao terminar</span></div>'),
    (f"<code>874f223</code> → <code>{ANT_HASH}</code> · ",
     f"<code>874f223</code> → <code>{HASH_NOVO}</code> · "),
    (f"· {ANT_TESTES} testes, 0 falhas, 0 pulados<br>",
     f"· {TESTES_NOVO} testes, 0 falhas, 0 pulados<br>"),
])

# o numero de commits do cabecalho e do rodape, medido e nao somado
s = io.open("RELATORIO-TRABALHO-AUTONOMO.html", encoding="utf-8", newline="").read()
import re
s, q1 = re.subn(r"<span><b>\d+</b> commits até o", f"<span><b>{n}</b> commits até o", s, count=1)
s, q2 = re.subn(r"</code> · \d+ commits ·", f"</code> · {n} commits ·", s, count=1)
assert q1 == 1 and q2 == 1, f"contagem de commits nao encontrada no relatorio: {q1}, {q2}"
io.open("RELATORIO-TRABALHO-AUTONOMO.html", "w", encoding="utf-8", newline="").write(s)

print(f"documentos: {TESTES_NOVO} testes, {n} commits, {HASH_NOVO}, {PASSADA_NOVA} passadas")

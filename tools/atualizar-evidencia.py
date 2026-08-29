# -*- coding: utf-8 -*-
"""Atualiza os NUMEROS de RESULTADO-SUITE.md e RELATORIO-TRABALHO-AUTONOMO.html.

NAO toca no DIARIO-AUTONOMO.md: o diario e prosa por dia, escrita a mao. A
primeira versao deste cabecalho dizia "os tres documentos", e a revisao externa
pegou -- roteiro que descreve errado o que faz e a mesma familia dos zeros
fabricados, num lugar diferente.

Existe porque eu ja errei essa atualizacao duas vezes a mao: uma vez o script
abortou no meio e nao gravou (o relatorio ficou dizendo 17 no cabecalho e
dezoito no rodape), e outra eu somei o commit de cabeca em vez de medir.

Aqui a contagem vem do git e TODA substituicao e conferida. As duas edicoes sao
montadas em memoria e so entao gravadas -- a primeira versao ja gravava o
RESULTADO-SUITE antes de olhar o relatorio, e uma falha no segundo deixava os
dois documentos em desacordo. As conferencias sao SystemExit, e nao assert,
porque assert desaparece sob `python -O`.

Uso:
  python tools/atualizar-evidencia.py HASH_NOVO HASH_ANT TESTES_NOVO TESTES_ANT
      "1 m 7 s" "vigesima terceira" 23 22 "29 de agosto de 2026"
"""
import io
import os
import re
import subprocess
import sys

os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE = "874f223"
SUITE = "RESULTADO-SUITE.md"
RELATORIO = "RELATORIO-TRABALHO-AUTONOMO.html"


def confere(cond, msg):
    if not cond:
        raise SystemExit("atualizar-evidencia: " + msg)


confere(len(sys.argv) == 10, f"esperava 9 argumentos, recebi {len(sys.argv) - 1}")

HASH_NOVO, ANT_HASH = sys.argv[1], sys.argv[2]
TESTES_NOVO, ANT_TESTES = sys.argv[3], sys.argv[4]
DUR = sys.argv[5]
ORD_NOVO = sys.argv[6]
PASSADA_NOVA, PASSADA_ANT = sys.argv[7], sys.argv[8]
DATA = sys.argv[9]

# ORD_ANT saiu. Ele exigia que o cabecalho corrente fosse literalmente
# "depois da <ordinal> revisao", e uma passada em que eu NAO chamei de revisao
# ("depois do automato testado") quebrou o roteiro. O rotulo da medicao que vai
# para o arquivo agora vem do proprio cabecalho, lido na hora -- o arquivo
# descreve o que ele dizia, e nao o que eu supus que dizia.

# check=True: sem ele, um git que falha devolveria contagem VAZIA, e ela seria
# gravada como se tivesse sido medida.
n = subprocess.run(["git", "rev-list", "--count", BASE + ".." + HASH_NOVO],
                   capture_output=True, text=True, check=True).stdout.strip()
confere(n.isdigit() and n != "0", f"contagem de commits invalida: {n!r}")


def aplica(caminho, pares):
    """Devolve o texto ja editado. NAO grava: quem grava e o final."""
    s = io.open(caminho, encoding="utf-8", newline="").read()
    for a, b in pares:
        confere(s.count(a) == 1,
                f"{caminho}: {s.count(a)} ocorrencias de {a[:70]!r}")
        s = s.replace(a, b, 1)
    return s


suite_bruto = io.open(SUITE, encoding="utf-8", newline="").read()
m = re.search(r"^## Medição corrente — (.+)$", suite_bruto, re.MULTILINE)
confere(m is not None, "nao achei o cabecalho da medicao corrente")
ROTULO_ANT = m.group(1).strip()

suite = aplica(SUITE, [
    (f"## Medição corrente — {ROTULO_ANT}",
     f"## Medição corrente — depois da {ORD_NOVO} revisão"),
    (f"| **Commit** | `{ANT_HASH}` — a árvore da solução .NET que foi medida |",
     f"| **Commit** | `{HASH_NOVO}` — a árvore da solução .NET que foi medida |"),
    (f"última antes do commit `{ANT_HASH}`, que contém",
     f"última antes do commit `{HASH_NOVO}`, que contém"),
])

nl = "\r\n" if "\r\n" in suite else "\n"
linha_ant = (f"Passed!  - Failed:     0, Passed:   {ANT_TESTES}, "
             f"Skipped:     0, Total:   {ANT_TESTES},")
confere(suite.count(linha_ant) == 1,
        f"linha da medicao anterior ({ANT_TESTES}) nao encontrada")

i = suite.index(linha_ant)
fim = suite.index(nl, i)
antiga = suite[i:fim]
nova = (f"Passed!  - Failed:     0, Passed:   {TESTES_NOVO}, Skipped:     0, "
        f"Total:   {TESTES_NOVO}, Duration: {DUR} - Iris.Tests.dll (net10.0)")
suite = suite[:i] + nova + suite[fim:]

# "### Medição da" nao serve mais como marca: os rotulos arquivados agora
# copiam o cabecalho corrente, e um deles e "### Medição depois do automato
# testado". A marca tem de ser o prefixo comum.
marca = "### Medição "
confere(marca in suite, "nao achei onde inserir a medicao anterior")
j = suite.index(marca)
bloco = nl.join([
    f"### Medição {ROTULO_ANT}",
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
suite = suite[:j] + bloco + suite[j:]

relatorio = aplica(RELATORIO, [
    (f"<span><b>{ANT_TESTES}</b> testes · 0 falhas · 0 pulados</span>",
     f"<span><b>{TESTES_NOVO}</b> testes · 0 falhas · 0 pulados</span>"),
    (f"commits até o <code>{ANT_HASH}</code></span>",
     f"commits até o <code>{HASH_NOVO}</code></span>"),
    (f"<span><b>{PASSADA_ANT}</b> passadas de revisão externa</span>",
     f"<span><b>{PASSADA_NOVA}</b> passadas de revisão externa</span>"),
    (f"<span><code>{BASE}</code> → <code>{ANT_HASH}</code></span>",
     f"<span><code>{BASE}</code> → <code>{HASH_NOVO}</code></span>"),
    (f'<div><span class="v">{ANT_TESTES}</span>'
     f'<span class="k">testes ao terminar</span></div>',
     f'<div><span class="v">{TESTES_NOVO}</span>'
     f'<span class="k">testes ao terminar</span></div>'),
    (f"<code>{BASE}</code> → <code>{ANT_HASH}</code> · ",
     f"<code>{BASE}</code> → <code>{HASH_NOVO}</code> · "),
    (f"· {ANT_TESTES} testes, 0 falhas, 0 pulados<br>",
     f"· {TESTES_NOVO} testes, 0 falhas, 0 pulados<br>"),
])

# O numero de commits do cabecalho e do rodape: medido, e nao somado de cabeca.
relatorio, q1 = re.subn(r"<span><b>\d+</b> commits até o",
                        f"<span><b>{n}</b> commits até o", relatorio, count=1)
relatorio, q2 = re.subn(r"</code> · \d+ commits ·",
                        f"</code> · {n} commits ·", relatorio, count=1)
confere(q1 == 1 and q2 == 1,
        f"contagem de commits nao encontrada no relatorio: {q1}, {q2}")

# SO AGORA GRAVA. As duas edicoes ja passaram por todas as conferencias.
io.open(SUITE, "w", encoding="utf-8", newline="").write(suite)
io.open(RELATORIO, "w", encoding="utf-8", newline="").write(relatorio)

print(f"documentos: {TESTES_NOVO} testes, {n} commits, {HASH_NOVO}, "
      f"{PASSADA_NOVA} passadas")

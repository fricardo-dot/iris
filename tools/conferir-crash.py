#!/usr/bin/env python3
"""Confere o estado do banco depois de cada ponto de crash, SEM usar o Iris.

Existe porque a §22.1 afirma que os quatro estados foram lidos por um leitor
independente do codigo do teste, e uma afirmacao dessas precisa de artefato:
sem script no repositorio, ninguem consegue auditar se a leitura independente
aconteceu mesmo.

E o motivo de ser independente e concreto. Os testes reabrem o banco com
CacheDatabase.Open, que roda o SchemaGate e a introspeccao - o mesmo codigo
que esta sob teste. Se ele tivesse um defeito capaz de mascarar o estado no
disco, o teste nao veria. Este script usa o sqlite3 do proprio Python: outra
biblioteca, outro processo, nenhuma linha de Iris.

A PRIMEIRA versao deste script so imprimia contagens e cobrava tres
invariantes fracos - geracao pareada com divida, cabeca com geracao, linhas
com encarnacoes. Isso nao confere nada: um escritor que publicasse tudo antes
de cada ponto passaria, e o defeito 'checkpoint-antes', que motivou o desenho
inteiro, tambem passaria - nenhum dos tres cobrava "cursor avancado com zero
linhas". Agora cada ponto tem ESTADO ESPERADO, e o ultimo cenario e o
CONTROLE NEGATIVO: liga o defeito e exige que a conferencia ACUSE.

Uso:
    python tools/conferir-crash.py
"""
import os
import shutil
import sqlite3
import subprocess
import sys
import tempfile
from pathlib import Path

# O harness grava 3 paginas de 2 linhas e publica.
TOTAL = 6

# O estado que cada ponto TEM de produzir. None = nao interessa.
#   morre   -> o processo devia morrer (exit != 0)?
ESPERADO = {
    "dentro-da-pagina-antes-do-commit": dict(
        morre=True, stage="aberta", cursor=None, linhas=0,
        geracoes=0, dividas=0, nao_drenadas=0, cabeca=None),
    "depois-do-commit-da-pagina": dict(
        morre=True, stage="varrendo", cursor="cursor-1", linhas=2,
        geracoes=0, dividas=0, nao_drenadas=0, cabeca=None),
    "dentro-da-publicacao-antes-do-commit": dict(
        morre=True, stage="varrendo", cursor="cursor-3", linhas=TOTAL,
        geracoes=0, dividas=0, nao_drenadas=0, cabeca=None),
    "depois-do-commit-da-publicacao": dict(
        morre=True, stage="publicada", cursor="cursor-3", linhas=TOTAL,
        geracoes=1, dividas=1, nao_drenadas=1, cabeca=1),
    "nenhum": dict(
        morre=False, stage="publicada", cursor="cursor-3", linhas=TOTAL,
        geracoes=1, dividas=1, nao_drenadas=1, cabeca=1),
}

CAMPOS = ("stage", "cursor", "linhas", "geracoes", "dividas", "nao_drenadas", "cabeca")


def raiz() -> Path:
    d = Path(__file__).resolve().parent
    while d != d.parent and not (d / "Iris.slnx").exists():
        d = d.parent
    if not (d / "Iris.slnx").exists():
        sys.exit("nao achei a raiz do repositorio")
    return d


def harness(r: Path) -> Path:
    base = r / "tools" / "Iris.CrashHarness" / "bin"
    achados = sorted(base.rglob("Iris.CrashHarness.exe"),
                     key=lambda p: p.stat().st_mtime, reverse=True)
    if not achados:
        sys.exit("harness nao compilado. Rode: dotnet build Iris.slnx")
    return achados[0]


def semear(db: str) -> None:
    """A fixture minima: um ambiente, um store, uma pasta."""
    c = sqlite3.connect(db)
    c.execute("PRAGMA foreign_keys=ON")
    c.execute("INSERT INTO environment_profile VALUES (1,'fp','teste',1,NULL,1,1)")
    c.execute("INSERT INTO store (store_key,provider_store_id) VALUES (1,'S')")
    c.execute("INSERT INTO folder (folder_key,store_key,provider_entry_id,name,"
              "published_generation_key,reconcile_epoch,stability) "
              "VALUES (1,1,'F',NULL,NULL,0,'estavel')")
    c.commit()
    c.close()


def estado(db: str) -> dict:
    c = sqlite3.connect(db)
    q = lambda s: c.execute(s).fetchone()[0]
    todos = lambda s: [x[0] for x in c.execute(s).fetchall()]
    e = {
        "stage": q("SELECT stage FROM scan_attempt"),
        "cursor": q("SELECT cursor FROM scan_attempt"),
        "linhas": q("SELECT COUNT(*) FROM scan_stage"),
        "encarnacoes": q("SELECT COUNT(*) FROM incarnation"),
        "geracoes": q("SELECT COUNT(*) FROM generation"),
        "dividas": q("SELECT COUNT(*) FROM publication_log"),
        "nao_drenadas": q("SELECT COUNT(*) FROM publication_log WHERE drained_at IS NULL"),
        "cabeca": q("SELECT published_generation_key FROM folder"),
        "ger_keys": sorted(todos("SELECT generation_key FROM generation")),
        "log_keys": sorted(todos("SELECT generation_key FROM publication_log")),
    }
    c.close()
    return e


def conferir(ponto: str, rc: int, e: dict, esp: dict) -> list:
    """Devolve a lista de divergencias. Vazia = tudo certo."""
    d = []

    if esp["morre"] and rc == 0:
        d.append(f"devia ter morrido, saiu com {rc}")
    if not esp["morre"] and rc != 0:
        d.append(f"nao devia ter morrido, saiu com {rc}")

    for campo in CAMPOS:
        alvo = esp[campo]
        real = e[campo]
        if campo == "cabeca":
            # So interessa presenca/ausencia: a chave e um rowid.
            if (alvo is None) != (real is None):
                d.append(f"cabeca: {real!r}, esperado {'ausente' if alvo is None else 'presente'}")
            continue
        if real != alvo:
            d.append(f"{campo}: {real!r}, esperado {alvo!r}")

    # Invariantes estruturais, alem do estado esperado.
    if e["linhas"] != e["encarnacoes"]:
        d.append(f"linhas encenadas ({e['linhas']}) != encarnacoes ({e['encarnacoes']})")
    if e["ger_keys"] != e["log_keys"]:
        # Contagem igual nao basta: as CHAVES tem de ser as mesmas.
        d.append(f"geracoes {e['ger_keys']} != divida {e['log_keys']}")
    if e["cabeca"] is not None and e["cabeca"] not in e["ger_keys"]:
        d.append(f"cabeca {e['cabeca']} nao aponta para geracao existente")
    if e["cursor"] is not None and e["linhas"] == 0:
        # O defeito 'checkpoint-antes' vive exatamente aqui: o cursor diz que
        # a pagina foi lida e nenhuma linha dela existe. A retomada comeca na
        # pagina seguinte e aquelas mensagens nunca mais sao lidas.
        d.append("cursor AVANCOU com zero linhas gravadas — perda silenciosa")

    return d


def rodar(exe: Path, ponto: str, defeito: str = "") -> tuple:
    tmp = tempfile.mkdtemp(prefix="iris-crash-")
    db = os.path.join(tmp, "c.db")
    try:
        # A primeira execucao so cria o schema (ela falha por falta de pasta).
        subprocess.run([str(exe), db, "1", "nenhum", "kill", "0", ""], capture_output=True)
        semear(db)
        p = subprocess.run([str(exe), db, "1", ponto, "kill", "0", defeito],
                           capture_output=True, text=True)
        if "Unhandled exception" in p.stderr:
            return None, p.stderr
        return p.returncode, estado(db)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main() -> int:
    r = raiz()
    exe = harness(r)
    print(f"harness: {exe.relative_to(r)}\n")

    largura = 36
    print(f"| {'ponto':{largura}} | {'stage':9} | {'cursor':8} | {'lin':3} | "
          f"{'ger':3} | {'div':3} | {'n-dren':6} | {'cabeca':6} | ok |")
    print("|" + "|".join("-" * n for n in (largura + 2, 11, 10, 5, 5, 5, 8, 8, 4)) + "|")

    falhas = 0
    for ponto, esp in ESPERADO.items():
        rc, e = rodar(exe, ponto)
        if rc is None:
            print(f"  ERRO: o harness explodiu em vez de morrer no ponto:\n{e}")
            falhas += 1
            continue

        d = conferir(ponto, rc, e, esp)
        print(f"| {ponto:{largura}} | {str(e['stage']):9} | {str(e['cursor']):8} | "
              f"{e['linhas']:3} | {e['geracoes']:3} | {e['dividas']:3} | "
              f"{e['nao_drenadas']:6} | {str(e['cabeca']):6} | {'ok' if not d else 'NAO':3}|")
        for x in d:
            print(f"      *** {x}")
        falhas += len(d)

    # ============================================================
    # CONTROLE NEGATIVO. Sem ele, tudo acima passaria identico num
    # escritor que nao grava nada — e, pior, passaria no proprio defeito
    # que motivou o desenho.
    print()
    print("=== controle negativo: defeito 'checkpoint-antes' ===")
    rc, e = rodar(exe, "dentro-da-pagina-antes-do-commit", defeito="checkpoint-antes")
    if rc is None:
        print(f"  ERRO: {e}")
        return 1

    d = conferir("dentro-da-pagina-antes-do-commit", rc,
                 e, ESPERADO["dentro-da-pagina-antes-do-commit"])
    print(f"  estado: stage={e['stage']} cursor={e['cursor']} linhas={e['linhas']}")
    if d:
        print(f"  ACUSOU {len(d)} divergencia(s), como devia:")
        for x in d:
            print(f"      - {x}")
    else:
        print("  *** NAO ACUSOU. A conferencia nao distingue o defeito que")
        print("  *** motivou o desenho — ela nao esta conferindo nada.")
        falhas += 1

    print()
    print("SEM DIVERGENCIA" if falhas == 0 else f"{falhas} DIVERGENCIA(S)")
    return 1 if falhas else 0


if __name__ == "__main__":
    raise SystemExit(main())

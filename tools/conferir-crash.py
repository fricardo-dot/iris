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

PONTOS = [
    "dentro-da-pagina-antes-do-commit",
    "depois-do-commit-da-pagina",
    "dentro-da-publicacao-antes-do-commit",
    "depois-do-commit-da-publicacao",
    "nenhum",
]


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
    r = {
        "stage": q("SELECT stage FROM scan_attempt"),
        "cursor": q("SELECT cursor FROM scan_attempt"),
        "linhas": q("SELECT COUNT(*) FROM scan_stage"),
        "encarnacoes": q("SELECT COUNT(*) FROM incarnation"),
        "geracoes": q("SELECT COUNT(*) FROM generation"),
        "dividas": q("SELECT COUNT(*) FROM publication_log"),
        "nao_drenadas": q("SELECT COUNT(*) FROM publication_log WHERE drained_at IS NULL"),
        "cabeca": q("SELECT published_generation_key FROM folder"),
    }
    c.close()
    return r


def main() -> int:
    r = raiz()
    exe = harness(r)
    print(f"harness: {exe.relative_to(r)}\n")

    cab = ("ponto", "saida", "stage", "cursor", "linhas", "ger", "div", "n-drenadas", "cabeca")
    print("| {:36} | {:5} | {:9} | {:8} | {:6} | {:3} | {:3} | {:10} | {:6} |".format(*cab))
    print("|" + "|".join("-" * n for n in (38, 7, 11, 10, 8, 5, 5, 12, 8)) + "|")

    falhas = 0
    for ponto in PONTOS:
        tmp = tempfile.mkdtemp(prefix="iris-crash-")
        db = os.path.join(tmp, "c.db")
        try:
            # A primeira execucao so cria o schema (ela falha por falta de pasta).
            subprocess.run([str(exe), db, "1", "nenhum", "kill", "0", ""],
                           capture_output=True)
            semear(db)

            p = subprocess.run([str(exe), db, "1", ponto, "kill", "0", ""],
                               capture_output=True, text=True)
            if "Unhandled exception" in p.stderr:
                print(f"  ERRO: o harness explodiu em vez de morrer no ponto:\n{p.stderr}")
                falhas += 1
                continue

            e = estado(db)
            print("| {:36} | {:5} | {:9} | {:8} | {:6} | {:3} | {:3} | {:10} | {:6} |".format(
                ponto, p.returncode, str(e["stage"]), str(e["cursor"]), e["linhas"],
                e["geracoes"], e["dividas"], e["nao_drenadas"], str(e["cabeca"])))

            # Os invariantes, cobrados aqui tambem - nao so na tabela.
            if e["geracoes"] != e["dividas"]:
                print("    *** geracao sem divida, ou divida sem geracao ***")
                falhas += 1
            if e["cabeca"] is None and e["geracoes"] != 0:
                print("    *** geracao publicada sem cabeca apontando para ela ***")
                falhas += 1
            if e["linhas"] != e["encarnacoes"]:
                print("    *** linhas encenadas != encarnacoes gravadas ***")
                falhas += 1
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    print()
    print("SEM DIVERGENCIA" if falhas == 0 else f"{falhas} DIVERGENCIA(S)")
    return 1 if falhas else 0


if __name__ == "__main__":
    raise SystemExit(main())

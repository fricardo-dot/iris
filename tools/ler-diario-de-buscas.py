# -*- coding: utf-8 -*-
"""
O DIARIO DE BUSCAS VIRANDO CANDIDATOS A CONSULTA POR SENTIDO.

===========================================================================
O QUE ELE PROCURA

Uma REFORMULACAO: duas buscas seguidas, perto no tempo, em que a primeira
achou pouco ou nada e a segunda achou -- e as duas nao compartilham
palavra. Isso e o par que a Fase 4 precisa:

    digitei "cobranca"  -> 0 achados
    digitei "fatura"    -> 3 achados     (dentro de 3 minutos, sem palavra em comum)

    => candidato: eu procuro por "cobranca" o que a caixa chama de "fatura"

Nao e prova. E CANDIDATO -- so o dono confirma que as duas buscas eram a
mesma intencao. Por isso a saida e uma lista para ele revisar, e nao um
numero para publicar.

===========================================================================
O QUE ELE NAO CONSEGUE VER

Busca que falhou e nao foi reformulada. Se voce procurou "cobranca", nao
achou e desistiu, o diario tem a linha com zero achados e nao tem com que
parear. Essas aparecem numa segunda lista -- "falhou e ninguem tentou de
novo" --, e elas SAO sinal, so que sem o outro lado do par.

===========================================================================
PRIVACIDADE

Le so o diario local, que por desenho tem apenas o termo digitado -- nunca
assunto, remetente ou EntryID. Nada sai da maquina.

Uso:  python tools/ler-diario-de-buscas.py [caminho-do-buscas.jsonl]
"""
import io
import json
import os
import sys
import unicodedata

# Duas buscas a mais de JANELA minutos uma da outra nao sao a mesma
# intencao -- sao duas sessoes. O numero e arbitrario no valor e nao no
# motivo, e esta aqui para ser discutido em vez de escondido.
JANELA_MINUTOS = 3

# "Achou pouco" e zero ou um. Dois ja e uma busca que funcionou o
# suficiente para a pessoa olhar antes de desistir.
POUCO = 1


def normalizar(s):
    if not s:
        return ""
    d = unicodedata.normalize("NFD", s)
    return unicodedata.normalize(
        "NFC", "".join(c for c in d if unicodedata.category(c) != "Mn")).lower()


def palavras(termo):
    return set(p for p in
               "".join(c if c.isalnum() else " " for c in normalizar(termo)).split()
               if len(p) >= 3)


def instante(iso):
    from datetime import datetime
    return datetime.fromisoformat(iso)


def ler(caminho):
    linhas = []
    ruins = 0
    with io.open(caminho, encoding="utf-8") as f:
        for l in f:
            l = l.strip()
            if not l:
                continue
            try:
                d = json.loads(l)
                d["_t"] = instante(d["quando"])
                linhas.append(d)
            except Exception:
                # LINHA QUEBRADA NAO DERRUBA A LEITURA, e nao some calada:
                # uma queda no meio de um append deixa meia linha, e o
                # arquivo continua servindo. O total vai no relatorio.
                ruins += 1
    return linhas, ruins


def main():
    caminho = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.environ.get("LOCALAPPDATA", ""), "Iris", "buscas.jsonl")

    if not os.path.exists(caminho):
        print("nao achei o diario em %s" % caminho)
        print("")
        print("Ele nasce na primeira busca feita no Iris. Se voce ainda nao")
        print("procurou nada desde que o registro foi ligado, e isso.")
        return

    buscas, ruins = ler(caminho)
    buscas.sort(key=lambda d: d["_t"])

    pares = []
    orfas = []

    for i, b in enumerate(buscas):
        achou_b = b.get("exatos", 0) + b.get("aproximados", 0)
        if achou_b > POUCO:
            continue

        # A PROXIMA busca, se houver, dentro da janela.
        seguinte = buscas[i + 1] if i + 1 < len(buscas) else None
        if seguinte is None:
            orfas.append(b)
            continue

        minutos = (seguinte["_t"] - b["_t"]).total_seconds() / 60.0
        if minutos > JANELA_MINUTOS:
            orfas.append(b)
            continue

        achou_s = seguinte.get("exatos", 0) + seguinte.get("aproximados", 0)
        if achou_s <= POUCO:
            continue     # as duas falharam: nao ha o lado que funcionou

        # PALAVRA EM COMUM DESQUALIFICA. "contrato" -> "contrato aditivo" e
        # a pessoa refinando a mesma busca, e nao trocando de vocabulario.
        # E o vocabulario e a coisa inteira que se quer medir.
        if palavras(b["termo"]) & palavras(seguinte["termo"]):
            continue

        pares.append((b, seguinte, minutos))

    print("DIARIO DE BUSCAS -> candidatos a consulta por sentido")
    print("buscas anotadas: %d%s" % (len(buscas),
                                     ("  (%d linha(s) ilegivel(is))" % ruins) if ruins else ""))
    print("")

    if pares:
        print("REFORMULACOES -- digitei uma coisa, nao achei, digitei outra e achei:")
        print("")
        for b, s, m in pares:
            print("  %-28s (%d achados)" % (repr(b["termo"]), b.get("exatos", 0) + b.get("aproximados", 0)))
            print("  -> %-25s (%d achados, %.1f min depois)"
                  % (repr(s["termo"]), s.get("exatos", 0) + s.get("aproximados", 0), m))
            print("")
        print("CADA UM DESTES E CANDIDATO, E NAO PROVA. So voce sabe se as duas")
        print("buscas eram a mesma intencao. As que forem, viram linha em")
        print("tools/consultas-por-sentido.json -- 'digitei' e a primeira, e o")
        print("'queria' e o assunto da mensagem que a segunda achou.")
    else:
        print("Nenhuma reformulacao ate agora.")
        print("")
        print("Isso NAO quer dizer que a busca por sentido funciona -- quer")
        print("dizer que ainda nao houve caso no diario. Ausencia de sinal em")
        print("amostra pequena e ausencia de amostra.")

    print("")
    if orfas:
        print("FALHARAM E NINGUEM TENTOU DE NOVO (%d):" % len(orfas))
        for b in orfas[:15]:
            print("  %-28s (%d achados)"
                  % (repr(b["termo"]), b.get("exatos", 0) + b.get("aproximados", 0)))
        if len(orfas) > 15:
            print("  ... e mais %d" % (len(orfas) - 15))
        print("")
        print("Estas SAO sinal, e sem o outro lado do par: a busca nao achou e")
        print("a pessoa desistiu. Se voce lembrar qual mensagem queria em")
        print("alguma delas, ela vira uma linha completa.")


if __name__ == "__main__":
    main()

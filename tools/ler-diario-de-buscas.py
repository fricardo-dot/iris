# -*- coding: utf-8 -*-
"""
O DIARIO DE BUSCAS VIRANDO CANDIDATOS A CONSULTA POR SENTIDO.

===========================================================================
O QUE ELE PROCURA

Um EPISODIO: buscas encadeadas, cada uma perto no tempo da anterior, em que
as primeiras acharam pouco ou nada e a ultima achou. Isso e o par que a
Fase 4 precisa:

    digitei "cobranca" -> 0 achados
    digitei "boleto"   -> 0 achados
    digitei "fatura"   -> 5 achados

    => candidato: eu procuro por "cobranca" o que a caixa chama de "fatura"

A PRIMEIRA VERSAO SO OLHAVA A BUSCA SEGUINTE, e perdia exatamente este
caso: "cobranca" nao virava par (porque "boleto" tambem falhou) nem orfa
(porque havia uma seguinte). Sumia. A revisao externa pegou.

===========================================================================
A REGRA DE VOCABULARIO, E POR QUE ELA MUDOU

Descartar o par quando as duas buscas compartilham QUALQUER palavra era
forte demais, e perdia o caso mais comum de todos:

    "contrato fornecedor" -> "acordo fornecedor"

A palavra em comum e o contexto; a trocada e a falha de vocabulario, que e
justamente o que se quer medir. A regra agora e outra: descarta quando a
busca seguinte CONTEM todas as palavras da anterior -- isso e refino
("contrato" -> "contrato aditivo"), e nao troca de palavra.

===========================================================================
O QUE ELE NAO CONSEGUE VER

Busca que falhou e nao foi reformulada. Se voce procurou "cobranca", nao
achou e desistiu, o diario tem a linha e nao tem com que parear. Essas
aparecem numa segunda lista -- e SAO sinal, so que com um lado so.

===========================================================================
PRIVACIDADE -- LEIA ANTES DE RODAR EM TELA COMPARTILHADA

Le so o diario local, que por desenho tem apenas o termo digitado. Nada sai
da maquina POR ESTE SCRIPT.

Mas ele IMPRIME OS TERMOS INTEIROS na saida, e termo de busca de caixa
corporativa costuma conter nome de pessoa, numero de contrato ou valor. A
saida vai para o terminal, e de la para onde voce mandar: arquivo,
compartilhamento de tela, anexo. Isso e exposicao, e o script nao tem como
impedi-la -- so avisar.

Uso:  python tools/ler-diario-de-buscas.py [caminho] [--pouco N] [--autoteste]
"""
import io
import json
import os
import sys
import unicodedata
from datetime import datetime, timedelta

# Duas buscas a mais de JANELA minutos uma da outra nao sao a mesma
# intencao -- sao duas sessoes. Arbitrario no valor, e nao no motivo.
JANELA_MINUTOS = 3

# "Achou pouco" e, por padrao, zero.
#
# A PRIMEIRA VERSAO USAVA 1, e a revisao externa apontou: numa caixa
# corporativa uma busca que devolve UMA mensagem costuma ser sucesso
# perfeito, e conta-la como falha fabrica candidato "1 -> muitos". Como o
# desenho nao registra clique, nao ha sinal para distinguir "um resultado
# certo" de "um resultado errado" -- entao o padrao passou a ser o unico
# valor sobre o qual nao ha duvida, e o resto e --pouco.
POUCO = 0


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


def e_refino(antes, depois):
    """A segunda busca e a primeira com mais palavras?

    Refino ("contrato" -> "contrato aditivo") nao e troca de vocabulario, e
    e o unico caso que precisa ser descartado. Compartilhar UMA palavra nao
    descarta: em "contrato fornecedor" -> "acordo fornecedor", a comum e o
    contexto e a trocada e o achado."""
    a, b = palavras(antes), palavras(depois)
    if not a:
        return False
    return a.issubset(b)


def achou(reg, pouco):
    return reg.get("exatos", 0) + reg.get("aproximados", 0) > pouco


def episodios(buscas, janela_min):
    """Parte a lista em blocos de buscas encadeadas no tempo."""
    saida, atual = [], []
    for b in buscas:
        if atual and (b["_t"] - atual[-1]["_t"]) > timedelta(minutes=janela_min):
            saida.append(atual)
            atual = []
        atual.append(b)
    if atual:
        saida.append(atual)
    return saida


def classificar(buscas, pouco=POUCO, janela_min=JANELA_MINUTOS):
    """Devolve (pares, orfas, refinos).

    pares:   (primeira_que_falhou, a_que_achou, minutos, cadeia_do_meio)
    orfas:   falharam e nao tiveram sucesso depois no episodio
    refinos: falharam, mas a que achou era so a mesma busca com mais palavras

    TODA BUSCA QUE FALHOU CAI EM UMA DAS TRES LISTAS, e isso e o conserto.
    A primeira versao descartava os refinos em silencio -- "contrato" -> 0
    seguido de "contrato aditivo" -> 3 nao virava par (por ser refino) nem
    orfa (porque houve sucesso): sumia. A revisao externa achou, e o pior e
    que o autoteste nao pegava, porque so conferia que NAO havia par.

    Sumir e a coisa que esta ferramenta existe para nao fazer."""
    pares, orfas, refinos = [], [], []

    for bloco in episodios(sorted(buscas, key=lambda d: d["_t"]), janela_min):
        falhas = []
        for b in bloco:
            if not achou(b, pouco):
                falhas.append(b)
                continue

            # ACHOU. As falhas acumuladas se pareiam com esta -- menos as
            # que sao so refino dela, que nao sao troca de vocabulario.
            candidatas = [f for f in falhas if not e_refino(f["termo"], b["termo"])]
            refinos.extend(f for f in falhas if e_refino(f["termo"], b["termo"]))
            if candidatas:
                primeira = candidatas[0]
                minutos = (b["_t"] - primeira["_t"]).total_seconds() / 60.0
                pares.append((primeira, b, minutos, candidatas[1:]))
            falhas = []

        orfas.extend(falhas)

    return pares, orfas, refinos


def ler(caminho):
    linhas, ruins = [], 0
    with io.open(caminho, encoding="utf-8") as f:
        for l in f:
            l = l.strip()
            if not l:
                continue
            try:
                d = json.loads(l)
                d["_t"] = datetime.fromisoformat(d["quando"])
                linhas.append(d)
            except Exception:
                # LINHA QUEBRADA NAO DERRUBA A LEITURA, e nao some calada:
                # uma queda no meio de um append deixa meia linha. O total
                # vai no relatorio.
                ruins += 1
    return linhas, ruins


# ==========================================================================

def autoteste():
    """Casos do proprio classificador, para ele nao ser a unica parte desta
    medicao sem teste. A suite do projeto e VB e nao roda Python."""
    def reg(minuto, termo, n):
        return {"quando": "x", "termo": termo, "exatos": n, "aproximados": 0,
                "_t": datetime(2026, 8, 30, 9, minuto, 0)}

    falhas = []

    def conferir(nome, ok):
        if not ok:
            falhas.append(nome)

    def total(p, o, r):
        """Toda falha tem de aparecer em alguma lista."""
        return len(o) + len(r) + sum(1 + len(x[3]) for x in p)

    # CADEIA: o caso que a primeira versao perdia.
    p, o, r = classificar([reg(0, "cobranca", 0), reg(1, "boleto", 0), reg(2, "fatura", 5)])
    conferir("cadeia produz um par", len(p) == 1)
    conferir("cadeia pareia a PRIMEIRA falha", p and p[0][0]["termo"] == "cobranca")
    conferir("cadeia guarda o meio", p and [x["termo"] for x in p[0][3]] == ["boleto"])
    conferir("cadeia nao deixa orfa", not o)

    conferir("cadeia nao perde falha", total(p, o, r) == 2)

    # REFINO nao e troca de vocabulario -- MAS NAO SOME.
    p, o, r = classificar([reg(0, "contrato", 0), reg(1, "contrato aditivo", 3)])
    conferir("refino nao vira par", not p)
    conferir("refino NAO SOME", len(r) == 1 and r[0]["termo"] == "contrato")
    conferir("refino nao vira orfa", not o)
    conferir("refino nao perde falha", total(p, o, r) == 1)

    # PALAVRA EM COMUM com troca E par -- o caso que a regra antiga perdia.
    p, o, r = classificar([reg(0, "contrato fornecedor", 0), reg(1, "acordo fornecedor", 4)])
    conferir("troca com contexto comum vira par", len(p) == 1)

    # FORA DA JANELA: episodios diferentes.
    p, o, r = classificar([reg(0, "cobranca", 0), reg(30, "fatura", 5)])
    conferir("fora da janela nao pareia", not p)
    conferir("fora da janela vira orfa", len(o) == 1)

    # DESISTIU: falhou e acabou.
    p, o, r = classificar([reg(0, "jacare bicicleta", 0)])
    conferir("desistencia vira orfa", len(o) == 1 and not p)

    # POUCO: com o padrao 0, uma busca com 1 achado e SUCESSO.
    p, o, r = classificar([reg(0, "unica", 1), reg(1, "outra", 5)])
    conferir("um achado conta como sucesso por padrao", not p and not o)
    p, o, r = classificar([reg(0, "unica", 1), reg(1, "outra", 5)], pouco=1)
    conferir("--pouco 1 muda a classificacao", len(p) == 1)

    if falhas:
        print("AUTOTESTE FALHOU em %d caso(s):" % len(falhas))
        for f in falhas:
            print("  " + f)
        raise SystemExit(1)
    print("autoteste: %d casos, todos passam" % 13)


def main():
    if "--autoteste" in sys.argv:
        autoteste()
        return

    pouco = POUCO
    if "--pouco" in sys.argv:
        pouco = int(sys.argv[sys.argv.index("--pouco") + 1])

    argumentos = []
    pular = False
    for a in sys.argv[1:]:
        if pular:
            pular = False
            continue
        if a == "--pouco":
            pular = True
            continue
        if not a.startswith("--"):
            argumentos.append(a)

    caminho = argumentos[0] if argumentos else os.path.join(
        os.environ.get("LOCALAPPDATA", ""), "Iris", "buscas.jsonl")

    if not os.path.exists(caminho):
        print("nao achei o diario em %s" % caminho)
        print("")
        print("Ele nasce na primeira busca feita no Iris. Se voce ainda nao")
        print("procurou nada desde que o registro foi ligado, e isso.")
        return

    buscas, ruins = ler(caminho)
    pares, orfas, refinos = classificar(buscas, pouco=pouco)

    print("DIARIO DE BUSCAS -> candidatos a consulta por sentido")
    print("buscas anotadas: %d%s | 'achou pouco' = ate %d resultado(s)"
          % (len(buscas),
             ("  (%d linha(s) ilegivel(is))" % ruins) if ruins else "", pouco))
    print("")
    print("A SAIDA ABAIXO CONTEM OS TERMOS QUE VOCE DIGITOU, inteiros. Numa")
    print("caixa corporativa isso costuma incluir nome de pessoa e numero de")
    print("contrato -- cuidado com tela compartilhada e com redirecionar para")
    print("arquivo.")
    print("")

    if pares:
        print("REFORMULACOES -- digitei uma coisa, nao achei, digitei outra e achei:")
        print("")
        for b, s, m, meio in pares:
            print("  %-28s (%d achados)"
                  % (repr(b["termo"]), b.get("exatos", 0) + b.get("aproximados", 0)))
            for x in meio:
                print("     %-25s (%d achados)"
                      % (repr(x["termo"]), x.get("exatos", 0) + x.get("aproximados", 0)))
            print("  -> %-25s (%d achados, %.1f min depois)"
                  % (repr(s["termo"]), s.get("exatos", 0) + s.get("aproximados", 0), m))
            print("")
        print("CADA UM DESTES E CANDIDATO, E NAO PROVA. So voce sabe se as buscas")
        print("de um bloco eram a mesma intencao. As que forem viram linha em")
        print("tools/consultas-por-sentido.json -- 'digitei' e a primeira, e o")
        print("'queria' e o assunto da mensagem que a ultima achou.")
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
        print("Estas SAO sinal, com um lado so: a busca nao achou e a pessoa")
        print("desistiu. Se voce lembrar qual mensagem queria em alguma delas,")
        print("ela vira uma linha completa.")

    if refinos:
        print("")
        print("REFINADAS, E NAO TROCADAS (%d):" % len(refinos))
        print("Falharam, e a busca seguinte era a MESMA com mais palavras --")
        print("refino, e nao troca de vocabulario. Nao entram como candidato, e")
        print("aparecem aqui para nao sumirem de toda saida.")
        for b in refinos[:10]:
            print("  %-28s (%d achados)"
                  % (repr(b["termo"]), b.get("exatos", 0) + b.get("aproximados", 0)))
        if len(refinos) > 10:
            print("  ... e mais %d" % (len(refinos) - 10))


if __name__ == "__main__":
    main()

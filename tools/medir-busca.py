# -*- coding: utf-8 -*-
"""
MEDIR A BUSCA TEXTUAL CONTRA O ACERVO REAL -- sem oraculo e sem egress.

===========================================================================
POR QUE ESTA FERRAMENTA EXISTE

O ESCOPO deixou a Fase 4 (busca semantica) parada com oito decisoes abertas,
e nomeou a condicao para reavaliar: "evidencia de que a busca textual
normalizada nao resolve". Ninguem tinha produzido essa evidencia -- a fase
estava parada por falta de MEDICAO, e nao por falta de opiniao.

O obstaculo declarado era nao haver oraculo: so o dono da caixa sabe o que
era relevante. Isso vale para consulta com julgamento ("mostre o que exigia
acao"), e NAO vale para uma classe inteira de consulta que se pode medir
sozinha:

    se voce digitar tres palavras do PROPRIO assunto de uma mensagem,
    a busca tem de achar aquela mensagem.

Isso e verificavel sem ninguem opinar. A mensagem e a resposta certa por
construcao. E dela sai um recall honesto por tipo de consulta.

===========================================================================
O QUE ESTA MEDIDA NAO ALCANCA, E E PRECISO DIZER

Consulta por SINONIMO ou PARAFRASE -- "cobranca" achando "fatura" -- e
exatamente onde embeddings ganhariam, e e exatamente o que esta medicao NAO
cobre, porque decidir que "fatura" responde "cobranca" e julgamento.

Entao esta ferramenta responde uma metade: **onde a busca textual falha por
motivo mecanico**. Se ela falhar muito ai, o conserto e barato e local
(morfologia, tolerancia a erro de digitacao) e nao precisa de nenhuma das
oito decisoes. Se ela nao falhar ai, sobra so a metade semantica -- e ai a
pergunta para o dono fica concreta em vez de abstrata.

===========================================================================
PRIVACIDADE

Le o cache local do dono, que por D1 tem so metadado -- assunto e
remetente, nunca corpo nem anexo. NADA sai da maquina, e o relatorio traz
**somente agregados**: contagens e percentuais. Nenhum assunto real, nenhum
remetente, nenhum EntryID e impresso ou gravado.

Uso:  python tools/medir-busca.py [caminho-do-cache.db]
"""
import os
import random
import re
import sqlite3
import sys
import unicodedata

# Semente fixa: a medicao tem de dar o mesmo numero duas vezes seguidas,
# senao nao e evidencia, e sim anedota.
SEMENTE = 20260829
AMOSTRA = 300


# ==========================================================================
# A BUSCA, COMO O IRIS A FAZ HOJE
#
# Reimplementada aqui de proposito, e a duplicacao e o ponto: se um dia o
# BuscaNoAcervo mudar e esta copia nao, a medicao passa a medir outra coisa.
# Por isso ha um teste na suite (BuscaMedidaTests) que compara as duas
# contra os mesmos casos -- ele quebra no dia em que divergirem.
# ==========================================================================

def normalizar(s):
    """Minusculas invariantes e sem diacritico. Espelha TermoDeBusca.Normalizar."""
    if not s:
        return ""
    decomposto = unicodedata.normalize("NFD", s)
    sem_marca = "".join(c for c in decomposto if unicodedata.category(c) != "Mn")
    return unicodedata.normalize("NFC", sem_marca).lower()


def radical(p):
    """Espelha TermoDeBusca.Radical: terminacoes em que o singular NAO e
    subcadeia do plural. Piso de cinco letras."""
    if not p or len(p) < 5:
        return p or ""
    for suf, troca in (("oes", "ao"), ("aes", "ao"), ("ais", "al"),
                       ("eis", "el"), ("ois", "ol"), ("uis", "ul")):
        if p.endswith(suf):
            return p[:-3] + troca
    if p.endswith("ns"):
        return p[:-2] + "m"
    if p.endswith("es"):
        return p[:-2]
    if p.endswith("s"):
        return p[:-1]
    return p


def distancia_ate_1(a, b):
    """Espelha TermoDeBusca.DistanciaAte1: para no segundo erro."""
    if a == b:
        return True
    ia = ib = erros = 0
    while ia < len(a) and ib < len(b):
        if a[ia] == b[ib]:
            ia += 1
            ib += 1
            continue
        erros += 1
        if erros > 1:
            return False
        if len(a) == len(b):
            ia += 1
            ib += 1
        elif len(a) > len(b):
            ia += 1
        else:
            ib += 1
    return erros + (len(a) - ia) + (len(b) - ib) <= 1


def parecidas(consulta, do_alvo):
    if not consulta or not do_alvo:
        return False
    if consulta in do_alvo:
        return True
    rc, ra = radical(consulta), radical(do_alvo)
    if len(rc) >= 4 and rc in ra:
        return True
    if len(consulta) >= 5 and abs(len(consulta) - len(do_alvo)) <= 1:
        return distancia_ate_1(consulta, do_alvo)
    return False


# Ligado por --sem-tolerancia. Existe para o ANTES e o DEPOIS sairem do
# MESMO gerador de consultas: o gerador foi corrigido em 29/08 (pluralizava
# palavra ja plural), e comparar o numero velho com o novo misturaria o
# conserto do gerador com o conserto da busca. Isso nao e comparacao, e
# propaganda.
TOLERANCIA = True


def grau(consulta, assunto, remetente):
    """0 = nenhum, 1 = exato, 2 = aproximado. Espelha TermoDeBusca.Grau."""
    alvo = normalizar("%s %s" % (assunto or "", remetente or ""))
    palavras = normalizar(consulta).split()
    if not palavras:
        return 0
    if all(p in alvo for p in palavras):
        return 1
    if not TOLERANCIA:
        return 0
    do_alvo = alvo.split()
    if do_alvo and all(any(parecidas(p, a) for a in do_alvo) for p in palavras):
        return 2
    return 0


def casa(consulta, assunto, remetente):
    """Acha de algum jeito -- exato ou aproximado."""
    return grau(consulta, assunto, remetente) != 0


# ==========================================================================
# AS CONSULTAS DERIVADAS
#
# Cada uma imita um jeito real de digitar, e cada uma tem resposta certa
# conhecida: a propria mensagem de onde ela saiu.
# ==========================================================================

PREFIXOS = re.compile(r"^\s*((re|res|enc|fw|fwd|encaminhada|em)\s*:\s*)+", re.I)
NAO_LETRA = re.compile(r"[^0-9A-Za-zÀ-ÿ]+")


def palavras_uteis(assunto):
    """Palavras do assunto que uma pessoa digitaria.

    Fora: os prefixos de resposta/encaminhamento (sintaxe, nao assunto) e as
    palavras de uma ou duas letras, que ninguem usa para procurar.
    """
    limpo = PREFIXOS.sub("", assunto or "")
    return [p for p in NAO_LETRA.split(limpo) if len(p) >= 3]


def com_erro_de_digitacao(palavra, rnd):
    """Troca UMA letra do meio. Nao a primeira: ninguem erra a inicial e
    continua achando que digitou certo."""
    if len(palavra) < 5:
        return None
    i = rnd.randrange(1, len(palavra) - 1)
    alfabeto = "abcdefghijklmnopqrstuvwxyz"
    trocada = rnd.choice([c for c in alfabeto if c != palavra[i].lower()])
    return palavra[:i] + trocada + palavra[i + 1:]


def flexionar(palavra):
    """Plural do portugues, so nos casos em que o singular NAO e subcadeia do
    plural -- que sao justamente os que a busca por subcadeia nao resolve.

    DUAS CORRECOES DE 29/08, e as duas vinham de o gerador ser ingenuo demais:

    1. DECIDE SOBRE A FORMA NORMALIZADA. Antes testava "ao" contra a palavra
       acentuada, e "reuniao" nunca casava porque no assunto ela e "reuniao"
       com til. O caso mais comum do portugues ficava de fora da medicao.

    2. NAO PLURALIZA O QUE JA E PLURAL. Antes "Dados" virava "dadoses" e
       "ATIVIDADES" virava "atividadeses" -- consultas que nenhuma pessoa
       digita. A busca acertava em nao achar, e a medicao contava isso como
       falha DELA. Um gerador que fabrica consulta impossivel mede o proprio
       gerador."""
    p = normalizar(palavra)
    if p.endswith("s"):
        return None               # ja esta no plural
    if p.endswith("ao"):          # reuniao -> reunioes
        return p[:-2] + "oes"
    if p.endswith("l"):           # contratual -> contratuais
        return p[:-1] + "is"
    if p.endswith("m"):           # armazem -> armazens
        return p[:-1] + "ns"
    if p.endswith("r") or p.endswith("z"):
        return p + "es"           # senhor -> senhores
    return None


def consultas(assunto, remetente, rnd):
    """Devolve [(nome_do_caso, consulta)] para uma mensagem."""
    ps = palavras_uteis(assunto)
    if not ps:
        return []

    saida = []
    tres = ps[:3]

    # 1. AS PROPRIAS PALAVRAS. Se isto falha, nada mais importa.
    saida.append(("exato", " ".join(tres)))

    # 2. SEM ACENTO. A normalizacao promete isto.
    saida.append(("sem_acento", normalizar(" ".join(tres))))

    # 3. CAIXA TROCADA.
    saida.append(("caixa_alta", " ".join(tres).upper()))

    # 4. FORA DE ORDEM. A conjuncao promete ser livre de ordem.
    if len(ps) >= 2:
        invertido = list(tres)
        invertido.reverse()
        saida.append(("fora_de_ordem", " ".join(invertido)))

    # 5. PEDACO DA PALAVRA. A busca por subcadeia promete isto.
    if len(ps[0]) >= 6:
        saida.append(("prefixo_da_palavra", ps[0][:5]))

    # 6. COM O REMETENTE JUNTO. O casamento e contra os dois campos.
    if remetente:
        primeiro = (palavras_uteis(remetente) or [""])[0]
        if primeiro:
            saida.append(("assunto_mais_remetente", ps[0] + " " + primeiro))

    # 7. ERRO DE DIGITACAO. Nada promete isto -- e por isso ele esta aqui.
    errada = com_erro_de_digitacao(ps[0], rnd)
    if errada:
        saida.append(("erro_de_digitacao", errada))

    # 8. FLEXAO. Idem: uma caixa em portugues flexiona, e subcadeia nao cobre.
    flex = flexionar(ps[0])
    if flex:
        saida.append(("flexao_de_numero", flex))

    return saida


# ==========================================================================

def main():
    global TOLERANCIA
    argumentos = [a for a in sys.argv[1:] if a != "--sem-tolerancia"]
    TOLERANCIA = "--sem-tolerancia" not in sys.argv

    caminho = argumentos[0] if argumentos else os.path.join(
        os.environ.get("LOCALAPPDATA", ""), "Iris", "cache.db")

    if not os.path.exists(caminho):
        raise SystemExit("cache nao encontrado em %s" % caminho)

    con = sqlite3.connect("file:%s?mode=ro" % caminho.replace("\\", "/"), uri=True)
    linhas = con.execute(
        "SELECT subject, sender_name FROM metadata_observation").fetchall()
    con.close()

    corpus = [(s or "", r or "") for s, r in linhas]
    if not corpus:
        raise SystemExit("o acervo esta vazio: nao ha o que medir")

    rnd = random.Random(SEMENTE)
    amostra = rnd.sample(corpus, min(AMOSTRA, len(corpus)))

    # achou / total / ambiguas, por caso
    placar = {}

    for assunto, remetente in amostra:
        for caso, consulta in consultas(assunto, remetente, rnd):
            g = grau(consulta, assunto, remetente)
            achou = g != 0
            # AMBIGUIDADE NAO E FALHA. Se a consulta casa com muitas outras
            # mensagens, achar a certa nao prova nada de qualidade -- mas
            # NAO achar prova defeito. Contamos separado para nao vender
            # precisao que nao medimos.
            quantas = sum(1 for a, r in corpus if casa(consulta, a, r))
            p = placar.setdefault(caso, {"achou": 0, "exato": 0, "total": 0, "ruido": 0})
            p["total"] += 1
            if achou:
                p["achou"] += 1
            if g == 1:
                p["exato"] += 1
            if quantas > 10:
                p["ruido"] += 1

    ordem = ["exato", "sem_acento", "caixa_alta", "fora_de_ordem",
             "prefixo_da_palavra", "assunto_mais_remetente",
             "erro_de_digitacao", "flexao_de_numero"]

    print("MEDICAO DA BUSCA TEXTUAL -- round-trip sobre o acervo local")
    print("segundo passe (tolerancia): %s" % ("LIGADO" if TOLERANCIA else "DESLIGADO"))
    print("mensagens no acervo: %d | amostradas: %d | semente: %d"
          % (len(corpus), len(amostra), SEMENTE))
    print("")
    print("%-24s %6s %6s %6s %9s" % ("caso", "exato", "aprox", "de", "recall"))
    print("-" * 56)
    for caso in ordem:
        p = placar.get(caso)
        if not p or p["total"] == 0:
            continue
        print("%-24s %6d %6d %6d %8.1f%%"
              % (caso, p["exato"], p["achou"] - p["exato"], p["total"],
                 100.0 * p["achou"] / p["total"]))

    print("")
    print("RUIDO (consultas que casam com mais de 10 mensagens):")
    for caso in ordem:
        p = placar.get(caso)
        if not p or p["total"] == 0:
            continue
        print("  %-22s %5.1f%%" % (caso, 100.0 * p["ruido"] / p["total"]))

    print("")
    print("NAO MEDIDO: consulta por sinonimo ou parafrase. E onde embeddings")
    print("ganhariam, e decidir que uma responde a outra e julgamento -- so o")
    print("dono da caixa pode dizer. Esta medicao cobre a metade mecanica.")


if __name__ == "__main__":
    main()

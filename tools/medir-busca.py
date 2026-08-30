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
import io
import os
import time
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
# Reimplementada aqui de proposito: a medicao precisa rodar sobre o SQLite
# sem subir a aplicacao. A duplicacao e um risco real -- se o BuscaNoAcervo
# mudar e esta copia nao, a medicao passa a medir outra coisa.
#
# ESTE COMENTARIO JA AFIRMOU UMA GARANTIA QUE NAO EXISTIA. Ele dizia que
# havia um teste comparando as duas implementacoes, e nao havia; a revisao
# externa de 29/08 pegou. Comentario que promete protecao inexistente e pior
# que nenhum comentario, porque quem le para de procurar.
#
# Agora existe, e e um arquivo de casos que os DOIS lados conferem:
#   tools/casos-de-busca.json
#   VB      -> BuscaMedidaTests.As_duas_implementacoes_concordam
#   Python  -> python tools/medir-busca.py --conferir
# Quem divergir falha sozinho.
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


def em_palavras(alvo):
    """Espelha TermoDeBusca.EmPalavras: parte por tudo o que nao e letra nem
    digito.

    Era uma lista de pontuacao ASCII dos dois lados, e a revisao externa
    mediu a diferenca: aspas curvas, travessao, reticencias e espaco nao
    separavel ficavam colados na palavra. Lista de pontuacao nunca termina --
    a pergunta certa e "isto e letra ou digito?"."""
    saida, atual = [], []
    for c in alvo or "":
        if c.isalnum():
            atual.append(c)
        elif atual:
            saida.append("".join(atual))
            atual = []
    if atual:
        saida.append("".join(atual))
    return saida


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
    do_alvo = em_palavras(alvo)
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


ERROS = ("substituicao", "insercao", "remocao", "transposicao")


def com_erro_de_digitacao(palavra, rnd, tipo):
    """UM erro de digitacao, do tipo pedido.

    A PRIMEIRA VERSAO SO FAZIA SUBSTITUICAO INTERNA, e a revisao externa
    mostrou por que isso e um numero inflado: substituicao interna em palavra
    longa e exatamente o caminho mais simples do DistanciaAte1. Medir so ela e
    perguntar ao algoritmo aquilo que ele responde melhor.

    Os quatro tipos cobrem o que uma pessoa faz de verdade -- e a transposicao
    esta aqui de proposito, porque ela e distancia DOIS para este algoritmo.
    Ela vai falhar, e falhar medido e melhor que passar por nao ter sido
    perguntado."""
    if len(palavra) < 5:
        return None

    alfabeto = "abcdefghijklmnopqrstuvwxyz"
    i = rnd.randrange(1, len(palavra) - 1)

    if tipo == "substituicao":
        c = rnd.choice([c for c in alfabeto if c != palavra[i].lower()])
        return palavra[:i] + c + palavra[i + 1:]
    if tipo == "insercao":
        return palavra[:i] + rnd.choice(alfabeto) + palavra[i:]
    if tipo == "remocao":
        return palavra[:i] + palavra[i + 1:]
    if tipo == "transposicao":
        if palavra[i] == palavra[i + 1]:
            return None
        return palavra[:i] + palavra[i + 1] + palavra[i] + palavra[i + 2:]
    return None


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


def sonda_fora_do_radical(palavra):
    """UMA transformacao morfologica escolhida a mao, que o radical NAO cobre.

    ISTO E UMA SONDA, E NAO UMA MEDIDA DE RECALL. A diferenca importa, e a
    primeira versao desta funcao a apagava: ela se chamava
    "flexao_fora_do_radical" e o relatorio publicava o resultado como se fosse
    recall de uma classe de consulta real.

    O problema e o mesmo do gerador de plural que ja foi corrigido uma vez:
    aplicar uma regra mecanica a qualquer palavra fabrica forma que ninguem
    digita. "contrato" -> "contratinho" e valido; a mesma regra sobre um nome
    proprio ou um estrangeirismo produz lixo. Sem lexico, o numero nao mede
    "diminutivos que uma pessoa escreveria" -- mede "esta transformacao que eu
    escolhi".

    Entao ela serve para UMA coisa, e o relatorio diz qual: mostrar que existe
    transformacao morfologica fora do alcance do radical. O tamanho do buraco
    ela nao mede.

    E ela NAO e evidencia sobre a metade semantica. Diminutivo e morfologia
    derivacional; sinonimo e outra coisa. O comentario anterior dizia que este
    numero era "o comeco da conversa sobre a metade semantica", e isso
    ultrapassava os dados -- a revisao externa pegou."""
    p = normalizar(palavra)
    if len(p) < 6 or p.endswith("s"):
        return None
    if p.endswith("o") or p.endswith("a"):
        return p[:-1] + "inh" + p[-1]
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

    # 7. ERRO DE DIGITACAO, nos quatro tipos, cada um contado a parte.
    #    Somados num numero so, a transposicao -- que e distancia 2 e nao pode
    #    ser achada -- ficaria escondida atras das outras tres.
    for tipo in ERROS:
        errada = com_erro_de_digitacao(ps[0], rnd, tipo)
        if errada:
            saida.append(("erro_" + tipo, errada))

    # 8. FLEXAO DE NUMERO, separada em DUAS medidas.
    #
    #    A revisao externa mostrou que somar as duas era quase tautologia: o
    #    gerador produzia exatamente as cinco terminacoes que o radical
    #    implementa, entao "100% em flexao" queria dizer "100% no que eu
    #    escolhi medir". Separado, o numero passa a dizer o que e:
    #
    #      flexao_de_numero        -- as familias que o radical cobre
    #      flexao_NAO_coberta      -- as que ele nao cobre, e nao esconde
    flex = flexionar(ps[0])
    if flex:
        saida.append(("flexao_de_numero", flex))

    fora = sonda_fora_do_radical(ps[0])
    if fora:
        saida.append(("sonda_diminutivo", fora))

    return saida


# ==========================================================================

def conferir():
    """Roda os casos compartilhados contra ESTA implementacao.

    A outra metade da mesma conferencia mora na suite VB. As duas leem o
    mesmo arquivo, entao divergencia aparece no lado que divergiu."""
    import json
    casos = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "casos-de-busca.json")
    with io.open(casos, encoding="utf-8") as f:
        dados = json.load(f)

    erros = []
    for c in dados["casos"]:
        obtido = grau(c["consulta"], c["assunto"], c["remetente"])
        if obtido != c["grau"]:
            # ascii() e nao %r: o console do Windows e cp1252, e um emoji
            # num caso de teste derrubava o relatorio de divergencia --
            # a ferramenta de diagnostico morrendo no diagnostico.
            erros.append("  %-24s vs %-28s  esperado %d, obtido %d"
                         % (ascii(c["consulta"]), ascii(c["assunto"]),
                            c["grau"], obtido))

    if erros:
        print("DIVERGIU em %d de %d casos:" % (len(erros), len(dados["casos"])))
        for e in erros:
            print(e)
        raise SystemExit(1)

    print("as %d regras conferem com tools/casos-de-busca.json"
          % len(dados["casos"]))


def por_sentido(corpus, caminho):
    """A METADE QUE NENHUM HARNESS MEDE SOZINHO.

    O round-trip deste arquivo deriva consultas do proprio assunto, entao ele
    so enxerga falha por LETRA. Falha por SENTIDO -- digitar "cobranca" e
    querer a mensagem que diz "fatura" -- exige alguem dizer qual era a
    mensagem certa. E julgamento, e so o dono da caixa tem.

    Esta funcao le esse julgamento de tools/consultas-por-sentido.json e
    responde a pergunta que decide a metade aberta da Fase 4:

        de N consultas por sentido, quantas a busca textual erra?

    Zero erro quer dizer que indexar nao compra nada aqui. Erro alto quer
    dizer que compra -- e ai as decisoes 4, 5, 6 e 7 do ESCOPO passam a
    valer a pena discutir. Antes disso, discuti-las e escolher a
    implementacao antes do requisito."""
    import json

    if not os.path.exists(caminho):
        print("nao achei %s" % caminho)
        print("")
        print("Copie tools/consultas-por-sentido.exemplo.json para esse nome e")
        print("troque os casos inventados pelos seus. Dez ou quinze bastam.")
        print("O arquivo esta no .gitignore: ele tem assunto real.")
        return

    with io.open(caminho, encoding="utf-8") as f:
        dados = json.load(f)

    consultas = dados.get("consultas", [])
    if not consultas:
        print("o arquivo nao tem nenhuma consulta")
        return

    achou_exato = achou_aprox = nao_achou = ambigua = nao_localizei = 0
    perdidas = []

    for c in consultas:
        digitei = c.get("digitei", "")
        queria = normalizar(c.get("queria", ""))
        if not digitei or not queria:
            continue

        # A MENSAGEM CERTA e a que contem o trecho informado. Se mais de uma
        # contem, o caso e ambiguo e NAO conta -- nem a favor nem contra.
        alvos = [(a, r) for a, r in corpus if queria in normalizar(a)]

        # ZERO NAO E AMBIGUIDADE. A primeira versao juntava os dois num
        # contador so, e sao coisas diferentes: zero quer dizer que o trecho
        # que voce escreveu nao casa com nenhuma mensagem do acervo -- erro de
        # digitacao no trecho, ou mensagem fora da janela varrida. Mais de uma
        # quer dizer que o trecho nao identifica. Colapsar as duas seria, de
        # novo, ausencia virando a mesma coisa que excesso.
        if len(alvos) == 0:
            nao_localizei += 1
            perdidas.append("  NAO LOCALIZEI a mensagem alvo: %s"
                            % ascii(c.get("queria", "")))
            continue
        if len(alvos) > 1:
            ambigua += 1
            perdidas.append("  AMBIGUA (%d mensagens casam o trecho): %s"
                            % (len(alvos), ascii(c.get("queria", ""))))
            continue

        g = grau(digitei, alvos[0][0], alvos[0][1])
        if g == 1:
            achou_exato += 1
        elif g == 2:
            achou_aprox += 1
        else:
            nao_achou += 1
            perdidas.append("  NAO ACHOU: digitei %s" % ascii(digitei))

    validas = achou_exato + achou_aprox + nao_achou
    print("BUSCA POR SENTIDO -- o julgamento que so o dono da caixa tem")
    print("consultas no arquivo: %d | validas: %d | ambiguas: %d | "
          "alvo nao localizado: %d"
          % (len(consultas), validas, ambigua, nao_localizei))
    print("")
    if validas:
        print("  achou exato       %3d  (%.0f%%)"
              % (achou_exato, 100.0 * achou_exato / validas))
        print("  achou aproximado  %3d  (%.0f%%)"
              % (achou_aprox, 100.0 * achou_aprox / validas))
        print("  NAO ACHOU         %3d  (%.0f%%)"
              % (nao_achou, 100.0 * nao_achou / validas))
    print("")
    for l in perdidas:
        print(l)
    print("")
    print("COMO LER ESTE NUMERO:")
    print("  NAO ACHOU baixo  -> indexar nao compra nada aqui, e a Fase 4")
    print("                      continua fechada com evidencia dos dois lados.")
    print("  NAO ACHOU alto   -> compra, e as decisoes 4/5/6/7 do ESCOPO")
    print("                      passam a valer a pena discutir.")
    print("")
    print("RESSALVA: a amostra e a que voce escreveu. Ela mede o que voce")
    print("lembrou de procurar, e nao a distribuicao das suas buscas reais.")
    print("Um numero destes decide o SENTIDO da resposta, nao a magnitude.")


def main():
    global TOLERANCIA

    if "--conferir" in sys.argv:
        conferir()
        return

    # Toda flag fora, e nao so a que eu lembrei: a primeira versao filtrava
    # so --sem-tolerancia, e --por-sentido virou "caminho do cache".
    argumentos = [a for a in sys.argv[1:] if not a.startswith("--")]
    TOLERANCIA = "--sem-tolerancia" not in sys.argv

    caminho = argumentos[0] if argumentos else os.path.join(
        os.environ.get("LOCALAPPDATA", ""), "Iris", "cache.db")

    if not os.path.exists(caminho):
        raise SystemExit("cache nao encontrado em %s" % caminho)

    con = sqlite3.connect("file:%s?mode=ro" % caminho.replace("\\", "/"), uri=True)

    # O MESMO CORPUS QUE A BUSCA PERCORRE, e nao a tabela inteira.
    #
    # A primeira versao lia metadata_observation direto, e a revisao externa
    # pegou: a busca le o MANIFESTO PUBLICADO -- associacoes com
    # generation_key nao nulo. Observacao de uma varredura ainda em curso, ou
    # de encarnacao sem associacao visivel, existe na tabela e NAO existe para
    # quem procura. Medir um corpus e concluir sobre o outro e o defeito que
    # esta ferramenta foi feita para nao cometer.
    #
    # A consulta espelha ManifestReader.LerManifesto.
    linhas = con.execute(
        "SELECT m.subject, m.sender_name "
        "FROM association a "
        "JOIN incarnation i ON i.item_key = a.item_key AND i.folder_key = a.folder_key "
        "LEFT JOIN metadata_observation m ON m.incarnation_key = i.incarnation_key "
        "WHERE a.generation_key IS NOT NULL").fetchall()
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

    ordem = (["exato", "sem_acento", "caixa_alta", "fora_de_ordem",
              "prefixo_da_palavra", "assunto_mais_remetente"] +
             ["erro_" + t for t in ERROS] +
             ["flexao_de_numero", "sonda_diminutivo"])

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

    # ==================================================================
    # QUANTO CUSTA UMA BUSCA, EM CIMA DESTE ACERVO.
    #
    # A revisao externa deixou aberto: o segundo passe roda item a item,
    # palavra a palavra, na thread da UI. Para 1.127 itens "provavelmente
    # continua aceitavel" -- e "provavelmente" nao e numero.
    #
    # O pior caso e a consulta que NAO acha nada: todo item falha no primeiro
    # passe e paga o segundo inteiro. E o que esta medido aqui.
    #
    # RESSALVA: Python e mais lento que .NET, entao este numero e um TETO
    # folgado, e nao o tempo da aplicacao. Ele responde "esta na casa dos
    # milissegundos ou dos segundos?", que e a pergunta que importa agora.
    pior = "jacarebicicletaxyz"
    t0 = time.perf_counter()
    for _ in range(10):
        for a, r in corpus:
            grau(pior, a, r)
    ms = (time.perf_counter() - t0) * 1000.0 / 10.0

    print("")
    print("PIOR CASO (consulta sem nenhum achado, acervo inteiro):")
    print("  %.1f ms por busca, sobre %d itens -- teto folgado, medido em"
          % (ms, len(corpus)))
    print("  Python, que e mais lento que o .NET da aplicacao.")

    print("")
    print("RUIDO (consultas que casam com mais de 10 mensagens):")
    for caso in ordem:
        p = placar.get(caso)
        if not p or p["total"] == 0:
            continue
        print("  %-22s %5.1f%%" % (caso, 100.0 * p["ruido"] / p["total"]))

    print("")
    print("sonda_diminutivo NAO E RECALL. E uma transformacao mecanica")
    print("escolhida a mao, sem lexico -- ela mostra que existe morfologia fora")
    print("do alcance do radical, e nao mede o tamanho desse buraco.")
    print("")
    print("NAO MEDIDO AQUI: consulta por sinonimo ou parafrase. E onde")
    print("embeddings ganhariam, e decidir que uma responde a outra e")
    print("julgamento -- so o dono da caixa pode dizer.")
    print("")
    print("Para medir essa metade:  python tools/medir-busca.py --por-sentido")

    # A metade que depende do julgamento do dono, quando ela existir.
    if "--por-sentido" in sys.argv:
        print("")
        print("=" * 58)
        print("")
        por_sentido(corpus, os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "consultas-por-sentido.json"))


if __name__ == "__main__":
    main()

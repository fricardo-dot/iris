# -*- coding: utf-8 -*-
"""
CONTROLE NEGATIVO — desliga cada guarda e confere o desfecho ESPERADO dela.

    python tools/controle-negativo.py     # 0 = tudo como esperado, 1 = divergiu

--------------------------------------------------------------------------
POR QUE ELE EXISTE

"Teste verde nao e prova": um teste que passa prova que aquele caminho vale,
nao que a propriedade vale. Um bloqueio sem controle negativo passa igual
quando a guarda existe e quando alguem a apagou.

--------------------------------------------------------------------------
NAO E "TODA MUTACAO TEM DE FICAR VERMELHA"

Cada cenario declara o desfecho que se espera dele, e o conjunto EXATO de
testes que devem cair:

  * as duas redundancias declaradas — cobertura na emissao e proveniencia no
    consumo — esperam VERDE sozinhas, porque a outra guarda ja segura. Elas
    estao aqui justamente para documentar que sao independentemente
    suficientes;
  * as duas juntas, e os demais cenarios, esperam VERMELHO, com os testes
    nomeados.

Desfecho diferente do esperado — nos dois sentidos, e inclusive um teste a
mais ou a menos na lista — e falha do roteiro, e o processo sai com 1.
Conferir o CONJUNTO, e nao "caiu alguem", importa porque a suite tem uma
intermitencia conhecida de SQLite: sem isso, uma flake satisfaria um cenario
que espera vermelho.

E o conjunto sozinho tambem nao basta: nos cenarios que esperam verde, o
conjunto vazio e igualzinho ao que sai quando o `dotnet test` morre antes de
rodar teste nenhum. Por isso o resumo e o codigo de saida entram na conta —
verde exige codigo 0, "Passed!" e pelo menos um teste executado; vermelho
exige codigo diferente de 0 e "Failed!".

--------------------------------------------------------------------------
POR QUE ELE FOTOGRAFA ANTES E CONFERE O HASH DEPOIS

A primeira versao refotografava o arquivo a cada edicao. Com duas edicoes no
mesmo arquivo, a segunda foto ja era da versao quebrada, e a restauracao
devolveu o Capability.vb SEM a chamada a Cobre — que foi commitada assim, com
a suite verde porque fora medida antes.

Agora cada arquivo e fotografado UMA vez, antes de qualquer edicao, e o
SHA-256 e conferido depois de restaurar. Se nao bater, o roteiro aborta.

Ferramenta de verificacao tambem precisa de verificacao, e a que edita codigo
de producao precisa mais que as outras.
"""
import hashlib
import io
import os
import re
import subprocess
import sys

# A raiz vem do proprio arquivo: caminho absoluto travado na maquina de quem
# escreveu faz o roteiro parar de existir para qualquer outro clone.
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

CAP = "src/Iris.Assist/Capability.vb"
PIPE = "src/Iris.Assist/ContentPipeline.vb"
VM = "src/Iris.App/ViewModels/AssistenteViewModel.vb"
TRANS = "src/Iris.Assist/AssistTransmitter.vb"
ATIV = "src/Iris.Assist/Activation.vb"
ORP = "src/Iris.Integration.Assist.Http/OpenRouterAssistantProvider.vb"
ADV = "tests/Iris.Tests/AdversarioPontaAPontaTests.vb"

ADVF = "FullyQualifiedName~AdversarioPontaAPontaTests"

COBRE = (CAP,
         "            If Not g.Cobre(envelope.Itens, envelope.Versoes) Then Return Nothing",
         "            ' desligado")
PROVENIENCIA = (CAP,
                """            If Not MesmosItens(envelope.Itens, c.Itens) OrElse
               Not MesmasVersoes(envelope.Versoes, c.Versoes) Then
                Return Recusar(CapabilityRefusal.ProveniencaDiferente)
            End If""",
                "            ' desligado")

CENARIOS = [
    dict(nome="cobertura no Emitir, sozinha",
         edicoes=[COBRE], filtro=ADVF,
         esperados=set(),
         porque="a proveniencia no consumo reconfere por conta propria"),

    dict(nome="proveniencia no Consumir, sozinha",
         edicoes=[PROVENIENCIA], filtro=ADVF,
         esperados=set(),
         porque="a cobertura na emissao ja segura"),

    dict(nome="cobertura E proveniencia",
         edicoes=[COBRE, PROVENIENCIA], filtro=ADVF,
         esperados={
             "Referencia_embutida_NAO_transmite",
             "HTML_hostil_tem_TRES_desfechos",
             "Anexo_que_APARECE_depois_da_classificacao_NAO_transmite",
             "Versao_que_MUDA_entre_classificar_e_montar_NAO_transmite",
             "Thread_montada_pela_METADE_NAO_transmite",
             "Selecao_que_MUDA_no_meio_NAO_transmite",
             "Classificacao_de_item_NAO_PEDIDO_nao_transmite",
         },
         porque="sem as duas, ate a recusa de conteudo manda envelope vazio"),

    dict(nome="anexo no pipeline",
         edicoes=[(PIPE,
                   """            If m.TemAnexo Is Nothing OrElse m.TemAnexo.Value Then
                Return Recusar(ContentRefusal.Anexo)
            End If""",
                   "            ' desligado")],
         filtro=ADVF,
         esperados={"Anexo_que_APARECE_depois_da_classificacao_NAO_transmite"},
         porque="fecha a corrida entre a visita da classificacao e a do corpo"),

    dict(nome="espaco em branco nao e resultado",
         edicoes=[(VM,
                   "                Return Not String.IsNullOrWhiteSpace(Resultado)",
                   "                Return Resultado.Length > 0")],
         filtro=ADVF,
         esperados={"Resposta_so_de_ESPACO_nao_e_resposta"},
         porque="tres espacos escapavam do aviso e eram aplicados no rascunho"),

    # ---- a ativacao ----

    dict(nome="clone do corpo preparado",
         edicoes=[(TRANS,
                   "                corpo = If(devolvido Is Nothing, Nothing, CType(devolvido.Clone(), Byte()))",
                   "                corpo = devolvido")],
         filtro="FullyQualifiedName~Corpo_ADULTERADO_depois_de_conferido_nao_sai",
         esperados={"Corpo_ADULTERADO_depois_de_conferido_nao_sai"},
         porque="sem a copia, o provedor adultera depois de o hash ter sido conferido"),

    dict(nome="lista de provedores nao pode ser vazia",
         edicoes=[(ATIV,
                   "            If ProvedoresPermitidos.Count = 0 Then Return False",
                   "            ' desligado")],
         filtro="FullyQualifiedName~Lista_de_provedores_VAZIA_e_incoerente",
         esperados={"Lista_de_provedores_VAZIA_e_incoerente"},
         porque="vazio voltaria a querer dizer 'qualquer provedor'"),

    dict(nome="o adaptador recusa ativacao de outro provedor",
         edicoes=[(ORP,
                   """            If Not Atende(ativacao) Then
                Throw New ArgumentException(
                    "esta ativação não é para o OpenRouter", NameOf(ativacao))
            End If""",
                   "            ' desligado")],
         filtro="FullyQualifiedName~Ativacao_de_OUTRO_provedor_nao_vira_adaptador",
         esperados={"Ativacao_de_OUTRO_provedor_nao_vira_adaptador"},
         porque="o protocolo e a credencial do OpenRouter iriam para outro endereco"),

    dict(nome="colecoes da autorizacao sao imutaveis",
         edicoes=[(ATIV,
                   "            Return Array.AsReadOnly(o.ToArray())",
                   "            Return o.ToList()")],
         filtro="FullyQualifiedName~As_colecoes_da_autorizacao_NAO_dao_para_mexer",
         esperados={"As_colecoes_da_autorizacao_NAO_dao_para_mexer"},
         porque="IReadOnlyList sobre List devolve a lista viva num TryCast"),

    dict(nome="selecao sempre B (a corrida nao acontece)",
         edicoes=[(ADV,
                   "                Dim classificouA = False\n",
                   "                Dim classificouA = True\n")],
         filtro="FullyQualifiedName~Selecao_que_MUDA_no_meio_NAO_transmite",
         esperados={"Selecao_que_MUDA_no_meio_NAO_transmite"},
         porque="prova que o teste depende do movimento, e nao do item B"),
]


def soma(rel):
    with open(os.path.join(RAIZ, rel), "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def ler(rel):
    return io.open(os.path.join(RAIZ, rel), encoding="utf-8").read()


def escrever(rel, texto):
    io.open(os.path.join(RAIZ, rel), "w", encoding="utf-8").write(texto)


def caidos(saida):
    """Os nomes dos testes que falharam, como conjunto."""
    return set(re.findall(r"^\s*Failed\s+([A-Za-z_][A-Za-z0-9_]*)", saida, re.M))


def executados_em(saida):
    """Quantos testes o runner disse ter executado. -1 se nem disse."""
    m = re.search(r"Total:\s*(\d+)", saida)
    return int(m.group(1)) if m else -1


def rodar(c, problemas):
    arquivos = sorted({a for a, _, _ in c["edicoes"]})
    antes = {a: ler(a) for a in arquivos}
    somas = {a: soma(a) for a in arquivos}

    try:
        atual = dict(antes)
        for arq, bom, quebrado in c["edicoes"]:
            n = atual[arq].count(bom)
            if n != 1:
                # ANCORA AMBIGUA E FALHA, e nao aviso. Se o codigo mudou de
                # forma, o cenario deixou de desligar o que dizia desligar — e
                # sair em silencio devolveria exit 0 sem ter verificado nada.
                print("!! ANCORA  -", c["nome"], "  (%d ocorrencias em %s)" % (n, arq))
                problemas.append(c["nome"] + ": ancora ambigua ou ausente")
                return
            atual[arq] = atual[arq].replace(bom, quebrado, 1)

        for a in arquivos:
            escrever(a, atual[a])

        r = subprocess.run(["dotnet", "test", "Iris.slnx", "--nologo", "--filter", c["filtro"]],
                           cwd=RAIZ, capture_output=True, text=True, timeout=1800)
        saida = (r.stdout or "") + (r.stderr or "")
        obtido = caidos(saida)
        codigo = r.returncode

    finally:
        for a in arquivos:
            escrever(a, antes[a])
        for a in arquivos:
            if soma(a) != somas[a]:
                print("!! RESTAURACAO FALHOU:", a)
                sys.exit(2)

    if "error BC" in saida:
        print("!! COMPILOU NAO -", c["nome"])
        problemas.append(c["nome"] + ": a mutacao nao compila")
        return

    # O RESUMO, E NAO SO OS NOMES.
    #
    # Nos cenarios que esperam VERDE, o conjunto de falhas vazio tambem e o
    # que sai quando o `dotnet test` morre antes de rodar teste nenhum — erro
    # de MSBuild, de SDK, de testhost, ou filtro que nao casa com nada. Sem
    # conferir que alguma coisa RODOU, "nao executou" viraria "passou".
    executados = executados_em(saida)
    esperados = c["esperados"]

    if not esperados:
        ok = (codigo == 0 and "Passed!" in saida and executados >= 1)
        if not ok:
            print("!! NAO RODOU -", c["nome"],
                  "  (codigo=%s, executados=%s)" % (codigo, executados))
            problemas.append(c["nome"] + ": esperava verde, e a execucao nao aconteceu")
            return
    else:
        if codigo == 0 or "Failed!" not in saida:
            print("!! SEM FALHA -", c["nome"],
                  "  (codigo=%s, executados=%s)" % (codigo, executados))
            problemas.append(c["nome"] + ": esperava vermelho, e nao houve falha")
            return

    if obtido == esperados:
        cor = "VERDE" if not esperados else "VERMELHO"
        print("OK", cor.ljust(9), "-", c["nome"])
        for t in sorted(obtido):
            print("        ", t)
        return

    print("!!", ("VERDE" if not obtido else "VERMELHO").ljust(9), "-", c["nome"])
    for t in sorted(esperados - obtido):
        print("         faltou cair:", t)
    for t in sorted(obtido - esperados):
        print("         caiu sem ser esperado:", t)
    problemas.append(c["nome"])


def main():
    problemas = []
    for c in CENARIOS:
        rodar(c, problemas)

    print()
    print("restaurado, com soma conferida")
    if problemas:
        print("FALHOU: desfecho diferente do esperado em —", ", ".join(problemas))
        return 1
    print("OK: todo cenario deu exatamente o desfecho esperado")
    return 0


if __name__ == "__main__":
    sys.exit(main())

# -*- coding: utf-8 -*-
"""
CONTROLE NEGATIVO — desliga cada guarda e confere que o teste dela FALHA.

    python tools/controle-negativo.py

--------------------------------------------------------------------------
POR QUE ELE EXISTE

"Teste verde nao e prova": um teste que passa prova que aquele caminho vale,
nao que a propriedade vale. Um bloqueio sem controle negativo passa igual
quando a guarda existe e quando alguem a apagou.

Este roteiro desliga uma guarda de cada vez, roda os testes que a cobrem, e
espera VERMELHO. Verde ali quer dizer que o teste prova outra coisa.

--------------------------------------------------------------------------
POR QUE ELE FOTOGRAFA ANTES E CONFERE O HASH DEPOIS

A versao anterior refotografava o arquivo a cada edicao. Com duas edicoes no
mesmo arquivo, a segunda foto ja era da versao quebrada, e a restauracao
devolveu o Capability.vb SEM a chamada a Cobre — que foi commitada assim, com
a suite verde porque fora medida antes.

Agora cada arquivo e fotografado UMA vez, antes de qualquer edicao, e o
SHA-256 e conferido depois de restaurar. Se nao bater, o roteiro aborta.

Ferramenta de verificacao tambem precisa de verificacao, e a que edita codigo
de producao precisa mais que as outras.
"""
import io, subprocess, sys, hashlib

divergencias = []   # cenarios cujo desfecho nao foi o esperado

R = "C:/Users/Ricardo/Documents/Iris/"
CAP  = "src/Iris.Assist/Capability.vb"
PIPE = "src/Iris.Assist/ContentPipeline.vb"
VM   = "src/Iris.App/ViewModels/AssistenteViewModel.vb"
ADV  = "tests/Iris.Tests/AdversarioPontaAPontaTests.vb"

def digest(p):
    return hashlib.sha256(io.open(R+p, "rb").read()).hexdigest()

def rodar(nome, edicoes, filtro, espera="vermelho"):
    """espera="vermelho": a mutacao TEM de derrubar teste.

    espera="verde": a mutacao NAO derruba, e isso e o ponto do cenario — ha
    outra guarda em serie que sozinha ja segura. Sao os dois casos de
    redundancia declarada (cobertura na emissao e proveniencia no consumo).
    Desfecho diferente do esperado, nos dois sentidos, e falha do roteiro.
    """
    arquivos = sorted({a for a, _, _ in edicoes})
    antes = {a: io.open(R+a, encoding="utf-8").read() for a in arquivos}
    somas = {a: digest(a) for a in arquivos}
    try:
        atual = dict(antes)
        for arq, bom, quebrado in edicoes:
            if atual[arq].count(bom) != 1:
                print("!! ancora ambigua:", nome, arq, atual[arq].count(bom)); return
            atual[arq] = atual[arq].replace(bom, quebrado, 1)
        for a in arquivos:
            io.open(R+a, "w", encoding="utf-8").write(atual[a])

        r = subprocess.run(["dotnet","test","Iris.slnx","--nologo","--filter",filtro],
                           cwd=R, capture_output=True, text=True, timeout=900)
        s = (r.stdout or "") + (r.stderr or "")
        caiu = [l.strip().replace("Failed ","").split("[")[0].strip()
                for l in s.splitlines() if "Failed " in l and "Failed!" not in l]
        obtido = "vermelho" if caiu else "verde"
        marca = "OK " if obtido == espera else "!! "
        print(marca + obtido.upper().ljust(8), "-", nome,
              "" if obtido == espera else "  (esperava " + espera + ")")
        if obtido != espera:
            divergencias.append(nome)
        for c in caiu: print("       ", c)
        if not caiu:
            tot = [l.strip() for l in s.splitlines() if "Passed!" in l or "Failed!" in l]
            print("       ", tot[-1] if tot else "?")
            # VERDE E FALHA DO ROTEIRO, e nao resultado. Uma mutacao que nao
            # derruba teste nenhum quer dizer que a guarda nao esta provada — ou
            # que ha outra guarda em serie, e nesse caso o cenario e que esta
            # mal montado. Imprimir e seguir treinaria quem le a ignorar.

    finally:
        for a in arquivos:
            io.open(R+a, "w", encoding="utf-8").write(antes[a])
        for a in arquivos:
            assert digest(a) == somas[a], "RESTAURACAO FALHOU: " + a
    return

ADVF = "FullyQualifiedName~AdversarioPontaAPontaTests"

rodar("cobertura no Emitir, sozinha",
      [(CAP, "            If Not g.Cobre(envelope.Itens, envelope.Versoes) Then Return Nothing",
             "            ' desligado")], ADVF, espera="verde")

rodar("proveniencia no Consumir, sozinha",
      [(CAP, """            If Not MesmosItens(envelope.Itens, c.Itens) OrElse
               Not MesmasVersoes(envelope.Versoes, c.Versoes) Then
                Return Recusar(CapabilityRefusal.ProveniencaDiferente)
            End If""", "            ' desligado")], ADVF, espera="verde")

rodar("cobertura E proveniencia",
      [(CAP, "            If Not g.Cobre(envelope.Itens, envelope.Versoes) Then Return Nothing",
             "            ' desligado"),
       (CAP, """            If Not MesmosItens(envelope.Itens, c.Itens) OrElse
               Not MesmasVersoes(envelope.Versoes, c.Versoes) Then
                Return Recusar(CapabilityRefusal.ProveniencaDiferente)
            End If""", "            ' desligado")], ADVF)

rodar("anexo no pipeline",
      [(PIPE, """            If m.TemAnexo Is Nothing OrElse m.TemAnexo.Value Then
                Return Recusar(ContentRefusal.Anexo)
            End If""", "            ' desligado")], ADVF)

rodar("espaco em branco nao e resultado",
      [(VM, "                Return Not String.IsNullOrWhiteSpace(Resultado)",
            "                Return Resultado.Length > 0")], ADVF)

rodar("selecao sempre B (a corrida nao acontece)",
      [(ADV, "                Dim classificouA = False\n",
             "                Dim classificouA = True\n")],
      "FullyQualifiedName~Selecao_que_MUDA_no_meio_NAO_transmite")

print("restaurado, com soma conferida")

if divergencias:
    print()
    print("FALHOU: desfecho diferente do esperado em —", ", ".join(divergencias))
    sys.exit(1)
print("OK: todo cenario deu o desfecho esperado")

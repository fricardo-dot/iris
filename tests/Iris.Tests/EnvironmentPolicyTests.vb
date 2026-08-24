Imports System.Linq
Imports Iris.Sync
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Q8 — a matriz de providers, e o que fazer fora dela.
'''
''' A matriz está incompleta e continua incompleta depois destes testes: a
''' §19.3 mediu que levantar as outras linhas custa horas e dezenas de GB
''' nesta máquina. O que estes testes cobram é a outra metade da resposta —
''' que ambiente não medido DEGRADE, em vez de ser tratado como
''' "provavelmente igual".
''' </summary>
<TestClass>
Public Class EnvironmentPolicyTests

    ''' <summary>
    ''' O token da janela medido em 2026-08-24. Nao e "1 mes" nem numero
    ''' nenhum de meses: e o valor CRU do perfil. Ver
    ''' <see cref="EnvironmentFingerprint.WindowToken"/>.
    ''' </summary>
    Private Const TokenMedido As String = "84-09-00-00"

    Private Shared Function Medido() As EnvironmentFingerprint
        Return New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, TokenMedido)
    End Function

    <TestMethod>
    Public Sub O_ambiente_medido_autoriza_tudo()
        Dim c = EnvironmentPolicy.Capacidades(Medido())
        Assert.IsTrue(c.PodeConcluirAusencia)
        Assert.IsTrue(c.PodeAfirmarCoberturaCompleta)
        Assert.IsTrue(c.PodeUsarIncremental)
        Assert.IsFalse(c.Degradado)
    End Sub

    ''' <summary>
    ''' O controle POSITIVO. Sem ele, uma política que degradasse SEMPRE
    ''' passaria em todos os testes de degradação abaixo — e seria inútil
    ''' sem que nenhum teste percebesse.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_positivo_a_politica_nao_degrada_sempre()
        Assert.IsFalse(EnvironmentPolicy.Capacidades(Medido()).Degradado,
            "se ate o ambiente medido degrada, a politica nao esta decidindo nada")
        Assert.IsTrue(EnvironmentPolicy.Matriz.Count > 0)
    End Sub

    <TestMethod>
    Public Sub Ambiente_fora_da_matriz_nao_pode_concluir_ausencia()
        Dim fora = New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "00-00-00-00")
        Dim c = EnvironmentPolicy.Capacidades(fora)
        Assert.IsFalse(c.PodeConcluirAusencia,
            "num ambiente nao medido, 'nao encontrei' nao distingue excluido de fora-da-janela")
        Assert.IsFalse(c.PodeAfirmarCoberturaCompleta)
        Assert.IsFalse(c.PodeUsarIncremental)
        StringAssert.Contains(c.Reason, "fora da matriz")
    End Sub

    <TestMethod>
    Public Sub Provider_desconhecido_degrada()
        Dim c = EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.Desconhecido, True, "1 mes"))
        Assert.IsTrue(c.Degradado)
    End Sub

    <TestMethod>
    Public Sub Ambiente_nulo_degrada()
        Assert.IsTrue(EnvironmentPolicy.Capacidades(Nothing).Degradado)
    End Sub

    ''' <summary>
    ''' Cached com janela não lida é o caso mais traiçoeiro: sabe-se que há
    ''' janela e não se sabe qual. Pior que não saber nada, porque parece
    ''' identificado.
    ''' </summary>
    <TestMethod>
    Public Sub Cached_com_janela_nao_lida_degrada()
        Dim c = EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
        Assert.IsFalse(c.PodeConcluirAusencia)
        StringAssert.Contains(c.Reason, "janela")
    End Sub

    ''' <summary>
    ''' PST não está medido — e não herda nada do Exchange.
    ''' </summary>
    <TestMethod>
    Public Sub PST_nao_herda_do_Exchange()
        Dim c = EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.PstLocal, False, "sem-janela"))
        Assert.IsTrue(c.Degradado, "PST nunca foi medido neste projeto")
    End Sub

    ' ==================================================================
    ' A janela faz parte da identidade do ambiente

    ''' <summary>
    ''' Mesma caixa, mesma conta, janela diferente: ambiente DIFERENTE.
    '''
    ''' É a §18.4 virada em código. A janela muda o que EXISTE, não só o que
    ''' custa: em 2026-08-22 o OOM alcançava 1.004 itens numa caixa de 17.668;
    ''' em 2026-08-24, com a janela maior, alcança 1.979 até 2024-10-09. Mesma
    ''' caixa, mesma conta, dois universos.
    ''' </summary>
    <TestMethod>
    Public Sub Trocar_a_janela_muda_o_ambiente()
        Dim um = New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, TokenMedido)
        Dim tres = New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "FF-FF-00-00")
        Assert.AreNotEqual(um.Value(), tres.Value())
        Assert.IsTrue(EnvironmentPolicy.ExigeReconciliacao(um, tres))
    End Sub

    ''' <summary>
    ''' AUMENTAR a janela também invalida, e este é o caso que a intuição
    ''' erra.
    '''
    ''' "Encolher esconde, aumentar só revela — aumentar é seguro" é falso:
    ''' aumentar revela itens que já haviam sido concluídos AUSENTES, e essa
    ''' conclusão anterior estava errada. Deixá-la de pé por ser "só um
    ''' aumento" preserva justamente o erro.
    ''' </summary>
    <TestMethod>
    Public Sub Aumentar_a_janela_tambem_exige_reconciliacao()
        Dim curta = New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, TokenMedido)
        Dim longa = New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "00-00-00-00")
        Assert.IsTrue(EnvironmentPolicy.ExigeReconciliacao(curta, longa),
            "aumentar revela itens ja concluidos ausentes — a conclusao anterior estava errada")
    End Sub

    <TestMethod>
    Public Sub Ambiente_igual_nao_exige_reconciliacao()
        Assert.IsFalse(EnvironmentPolicy.ExigeReconciliacao(Medido(), Medido()),
            "senao toda abertura reconciliaria a caixa inteira")
    End Sub

    <TestMethod>
    Public Sub Ambiente_desconhecido_de_um_lado_exige_reconciliacao()
        Assert.IsTrue(EnvironmentPolicy.ExigeReconciliacao(Nothing, Medido()))
        Assert.IsTrue(EnvironmentPolicy.ExigeReconciliacao(Medido(), Nothing))
    End Sub

    ''' <summary>
    ''' Toda linha da matriz aponta para onde foi medida.
    '''
    ''' Sem isso, "ambiente suportado" vira lista que cresce por
    ''' conveniência: alguém acrescenta uma linha para destravar um caso e
    ''' ninguém consegue perguntar depois "medido onde?".
    ''' </summary>
    <TestMethod>
    Public Sub Toda_linha_da_matriz_tem_evidencia()
        For Each linha In EnvironmentPolicy.Matriz
            Assert.IsFalse(String.IsNullOrWhiteSpace(linha.Evidence),
                $"linha sem evidencia: {linha.Fingerprint.Value()}")
            StringAssert.Contains(linha.Evidence, "§",
                "a evidencia precisa apontar a secao do FASE2.md")
        Next
    End Sub

    ''' <summary>
    ''' A matriz não tem linhas duplicadas — duas linhas para a mesma
    ''' impressão digital tornariam <c>Medido</c> dependente da ordem.
    ''' </summary>
    ''' <summary>
    ''' A linha medida registra o ALCANCE observado — não fica em Nothing.
    '''
    ''' Uma linha sem alcance seria uma medição sem medida: diria "este
    ''' ambiente foi verificado" sem dizer até onde se enxergou nele.
    ''' </summary>
    <TestMethod>
    Public Sub A_linha_medida_registra_o_alcance_observado()
        Dim linha = EnvironmentPolicy.Medido(Medido())
        Assert.IsNotNull(linha, "a linha medida sumiu da matriz")
        Assert.IsTrue(linha.AlcanceMedido.HasValue,
            "medicao sem alcance nao diz ate onde se enxergou")
        Assert.IsTrue(linha.AlcanceMedido.Value < Date.Today)
    End Sub

    <TestMethod>
    Public Sub A_matriz_nao_tem_impressao_repetida()
        Dim chaves = EnvironmentPolicy.Matriz.Select(Function(x) x.Fingerprint.Value()).ToList()
        Assert.AreEqual(chaves.Count, chaves.Distinct().Count())
    End Sub

End Class

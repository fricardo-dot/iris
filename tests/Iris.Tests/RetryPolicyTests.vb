Imports Iris.Sync
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Retry, backoff e justiça da fila — adversários 55 a 61.
'''
''' O que estes testes protegem é um laço infinito. A §17.1 diz "discordou,
''' descarte e repita", e numa Caixa de Entrada que recebe mensagem sozinha
''' a discordância pode ser <b>permanente</b>. Um Iris que insiste para
''' sempre numa pasta movimentada trava a fila da STA e não entrega nada —
''' pior que um que admite não ter conseguido.
''' </summary>
<TestClass>
Public Class RetryPolicyTests

    Private Shared ReadOnly T0 As New DateTimeOffset(2026, 8, 24, 9, 0, 0, TimeSpan.Zero)

    <TestMethod>
    Public Sub Primeira_tentativa_e_imediata()
        Dim r = RetryPolicy.Decidir(0, Nothing, T0, emVoo:=False)
        Assert.AreEqual(RetryDecision.Tentar, r.Decision)
    End Sub

    ''' <summary>
    ''' Adversário 55: a quarta consecutiva não acontece.
    ''' </summary>
    <TestMethod>
    Public Sub Tres_tentativas_e_a_quarta_degrada()
        For falhas = 0 To 2
            Dim r = RetryPolicy.Decidir(falhas, T0.AddHours(-1), T0, emVoo:=False)
            Assert.AreNotEqual(RetryDecision.Degradar, r.Decision,
                $"com {falhas} falhas ainda pode tentar")
        Next
        Dim quarta = RetryPolicy.Decidir(3, T0.AddHours(-1), T0, emVoo:=False)
        Assert.AreEqual(RetryDecision.Degradar, quarta.Decision)
        StringAssert.Contains(quarta.Reason, "sem convergir")
    End Sub

    <TestMethod>
    Public Sub Backoff_e_30s_2min_8min_e_teto_de_15min()
        Assert.AreEqual(TimeSpan.FromSeconds(30), RetryPolicy.EsperaApos(1))
        Assert.AreEqual(TimeSpan.FromMinutes(2), RetryPolicy.EsperaApos(2))
        Assert.AreEqual(TimeSpan.FromMinutes(8), RetryPolicy.EsperaApos(3))
        For n = 4 To 50
            Assert.AreEqual(TimeSpan.FromMinutes(15), RetryPolicy.EsperaApos(n),
                            $"a {n}a espera deveria estar no teto")
        Next
    End Sub

    <TestMethod>
    Public Sub Antes_do_backoff_nao_tenta()
        Dim falhou = T0
        Dim r = RetryPolicy.Decidir(1, falhou, falhou.AddSeconds(29), emVoo:=False)
        Assert.AreEqual(RetryDecision.Aguardar, r.Decision)
        Assert.AreEqual(falhou.AddSeconds(30), r.NotBefore)

        Dim depois = RetryPolicy.Decidir(1, falhou, falhou.AddSeconds(30), emVoo:=False)
        Assert.AreEqual(RetryDecision.Tentar, depois.Decision)
    End Sub

    ''' <summary>
    ''' Adversário 58: duas varreduras concorrentes da MESMA pasta produzem
    ''' exatamente o corte fraturado que a §16.1 mediu.
    ''' </summary>
    <TestMethod>
    Public Sub Nunca_dois_retries_da_mesma_pasta_em_voo()
        For falhas = 0 To 5
            Dim r = RetryPolicy.Decidir(falhas, T0.AddHours(-1), T0, emVoo:=True)
            Assert.AreEqual(RetryDecision.Aguardar, r.Decision, $"com {falhas} falhas")
        Next
    End Sub

    ''' <summary>
    ''' Adversário 61: máquina que hibernou, ou horário que mudou, não pode
    ''' liberar uma rajada de retries atrasados de uma vez. O backoff é
    ''' sobre ESPAÇAR, e um salto no relógio não desfaz isso.
    ''' </summary>
    <TestMethod>
    Public Sub Salto_do_relogio_nao_libera_rajada()
        For Each salto In {TimeSpan.FromHours(1), TimeSpan.FromDays(1), TimeSpan.FromDays(30)}
            Assert.AreEqual(1, RetryPolicy.TentativasAutorizadasApos(salto),
                $"um salto de {salto} nao autoriza mais de uma tentativa")
        Next
    End Sub

    ''' <summary>
    ''' Falhou mas ninguém anotou quando. Fail-closed: espera. O contrário
    ''' — tentar porque não se sabe — vira busy-loop na primeira vez que o
    ''' registro do instante falhar.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_instante_da_ultima_falha_espera()
        Dim r = RetryPolicy.Decidir(1, Nothing, T0, emVoo:=False)
        Assert.AreEqual(RetryDecision.Aguardar, r.Decision)
    End Sub

    ''' <summary>
    ''' Adversário 59: sucesso zera tudo. Sem isto, uma pasta que falhou
    ''' três vezes de manhã continuaria degradada à tarde.
    ''' </summary>
    <TestMethod>
    Public Sub Sucesso_zera_o_contador()
        Assert.AreEqual(RetryDecision.Tentar,
                        RetryPolicy.Decidir(0, T0, T0.AddDays(1), emVoo:=False).Decision)
        Assert.AreEqual(TimeSpan.Zero, RetryPolicy.EsperaApos(0))
    End Sub

End Class

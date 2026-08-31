Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A PRIORIDADE DA FILA — Fase 9.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTE ARQUIVO PRENDE</b>
'''
''' Que a ordem seja <b>conferível</b>. Uma nota que o dono não consegue
''' destrinchar é pior do que nenhuma ordenação: ele passa a obedecer a um
''' número que não sabe de onde vem, e quando discordar não terá do que
''' discordar.
'''
''' ------------------------------------------------------------------
''' <b>O CONTROLE NEGATIVO</b>
'''
''' <see cref="O_total_e_sempre_a_soma_das_parcelas_MOSTRADAS"/>. Sem ele, uma
''' parcela escondida — um ajuste, um empurrãozinho — passaria despercebida em
''' todo o resto daqui, e a explicação na tela não bateria com a ordem.
'''
''' E <see cref="A_parcela_de_ZERO_dias_aparece"/>: omitir a de zero pareceria
''' mais limpo e esconderia a informação mais útil da explicação — <i>esta linha
''' está aqui apesar de estar esperando há zero dias</i>.
''' </summary>
<TestClass>
Public Class PrioridadeDaFilaTests

    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 31, 12, 0, 0, TimeSpan.Zero)

    ' ==================================================================
    ' A CONTA É ABERTA

    ''' <summary>
    ''' <b>O controle negativo.</b> O total é a soma das parcelas que a tela
    ''' mostra — não há ajuste escondido, e não pode haver: uma parcela invisível
    ''' faria a explicação não bater com a ordem, e o dono descobriria isso
    ''' comparando duas linhas e não entendendo.
    ''' </summary>
    <TestMethod>
    Public Sub O_total_e_sempre_a_soma_das_parcelas_MOSTRADAS()
        Dim p = PrioridadeDaFila.Pontuar(dias:=12, rotulo:="precisa_de_mim",
                                         regrasCasadas:=2, pessoaProxima:=True,
                                         prazo:=Agora.AddDays(-1), agora:=Agora)

        Assert.AreEqual(p.Parcelas.Sum(Function(x) x.Valor), p.Total, 0.0001)

        ' E a explicacao mostra TODAS elas.
        For Each parcela In p.Parcelas
            StringAssert.Contains(p.Explicar(), parcela.Frase)
        Next
    End Sub

    <TestMethod>
    Public Sub A_conta_bate_com_papel_e_caneta()
        Dim p = PrioridadeDaFila.Pontuar(dias:=10, rotulo:="precisa_de_mim",
                                         regrasCasadas:=1)

        ' 10 dias + 20 de "espera voce" + 5 da regra = 35.
        Assert.AreEqual(35.0, p.Total, 0.0001)
    End Sub

    ''' <summary>
    ''' A parcela de zero dias aparece. Omiti-la pareceria mais limpo e
    ''' esconderia a informação mais útil da explicação: esta linha está aqui
    ''' <i>apesar</i> de não estar esperando há tempo nenhum.
    ''' </summary>
    <TestMethod>
    Public Sub A_parcela_de_ZERO_dias_aparece()
        Dim p = PrioridadeDaFila.Pontuar(dias:=0, rotulo:="fyi", regrasCasadas:=0)

        Assert.AreEqual(1, p.Parcelas.Count)
        StringAssert.Contains(p.Explicar(), "0 dia(s) esperando")
    End Sub

    ''' <summary>
    ''' A explicação é em português e termina no total. Nome interno explica para
    ''' quem escreveu o código, e para mais ninguém.
    ''' </summary>
    <TestMethod>
    Public Sub A_explicacao_e_em_portugues_e_fecha_no_total()
        Dim p = PrioridadeDaFila.Pontuar(dias:=3, rotulo:="precisa_de_mim",
                                         regrasCasadas:=0)

        StringAssert.Contains(p.Explicar(), "alguém espera uma resposta sua")
        StringAssert.Contains(p.Explicar(), "total: 23")
        Assert.IsFalse(p.Explicar().Contains("espera_resposta"))
    End Sub

    ' ==================================================================
    ' AS PARCELAS

    <TestMethod>
    Public Sub Esperar_resposta_vale_mais_que_uma_espera_razoavel()
        Dim comRotulo = PrioridadeDaFila.Pontuar(0, "precisa_de_mim", 0)
        Dim semRotulo = PrioridadeDaFila.Pontuar(14, "fyi", 0)

        Assert.IsTrue(comRotulo.Total > semRotulo.Total,
            "duas semanas de FYI passaram na frente de quem espera resposta")
    End Sub

    ''' <summary>
    ''' A regra do dono vale <b>menos</b> que o rótulo, e de propósito: ela diz
    ''' sobre o que é a mensagem, não que alguém está esperando. Uma reclamação
    ''' de cliente já respondida não é urgente por ser reclamação.
    ''' </summary>
    <TestMethod>
    Public Sub A_regra_do_dono_pesa_menos_que_esperar_resposta()
        Assert.IsTrue(PrioridadeDaFila.PorRegraDoDono < PrioridadeDaFila.PorEsperarResposta)
    End Sub

    <TestMethod>
    Public Sub Duas_regras_casadas_pesam_o_dobro_de_uma()
        Dim uma = PrioridadeDaFila.Pontuar(0, "fyi", 1).Total
        Dim duas = PrioridadeDaFila.Pontuar(0, "fyi", 2).Total

        Assert.AreEqual(PrioridadeDaFila.PorRegraDoDono, duas - uma, 0.0001)
    End Sub

    ''' <summary>
    ''' <b>Sem prazo não há parcela de prazo.</b> <c>Nothing</c> não vale como
    ''' "sem pressa": vale como "ninguém disse", e transformar silêncio em
    ''' folga seria uma afirmação.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_prazo_nao_ha_parcela_de_prazo()
        Dim p = PrioridadeDaFila.Pontuar(0, "fyi", 0, agora:=Agora)

        Assert.IsFalse(p.Parcelas.Any(Function(x) x.Nome = "prazo"))
    End Sub

    <TestMethod>
    Public Sub Prazo_vencido_pesa_mais_que_prazo_desta_semana()
        Dim vencido = PrioridadeDaFila.Pontuar(0, "fyi", 0, prazo:=Agora.AddDays(-1),
                                               agora:=Agora)
        Dim perto = PrioridadeDaFila.Pontuar(0, "fyi", 0, prazo:=Agora.AddDays(2),
                                             agora:=Agora)

        Assert.IsTrue(vencido.Total > perto.Total)
    End Sub

    <TestMethod>
    Public Sub Prazo_longe_nao_vira_parcela()
        Dim p = PrioridadeDaFila.Pontuar(0, "fyi", 0,
                                         prazo:=Agora.AddDays(PrioridadeDaFila.DiasDePrazoPerto + 1),
                                         agora:=Agora)

        Assert.IsFalse(p.Parcelas.Any(Function(x) x.Nome = "prazo"))
    End Sub

    ''' <summary>
    ''' Prazo sem relógio não vira parcela nenhuma. Comparar um prazo com um
    ''' "agora" que ninguém passou exigiria ler o relógio aqui dentro — e uma
    ''' nota que muda sozinha entre duas leituras não é conferível.
    ''' </summary>
    <TestMethod>
    Public Sub Prazo_sem_relogio_nao_vira_parcela()
        Dim p = PrioridadeDaFila.Pontuar(0, "fyi", 0, prazo:=Agora.AddDays(-10))

        Assert.IsFalse(p.Parcelas.Any(Function(x) x.Nome = "prazo"))
    End Sub

    ''' <summary>
    ''' <b>Sem saber, não afirmar.</b> "Pessoa próxima" é decisão de quem chama,
    ''' e o padrão é não pontuar — o contrário faria toda linha ganhar dez
    ''' pontos por uma informação que ninguém tem.
    ''' </summary>
    <TestMethod>
    Public Sub Pessoa_proxima_e_FALSO_por_padrao()
        Dim p = PrioridadeDaFila.Pontuar(0, "fyi", 0)

        Assert.IsFalse(p.Parcelas.Any(Function(x) x.Nome = "pessoa"))
    End Sub

    <TestMethod>
    Public Sub Dias_negativos_nao_viram_pontos_negativos()
        Dim p = PrioridadeDaFila.Pontuar(-5, "fyi", 0)

        Assert.AreEqual(0.0, p.Total, 0.0001)
    End Sub

    ' ==================================================================
    ' A ESCALA

    ''' <summary>
    ''' <b>Linear, e não exponencial.</b> Uma curva que dispara faria a linha de
    ''' 60 dias esmagar todas as outras para sempre — e uma pendência que nunca
    ''' sai de cima é uma pendência que o dono aprende a ignorar.
    ''' </summary>
    <TestMethod>
    Public Sub A_espera_cresce_LINEARMENTE()
        Dim dez = PrioridadeDaFila.Pontuar(10, "fyi", 0).Total
        Dim vinte = PrioridadeDaFila.Pontuar(20, "fyi", 0).Total
        Dim trinta = PrioridadeDaFila.Pontuar(30, "fyi", 0).Total

        Assert.AreEqual(vinte - dez, trinta - vinte, 0.0001)
    End Sub

    ''' <summary>
    ''' <b>Os pesos são congelados aqui de propósito.</b> Nenhum deles saiu de
    ''' medição — são palpites com justificativa. Quando houver dado para
    ''' calibrá-los, este teste falha e obriga quem os mudou a dizer por quê.
    ''' </summary>
    <TestMethod>
    Public Sub Os_pesos_de_hoje_sao_ESTES()
        Assert.AreEqual(1.0, PrioridadeDaFila.PorDia, 0.0001)
        Assert.AreEqual(20.0, PrioridadeDaFila.PorEsperarResposta, 0.0001)
        Assert.AreEqual(5.0, PrioridadeDaFila.PorRegraDoDono, 0.0001)
        Assert.AreEqual(10.0, PrioridadeDaFila.PorPessoaProxima, 0.0001)
        Assert.AreEqual(30.0, PrioridadeDaFila.PorPrazoVencido, 0.0001)
        Assert.AreEqual(10.0, PrioridadeDaFila.PorPrazoPerto, 0.0001)
        Assert.AreEqual(7, PrioridadeDaFila.DiasDePrazoPerto)
    End Sub

End Class

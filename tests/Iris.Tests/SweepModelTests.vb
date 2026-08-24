Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Sync
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' A máquina de estados da varredura — adversários 1 a 26 e 47 a 66.
'''
''' O que estes testes protegem: <b>publicar metade</b>. A §16.1 mediu que
''' um único `Move` durante a varredura trunca a `Table` em 16 itens, três
''' vezes seguidas, com `EndOfTable` normal e sem erro. Uma varredura assim
''' entrega 60% da pasta parecendo completa — e se ela publicar, o que
''' faltou vira suspeita, e depois ausência.
'''
''' Por isso a divisão de autoridade: <b>S6 rejeita e não valida</b>.
''' Contagens concordando não provam nada; contagens discordando provam que
''' a varredura não vale.
''' </summary>
<TestClass>
Public Class SweepModelTests

    Private Shared Function Universo(Optional cutoff As String = "1m",
                                     Optional alg As Integer = 1,
                                     Optional amb As String = "exchange-cached-1m") As SweepUniverse
        Return New SweepUniverse("S1", "F1", "", cutoff, alg, amb)
    End Function

    ''' <summary>Uma varredura completa e correta, para partir dela.</summary>
    Private Shared Function Ate(estagio As AttemptStage,
                                Optional chaves As String() = Nothing,
                                Optional epoca As Long = 7) As SweepAttempt
        Dim ks = If(chaves, New String() {"A", "B", "C"})
        Dim a = SweepModel.Abrir(Universo(), epoca, 1, True).State
        If estagio = AttemptStage.Aberta Then Return a
        a = SweepModel.ContagemInicial(a, ks.Length, Universo()).State
        If estagio = AttemptStage.ContagemInicialLida Then Return a
        a = SweepModel.Pagina(a, ks, "cur1", Universo()).State
        If estagio = AttemptStage.Varrendo Then Return a
        a = SweepModel.ContagemFinal(a, ks.Length, Universo()).State
        Return a
    End Function

    ' ==================================================================
    ' Controle positivo (adversário 1)
    ' ==================================================================

    <TestMethod>
    Public Sub Caminho_feliz_publica_exatamente_uma_geracao()
        Dim a = Ate(AttemptStage.ContagemFinalLida)
        Dim r = SweepModel.Publicar(a, 7)

        Assert.IsFalse(r.Rejected, r.Rejection)
        Assert.AreEqual(AttemptStage.Publicada, r.State.Stage)
        CollectionAssert.AreEquivalent(
            New SweepCommand() {SweepCommand.PublicarGeracao,
                                SweepCommand.MarcarNaoVistosComoSuspeitos,
                                SweepCommand.EmitirPublicationLog},
            r.Commands.ToArray())
    End Sub

    ''' <summary>
    ''' Sem este teste, uma máquina que NUNCA publica passaria em todos os
    ''' testes de segurança. É o mesmo controle positivo que faltou no
    ''' `Restrict` da §16.5 e me fez publicar duas conclusões opostas.
    ''' </summary>
    <TestMethod>
    Public Sub A_maquina_realmente_publica_alguma_coisa()
        Assert.AreEqual(AttemptStage.Publicada,
                        SweepModel.Publicar(Ate(AttemptStage.ContagemFinalLida), 7).State.Stage)
    End Sub

    ' ==================================================================
    ' As rejeições do S6 (adversários 4 a 10)
    ' ==================================================================

    <TestMethod>
    Public Sub S6_rejeita_lidas_diferente_de_contagem_inicial()
        ' Contagem inicial diz 5, mas so 3 linhas vieram — a truncagem da §16.1.
        Dim a = SweepModel.Abrir(Universo(), 7, 1, True).State
        a = SweepModel.ContagemInicial(a, 5, Universo()).State
        a = SweepModel.Pagina(a, {"A", "B", "C"}, "c", Universo()).State
        a = SweepModel.ContagemFinal(a, 5, Universo()).State

        Dim r = SweepModel.Publicar(a, 7)
        Assert.IsTrue(r.Rejected, "deveria rejeitar: leu 3 de 5")
        Assert.AreEqual(AttemptStage.Descartada, r.State.Stage)
    End Sub

    <TestMethod>
    Public Sub S6_rejeita_contagem_final_diferente()
        Dim a = SweepModel.Abrir(Universo(), 7, 1, True).State
        a = SweepModel.ContagemInicial(a, 3, Universo()).State
        a = SweepModel.Pagina(a, {"A", "B", "C"}, "c", Universo()).State
        a = SweepModel.ContagemFinal(a, 4, Universo()).State
        Assert.IsTrue(SweepModel.Publicar(a, 7).Rejected, "chegou item durante a varredura")
    End Sub

    <TestMethod>
    Public Sub S6_rejeita_chave_vazia()
        Dim a = Ate(AttemptStage.ContagemInicialLida)
        Dim r = SweepModel.Pagina(a, {"A", "", "C"}, "c", Universo())
        Assert.IsTrue(r.Rejected)
        StringAssert.Contains(r.Rejection, "chave")
    End Sub

    ''' <summary>
    ''' A mesma chave em páginas diferentes. Pode ser cursor repetindo ou
    ''' fonte instável; das duas, a varredura não vale. E a §16.1 mostrou
    ''' que a `Table` faz coisas estranhas sob mutação.
    ''' </summary>
    <TestMethod>
    Public Sub S6_rejeita_chave_repetida_entre_paginas()
        Dim a = Ate(AttemptStage.ContagemInicialLida)
        a = SweepModel.Pagina(a, {"A", "B"}, "c1", Universo()).State
        Dim r = SweepModel.Pagina(a, {"B", "C"}, "c2", Universo())
        Assert.IsTrue(r.Rejected)
        StringAssert.Contains(r.Rejection, "repetida")
    End Sub

    <TestMethod>
    Public Sub S6_rejeita_universo_diferente_entre_contagem_e_manifesto()
        Dim a = Ate(AttemptStage.ContagemInicialLida)
        Dim r = SweepModel.Pagina(a, {"A"}, "c", Universo(cutoff:="3m"))
        Assert.IsTrue(r.Rejected)
        StringAssert.Contains(r.Rejection, "universo")
    End Sub

    ''' <summary>
    ''' Adversário 20: a mutação balanceada. Sai um, entra outro, e as três
    ''' contagens continuam iguais — o S6 NÃO pega. É por isso que
    ''' concordância não valida, e por isso o S7 existe em separado.
    ''' </summary>
    <TestMethod>
    Public Sub Mutacao_balanceada_PASSA_no_S6_e_isso_e_esperado()
        ' 3 antes, 3 depois, 3 lidas — mas "B" saiu e "D" entrou.
        Dim a = SweepModel.Abrir(Universo(), 7, 1, True).State
        a = SweepModel.ContagemInicial(a, 3, Universo()).State
        a = SweepModel.Pagina(a, {"A", "C", "D"}, "c", Universo()).State
        a = SweepModel.ContagemFinal(a, 3, Universo()).State

        Dim r = SweepModel.Publicar(a, 7)
        Assert.IsFalse(r.Rejected,
            "o S6 nao pega mutacao balanceada, e afirmar que pega seria mentira")
        ' O que salva "B" e nao chegar a ausencia por aqui: ele vira
        ' SUSPEITO, e so verificacao individual com cobertura decide.
        Assert.IsTrue(r.Commands.Contains(SweepCommand.MarcarNaoVistosComoSuspeitos))
    End Sub

    ''' <summary>
    ''' Adversário 21: a pasta parcial com zero. `rows = before = after = 0`
    ''' passa no S6 com folga — e é exatamente a §19.2.
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_vazia_publica_mas_nao_confirma_ausencia()
        Dim a = SweepModel.Abrir(Universo(), 7, 1, True).State
        a = SweepModel.ContagemInicial(a, 0, Universo()).State
        a = SweepModel.ContagemFinal(a, 0, Universo()).State

        Dim r = SweepModel.Publicar(a, 7)
        Assert.IsFalse(r.Rejected, "zero e zero e zero: o S6 aceita")
        Assert.IsFalse(r.Commands.Any(Function(c) c.ToString().Contains("Ausen")),
            "nenhum comando de ausencia sai daqui — quem decide isso e o S7")
    End Sub

    ' ==================================================================
    ' Fronteiras (adversários 11 a 15)
    ' ==================================================================

    <TestMethod>
    Public Sub Cancelamento_em_qualquer_fronteira_nao_publica_metade()
        For Each estagio In {AttemptStage.Aberta, AttemptStage.ContagemInicialLida,
                             AttemptStage.Varrendo, AttemptStage.ContagemFinalLida}
            Dim r = SweepModel.Cancelar(Ate(estagio), "usuario cancelou")
            Assert.AreEqual(AttemptStage.Descartada, r.State.Stage, $"em {estagio}")
            Assert.IsFalse(r.Commands.Contains(SweepCommand.PublicarGeracao), $"em {estagio}")
            Assert.IsTrue(r.Commands.Contains(SweepCommand.DescartarTentativa), $"em {estagio}")
        Next
    End Sub

    <TestMethod>
    Public Sub Falha_da_fonte_em_qualquer_fronteira_descarta()
        For Each estagio In {AttemptStage.Aberta, AttemptStage.ContagemInicialLida,
                             AttemptStage.Varrendo, AttemptStage.ContagemFinalLida}
            Dim r = SweepModel.Falhar(Ate(estagio), "RPC_S_SERVER_UNAVAILABLE")
            Assert.AreEqual(AttemptStage.Descartada, r.State.Stage, $"em {estagio}")
            Assert.IsTrue(r.Commands.Contains(SweepCommand.AgendarRetry), $"em {estagio}")
        Next
    End Sub

    ' ==================================================================
    ' Ambiente e universo (adversários 16 a 19)
    ' ==================================================================

    ''' <summary>D2: ambiente fora da allowlist nem abre tentativa.</summary>
    <TestMethod>
    Public Sub Ambiente_nao_suportado_nem_comeca()
        Dim r = SweepModel.Abrir(Universo(amb:="pst-desconhecido"), 7, 1, ambienteSuportado:=False)
        Assert.IsTrue(r.Rejected)
        Assert.IsNull(r.State, "nem chega a existir tentativa")
        StringAssert.Contains(r.Rejection, "allowlist")
    End Sub

    <TestMethod>
    Public Sub Mudanca_de_universo_no_meio_abandona()
        For Each mudado In {Universo(cutoff:="3m"), Universo(alg:=2), Universo(amb:="outro")}
            Dim a = Ate(AttemptStage.Varrendo)
            Dim r = SweepModel.ContagemFinal(a, 3, mudado)
            Assert.IsTrue(r.Rejected, "universo mudou e a tentativa continuou")
        Next
    End Sub

    ' ==================================================================
    ' Fencing e idempotência (adversários 24 a 26, 62 a 64)
    ' ==================================================================

    ''' <summary>
    ''' Adversário 24, que é o item 10 do antigo §8: tentativa velha termina
    ''' DEPOIS de a nova ter publicado. O CAS da velha falha.
    ''' </summary>
    <TestMethod>
    Public Sub Geracao_velha_nao_sobrescreve_a_nova()
        Dim velha = Ate(AttemptStage.ContagemFinalLida, epoca:=7)
        ' Enquanto ela corria, outra publicou e a epoca da pasta avancou.
        Dim r = SweepModel.Publicar(velha, epocaCorrenteDaPasta:=8)

        Assert.IsTrue(r.Rejected, "a tentativa da epoca 7 nao pode publicar sobre a 8")
        StringAssert.Contains(r.Rejection, "epoca")
        Assert.AreEqual(AttemptStage.Descartada, r.State.Stage)
    End Sub

    <TestMethod>
    Public Sub Republicar_a_mesma_tentativa_e_idempotente()
        Dim r1 = SweepModel.Publicar(Ate(AttemptStage.ContagemFinalLida), 7)
        Dim r2 = SweepModel.Publicar(r1.State, 7)
        Assert.IsFalse(r2.Rejected)
        Assert.AreEqual(0, r2.Commands.Count,
            "republicar nao emite segundo comando: nao cria segunda geracao nem segundo evento")
    End Sub

    <TestMethod>
    Public Sub Geracao_publicada_rejeita_mutacao_posterior()
        Dim pub = SweepModel.Publicar(Ate(AttemptStage.ContagemFinalLida), 7).State
        Assert.IsTrue(SweepModel.Pagina(pub, {"Z"}, "c", Universo()).Rejected)
        Assert.IsTrue(SweepModel.ContagemFinal(pub, 3, Universo()).Rejected)
    End Sub

    <TestMethod>
    Public Sub Tentativa_descartada_nunca_volta_a_publicavel()
        Dim descartada = SweepModel.Cancelar(Ate(AttemptStage.Varrendo), "x").State
        Assert.IsTrue(SweepModel.Pagina(descartada, {"Z"}, "c", Universo()).Rejected)
        Assert.IsTrue(SweepModel.Publicar(descartada, 7).Rejected)
    End Sub

    ''' <summary>
    ''' Adversário 64: o publication_log só sai JUNTO da publicação. Se
    ''' saísse antes, um crash entre os dois reprocessaria uma publicação
    ''' que nunca houve.
    ''' </summary>
    <TestMethod>
    Public Sub PublicationLog_so_sai_junto_da_publicacao()
        Dim comLog = SweepModel.Publicar(Ate(AttemptStage.ContagemFinalLida), 7)
        Assert.IsTrue(comLog.Commands.Contains(SweepCommand.EmitirPublicationLog))
        Assert.IsTrue(comLog.Commands.Contains(SweepCommand.PublicarGeracao))

        For Each estagio In {AttemptStage.Aberta, AttemptStage.ContagemInicialLida, AttemptStage.Varrendo}
            Dim r = SweepModel.Cancelar(Ate(estagio), "x")
            Assert.IsFalse(r.Commands.Contains(SweepCommand.EmitirPublicationLog), $"em {estagio}")
        Next
    End Sub

    ' ==================================================================
    ' Retomada (adversários 47 a 53)
    ' ==================================================================

    <TestMethod>
    Public Sub Retomar_com_universo_diferente_abandona()
        Dim a = Ate(AttemptStage.Varrendo)
        Assert.IsTrue(SweepModel.Retomar(a, Universo(cutoff:="3m"), 7).Rejected)
    End Sub

    <TestMethod>
    Public Sub Retomar_com_epoca_diferente_abandona()
        Dim a = Ate(AttemptStage.Varrendo, epoca:=7)
        Assert.IsTrue(SweepModel.Retomar(a, Universo(), 9).Rejected)
    End Sub

    ''' <summary>
    ''' Adversário 48: cursor avançado sem a página correspondente é estado
    ''' inválido. Se fosse aceito, a retomada continuaria de um ponto cujas
    ''' linhas ninguém guardou — pulando-as em silêncio.
    ''' </summary>
    <TestMethod>
    Public Sub Cursor_avancado_sem_pagina_e_estado_invalido()
        Dim a = Ate(AttemptStage.ContagemInicialLida)
        ' Fabrica o estado corrompido: cursor sem linha nenhuma.
        Dim corrompido = SweepModel.Pagina(a, New String() {}, "cursor-orfao", Universo()).State
        Assert.IsTrue(SweepModel.Retomar(corrompido, Universo(), 7).Rejected)
    End Sub

    <TestMethod>
    Public Sub Retomada_nunca_emite_confirmacao_de_ausencia()
        Dim a = Ate(AttemptStage.Varrendo)
        Dim r = SweepModel.Retomar(a, Universo(), 7)
        Assert.IsFalse(r.Commands.Any(Function(c) c.ToString().Contains("Ausen")))
    End Sub

    <TestMethod>
    Public Sub Repetir_a_ultima_pagina_staged_nao_duplica_chaves()
        Dim a = Ate(AttemptStage.ContagemInicialLida)
        a = SweepModel.Pagina(a, {"A", "B"}, "c1", Universo()).State
        Assert.AreEqual(2, a.DistinctKeys)
        ' Reentrega da MESMA pagina: as chaves ja estao staged, entao isto e
        ' repeticao, e repeticao e rejeitada — o replay tem de ser detectado
        ' pela camada transacional, nao aceito como pagina nova.
        Assert.IsTrue(SweepModel.Pagina(a, {"A", "B"}, "c1", Universo()).Rejected)
    End Sub

End Class

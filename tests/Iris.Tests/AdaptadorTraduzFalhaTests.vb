Imports System.Threading
Imports Iris.Core
Imports Iris.Integration.Outlook
Imports Iris.Model
Imports Iris.Sync
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A tradução do adaptador, exercitada de ponta a ponta.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ISTO EXISTE, E POR QUE O TESTE ANTERIOR NÃO BASTAVA</b>
'''
''' O <c>OutlookSweepSource</c> traduz o <c>OperationResult</c> do broker em
''' exceção, e é aí que se decide o desfecho da varredura:
'''
'''   • <c>ErrorKind.Cancelled</c> → <c>OperationCanceledException</c>, que o
'''     runner conclui como <c>Cancelada</c>;
'''   • qualquer outro <c>ErrorKind</c> → <c>SourceUnavailableException</c>,
'''     que o runner conclui como <c>Falhou</c>, com a causa preservada.
'''
''' Eu tinha escrito um teste para a primeira regra usando uma <b>fonte falsa
''' que já lançava <c>OperationCanceledException</c></b> — ou seja, ele
''' exercitava o <c>Catch</c> do runner, que já existia, e <b>não</b> a linha
''' nova do adaptador. O Codex pegou: apagando a conversão, aquele teste
''' continuava verde e a regressão passava.
'''
''' Aqui o caminho é o inteiro — broker falso, adaptador real, runner real.
''' </summary>
<TestClass>
Public Class AdaptadorTraduzFalhaTests

    Private Shared Function Universo() As SweepUniverse
        Return New SweepUniverse("store-1", "pasta-1", "f", Nothing, 1, "amb-1")
    End Function

    ''' <summary>Broker que responde a paginação com a falha pedida.</summary>
    Private Shared Function Recusando(k As ErrorKind) As FakeBroker
        Return New FakeBroker() With {
            .RespostaDaPagina = OperationResult(Of MessagePage).Fail(k, "de proposito")}
    End Function

    Private Shared Function Rodar(b As FakeBroker) As SweepResult
        Dim fonte As New OutlookSweepSource(b, New FolderKey("store-1", "pasta-1"),
                                            Universo(), 1)
        Dim cap = EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
        Return New SweepRunner(fonte, New DestinoFalso(), 2).
               Executar(Universo(), 0, 1, cap, CancellationToken.None)
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <c>Cancelled</c> do broker vira <b>cancelamento</b> — e este teste
    ''' falha se alguém apagar a conversão no adaptador.
    ''' </summary>
    <TestMethod>
    Public Sub Cancelled_do_broker_conclui_Cancelada()
        Dim r = Rodar(Recusando(ErrorKind.Cancelled))

        Assert.AreEqual(SweepConclusion.Cancelada, r.Conclusion,
            "sem a conversao no adaptador isto sai Falhou, e 'o usuario mandou " &
            "parar' fica indistinguivel de 'a varredura quebrou'. Motivo: " & r.Motivo)
        Assert.IsFalse(r.CausaDaFonte.HasValue,
                       "cancelamento nao e recusa classificada")
    End Sub

    ''' <summary>
    ''' O contraponto, e sem ele o de cima passaria num adaptador que
    ''' converte <b>tudo</b> em cancelamento — o que esconderia toda falha real
    ''' atrás de um desfecho que ninguém investiga.
    ''' </summary>
    <DataTestMethod>
    <DataRow(ErrorKind.Busy)>
    <DataRow(ErrorKind.NotConnected)>
    <DataRow(ErrorKind.Denied)>
    <DataRow(ErrorKind.Unexpected)>
    <DataRow(ErrorKind.Stale)>
    Public Sub O_resto_conclui_Falhou_com_a_causa_preservada(k As ErrorKind)
        Dim r = Rodar(Recusando(k))

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion,
                        $"{k} nao pode virar cancelamento: " & r.Motivo)
        Assert.IsTrue(r.CausaDaFonte.HasValue, "a causa tem de atravessar o adaptador")
        Assert.AreEqual(k, r.CausaDaFonte.Value)
    End Sub

    ''' <summary>
    ''' Controle: a varredura nem chega a publicar nesses casos. Sem isto, um
    ''' adaptador que engolisse a falha e devolvesse página vazia passaria nos
    ''' dois testes de cima ao custo de publicar uma pasta vazia por cima de
    ''' uma cheia.
    ''' </summary>
    <TestMethod>
    Public Sub Nenhuma_recusa_chega_a_publicar()
        For Each k In {ErrorKind.Cancelled, ErrorKind.Busy, ErrorKind.Denied}
            Assert.IsFalse(Rodar(Recusando(k)).Publicou,
                           $"{k} nao pode publicar nada")
        Next
    End Sub

End Class

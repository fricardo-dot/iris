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


    ''' ==================================================================

    ''' <summary>
    ''' <b>O <c>MessageClass</c> que chega ao cache e o MEDIDO, e nao uma
    ''' constante.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' Ate 28/08/2026 o <c>Traduzir</c> gravava <c>"IPM.Note"</c> fixo. Nenhum
    ''' teste percebeu, e nenhum teria: o filtro da paginacao so deixa passar
    ''' linha que COMECA com <c>IPM.Note</c>, entao a constante coincidia com a
    ''' verdade em todos os casos comuns.
    '''
    ''' Quem encontrou foi a MEDICAO do acervo real: 1.123 linhas e <b>uma</b>
    ''' classe distinta. Numero limpo demais para ser medida.
    '''
    ''' Este teste usa uma subclasse — <c>IPM.Note.SMIME</c> — que passa pelo
    ''' filtro e <b>nao</b> e igual a constante. E o unico formato em que a
    ''' diferenca entre "medido" e "afirmado" fica visivel de fora.
    ''' </summary>
    <TestMethod>
    Public Sub MessageClass_chega_ao_destino_como_veio_do_broker()
        Dim b As New FakeBroker()
        b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
            New MessagePage With {
                .Items = New List(Of MailSummary)() From {
                    New MailSummary With {
                        .Key = New ItemKey("E-1", "store-1"),
                        .Subject = "assinada", .SenderName = "quem",
                        .ReceivedTime = New DateTimeOffset(2026, 8, 28, 9, 0, 0, TimeSpan.Zero),
                        .MessageClass = "IPM.Note.SMIME"},
                    New MailSummary With {
                        .Key = New ItemKey("E-2", "store-1"),
                        .Subject = "comum", .SenderName = "quem",
                        .ReceivedTime = New DateTimeOffset(2026, 8, 28, 9, 1, 0, TimeSpan.Zero),
                        .MessageClass = "IPM.Note"}},
                .NextCursor = Nothing,
                .TotalAtStart = 2})

        Dim destino As New DestinoFalso()
        Dim fonte As New OutlookSweepSource(b, New FolderKey("pasta-1", "store-1"),
                                            Universo(), 1)
        Dim cap = EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
        Dim r = New SweepRunner(fonte, destino, 10).
                Executar(Universo(), 0, 1, cap, CancellationToken.None)

        Assert.AreEqual(2, destino.LinhasGravadas.Count,
            $"controle: as duas linhas tinham de chegar ao destino. motivo: {r.Motivo}")

        Dim classes = destino.LinhasGravadas.Select(Function(l) l.MessageClass).OrderBy(Function(x) x).ToList()
        Assert.AreEqual("IPM.Note|IPM.Note.SMIME", String.Join("|", classes),
            "a classe gravada nao e a que o broker mediu")
    End Sub

End Class

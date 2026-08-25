Imports System.Threading
Imports Iris.Model
Imports Iris.Sync
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>Falha da fonte carrega a causa, e a causa é um enum.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ISTO EXISTE</b>
'''
''' Os testes de importação real rodam contra a caixa VIVA do usuário, e o
''' Outlook pode recusar uma chamada a qualquer momento. Eles precisam aceitar
''' esse desfecho sem aceitar <b>defeito meu</b> junto — e a primeira tentativa
''' de separar os dois foi por substring da mensagem de erro.
'''
''' O Codex derrubou: <c>"GetMessagePageAsync falhou: {Kind}"</c> era emitida
''' para TODO <c>ErrorKind</c>, então <c>Unexpected</c>, <c>Stale</c>,
''' <c>Denied</c> e <c>NotImplemented</c> passavam de soluço do ambiente. E
''' <c>"TotalAtStart"</c>, que é violação de contrato da fonte, estava na
''' allowlist.
'''
''' A regra agora é estrutural, e este arquivo é o que a prova. Sem ele, a
''' regra viveria só dentro de um teste que <b>só falha quando o Outlook tem
''' soluço</b> — ou seja, nunca seria exercitada de propósito.
''' </summary>
<TestClass>
Public Class CausaDaFalhaTests

    Private Shared Function Universo() As SweepUniverse
        Return New SweepUniverse("store-1", "pasta-1", "f", Nothing, 1, "amb-1")
    End Function

    Private Shared Function Rodar(f As FonteFalsaMutavel,
                                  Optional d As DestinoFalso = Nothing) As SweepResult
        Dim cap = EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
        Return New SweepRunner(f, If(d, New DestinoFalso()), 2).
               Executar(Universo(), 0, 1, cap, CancellationToken.None)
    End Function

    ' ==================================================================

    ''' <summary>
    ''' Recusa classificada da fonte chega no resultado <b>com o enum</b> — em
    ''' qualquer <c>ErrorKind</c>, não só nos toleráveis.
    ''' </summary>
    <DataTestMethod>
    <DataRow(ErrorKind.Busy)>
    <DataRow(ErrorKind.NotConnected)>
    <DataRow(ErrorKind.Unexpected)>
    <DataRow(ErrorKind.Stale)>
    <DataRow(ErrorKind.Denied)>
    <DataRow(ErrorKind.NotImplemented)>
    Public Sub Recusa_da_fonte_preserva_o_ErrorKind(k As ErrorKind)
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c") With {.RecusarNaPagina = k}

        Dim r = Rodar(f)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion, r.Motivo)
        Assert.IsTrue(r.CausaDaFonte.HasValue, "a causa nao pode sumir no caminho")
        Assert.AreEqual(k, r.CausaDaFonte.Value)
    End Sub

    ''' <summary>Vale também quando a recusa acontece na CONTAGEM.</summary>
    <TestMethod>
    Public Sub Recusa_na_contagem_tambem_preserva()
        Dim f As New FonteFalsaMutavel(Universo(), "a") With {.RecusarNaContagem = ErrorKind.Busy}

        Dim r = Rodar(f)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.AreEqual(ErrorKind.Busy, r.CausaDaFonte.Value)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo, e o mais importante do arquivo.</b>
    '''
    ''' Exceção comum da fonte — que é como sai violação de contrato, incluindo
    ''' a página sem <c>TotalAtStart</c> — chega SEM causa. É isso que faz
    ''' defeito meu reprovar num teste que tolera soluço do ambiente.
    '''
    ''' Se um dia alguém "melhorar" o runner classificando toda exceção da
    ''' fonte como recusa, este teste cai — e tem de cair.
    ''' </summary>
    <TestMethod>
    Public Sub Excecao_COMUM_da_fonte_chega_SEM_causa()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c") With {.LancarNaPagina = 1}

        Dim r = Rodar(f)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.IsFalse(r.CausaDaFonte.HasValue,
            "defeito de contrato NAO pode chegar classificado: seria desculpavel " &
            "num teste que tolera soluco do ambiente. Motivo: " & r.Motivo)
    End Sub

    ''' <summary>
    ''' <b>O DESTINO lançando o mesmo tipo NÃO vira causa da fonte.</b>
    '''
    ''' O <c>Catch</c> do runner cobre o método inteiro, e o destino é chamado
    ''' lá dentro — <c>AbrirTentativa</c>, <c>GravarPagina</c>,
    ''' <c>EpocaCorrente</c>, <c>Publicar</c>. Classificar pelo TIPO fazia um
    ''' defeito de persistência sair como recusa do provider, e daí como
    ''' "soluço do Outlook" tolerado no teste de importação real.
    '''
    ''' Quem marca a origem é a CHAMADA, não o tipo. Este teste é o que cobra
    ''' isso, e ele falha na versão anterior da correção.
    ''' </summary>
    <TestMethod>
    Public Sub Recusa_vinda_do_DESTINO_nao_e_causa_da_fonte()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c")
        Dim d As New DestinoFalso() With {.RecusarAoGravar = ErrorKind.Busy}

        Dim r = Rodar(f, d)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.IsFalse(r.CausaDaFonte.HasValue,
            "o destino lancou, nao a fonte — o TIPO nao pode decidir a origem. " &
            "Motivo: " & r.Motivo)
    End Sub

    ''' <summary>
    ''' <c>Cancelled</c> vindo da fonte é <b>cancelamento</b>, não falha.
    '''
    ''' O broker devolve <c>ErrorKind.Cancelled</c> quando o token já caiu, e o
    ''' adaptador o converte em <c>OperationCanceledException</c>. Sem essa
    ''' conversão, "o usuário mandou parar" sairia como <c>Falhou</c> — dois
    ''' desfechos com significados diferentes para quem lê o log depois.
    '''
    ''' Aqui a fonte falsa lança direto: isto cobra o <b>desfecho do runner</b>,
    ''' e só isso. A conversão no adaptador — que é a linha nova, e a que
    ''' regride se alguém a apagar — está coberta em
    ''' <see cref="AdaptadorTraduzFalhaTests"/>, com broker falso e caminho
    ''' inteiro. Foi exatamente esse buraco que o Codex pegou nesta primeira
    ''' versão: o teste afirmava impedir uma regressão que não tocava.
    ''' </summary>
    <TestMethod>
    Public Sub Cancelamento_da_fonte_e_Cancelada_e_nao_Falhou()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c") With {
            .LancarCancelamentoNaPagina = 1}

        Dim r = Rodar(f)

        Assert.AreEqual(SweepConclusion.Cancelada, r.Conclusion, r.Motivo)
        Assert.IsFalse(r.CausaDaFonte.HasValue,
                       "cancelamento nao e recusa classificada da fonte")
    End Sub

    ''' <summary>Publicação limpa não inventa causa nenhuma.</summary>
    <TestMethod>
    Public Sub Varredura_que_publica_nao_tem_causa()
        Dim r = Rodar(New FonteFalsaMutavel(Universo(), "a", "b", "c"))

        Assert.IsTrue(r.Publicou, r.Motivo)
        Assert.IsFalse(r.CausaDaFonte.HasValue)
    End Sub

    ' ==================================================================

    ''' <summary>
    ''' <b>O pino da política.</b> Fixa, <c>ErrorKind</c> por <c>ErrorKind</c>,
    ''' o que o ambiente produz sozinho.
    '''
    ''' A regra vive num lugar só — <see cref="ErrorPolicy.Transitorio"/> — e
    ''' <c>IsRetryable</c>, <c>DoAmbiente</c> e o teste de importação real
    ''' todos a consultam. Antes eram três cópias escritas à mão, e eu tinha
    ''' posto um teste comparando <b>duas</b> delas e chamado de resolvido: o
    ''' Codex mostrou a terceira, que era justamente a que decidia se uma
    ''' falha reprovava.
    '''
    ''' Com uma cópia só, comparar cópias virou tautologia. O que resta a
    ''' cobrar é o <b>conteúdo</b> da política — e é isto: mexer nela quebra
    ''' este teste, que é onde a decisão tem de ser tomada de novo.
    '''
    ''' <c>Ambiguous</c> e <c>Cancelled</c> aparecem na lista dos negados de
    ''' propósito: são os dois que alguém tolera sem pensar. Repetir
    ''' <c>Ambiguous</c> é o pior que se pode fazer, e <c>Cancelled</c> tem
    ''' desfecho próprio.
    ''' </summary>
    <TestMethod>
    Public Sub A_politica_do_que_e_transitorio_esta_pinada()
        Dim toleram = {ErrorKind.Busy, ErrorKind.NotConnected}

        For Each k As ErrorKind In [Enum].GetValues(GetType(ErrorKind))
            If k = ErrorKind.None Then Continue For
            Assert.AreEqual(toleram.Contains(k), ErrorPolicy.Transitorio(k),
                $"{k}: a politica do que o ambiente produz sozinho mudou")
        Next

        Assert.IsFalse(ErrorPolicy.Transitorio(ErrorKind.Ambiguous),
                       "Ambiguous NUNCA e tolerado: repetir e o pior que se pode fazer")
        Assert.IsFalse(ErrorPolicy.Transitorio(ErrorKind.Cancelled),
                       "Cancelled tem desfecho proprio, nao e soluco do ambiente")
    End Sub

    ''' <summary>
    ''' E as duas fachadas realmente delegam — não voltaram a ter cópia
    ''' própria por baixo.
    ''' </summary>
    <TestMethod>
    Public Sub As_fachadas_delegam_a_politica()
        For Each k As ErrorKind In [Enum].GetValues(GetType(ErrorKind))
            If k = ErrorKind.None Then Continue For
            Dim esperado = ErrorPolicy.Transitorio(k)
            Assert.AreEqual(esperado, OperationResult(Of Integer).Fail(k, "x").IsRetryable,
                            $"{k}: IsRetryable divergiu da politica")
            Assert.AreEqual(esperado, New SourceUnavailableException(k, "x").DoAmbiente,
                            $"{k}: DoAmbiente divergiu da politica")
        Next
    End Sub

End Class

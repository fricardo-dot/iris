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

    Private Shared Function Rodar(f As FonteFalsaMutavel) As SweepResult
        Dim cap = EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
        Return New SweepRunner(f, New DestinoFalso(), 2).
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

    ''' <summary>Publicação limpa não inventa causa nenhuma.</summary>
    <TestMethod>
    Public Sub Varredura_que_publica_nao_tem_causa()
        Dim r = Rodar(New FonteFalsaMutavel(Universo(), "a", "b", "c"))

        Assert.IsTrue(r.Publicou, r.Motivo)
        Assert.IsFalse(r.CausaDaFonte.HasValue)
    End Sub

    ' ==================================================================

    ''' <summary>
    ''' A lista do que o ambiente produz sozinho é <b>exatamente</b> a do
    ''' <c>OperationResult.IsRetryable</c>, e não uma cópia que envelhece à parte.
    '''
    ''' Este teste existe porque a duplicata é a falha silenciosa aqui: alguém
    ''' acrescenta um <c>ErrorKind</c> transitório num lugar, esquece o outro, e
    ''' a divergência só aparece como intermitência meses depois.
    ''' </summary>
    <TestMethod>
    Public Sub DoAmbiente_bate_com_Retryable_em_TODO_ErrorKind()
        For Each k As ErrorKind In [Enum].GetValues(GetType(ErrorKind))
            If k = ErrorKind.None Then Continue For

            Dim retryable = OperationResult(Of Integer).Fail(k, "x").IsRetryable
            Dim doAmbiente = New SourceUnavailableException(k, "x").DoAmbiente

            Assert.AreEqual(retryable, doAmbiente,
                $"{k}: as duas nocoes de 'transitorio' divergiram")
        Next
    End Sub

    ''' <summary>
    ''' Controle: a lista não é vazia nem universal. Sem isto, um
    ''' <c>DoAmbiente</c> que devolvesse sempre <c>False</c> passaria no teste
    ''' de cima junto com um <c>Retryable</c> igualmente quebrado.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_a_lista_do_ambiente_nao_e_vazia_nem_universal()
        Assert.IsTrue(New SourceUnavailableException(ErrorKind.Busy, "x").DoAmbiente,
                      "Busy TEM de ser tolerado")
        Assert.IsTrue(New SourceUnavailableException(ErrorKind.NotConnected, "x").DoAmbiente,
                      "NotConnected TEM de ser tolerado")
        Assert.IsFalse(New SourceUnavailableException(ErrorKind.Unexpected, "x").DoAmbiente,
                       "Unexpected NAO pode ser tolerado")
        Assert.IsFalse(New SourceUnavailableException(ErrorKind.Denied, "x").DoAmbiente,
                       "Denied NAO pode ser tolerado")
    End Sub

End Class

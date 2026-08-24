Imports Iris.Core
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' A regra que decide se uma falha pode ser REPETIDA.
'''
''' É a mais consequente do projeto e viveu até agora sem teste nenhum,
''' dentro de um arquivo de 852 linhas que só compila em Windows e só roda
''' com Outlook aberto. Saiu para o Core exatamente por isso.
'''
''' O que está em jogo: repetir um <c>Send</c> que já saiu manda a mesma
''' mensagem duas vezes, e não existe desfazer.
''' </summary>
<TestClass>
Public Class FailurePolicyTests

    ' ================================================================
    ' A regra que importa
    ' ================================================================

    ''' <summary>
    ''' Mutação cuja tentativa começou é AMBÍGUA, e ambiguidade vence
    ''' qualquer HRESULT — inclusive os que sozinhos seriam retentáveis.
    '''
    ''' É o buraco que já existiu neste projeto: um <c>Send</c> estourando
    ''' com <c>RPC_E_DISCONNECTED</c> depois de a mensagem sair virava
    ''' <c>NotConnected</c>, cujo <c>IsRetryable</c> é True. O código
    ''' convidava a reenviar no único caso em que reenviar duplica.
    ''' </summary>
    <TestMethod>
    Public Sub Mutacao_iniciada_e_ambigua_para_qualquer_HRESULT()
        For Each hr In TodosOsHResults()
            Dim kind = OutlookFailurePolicy.ClassifyFailure(
                hr, isMutation:=True, mutationAttemptStarted:=True)

            Assert.AreEqual(ErrorKind.Ambiguous, kind,
                $"HRESULT 0x{hr:X8} deveria virar Ambiguous numa mutação já iniciada.")
        Next

        ' Sem HRESULT nenhum — exceção que não é COMException — também.
        Assert.AreEqual(ErrorKind.Ambiguous,
            OutlookFailurePolicy.ClassifyFailure(Nothing, isMutation:=True,
                                                 mutationAttemptStarted:=True))
    End Sub

    ''' <summary>
    ''' Controle negativo do teste acima, e ele é essencial: sem isto, uma
    ''' política que devolvesse <c>Ambiguous</c> para TUDO passaria — e
    ''' travaria o Iris inteiro, porque ambíguo nunca é retentável.
    ''' </summary>
    <TestMethod>
    Public Sub Leitura_com_o_mesmo_HRESULT_nao_e_ambigua()
        For Each hr In TodosOsHResults()
            Dim kind = OutlookFailurePolicy.ClassifyFailure(
                hr, isMutation:=False, mutationAttemptStarted:=True)

            Assert.AreNotEqual(ErrorKind.Ambiguous, kind,
                $"HRESULT 0x{hr:X8} numa LEITURA não pode virar Ambiguous.")
        Next
    End Sub

    ''' <summary>
    ''' Segundo controle negativo: mutação que NÃO chegou a começar também
    ''' não é ambígua. Se fosse, toda falha de conexão anterior à chamada
    ''' bloquearia o rascunho sem motivo.
    ''' </summary>
    <TestMethod>
    Public Sub Mutacao_que_nao_comecou_nao_e_ambigua()
        Assert.AreEqual(ErrorKind.NotConnected,
            OutlookFailurePolicy.ClassifyFailure(
                OutlookFailurePolicy.RPC_E_DISCONNECTED,
                isMutation:=True, mutationAttemptStarted:=False))

        Assert.AreEqual(ErrorKind.Busy,
            OutlookFailurePolicy.ClassifyFailure(
                OutlookFailurePolicy.RPC_E_SERVERCALL_RETRYLATER,
                isMutation:=True, mutationAttemptStarted:=False))
    End Sub

    ''' <summary>
    ''' O que a ambiguidade significa na prática: não é retentável. Sem
    ''' isto, o resto da regra não protege ninguém.
    ''' </summary>
    <TestMethod>
    Public Sub Ambiguo_nunca_e_retentavel()
        Dim ambiguo = OperationResult(Of Boolean).Fail(ErrorKind.Ambiguous, "teste")
        Assert.IsTrue(ambiguo.IsAmbiguous)
        Assert.IsFalse(ambiguo.IsRetryable)

        ' E o contraste que dá sentido à afirmação.
        Assert.IsTrue(OperationResult(Of Boolean).Fail(ErrorKind.NotConnected, "").IsRetryable)
        Assert.IsTrue(OperationResult(Of Boolean).Fail(ErrorKind.Busy, "").IsRetryable)
    End Sub

    ' ================================================================
    ' Tradução de cada HRESULT
    ' ================================================================

    <TestMethod>
    Public Sub Ocupado_e_ocupado()
        Assert.AreEqual(ErrorKind.Busy, Ler(OutlookFailurePolicy.RPC_E_CALL_REJECTED))
        Assert.AreEqual(ErrorKind.Busy, Ler(OutlookFailurePolicy.RPC_E_SERVERCALL_RETRYLATER))
    End Sub

    <TestMethod>
    Public Sub Sessao_morta_e_NotConnected()
        Assert.AreEqual(ErrorKind.NotConnected, Ler(OutlookFailurePolicy.RPC_E_DISCONNECTED))
        Assert.AreEqual(ErrorKind.NotConnected, Ler(OutlookFailurePolicy.RPC_S_SERVER_UNAVAILABLE))
        Assert.AreEqual(ErrorKind.NotConnected, Ler(OutlookFailurePolicy.CO_E_OBJNOTCONNECTED))

        ' Acrescentados no 1.6: o Outlook pode morrer por estes também, e
        ' antes eles caíam em Unexpected.
        Assert.AreEqual(ErrorKind.NotConnected, Ler(OutlookFailurePolicy.RPC_E_SERVER_DIED))
        Assert.AreEqual(ErrorKind.NotConnected, Ler(OutlookFailurePolicy.RPC_E_SERVER_DIED_DNE))
        Assert.AreEqual(ErrorKind.NotConnected, Ler(OutlookFailurePolicy.RPC_S_CALL_FAILED))
        Assert.AreEqual(ErrorKind.NotConnected, Ler(OutlookFailurePolicy.RPC_E_INVALID_OBJECT))
    End Sub

    <TestMethod>
    Public Sub Acesso_negado_e_Denied()
        Assert.AreEqual(ErrorKind.Denied, Ler(OutlookFailurePolicy.E_ACCESSDENIED))
    End Sub

    ''' <summary>
    ''' HRESULT que ninguém previu não pode ser chutado para um lado
    ''' cômodo. <c>Unexpected</c> não é retentável e não é ambíguo: fica
    ''' visível, que é o que se quer de algo desconhecido.
    ''' </summary>
    <TestMethod>
    Public Sub Desconhecido_e_Unexpected()
        Assert.AreEqual(ErrorKind.Unexpected, Ler(&H12345678))
        Assert.AreEqual(ErrorKind.Unexpected, Ler(0))

        Assert.AreEqual(ErrorKind.Unexpected,
            OutlookFailurePolicy.ClassifyFailure(Nothing, isMutation:=False,
                                                 mutationAttemptStarted:=False),
            "Exceção que não é COMException não tem HRESULT para classificar.")
    End Sub

    ' ================================================================
    ' Morte de sessão e reanexação
    ' ================================================================

    <TestMethod>
    Public Sub Reconhece_os_HRESULTs_de_sessao_morta()
        Assert.IsTrue(OutlookFailurePolicy.IsSessionDead(OutlookFailurePolicy.RPC_E_DISCONNECTED))
        Assert.IsTrue(OutlookFailurePolicy.IsSessionDead(OutlookFailurePolicy.RPC_E_SERVER_DIED))
        Assert.IsTrue(OutlookFailurePolicy.IsSessionDead(OutlookFailurePolicy.RPC_S_CALL_FAILED))
    End Sub

    ''' <summary>
    ''' Ocupado NÃO é morte. Derrubar a sessão numa recusa transitória
    ''' trocaria uma espera de segundos por uma reconexão inteira — e o
    ''' Outlook recusa chamada o tempo todo quando está sincronizando.
    ''' </summary>
    <TestMethod>
    Public Sub Ocupado_nao_e_sessao_morta()
        Assert.IsFalse(OutlookFailurePolicy.IsSessionDead(OutlookFailurePolicy.RPC_E_CALL_REJECTED))
        Assert.IsFalse(OutlookFailurePolicy.IsSessionDead(OutlookFailurePolicy.RPC_E_SERVERCALL_RETRYLATER))
        Assert.IsTrue(OutlookFailurePolicy.IsBusy(OutlookFailurePolicy.RPC_E_CALL_REJECTED))
    End Sub

    ''' <summary>
    ''' <c>MAPI_E_NETWORK_ERROR</c> não é morte de sessão: pode ser Outlook
    ''' vivo com o Exchange fora do ar. Derrubar a sessão aí trocaria um
    ''' problema de rede por uma reconexão que não resolve nada.
    ''' </summary>
    <TestMethod>
    Public Sub Erro_de_rede_nao_derruba_a_sessao()
        Const MAPI_E_NETWORK_ERROR As Integer = &H80040115
        Assert.IsFalse(OutlookFailurePolicy.IsSessionDead(MAPI_E_NETWORK_ERROR))
    End Sub

    ''' <summary>
    ''' Falha desconhecida preserva a sessão na primeira, mas não para
    ''' sempre.
    '''
    ''' O comportamento anterior era devolver "ocupado" e preservar o RCW
    ''' indefinidamente. Se o RCW estivesse morto com um código não
    ''' previsto, o Iris ficava ocupado eternamente: sem erro, sem
    ''' reconexão e sem sinal nenhum.
    ''' </summary>
    <TestMethod>
    Public Sub Falha_desconhecida_preserva_a_sessao_no_comeco_e_desiste_depois()
        Assert.IsFalse(OutlookFailurePolicy.ShouldReattachAfterUnknown(1),
            "Recusa transitória não é morte.")
        Assert.IsFalse(OutlookFailurePolicy.ShouldReattachAfterUnknown(2))

        Assert.IsTrue(OutlookFailurePolicy.ShouldReattachAfterUnknown(
            OutlookFailurePolicy.ProbesAteReanexar),
            "Insistir para sempre deixa o Iris ocupado eternamente.")
        Assert.IsTrue(OutlookFailurePolicy.ShouldReattachAfterUnknown(99))
    End Sub

    ' ================================================================

    Private Shared Function Ler(hresult As Integer) As ErrorKind
        Return OutlookFailurePolicy.ClassifyFailure(hresult, isMutation:=False,
                                                    mutationAttemptStarted:=False)
    End Function

    ''' <summary>
    ''' Todos os HRESULTs que a política conhece, mais um desconhecido. A
    ''' regra da ambiguidade tem de valer para o conjunto inteiro — testar
    ''' com um só deixaria passar uma implementação que só cobrisse aquele.
    ''' </summary>
    Private Shared Function TodosOsHResults() As Integer()
        Return New Integer() {
            OutlookFailurePolicy.RPC_E_CALL_REJECTED,
            OutlookFailurePolicy.RPC_E_SERVERCALL_RETRYLATER,
            OutlookFailurePolicy.RPC_E_DISCONNECTED,
            OutlookFailurePolicy.RPC_S_SERVER_UNAVAILABLE,
            OutlookFailurePolicy.CO_E_OBJNOTCONNECTED,
            OutlookFailurePolicy.RPC_E_SERVER_DIED,
            OutlookFailurePolicy.RPC_E_SERVER_DIED_DNE,
            OutlookFailurePolicy.RPC_S_CALL_FAILED,
            OutlookFailurePolicy.RPC_E_INVALID_OBJECT,
            OutlookFailurePolicy.E_ACCESSDENIED,
            &H12345678
        }
    End Function

    ''' <summary>
    ''' Os HRESULTs que o ANEXAR encontra, e que antes tinham uma segunda
    ''' tabela em ComInterop.GetRunningInstance discordando desta.
    '''
    ''' RPC_E_DISCONNECTED era Busy la e sessao MORTA aqui. Chamar de
    ''' "ocupado" uma sessao morta faz a UI prometer reconexao automatica
    ''' que nao vem. Agora o ComInterop DELEGA para ca, entao este teste
    ''' protege os dois.
    ''' </summary>
    <TestMethod>
    Public Sub HRESULTs_do_anexar_sao_classificados_de_um_jeito_so()
        ''' Ocupado de verdade: da para tentar de novo.
        Assert.AreEqual(ErrorKind.Busy,
            OutlookFailurePolicy.ClassifyFailure(OutlookFailurePolicy.RPC_E_CALL_REJECTED, False, False))
        Assert.AreEqual(ErrorKind.Busy,
            OutlookFailurePolicy.ClassifyFailure(OutlookFailurePolicy.RPC_E_SERVERCALL_RETRYLATER, False, False))

        ''' Sessao morta. NAO e ocupado.
        Assert.AreEqual(ErrorKind.NotConnected,
            OutlookFailurePolicy.ClassifyFailure(OutlookFailurePolicy.RPC_E_DISCONNECTED, False, False),
            "RPC_E_DISCONNECTED e sessao morta, nao ocupada")
    End Sub

    ''' <summary>
    ''' O HRESULT que a FASE2 secao 21.1 mediu de verdade: fechei o Outlook
    ''' com uma Table aberta e o GetArray seguinte lancou 0x800706BA.
    '''
    ''' Foi a primeira validacao deste classificador contra um caso real -
    ''' ate entao ele so tinha teste sintetico. Este teste existe para que
    ''' o caso medido nao possa regredir em silencio.
    ''' </summary>
    <TestMethod>
    Public Sub Morte_do_Outlook_no_meio_da_varredura_e_NotConnected()
        Const MEDIDO_EM_21_1 As Integer = &H800706BA
        Assert.AreEqual(OutlookFailurePolicy.RPC_S_SERVER_UNAVAILABLE, MEDIDO_EM_21_1)
        Assert.AreEqual(ErrorKind.NotConnected,
            OutlookFailurePolicy.ClassifyFailure(MEDIDO_EM_21_1, False, False))
        Assert.IsTrue(OutlookFailurePolicy.IsSessionDead(MEDIDO_EM_21_1),
            "o watchdog precisa reconhecer isto como sessao morta e reanexar")

        ''' E continua Ambiguous se uma MUTACAO ja tinha comecado: a regra de
        ''' que mutacao vence qualquer HRESULT nao pode ser furada por este.
        Assert.AreEqual(ErrorKind.Ambiguous,
            OutlookFailurePolicy.ClassifyFailure(MEDIDO_EM_21_1, isMutation:=True, mutationAttemptStarted:=True))
    End Sub
End Class

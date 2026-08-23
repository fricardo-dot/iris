Imports Iris.Model

Namespace Global.Iris.Core

    ''' <summary>
    ''' Traduz falha da borda COM em <see cref="ErrorKind"/>.
    '''
    ''' O nome é específico de propósito. <see cref="AddressPolicy"/> é
    ''' política de DOMÍNIO, usada por duas camadas por decisão de produto.
    ''' Esta aqui é tradução da borda COM: mora no Core por testabilidade, e
    ''' chamá-la de "política" no mesmo tom seria arrumação estética.
    '''
    ''' Vive fora do broker porque dentro dele nenhum teste a alcançava — são
    ''' 852 linhas que só compilam em Windows e só rodam com Outlook aberto —
    ''' e porque a regra que ela carrega é a mais consequente do projeto:
    ''' decide se uma falha pode ser repetida.
    '''
    ''' Função pura. O broker continua dono de observar a fase da operação e
    ''' de escrever o log; aqui só entra fato, e sai decisão.
    ''' </summary>
    Public NotInheritable Class OutlookFailurePolicy

        Private Sub New()
        End Sub

        ' HRESULTs, com nome. Em hexadecimal solto ninguém confere nada.
        Public Const RPC_E_CALL_REJECTED As Integer = &H80010001
        Public Const RPC_E_SERVERCALL_RETRYLATER As Integer = &H8001010A
        Public Const RPC_E_DISCONNECTED As Integer = &H80010108
        Public Const RPC_S_SERVER_UNAVAILABLE As Integer = &H800706BA
        Public Const CO_E_OBJNOTCONNECTED As Integer = &H800401FD
        Public Const E_ACCESSDENIED As Integer = &H80070005

        ''' <summary>
        ''' Servidor COM morreu. Diferente de "ocupado": aqui o RCW guardado
        ''' não serve mais, e insistir nele não resolve nunca.
        ''' </summary>
        Public Const RPC_E_SERVER_DIED As Integer = &H80010007
        Public Const RPC_E_SERVER_DIED_DNE As Integer = &H80010012
        Public Const RPC_S_CALL_FAILED As Integer = &H800706BE
        Public Const RPC_E_INVALID_OBJECT As Integer = &H80010114

        ''' <summary>
        ''' Classifica uma falha.
        ''' </summary>
        ''' <param name="mutationAttemptStarted">
        ''' O delegate da operação COMEÇOU a rodar. Note o nome: não é "o
        ''' efeito aconteceu", nem "a chamada COM mutante começou" — é só que
        ''' a tentativa entrou. É conservador de propósito, e o nome não
        ''' promete precisão que o código não tem.
        '''
        ''' Este fato é fornecido pelo ORQUESTRADOR, e precisa ser local à
        ''' invocação. Enquanto foi um campo do broker, uma operação
        ''' concorrente zerava o campo entre a falha de um <c>Send</c> e a
        ''' classificação dela — e o envio que talvez tenha saído era
        ''' devolvido como retentável.
        ''' </param>
        Public Shared Function ClassifyFailure(hresult As Integer?,
                                               isMutation As Boolean,
                                               mutationAttemptStarted As Boolean) As ErrorKind

            ' Mutação que começou vence QUALQUER HRESULT, inclusive os que
            ' sozinhos seriam NotConnected ou Busy — que são justamente os
            ' retentáveis. Era exatamente esse o buraco: um Send estourando
            ' com RPC_E_DISCONNECTED depois de a mensagem sair virava
            ' NotConnected, cujo IsRetryable é True, e o código convidava a
            ' reenviar no único caso em que reenviar duplica.
            If isMutation AndAlso mutationAttemptStarted Then Return ErrorKind.Ambiguous

            If Not hresult.HasValue Then Return ErrorKind.Unexpected

            Select Case hresult.Value
                Case RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER
                    Return ErrorKind.Busy

                Case RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, CO_E_OBJNOTCONNECTED,
                     RPC_E_SERVER_DIED, RPC_E_SERVER_DIED_DNE, RPC_S_CALL_FAILED,
                     RPC_E_INVALID_OBJECT
                    Return ErrorKind.NotConnected

                Case E_ACCESSDENIED
                    Return ErrorKind.Denied

                Case Else
                    Return ErrorKind.Unexpected
            End Select
        End Function

        ''' <summary>
        ''' Este HRESULT quer dizer que a sessão COM morreu e o RCW guardado
        ''' não serve mais?
        '''
        ''' Lista nunca é prova completa — sempre vai existir um código que
        ''' ninguém previu. Por isso a robustez de verdade não está aqui, e
        ''' sim no limiar de <see cref="ShouldReattachAfterUnknown"/>: o que
        ''' não é reconhecido não fica preso para sempre.
        '''
        ''' <c>MAPI_E_NETWORK_ERROR</c> NÃO entra: pode ser Outlook vivo com
        ''' Exchange indisponível, e derrubar a sessão nesse caso trocaria uma
        ''' falha temporária por uma reconexão desnecessária.
        ''' </summary>
        Public Shared Function IsSessionDead(hresult As Integer) As Boolean
            Select Case hresult
                Case RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, CO_E_OBJNOTCONNECTED,
                     RPC_E_SERVER_DIED, RPC_E_SERVER_DIED_DNE, RPC_S_CALL_FAILED,
                     RPC_E_INVALID_OBJECT
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Public Shared Function IsBusy(hresult As Integer) As Boolean
            Return hresult = RPC_E_CALL_REJECTED OrElse hresult = RPC_E_SERVERCALL_RETRYLATER
        End Function

        ''' <summary>
        ''' Quantos probes seguidos com falha desconhecida antes de largar o
        ''' RCW e tentar reanexar.
        ''' </summary>
        Public Const ProbesAteReanexar As Integer = 3

        ''' <summary>
        ''' Falha que não foi reconhecida: preservar a sessão ou largar e
        ''' reanexar?
        '''
        ''' O comportamento anterior era devolver "ocupado" e preservar o RCW,
        ''' para sempre. Se o RCW estivesse de fato morto com um código não
        ''' previsto, o Iris ficava ocupado eternamente — sem erro, sem
        ''' reconexão e sem sinal nenhum para o usuário.
        '''
        ''' Preservar na primeira é certo: recusa transitória não é morte.
        ''' Insistir indefinidamente é que não.
        ''' </summary>
        ''' <param name="probesSeguidosComFalha">
        ''' Quantos probes consecutivos falharam sem classificação, contando
        ''' este.
        ''' </param>
        Public Shared Function ShouldReattachAfterUnknown(probesSeguidosComFalha As Integer) As Boolean
            Return probesSeguidosComFalha >= ProbesAteReanexar
        End Function

    End Class

End Namespace

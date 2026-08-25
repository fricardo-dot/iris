Namespace Global.Iris.Model

    ''' <summary>
    ''' <b>O único lugar que decide o que o ambiente produz sozinho.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO É UM MÓDULO E NÃO TRÊS CÓPIAS</b>
    '''
    ''' A lista <c>Busy</c> + <c>NotConnected</c> apareceu escrita à mão em três
    ''' lugares: <c>OperationResult.IsRetryable</c>,
    ''' <c>SourceUnavailableException.DoAmbiente</c> e a asserção de um teste de
    ''' integração. Eu tinha posto um teste comparando os dois primeiros e
    ''' chamado de resolvido — o Codex mostrou o terceiro, que o teste não
    ''' cobria e que era justamente o que decidia se uma falha reprovava.
    '''
    ''' Três cópias de uma regra divergem em silêncio: alguém acrescenta um
    ''' <see cref="ErrorKind"/> transitório num lugar, esquece nos outros, e a
    ''' divergência só aparece como intermitência meses depois. Comparar cópias
    ''' com teste adia o problema; ter <b>uma</b> cópia o remove.
    ''' </summary>
    Public Module ErrorPolicy

        ''' <summary>
        ''' O ambiente pode ter produzido isto sozinho, e pode voltar ao normal
        ''' sem ninguém consertar nada.
        '''
        ''' <c>Ambiguous</c> fica de fora <b>por desenho</b>: não se sabe se o
        ''' efeito ocorreu, e repetir é o pior que se pode fazer.
        ''' <c>Cancelled</c> também: cancelamento é decisão de quem chamou, tem
        ''' desfecho próprio, e tratá-lo como soluço do ambiente esconderia um
        ''' cancelamento saindo classificado de falha.
        ''' </summary>
        Public Function Transitorio(kind As ErrorKind) As Boolean
            Return kind = ErrorKind.Busy OrElse kind = ErrorKind.NotConnected
        End Function

    End Module

End Namespace

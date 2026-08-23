Imports Iris.Model

Namespace Global.Iris.Core

    ''' <summary>
    ''' Quando o corpo de uma mensagem pode ser lido.
    '''
    ''' Vive aqui, e não dentro do <c>MessageReading</c>, pelo mesmo motivo
    ''' que o <see cref="OutlookFailurePolicy"/>: regra consequente enterrada
    ''' num módulo que só roda com COM é regra que nenhum teste alcança.
    '''
    ''' A regra tem UMA linha e um único jeito errado de escrevê-la.
    ''' Escrever <c>state = Restricted</c> para bloquear deixa
    ''' <c>Unknown</c> passar — e <c>Unknown</c> é exatamente o caso em que
    ''' a leitura de <c>Permission</c> FALHOU. Um gate que falha aberto é
    ''' pior que não ter gate: ele dá a impressão de que existe proteção.
    ''' </summary>
    Public NotInheritable Class ProtectionPolicy

        ''' <summary>
        ''' Só <c>Unprotected</c> libera. Qualquer outra coisa — inclusive
        ''' um valor futuro que ninguém previu aqui — bloqueia.
        ''' </summary>
        Public Shared Function CanReadBody(state As ProtectionState) As Boolean
            Return state = ProtectionState.Unprotected
        End Function

        ''' <summary>
        ''' Vale para log e para IA também, não só para a tela. R11: conteúdo
        ''' protegido não sai do processo.
        ''' </summary>
        Public Shared Function CanSendToAi(state As ProtectionState) As Boolean
            Return CanReadBody(state)
        End Function

        Public Shared Function DescribeBlock(state As ProtectionState) As String
            Select Case state
                Case ProtectionState.Restricted
                    Return "Mensagem protegida: o corpo não é entregue pelo Outlook."
                Case ProtectionState.Unknown
                    Return "Não foi possível confirmar se a mensagem é protegida. " &
                           "O corpo não será lido."
                Case Else
                    Return ""
            End Select
        End Function

    End Class

End Namespace

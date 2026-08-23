Namespace Global.Iris.Model

    ''' <summary>
    ''' Quão completa é UMA parte de uma mensagem.
    '''
    ''' Existe porque "lista vazia" era ambíguo: não dava para distinguir
    ''' "esta mensagem não tem destinatários" de "não deu para ler os
    ''' destinatários". As duas apareciam iguais na tela, e a segunda é a que
    ''' torna perigoso responder.
    '''
    ''' Por COMPONENTE, e não um estado só para a mensagem inteira. Corpo,
    ''' destinatários e anexos falham de forma independente — cada um é uma
    ''' chamada COM diferente, e uma pode ser negada por política enquanto as
    ''' outras vêm inteiras. Um estado global obrigaria a escolher entre
    ''' rebaixar a mensagem toda por causa de uma parte, ou esconder que
    ''' aquela parte falhou.
    ''' </summary>
    Public Enum PartState
        ''' <summary>Lido por inteiro. Lista vazia aqui significa vazia mesmo.</summary>
        Complete

        ''' <summary>
        ''' Veio uma parte. O que está na tela é verdade, mas está
        ''' incompleto — e é o pior caso para quem confia no que vê.
        '''
        ''' NÃO se chama "Partial": é palavra reservada do VB, e como membro
        ''' de enum quebra a análise do arquivo inteiro com erros que apontam
        ''' para linhas sem relação nenhuma.
        ''' </summary>
        Incomplete

        ''' <summary>Não deu para ler nada desta parte.</summary>
        Unavailable
    End Enum

    ''' <summary>
    ''' O estado de uma parte, com o motivo quando ela não veio inteira.
    ''' </summary>
    Public NotInheritable Class PartStatus

        Public ReadOnly Property State As PartState
        Public ReadOnly Property Reason As ErrorKind

        ''' <summary>
        ''' Quantos itens ERAM esperados, quando dá para saber. Ler
        ''' <c>Recipients.Count</c> costuma funcionar mesmo quando ler um dos
        ''' destinatários falha, e a diferença entre esperado e obtido é o que
        ''' permite dizer "3 de 5" em vez de "alguns".
        ''' </summary>
        Public ReadOnly Property Expected As Integer

        Public ReadOnly Property Obtained As Integer

        Private Sub New(state As PartState, reason As ErrorKind,
                        expected As Integer, obtained As Integer)
            Me.State = state
            Me.Reason = reason
            Me.Expected = expected
            Me.Obtained = obtained
        End Sub

        Public Shared ReadOnly Property Full As PartStatus
            Get
                Return New PartStatus(PartState.Complete, ErrorKind.None, 0, 0)
            End Get
        End Property

        Public Shared Function CompleteWith(count As Integer) As PartStatus
            Return New PartStatus(PartState.Complete, ErrorKind.None, count, count)
        End Function

        Public Shared Function IncompleteWith(expected As Integer, obtained As Integer,
                                           reason As ErrorKind) As PartStatus
            Return New PartStatus(PartState.Incomplete, reason, expected, obtained)
        End Function

        Public Shared Function Missing(reason As ErrorKind) As PartStatus
            Return New PartStatus(PartState.Unavailable, reason, 0, 0)
        End Function

        ''' <summary>
        ''' Dá para confiar nesta parte para tomar uma decisão irreversível?
        '''
        ''' Só <c>Complete</c> serve. <c>Incomplete</c> é justamente o caso
        ''' traiçoeiro: há conteúdo na tela, ele está certo, e mesmo assim
        ''' agir sobre ele significa agir sobre menos do que existe.
        ''' </summary>
        Public ReadOnly Property IsTrustworthy As Boolean
            Get
                Return State = PartState.Complete
            End Get
        End Property

        Public Overrides Function ToString() As String
            If State = PartState.Complete Then Return "completo"
            If State = PartState.Unavailable Then Return $"indisponivel ({Reason})"
            Return $"parcial {Obtained}/{Expected} ({Reason})"
        End Function

    End Class

End Namespace

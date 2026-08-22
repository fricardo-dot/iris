Namespace Global.Iris.Model

    ''' <summary>
    ''' Resultado de uma operação do broker.
    '''
    ''' Existe em vez de exceções porque quase toda falha aqui é ESPERADA e
    ''' informativa: o Outlook está fechado, o item foi movido por baixo, a
    ''' página é de uma geração vencida. Isso é fluxo normal, não excepcional,
    ''' e a UI precisa reagir diferente para cada caso sem ler texto de
    ''' exceção nem HRESULT.
    '''
    ''' <see cref="ErrorKind.Ambiguous"/> merece cuidado especial: significa
    ''' que a operação pode ter surtido efeito. Nunca repetir automaticamente.
    ''' </summary>
    Public NotInheritable Class OperationResult(Of T)

        Public ReadOnly Property Value As T
        Public ReadOnly Property Kind As ErrorKind

        ''' <summary>
        ''' DIAGNOSTICO, nao mensagem de usuario. Nao e localizado, e nunca
        ''' contem corpo, assunto, endereco ou caminho. A UI deve traduzir o
        ''' <see cref="Kind"/>; exibir Detail direto na tela vaza detalhe
        ''' tecnico e, pior, pode vazar dado da caixa.
        ''' </summary>
        Public ReadOnly Property Detail As String

        Private Sub New(value As T, kind As ErrorKind, detail As String)
            Me.Value = value
            Me.Kind = kind
            Me.Detail = If(detail, String.Empty)
        End Sub

        Public ReadOnly Property Succeeded As Boolean
            Get
                Return Kind = ErrorKind.None
            End Get
        End Property

        ''' <summary>
        ''' A operação pode ter surtido efeito. Repetir é proibido.
        ''' </summary>
        Public ReadOnly Property IsAmbiguous As Boolean
            Get
                Return Kind = ErrorKind.Ambiguous
            End Get
        End Property

        ''' <summary>
        ''' Faz sentido tentar de novo? Falso para Ambiguous, por desenho.
        ''' </summary>
        Public ReadOnly Property IsRetryable As Boolean
            Get
                Return Kind = ErrorKind.Busy OrElse Kind = ErrorKind.NotConnected
            End Get
        End Property

        Public Shared Function Ok(value As T) As OperationResult(Of T)
            Return New OperationResult(Of T)(value, ErrorKind.None, String.Empty)
        End Function

        Public Shared Function Fail(kind As ErrorKind, detail As String) As OperationResult(Of T)
            If kind = ErrorKind.None Then
                Throw New ArgumentException("Falha precisa de um ErrorKind.", NameOf(kind))
            End If
            Return New OperationResult(Of T)(Nothing, kind, detail)
        End Function

        Public Overrides Function ToString() As String
            Return If(Succeeded, "ok", $"{Kind}: {Detail}")
        End Function
    End Class

End Namespace

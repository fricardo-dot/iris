Namespace Global.Iris.Model

    ''' <summary>
    ''' <b>Tem anexo?</b> — a resposta com o "não sei" preservado.
    '''
    ''' O portão nega mensagem com anexo, e para negar ele precisa saber. A
    ''' classificação lia isto como <c>False</c> fixo, o que não era uma
    ''' suposição conservadora: era uma <b>afirmação falsa</b> de que não havia
    ''' anexo, feita no lugar exato onde o portão decide.
    '''
    ''' <c>Tem = Nothing</c> quer dizer que a contagem não foi possível — guarda
    ''' do Object Model, item de classe inesperada, erro de COM. Quem decide
    ''' trata <c>Nothing</c> como <b>tem</b>: "não consegui contar" nunca vira
    ''' prova de ausência.
    ''' </summary>
    Public NotInheritable Class AttachmentPresence

        Public ReadOnly Property Item As ItemKey
        ''' <summary><c>Nothing</c> = não deu para saber.</summary>
        Public ReadOnly Property Tem As Boolean?

        Public Sub New(item As ItemKey, tem As Boolean?)
            Me.Item = item
            Me.Tem = tem
        End Sub

    End Class

End Namespace

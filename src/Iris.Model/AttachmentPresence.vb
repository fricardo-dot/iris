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
        ''' <summary>
        ''' Tem anexo <b>de verdade</b> — arquivo que alguém mandou.
        ''' <c>Nothing</c> = não deu para saber, e o portão fecha.
        ''' </summary>
        Public ReadOnly Property Tem As Boolean?

        ''' <summary>
        ''' Quantas imagens <b>embutidas</b> — as que o corpo referencia por
        ''' <c>cid:</c>, tipicamente logo de assinatura. Não bloqueiam; são
        ''' declaradas no resumo, porque uma captura de tela colada no corpo é
        ''' embutida do mesmo jeito.
        ''' </summary>
        Public ReadOnly Property Embutidas As Integer?

        Public Sub New(item As ItemKey, tem As Boolean?,
                       Optional embutidas As Integer? = Nothing)
            Me.Item = item
            Me.Tem = tem
            Me.Embutidas = embutidas
        End Sub

    End Class

End Namespace

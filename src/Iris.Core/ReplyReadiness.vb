Imports Iris.Model

Namespace Global.Iris.Core

    ''' <summary>
    ''' O que uma leitura incompleta impede.
    '''
    ''' A regra não é visual. Sinalizar leitura parcial com um ícone e deixar
    ''' tudo funcionando seria o pior dos mundos: o usuário vê um aviso que
    ''' não sabe interpretar e age assim mesmo.
    '''
    ''' A pergunta que importa é: esta parte é INSUMO de alguma ação
    ''' irreversível? Se for, incompleta bloqueia.
    ''' </summary>
    Public NotInheritable Class ReplyReadiness

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Responder e responder a todos dependem dos destinatários lidos.
        '''
        ''' Com a lista incompleta, "Responder a todos" responde a MENOS
        ''' gente do que a mensagem tinha — e ninguém percebe, porque o
        ''' resultado parece normal: uma resposta foi enviada, para pessoas
        ''' reais. O que falta é invisível por definição.
        '''
        ''' Bloqueia também quando os destinatários estão indisponíveis: aí
        ''' o Reply do Outlook montaria a resposta a partir de dado que o
        ''' Iris não conseguiu conferir.
        ''' </summary>
        Public Shared Function CanReply(recipients As PartStatus) As Boolean
            If recipients Is Nothing Then Return False
            Return recipients.IsTrustworthy
        End Function

        ''' <summary>
        ''' Encaminhar leva os ANEXOS junto. Com a lista incompleta, o
        ''' usuário não consegue conferir o que está mandando para fora — e
        ''' anexo é exatamente o que costuma não poder sair.
        '''
        ''' Note que o corpo NÃO entra aqui. Corpo incompleto é visível: o
        ''' texto aparece truncado e a pessoa percebe. Anexo que não foi
        ''' lido não deixa rastro nenhum na tela.
        ''' </summary>
        Public Shared Function CanForward(attachments As PartStatus) As Boolean
            If attachments Is Nothing Then Return False
            Return attachments.IsTrustworthy
        End Function

        ''' <summary>
        ''' Corpo incompleto NÃO bloqueia responder.
        '''
        ''' É uma escolha, e vale registrar o porquê: o corpo aparece na
        ''' tela, então uma leitura truncada é perceptível, e responder a uma
        ''' mensagem cujo texto você leu pela metade é uma decisão que cabe a
        ''' quem responde. Já a lista de destinatários é insumo que o usuário
        ''' não confere linha a linha antes de clicar.
        ''' </summary>
        Public Shared Function BodyBlocksReply(body As PartStatus) As Boolean
            Return False
        End Function

        ''' <summary>
        ''' Explica ao usuário o que está faltando e por que isso importa.
        ''' Aviso sem consequência clara é ruído; consequência sem
        ''' explicação é frustração.
        ''' </summary>
        Public Shared Function DescribeBlock(part As String, status As PartStatus) As String
            If status Is Nothing OrElse status.IsTrustworthy Then Return ""

            If status.State = PartState.Unavailable Then
                Return $"Não foi possível ler {part} desta mensagem."
            End If

            Return $"Só foi possível ler {status.Obtained} de {status.Expected} {part}."
        End Function

    End Class

End Namespace

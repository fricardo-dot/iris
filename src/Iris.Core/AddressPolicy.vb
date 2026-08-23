Imports System.Collections.Generic
Imports Iris.Model

Namespace Global.Iris.Core

    ''' <summary>
    ''' Que endereço serve para o usuário CONFERIR antes de um envio.
    '''
    ''' Vive no Core, e não escondida na camada COM, por dois motivos.
    '''
    ''' Primeiro, é testável: era um método privado dentro de
    ''' <c>DraftWriting</c>, e uma regra que decide se uma mensagem pode sair
    ''' não pode ser a única do projeto que nenhum teste alcança.
    '''
    ''' Segundo, vale nas duas pontas. Quem lê do Outlook usa para decidir se
    ''' marca o destinatário como resolvido; o compositor usa de novo, antes
    ''' de deixar enviar. Duplicar a checagem é de propósito: para a única
    ''' operação sem desfazer, a garantia não deve depender de uma camada só
    ''' ter feito o trabalho direito.
    ''' </summary>
    Public NotInheritable Class AddressPolicy

        Private Sub New()
        End Sub

        ''' <summary>
        ''' O endereço é SMTP reconhecível por uma pessoa?
        '''
        ''' O Outlook resolve nomes internos para
        ''' <c>/O=EMPRESA/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=...</c>. Ele
        ''' considera isso resolvido, e do ponto de vista dele está. Do ponto
        ''' de vista de quem lê a tela de confirmação, não: ninguém reconhece
        ''' o próprio colega nessa string, e conferir é a única função
        ''' daquela tela.
        ''' </summary>
        Public Shared Function IsUsableSmtp(address As String) As Boolean
            If String.IsNullOrWhiteSpace(address) Then Return False

            Dim limpo = address.Trim()

            ' Endereço Exchange legado, no formato X.500.
            If limpo.StartsWith("/", StringComparison.Ordinal) Then Return False

            ' Espaço no meio não é endereço. Sem isto, "Fulano de Tal
            ' @empresa.com" passava — tinha arroba e domínio com ponto.
            For Each c In limpo
                If Char.IsWhiteSpace(c) Then Return False
            Next

            ' Precisa de arroba, com algo antes e algo depois.
            Dim arroba = limpo.IndexOf("@"c)
            If arroba <= 0 OrElse arroba = limpo.Length - 1 Then Return False

            ' Uma só. "a@b@c" não é endereço.
            If limpo.IndexOf("@"c, arroba + 1) >= 0 Then Return False

            ' Domínio com ponto, com algo antes e algo depois dele.
            '
            ' Isto REPROVA domínio de rótulo único, como "fulano@intranet",
            ' que existe em rede interna e é endereço válido. É deliberado:
            ' a regra não decide se o e-mail funciona, decide se o usuário
            ' consegue CONFERIR para onde vai. Errar para o lado de bloquear
            ' custa um envio a mais pelo Outlook; errar para o outro custa
            ' uma mensagem na caixa de quem não devia. Registrado como
            ' dívida na FASE1, seção 11.
            Dim dominio = limpo.Substring(arroba + 1)
            If dominio.IndexOf("..", StringComparison.Ordinal) >= 0 Then Return False

            Dim ponto = dominio.IndexOf("."c)
            Return ponto > 0 AndAlso ponto < dominio.Length - 1
        End Function

        ''' <summary>
        ''' Todo destinatário está resolvido E com endereço conferível.
        ''' Lista vazia devolve False: não há para quem mandar.
        ''' </summary>
        Public Shared Function AllRecipientsUsable(recipients As IEnumerable(Of RecipientInfo)) As Boolean
            If recipients Is Nothing Then Return False

            Dim algum = False
            For Each r In recipients
                If r Is Nothing Then Return False
                algum = True
                If Not r.Resolved Then Return False
                If Not IsUsableSmtp(r.Address) Then Return False
            Next

            Return algum
        End Function

        ''' <summary>
        ''' Os que não passam, para a mensagem de erro dizer QUAIS.
        ''' "Há destinatários não reconhecidos" sem dizer quais obriga o
        ''' usuário a caçar no escuro.
        ''' </summary>
        Public Shared Function Unusable(recipients As IEnumerable(Of RecipientInfo)) As List(Of RecipientInfo)
            Dim ruins As New List(Of RecipientInfo)()
            If recipients Is Nothing Then Return ruins

            For Each r In recipients
                If r Is Nothing Then Continue For
                If Not r.Resolved OrElse Not IsUsableSmtp(r.Address) Then ruins.Add(r)
            Next

            Return ruins
        End Function

    End Class

End Namespace

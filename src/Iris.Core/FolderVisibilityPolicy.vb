Imports System.Linq
Imports Iris.Model

Namespace Global.Iris.Core

    ''' <summary>
    ''' Quais pastas o Iris mostra.
    '''
    ''' Vive no Core, não no ViewModel: é política de APLICAÇÃO, não detalhe
    ''' visual. Aqui ela é testável sem WPF e reutilizável se um dia houver
    ''' outra interface.
    '''
    ''' E não vive no broker: o broker relata o que o Outlook tem —
    ''' <c>IsHidden</c>, tipo de conteúdo, contagens — e quem decide o que
    ''' fazer com isso é esta camada. Filtrar lá embaixo enterraria uma
    ''' decisão de produto dentro do acesso a dados.
    ''' </summary>
    Public NotInheritable Class FolderVisibilityPolicy

        ''' <summary>
        ''' <b>Só mostra pasta que o Iris sabe abrir.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>ISTO ERA <c>MailOnly</c>, E O NOME VIROU MENTIRA EM 28/08/2026</b>
        '''
        ''' A regra sempre foi "não ofereça porta que não abre". Enquanto só
        ''' havia leitura de e-mail, "só correio" e "só o que abre" eram a mesma
        ''' coisa, e o nome mais curto ganhou.
        '''
        ''' A Fase 6 passou a ler calendário — e a pasta de calendário continuou
        ''' invisível na árvore. A leitura funcionava, os testes contra o Outlook
        ''' real passavam, e <b>o usuário não tinha como chegar nela</b>: os
        ''' testes achavam o calendário pelo broker, contornando a árvore.
        '''
        ''' Foi a revisão externa que pegou. É o erro mais comum deste projeto —
        ''' código que existe e não está ligado a nada — e desta vez ele estava
        ''' um nível acima do de costume: não era o método sem chamador, era a
        ''' funcionalidade inteira sem porta.
        '''
        ''' O nome agora descreve a regra, e a lista descreve o que o Iris sabe
        ''' fazer. Contatos, Tarefas e Observações continuam de fora <b>porque
        ''' ainda não há tela para elas</b>, e não por serem o que são.
        ''' </summary>
        Public Property SoOQueAbre As Boolean = True

        ''' <summary>
        ''' O que o Iris sabe abrir hoje. Cresce quando uma fase entrega a tela,
        ''' e não quando entrega a leitura — porque é a tela que faz a pasta
        ''' deixar de ser uma porta fechada.
        ''' </summary>
        Private Shared ReadOnly Abriveis As FolderContentKind() = {
            FolderContentKind.Mail,
            FolderContentKind.Calendar
        }

        ''' <summary>
        ''' Pastas com PR_ATTR_HIDDEN são internas do Outlook, como
        ''' "Conversation Action Settings" e "Quick Step Settings". O próprio
        ''' Outlook não as mostra. Fica configurável porque um modo de
        ''' diagnóstico pode querer vê-las.
        ''' </summary>
        Public Property IncludeHidden As Boolean = False

        Public Function IsVisible(folder As FolderInfo) As Boolean
            If folder Is Nothing Then Return False
            If folder.IsHidden AndAlso Not IncludeHidden Then Return False
            If SoOQueAbre AndAlso Not Abriveis.Contains(folder.ContentKind) Then Return False
            Return True
        End Function

    End Class

End Namespace

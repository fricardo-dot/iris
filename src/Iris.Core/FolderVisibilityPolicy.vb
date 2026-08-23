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
        ''' Fase 1 é só e-mail. Calendário, Contatos, Tarefas e Observações
        ''' existem no store, mas mostrá-los agora seria oferecer uma porta
        ''' que não abre — eles voltam nas Fases 5 a 7.
        ''' </summary>
        Public Property MailOnly As Boolean = True

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
            If MailOnly AndAlso folder.ContentKind <> FolderContentKind.Mail Then Return False
            Return True
        End Function

    End Class

End Namespace

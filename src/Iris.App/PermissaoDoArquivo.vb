Imports System.IO
Imports System.Linq
Imports System.Security.AccessControl
Imports System.Security.Principal

Namespace Global.Iris.App

    ''' <summary>
    ''' <b>Quem é dono do arquivo de ativação, e quem pode escrever nele.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO EXISTE</b>
    '''
    ''' O carregador dizia, no próprio comentário, que não conferia dono nem
    ''' permissão — e justificava com "o caminho fica sob <c>%LOCALAPPDATA%</c>,
    ''' que o Windows já protege por usuário".
    '''
    ''' <b>A justificativa era falsa nesta máquina.</b> O Codex rodou
    ''' <c>icacls</c> e achou controle total herdado por um SID de usuário não
    ''' resolvido e por um SID de <i>capability</i>. "Já é protegido" era uma
    ''' suposição sobre um sistema real, e suposição sobre sistema real se
    ''' confere ou se abandona.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>PELO HANDLE, E NÃO PELO CAMINHO</b>
    '''
    ''' A conferência recebe o <c>FileStream</c> já aberto. Conferir pelo
    ''' caminho e ler depois deixaria a janela clássica: o arquivo conferido e o
    ''' arquivo lido podem ser outros. Aqui os dois são o mesmo objeto do
    ''' sistema, porque é o mesmo handle.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE PASSA</b>
    '''
    ''' O dono tem de ser <b>você</b>. E ninguém além de você, do
    ''' <c>SYSTEM</c> e dos <c>Administradores</c> pode escrever — quem pode
    ''' escrever pode trocar a autorização, e trocar a autorização é escolher
    ''' para onde o seu e-mail vai.
    '''
    ''' Administrador entra na lista porque quem é administrador da máquina já
    ''' pode tudo, inclusive trocar o próprio Iris. Excluí-lo seria teatro.
    ''' </summary>
    Public NotInheritable Class PermissaoDoArquivo

        ''' <summary>
        ''' A conferência que o <c>ActivationLoader</c> recebe. <c>False</c> faz
        ''' a ativação ser recusada com <c>PermissaoRuim</c>.
        '''
        ''' Confere <b>duas</b> coisas: o arquivo, pelo handle já aberto, e o
        ''' <b>diretório</b> que o contém.
        ''' </summary>
        Public Shared Function SoMinha(fs As FileStream) As Boolean
            If fs Is Nothing Then Return False

            Try
                Dim meuSid = WindowsIdentity.GetCurrent().User
                If meuSid Is Nothing Then Return False

                ' O ARQUIVO, pelo HANDLE. Conferir por caminho e ler depois
                ' deixaria a janela em que os dois sao objetos diferentes.
                If Not Limpo(fs.GetAccessControl(), meuSid, DoArquivo) Then Return False

                ' E O DIRETORIO, que e o furo que a conferencia do arquivo NAO
                ' fecha.
                '
                ' Quem tem CreateFiles e DeleteSubdirectoriesAndFiles na pasta
                ' apaga o ativacao.json e poe outro no lugar, mesmo que a ACL do
                ' arquivo esteja impecavel. O handle protege a carga que ja foi
                ' lida; ele nao protege o NOME entre uma execucao e a seguinte.
                Dim pasta = Path.GetDirectoryName(fs.Name)
                If String.IsNullOrEmpty(pasta) Then Return False

                Dim di As New DirectoryInfo(pasta)
                If Not di.Exists Then Return False
                Return Limpo(di.GetAccessControl(), meuSid, DaPasta)

            Catch
                ' Falha fechada: nao consegui conferir e nao vou supor.
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Dono é o usuário, e ninguém a mais tem direito perigoso.
        ''' </summary>
        Private Shared Function Limpo(seguranca As FileSystemSecurity,
                                      meuSid As SecurityIdentifier,
                                      perigosos As FileSystemRights) As Boolean

            Dim dono = TryCast(seguranca.GetOwner(GetType(SecurityIdentifier)),
                               SecurityIdentifier)
            If dono Is Nothing OrElse Not dono.Equals(meuSid) Then Return False

            Dim permitidos As New HashSet(Of SecurityIdentifier) From {meuSid}
            permitidos.Add(New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing))
            permitidos.Add(New SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid,
                                                  Nothing))

            For Each regra As FileSystemAccessRule In
                seguranca.GetAccessRules(True, True, GetType(SecurityIdentifier))

                If regra.AccessControlType <> AccessControlType.Allow Then Continue For

                ' ACE so-de-heranca nao vale para o proprio objeto: ela e um
                ' molde para o que nascer dentro. CREATOR OWNER quase sempre
                ' aparece assim, e conta-la reprovaria pasta normal — e uma
                ' barreira que ninguem consegue satisfazer vira uma barreira
                ' desligada.
                If (regra.PropagationFlags And PropagationFlags.InheritOnly) <> 0 Then Continue For

                If (regra.FileSystemRights And perigosos) = 0 Then Continue For

                Dim quem = TryCast(regra.IdentityReference, SecurityIdentifier)
                If quem Is Nothing OrElse Not permitidos.Contains(quem) Then Return False
            Next

            Return True
        End Function

        ''' <summary>
        ''' Os direitos que mudam o <b>conteúdo</b> do arquivo.
        '''
        ''' <c>WriteAttributes</c> e <c>ReadPermissions</c> ficam de fora: tratar
        ''' tudo como escrita reprovaria ACL normal.
        ''' </summary>
        Private Shared ReadOnly DoArquivo As FileSystemRights =
            FileSystemRights.WriteData Or
            FileSystemRights.AppendData Or
            FileSystemRights.Delete Or
            FileSystemRights.ChangePermissions Or
            FileSystemRights.TakeOwnership

        ''' <summary>
        ''' Os direitos que permitem <b>trocar o arquivo</b> sem tocar nele.
        '''
        ''' <c>CreateFiles</c> tem o mesmo valor de <c>WriteData</c> e
        ''' <c>CreateDirectories</c> o mesmo de <c>AppendData</c> — num diretório
        ''' eles querem dizer "pode pôr coisa aqui dentro".
        ''' <c>DeleteSubdirectoriesAndFiles</c> é o que faltava: ele permite
        ''' apagar o <c>ativacao.json</c> <b>sem</b> ter direito nenhum sobre o
        ''' arquivo.
        ''' </summary>
        Private Shared ReadOnly DaPasta As FileSystemRights =
            FileSystemRights.CreateFiles Or
            FileSystemRights.CreateDirectories Or
            FileSystemRights.DeleteSubdirectoriesAndFiles Or
            FileSystemRights.Delete Or
            FileSystemRights.ChangePermissions Or
            FileSystemRights.TakeOwnership

    End Class

End Namespace

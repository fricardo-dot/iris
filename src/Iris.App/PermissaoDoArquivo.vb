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
        ''' </summary>
        Public Shared Function SoMinha(fs As FileStream) As Boolean
            If fs Is Nothing Then Return False

            Try
                Dim eu = WindowsIdentity.GetCurrent()
                Dim meuSid = eu.User
                If meuSid Is Nothing Then Return False

                Dim seguranca = fs.GetAccessControl()

                ' O DONO. Dono diferente quer dizer que o arquivo foi posto ali
                ' por outra conta — e uma autorizacao que outra conta escreveu
                ' nao e a sua autorizacao.
                Dim dono = TryCast(seguranca.GetOwner(GetType(SecurityIdentifier)),
                                   SecurityIdentifier)
                If dono Is Nothing OrElse Not dono.Equals(meuSid) Then Return False

                ' QUEM PODE ESCREVER. Nao basta o dono estar certo: uma ACL
                ' herdada pode dar escrita a mais gente, e foi exatamente isso
                ' que apareceu nesta maquina.
                Dim permitidos As New HashSet(Of SecurityIdentifier) From {meuSid}
                permitidos.Add(New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing))
                permitidos.Add(New SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid,
                                                      Nothing))

                Dim regras = seguranca.GetAccessRules(True, True, GetType(SecurityIdentifier))
                For Each regra As FileSystemAccessRule In regras
                    If regra.AccessControlType <> AccessControlType.Allow Then Continue For
                    If Not Escreve(regra.FileSystemRights) Then Continue For

                    Dim quem = TryCast(regra.IdentityReference, SecurityIdentifier)
                    If quem Is Nothing OrElse Not permitidos.Contains(quem) Then Return False
                Next

                Return True

            Catch
                ' Falha fechada: nao consegui conferir e nao vou supor.
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Este direito permite mudar o conteúdo?
        '''
        ''' <c>WriteData</c> e <c>AppendData</c> mudam o arquivo;
        ''' <c>WriteAttributes</c> e <c>ReadPermissions</c> não. Tratar tudo
        ''' como escrita reprovaria ACL normal e o usuário desligaria a
        ''' conferência — uma barreira que ninguém consegue satisfazer vira uma
        ''' barreira desligada.
        ''' </summary>
        Private Shared Function Escreve(d As FileSystemRights) As Boolean
            Dim perigosos = FileSystemRights.WriteData Or
                            FileSystemRights.AppendData Or
                            FileSystemRights.Delete Or
                            FileSystemRights.ChangePermissions Or
                            FileSystemRights.TakeOwnership
            Return (d And perigosos) <> 0
        End Function

    End Class

End Namespace

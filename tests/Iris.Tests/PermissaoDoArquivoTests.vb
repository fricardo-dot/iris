Imports System.Collections.Generic
Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports Iris.App
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace Global.Iris.Tests

    ''' <summary>
    ''' <b>Quem pode escrever no arquivo de ativação.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ESTA CONFERÊNCIA EXISTE</b>
    '''
    ''' O carregador dizia, no comentário, que não conferia permissão, e
    ''' justificava com "<c>%LOCALAPPDATA%</c> já é protegido por usuário".
    ''' <b>Era falso na máquina real</b>: o arquivo de ativação tinha controle
    ''' total herdado por um SID de <i>capability</i> e por um SID de outra
    ''' máquina, órfão de um perfil antigo.
    '''
    ''' A lição não é sobre ACL: é que "o sistema já cuida disso" é uma
    ''' suposição sobre um sistema real, e suposição sobre sistema real se
    ''' confere ou se abandona.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>OS TESTES CONSTROEM A ACL QUE MEDEM</b>
    '''
    ''' Nada aqui depende da ACL herdada do <c>%TEMP%</c> — que é justamente o
    ''' que costuma trazer as entradas a mais. Cada teste corta a herança e põe
    ''' exatamente as regras que quer medir.
    ''' </summary>
    <TestClass>
    Public Class PermissaoDoArquivoTests

        Private ReadOnly _arquivos As New List(Of String)()

        <TestCleanup>
        Public Sub Limpar()
            For Each f In _arquivos
                Try
                    If File.Exists(f) Then File.Delete(f)
                Catch
                End Try
            Next
            _arquivos.Clear()
        End Sub

        Private Shared ReadOnly Property Eu As SecurityIdentifier
            Get
                Return WindowsIdentity.GetCurrent().User
            End Get
        End Property

        ''' <summary>
        ''' Um arquivo com ACL <b>construída aqui</b>: herança cortada, e só
        ''' quem o teste mandar.
        ''' </summary>
        Private Function Arquivo(ParamArray extras As SecurityIdentifier()) As String
            Dim caminho = Path.Combine(Path.GetTempPath(),
                                       "iris-perm-" & Guid.NewGuid().ToString("N") & ".json")
            File.WriteAllText(caminho, "{}")
            _arquivos.Add(caminho)

            Dim fi As New FileInfo(caminho)
            Dim acl = fi.GetAccessControl()
            acl.SetAccessRuleProtection(True, False)   ' corta a heranca

            For Each regra In acl.GetAccessRules(True, False, GetType(SecurityIdentifier)).
                              Cast(Of FileSystemAccessRule)().ToList()
                acl.RemoveAccessRule(regra)
            Next

            acl.SetOwner(Eu)
            acl.AddAccessRule(New FileSystemAccessRule(
                Eu, FileSystemRights.FullControl, AccessControlType.Allow))
            For Each sid In extras
                acl.AddAccessRule(New FileSystemAccessRule(
                    sid, FileSystemRights.FullControl, AccessControlType.Allow))
            Next

            fi.SetAccessControl(acl)
            Return caminho
        End Function

        Private Shared Function Passa(caminho As String) As Boolean
            Using fs As New FileStream(caminho, FileMode.Open, FileAccess.Read, FileShare.Read)
                Return PermissaoDoArquivo.SoMinha(fs)
            End Using
        End Function

        ' ==================================================================

        ''' <summary>
        ''' Controle: só eu, e passa.
        '''
        ''' Sem ele, uma conferência que recusasse tudo passaria em todos os
        ''' testes abaixo — e a IA ficaria desligada para sempre, pelo motivo
        ''' errado.
        ''' </summary>
        <TestMethod>
        Public Sub Controle_so_EU_e_passa()
            Assert.IsTrue(Passa(Arquivo()))
        End Sub

        ''' <summary>
        ''' <b>SYSTEM e Administradores não reprovam.</b>
        '''
        ''' Quem é administrador da máquina já pode trocar o próprio Iris.
        ''' Excluí-los seria teatro, e uma barreira que ninguém consegue
        ''' satisfazer vira uma barreira desligada.
        ''' </summary>
        <TestMethod>
        Public Sub SYSTEM_e_Administradores_NAO_reprovam()
            Assert.IsTrue(Passa(Arquivo(
                New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing),
                New SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, Nothing))))
        End Sub

        ''' <summary>
        ''' <b>Mais alguém com escrita reprova.</b>
        '''
        ''' <c>Todos</c> é o caso extremo e o mais fácil de reconhecer; o caso
        ''' real desta máquina era mais discreto — um SID que nem resolve.
        ''' Quem pode escrever pode trocar a autorização, e trocar a autorização
        ''' é escolher para onde o e-mail vai.
        ''' </summary>
        <TestMethod>
        Public Sub Mais_alguem_com_ESCRITA_reprova()
            Assert.IsFalse(Passa(Arquivo(
                New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing))))
        End Sub

        ''' <summary>
        ''' <b>Um SID que não resolve reprova igual.</b>
        '''
        ''' É o caso concreto que apareceu: ACE órfã, de outra máquina, com
        ''' controle total. Ninguém a assume hoje — e "hoje" não é uma
        ''' propriedade de segurança.
        ''' </summary>
        <TestMethod>
        Public Sub SID_que_NAO_RESOLVE_reprova_igual()
            Assert.IsFalse(Passa(Arquivo(
                New SecurityIdentifier("S-1-5-21-2127633640-1427687438-3421751345-1000"))))
        End Sub

        ''' <summary>
        ''' <b>Leitura por outra conta não reprova.</b>
        '''
        ''' A regra é sobre <b>escrever</b>. Reprovar por leitura tornaria a
        ''' conferência impossível de satisfazer na prática, e o usuário
        ''' desligaria a proteção inteira.
        ''' </summary>
        <TestMethod>
        Public Sub LEITURA_por_outra_conta_nao_reprova()
            Dim caminho = Arquivo()
            Dim fi As New FileInfo(caminho)
            Dim acl = fi.GetAccessControl()
            acl.AddAccessRule(New FileSystemAccessRule(
                New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing),
                FileSystemRights.Read, AccessControlType.Allow))
            fi.SetAccessControl(acl)

            Assert.IsTrue(Passa(caminho))
        End Sub

        ''' <summary><b>Sem fluxo, não passa.</b> Falha fechada.</summary>
        <TestMethod>
        Public Sub Sem_fluxo_NAO_passa()
            Assert.IsFalse(PermissaoDoArquivo.SoMinha(Nothing))
        End Sub

    End Class

End Namespace

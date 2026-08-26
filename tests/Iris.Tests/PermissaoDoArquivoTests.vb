Imports System.Collections.Generic
Imports System.Diagnostics
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
        Private ReadOnly _pastas As New List(Of String)()

        <TestCleanup>
        Public Sub Limpar()
            For Each caminho In _arquivos
                Try
                    If File.Exists(caminho) Then File.Delete(caminho)
                Catch
                End Try
            Next
            For Each caminho In _pastas
                Try
                    If Directory.Exists(caminho) Then Directory.Delete(caminho, True)
                Catch
                End Try
            Next
            _arquivos.Clear()
            _pastas.Clear()
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
        ''' <summary>
        ''' Uma pasta com ACL construída aqui: herança cortada, dono eu, e só eu.
        '''
        ''' Precisa existir porque a conferência passou a olhar o <b>diretório</b>
        ''' também — e o <c>%TEMP%</c> tem exatamente as entradas a mais que o
        ''' resto deste arquivo existe para detectar.
        ''' </summary>
        ''' <summary>
        ''' <b>Duas</b> pastas limpas encaixadas, e devolve a de dentro.
        '''
        ''' Duas porque a conferência olha três níveis: arquivo, pasta e
        ''' <b>pasta-mãe</b>. Uma pasta só deixaria o <c>%TEMP%</c> como avó — e o
        ''' <c>%TEMP%</c> tem exatamente as entradas a mais que este arquivo
        ''' existe para detectar, o que faria todo controle daqui reprovar por um
        ''' motivo que não é o dele.
        ''' </summary>
        Private Function PastaLimpa() As String
            Dim raiz = Limpar(Path.Combine(Path.GetTempPath(),
                                           "iris-perm-" & Guid.NewGuid().ToString("N")))
            Return Limpar(Path.Combine(raiz, "dentro"))
        End Function

        ''' <summary>Cria a pasta com herança cortada, dono eu, e só eu.</summary>
        Private Function Limpar(pasta As String) As String
            Directory.CreateDirectory(pasta)
            _pastas.Insert(0, pasta)   ' a de dentro apaga primeiro

            Dim di As New DirectoryInfo(pasta)
            Dim acl = di.GetAccessControl()
            acl.SetAccessRuleProtection(True, False)
            For Each regra In acl.GetAccessRules(True, False, GetType(SecurityIdentifier)).
                              Cast(Of FileSystemAccessRule)().ToList()
                acl.RemoveAccessRule(regra)
            Next
            acl.SetOwner(Eu)
            acl.AddAccessRule(New FileSystemAccessRule(
                Eu, FileSystemRights.FullControl,
                InheritanceFlags.ObjectInherit Or InheritanceFlags.ContainerInherit,
                PropagationFlags.None, AccessControlType.Allow))
            di.SetAccessControl(acl)
            Return pasta
        End Function

        Private Function Arquivo(ParamArray extras As SecurityIdentifier()) As String
            Dim caminho = Path.Combine(PastaLimpa(), "ativacao.json")
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

        ''' <summary>
        ''' <b>ACE perigosa SÓ no diretório reprova — mesmo com o arquivo
        ''' impecável.</b>
        '''
        ''' É o furo que a conferência do arquivo não fecha, e que os outros
        ''' testes deste arquivo não alcançavam. Quem tem <c>CreateFiles</c> e
        ''' <c>DeleteSubdirectoriesAndFiles</c> na pasta <b>apaga o
        ''' <c>ativacao.json</c> e põe outro no lugar</b>, sem precisar de
        ''' direito nenhum sobre o arquivo.
        '''
        ''' A ACE aqui é <b>sem herança</b> de propósito: a DACL do arquivo fica
        ''' limpa, e é justamente por isso que olhar só o arquivo não bastava. O
        ''' handle protege a carga que já foi lida; ele não protege o <b>nome</b>
        ''' entre uma execução e a seguinte.
        ''' </summary>
        <TestMethod>
        Public Sub ACE_perigosa_SO_no_diretorio_reprova()
            Dim caminho = Arquivo()
            Dim di As New DirectoryInfo(Path.GetDirectoryName(caminho))

            Dim acl = di.GetAccessControl()
            acl.AddAccessRule(New FileSystemAccessRule(
                New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing),
                FileSystemRights.CreateFiles Or
                FileSystemRights.DeleteSubdirectoriesAndFiles,
                InheritanceFlags.None,          ' NAO herda: o arquivo continua limpo
                PropagationFlags.None,
                AccessControlType.Allow))
            di.SetAccessControl(acl)

            ' O arquivo continua com a ACL de antes — o teste so mede alguma
            ' coisa se isso for verdade.
            Dim doArquivo = New FileInfo(caminho).GetAccessControl().
                            GetAccessRules(True, True, GetType(SecurityIdentifier)).
                            Cast(Of FileSystemAccessRule)().
                            Any(Function(r) r.IdentityReference.Value =
                                            New SecurityIdentifier(WellKnownSidType.WorldSid,
                                                                   Nothing).Value)
            Assert.IsFalse(doArquivo, "a ACE nao devia ter herdado para o arquivo")

            Assert.IsFalse(Passa(caminho),
                           "quem pode trocar o arquivo na pasta escolhe a autorizacao")
        End Sub

        ''' <summary>
        ''' <b>Por que não há teste separado para o dono do diretório.</b>
        '''
        ''' Escrevi um, e ele não provava nada: usava o <c>hosts</c>, cujo
        ''' <b>arquivo</b> também pertence a Administradores — então a
        ''' conferência do arquivo já reprovava antes de a do diretório ser
        ''' consultada. O controle negativo pegou: desligando a conferência do
        ''' diretório, ele continuava vermelho.
        '''
        ''' Isolar o caso exigiria um diretório pertencente a outro principal e
        ''' com DACL limpa no resto, e trocar dono de diretório precisa de
        ''' privilégio que uma sessão comum não tem.
        '''
        ''' O que sustenta essa metade: <c>Limpo()</c> é <b>uma função só</b>,
        ''' usada para o arquivo e para a pasta, e o ramo do dono é literalmente
        ''' o mesmo código — coberto por <see cref="Dono_DIFERENTE_reprova"/>. O
        ''' que é específico do diretório é o <b>conjunto de direitos</b>, e isso
        ''' tem prova própria logo acima.
        ''' </summary>

        ''' <summary>
        ''' <b>ACE de escrita HERDADA do diretório reprova.</b>
        '''
        ''' É o caso real, e o que os outros testes deste arquivo não alcançam:
        ''' todos eles cortam a herança para construir a ACL que medem, e a
        ''' herança é justamente de onde veio o problema na máquina do usuário.
        '''
        ''' Aqui o diretório recebe uma ACE <b>herdável</b> e o arquivo nasce
        ''' dentro dele. Ninguém escreveu nada na ACL do arquivo — e mesmo assim
        ''' há mais gente podendo escrever.
        ''' </summary>
        <TestMethod>
        Public Sub ACE_HERDADA_do_diretorio_reprova()
            Dim pasta = Limpar(Path.Combine(PastaLimpa(), "com-heranca"))

            Dim di As New DirectoryInfo(pasta)
            Dim aclDir = di.GetAccessControl()
            aclDir.AddAccessRule(New FileSystemAccessRule(
                New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing),
                FileSystemRights.Modify,
                InheritanceFlags.ObjectInherit Or InheritanceFlags.ContainerInherit,
                PropagationFlags.None,
                AccessControlType.Allow))
            di.SetAccessControl(aclDir)

            Dim caminho = Path.Combine(pasta, "ativacao.json")
            File.WriteAllText(caminho, "{}")

            Assert.IsFalse(Passa(caminho),
                           "a ACE veio da heranca, e heranca conta igual")
        End Sub

        ''' <summary>
        ''' Controle da herança: diretório <b>sem</b> a ACE extra, arquivo passa.
        '''
        ''' Sem ele, o teste de cima passaria por qualquer motivo — inclusive
        ''' porque arquivo em <c>%TEMP%</c> nunca passa.
        ''' </summary>
        <TestMethod>
        Public Sub Controle_diretorio_LIMPO_o_arquivo_passa()
            ' A cadeia inteira limpa: avo, pasta e arquivo. O arquivo herda a
            ' ACL da pasta e ninguem mexe nele.
            Dim caminho = Path.Combine(PastaLimpa(), "ativacao.json")
            File.WriteAllText(caminho, "{}")

            Assert.IsTrue(Passa(caminho))
        End Sub

        ''' <summary>
        ''' <b>Cada direito de escrita reprova sozinho.</b>
        '''
        ''' Os outros testes davam <c>FullControl</c> ao intruso, e isso os
        ''' deixava passar mesmo se <c>Escreve()</c> esquecesse um direito — bastaria
        ''' reconhecer <b>um</b> deles. Aqui cada um aparece isolado, e um
        ''' esquecimento vira uma linha vermelha com nome.
        ''' </summary>
        <TestMethod>
        Public Sub Cada_direito_de_escrita_reprova_SOZINHO()
            For Each direito In {FileSystemRights.WriteData,
                                 FileSystemRights.AppendData,
                                 FileSystemRights.Delete,
                                 FileSystemRights.ChangePermissions,
                                 FileSystemRights.TakeOwnership}
                Dim caminho = Arquivo()
                Dim fi As New FileInfo(caminho)
                Dim acl = fi.GetAccessControl()
                acl.AddAccessRule(New FileSystemAccessRule(
                    New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing),
                    direito, AccessControlType.Allow))
                fi.SetAccessControl(acl)

                Assert.IsFalse(Passa(caminho), $"{direito} sozinho tinha de reprovar")
            Next
        End Sub

        ''' <summary>
        ''' <b>Dono diferente reprova, mesmo sem ACE sobrando.</b>
        '''
        ''' Quem é dono pode mudar a ACL a qualquer momento, então conferir só as
        ''' ACEs deixaria a proteção valendo até o dono decidir o contrário.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE UM ARQUIVO DO SISTEMA, E NÃO UM QUE O TESTE CRIA</b>
        '''
        ''' A primeira versão trocava o dono de um arquivo próprio e caía em
        ''' <c>Inconclusive</c> quando não conseguia — que é <b>sempre</b>, numa
        ''' sessão não elevada: o SID de Administradores entra no token como
        ''' "deny only" e <c>SetOwner</c> falha. Um teste que se desliga sozinho e
        ''' reporta verde é pior que teste nenhum.
        '''
        ''' O <c>hosts</c> é legível por qualquer usuário e pertence ao sistema.
        ''' Não precisa de privilégio nenhum, e o dono é de verdade outro.
        ''' </summary>
        <TestMethod>
        Public Sub Dono_DIFERENTE_reprova()
            Const doSistema = "C:\Windows\System32\drivers\etc\hosts"
            If Not File.Exists(doSistema) Then
                Assert.Inconclusive("esta maquina nao tem o arquivo hosts no lugar de sempre")
            End If

            Dim dono = TryCast(New FileInfo(doSistema).GetAccessControl().
                               GetOwner(GetType(SecurityIdentifier)), SecurityIdentifier)
            Assert.IsNotNull(dono)
            Assert.AreNotEqual(Eu, dono,
                               "o teste so mede alguma coisa se o dono for outro mesmo")

            Assert.IsFalse(Passa(doSistema))
        End Sub

        ''' <summary>
        ''' <b>ACE perigosa só na PASTA-MÃE reprova.</b>
        '''
        ''' O terceiro nível, e o último que se confere. Quem tem
        ''' <c>CreateDirectories</c> e <c>DeleteSubdirectoriesAndFiles</c> na
        ''' mãe <b>renomeia a pasta inteira</b> e põe outra no lugar, com outra
        ''' ativação dentro — sem nunca tocar no arquivo nem na pasta que o Iris
        ''' confere.
        '''
        ''' A ACE é <b>sem herança</b>: o teste confere que ela não chegou nem à
        ''' pasta nem ao arquivo, e mesmo assim a conferência recusa.
        ''' </summary>
        <TestMethod>
        Public Sub ACE_perigosa_so_na_PASTA_MAE_reprova()
            Dim caminho = Arquivo()
            Dim pasta As New DirectoryInfo(Path.GetDirectoryName(caminho))
            Dim mae = pasta.Parent

            Dim todos = New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing)
            Dim acl = mae.GetAccessControl()
            acl.AddAccessRule(New FileSystemAccessRule(
                todos,
                FileSystemRights.CreateDirectories Or
                FileSystemRights.DeleteSubdirectoriesAndFiles,
                InheritanceFlags.None,          ' NAO desce
                PropagationFlags.None,
                AccessControlType.Allow))
            mae.SetAccessControl(acl)

            Assert.IsFalse(TemAce(pasta.FullName, todos), "nao devia ter descido para a pasta")
            Assert.IsFalse(TemAce(caminho, todos), "nem para o arquivo")

            Assert.IsFalse(Passa(caminho),
                           "quem troca a pasta inteira escolhe a autorizacao")
        End Sub

        ''' <summary>
        ''' <b>Pasta que é junction não passa.</b>
        '''
        ''' Um ponto de nova análise aponta para outro lugar — e o outro lugar
        ''' ninguém conferiu: nem o dono dele, nem quem escreve lá. Conferir a
        ''' ACL do link diria coisas sobre o link, e o conteúdo viria de outro
        ''' sítio.
        ''' </summary>
        <TestMethod>
        Public Sub Pasta_que_e_JUNCTION_nao_passa()
            Dim real = PastaLimpa()
            Dim caminho = Path.Combine(real, "ativacao.json")
            File.WriteAllText(caminho, "{}")
            Assert.IsTrue(Passa(caminho), "controle: pela pasta de verdade, passa")

            ' JUNCTION, e nao symlink: symlink exige privilegio e junction
            ' NAO. Um teste que usasse symlink mediria um vetor que quase
            ' ninguem consegue usar, e deixaria de fora o que qualquer conta
            ' cria.
            Dim link = Path.Combine(Path.GetDirectoryName(real), "atalho")
            Using cmd = Process.Start(New ProcessStartInfo("cmd.exe") With {
                .Arguments = $"/c mklink /J ""{link}"" ""{real}""",
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True})
                cmd.WaitForExit(10000)
                Assert.AreEqual(0, cmd.ExitCode,
                                "nao consegui criar a junction: " & cmd.StandardError.ReadToEnd())
            End Using
            _pastas.Insert(0, link)

            Assert.IsTrue((New DirectoryInfo(link).Attributes And
                           FileAttributes.ReparsePoint) <> 0,
                          "o teste so mede alguma coisa se isto for reparse point")

            Assert.IsFalse(Passa(Path.Combine(link, "ativacao.json")),
                           "o mesmo arquivo, alcancado por link, nao pode passar")
        End Sub

        ''' <summary>Alguém com esse SID aparece na ACL deste objeto?</summary>
        Private Shared Function TemAce(caminho As String, quem As SecurityIdentifier) As Boolean
            Dim seguranca As FileSystemSecurity =
                If(Directory.Exists(caminho),
                   CType(New DirectoryInfo(caminho).GetAccessControl(), FileSystemSecurity),
                   New FileInfo(caminho).GetAccessControl())

            Return seguranca.GetAccessRules(True, True, GetType(SecurityIdentifier)).
                   Cast(Of FileSystemAccessRule)().
                   Any(Function(r) r.IdentityReference.Value = quem.Value)
        End Function

        ''' <summary><b>Sem fluxo, não passa.</b> Falha fechada.</summary>
        <TestMethod>
        Public Sub Sem_fluxo_NAO_passa()
            Assert.IsFalse(PermissaoDoArquivo.SoMinha(Nothing))
        End Sub

    End Class

End Namespace

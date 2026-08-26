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
    ''' <b>ONDE A CONFIANÇA COMEÇA</b>
    '''
    ''' Três níveis, e depois <b>para</b>: o arquivo, a pasta que o contém, e a
    ''' <b>pasta-mãe</b> — que é quem controla o <i>nome</i> da pasta. Sem o
    ''' terceiro, quem pudesse criar e apagar em <c>%LOCALAPPDATA%</c> renomearia
    ''' a pasta <c>Iris</c> e poria outra no lugar, com outra ativação dentro.
    '''
    ''' Acima disso não se confere, e isso é <b>decisão declarada</b> e não
    ''' esquecimento: a recursão não tem fim natural. O limite fica escrito aqui
    ''' para que uma revisão futura discuta o limite, e não descubra a ausência
    ''' dele.
    '''
    ''' A âncora usa um conjunto de direitos <b>mais estreito</b> que os outros
    ''' dois — ver <see cref="DaAncora"/>. É o que faz a regra ser satisfazível
    ''' numa máquina real em vez de exigir cirurgia de ACL no perfil.
    '''
    ''' <b>Junction também não.</b> Uma pasta que é ponto de nova análise aponta
    ''' para outro lugar, e o outro lugar ninguém conferiu.
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

                ' JUNCTION NAO. Uma pasta que e ponto de nova analise aponta
                ' para outro lugar, e o outro lugar ninguem conferiu — nem o
                ' dono dele, nem quem escreve la.
                If di.LinkTarget IsNot Nothing OrElse
                   (di.Attributes And FileAttributes.ReparsePoint) <> 0 Then
                    Return False
                End If

                If Not Limpo(di.GetAccessControl(), meuSid, DaPasta) Then Return False

                ' E A PASTA-MAE, QUE E ONDE A CONFIANCA COMECA.
                '
                ' Quem tem CreateDirectories e DeleteSubdirectoriesAndFiles em
                ' %LOCALAPPDATA% renomeia a pasta Iris e poe outra no lugar —
                ' ou uma junction — antes da proxima abertura. Conferir a pasta
                ' e nao a mae fecharia a porta e deixaria a parede.
                Dim mae = di.Parent
                If mae Is Nothing OrElse Not mae.Exists Then Return False

                ' Junction na mae tambem nao: ela decide onde a pasta mora.
                If mae.LinkTarget IsNot Nothing OrElse
                   (mae.Attributes And FileAttributes.ReparsePoint) <> 0 Then
                    Return False
                End If

                ' A ANCORA usa um conjunto de direitos MAIS ESTREITO, e o dono
                ' pode ser o sistema. Ver o doc de DaAncora para por que criar
                ' nao e o mesmo que substituir.
                Return Limpo(mae.GetAccessControl(), meuSid, DaAncora,
                             donoTemDeSerEu:=False)

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
                                      perigosos As FileSystemRights,
                                      Optional donoTemDeSerEu As Boolean = True) As Boolean

            Dim permitidos As New HashSet(Of SecurityIdentifier) From {meuSid}
            permitidos.Add(New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing))
            permitidos.Add(New SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid,
                                                  Nothing))

            Dim dono = TryCast(seguranca.GetOwner(GetType(SecurityIdentifier)),
                               SecurityIdentifier)
            If dono Is Nothing Then Return False
            If donoTemDeSerEu Then
                If Not dono.Equals(meuSid) Then Return False
            ElseIf Not permitidos.Contains(dono) Then
                Return False
            End If

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
        ''' Os direitos que permitem <b>substituir a pasta</b> — e só eles.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>CRIAR NÃO É SUBSTITUIR, E ESSA DISTINÇÃO É O QUE TORNA A ÂNCORA
        ''' ALCANÇÁVEL</b>
        '''
        ''' Na âncora eu usava o mesmo conjunto da pasta, e com ele
        ''' <b>nenhuma localização padrão do Windows passava</b>: medi
        ''' <c>%LOCALAPPDATA%</c>, <c>%USERPROFILE%</c> e <c>C:\</c>, e os três
        ''' reprovavam. Uma barreira que ninguém consegue satisfazer é uma
        ''' barreira que alguém desliga.
        '''
        ''' O erro estava no conjunto. Quem tem <c>CreateDirectories</c> numa
        ''' raiz consegue criar pastas <b>irmãs</b> — e não consegue apagar nem
        ''' renomear uma pasta que já existe e está protegida. Para substituí-la
        ''' é preciso <c>DeleteSubdirectoriesAndFiles</c>, <c>Delete</c> no
        ''' filho, ou poder mudar a ACL.
        '''
        ''' Então a âncora recusa exatamente esses, e <b>tolera criar</b>. Com
        ''' isso <c>%ProgramData%</c> passa — <c>Users</c> tem ali só
        ''' <c>WD,AD</c> — e o Iris ganha uma raiz onde a autorização pode morar
        ''' sem exigir cirurgia de ACL no perfil de ninguém.
        ''' </summary>
        Private Shared ReadOnly DaAncora As FileSystemRights =
            FileSystemRights.DeleteSubdirectoriesAndFiles Or
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

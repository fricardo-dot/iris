Imports System.Runtime.InteropServices

Namespace Global.Iris.App

    ''' <summary>
    ''' <b>A chave do provedor, lida do Gerenciador de Credenciais do Windows.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE AQUI, E NÃO NUM ARQUIVO OU NUMA VARIÁVEL</b>
    '''
    ''' O segredo fica cifrado pela conta de login: copiado para outra máquina,
    ''' é lixo ilegível. Variável de ambiente aparece em dump de processo e em
    ''' ferramenta de diagnóstico; arquivo entra em backup e em sincronização de
    ''' pasta sem ninguém pedir.
    '''
    ''' Para gravar, uma vez, no prompt do usuário:
    ''' <code>
    ''' cmdkey /generic:Iris/OpenRouter /user:iris /pass
    ''' </code>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTA CLASSE GARANTE, E O QUE ELA NÃO PODE GARANTIR</b>
    '''
    ''' Garante: o buffer nativo é liberado sempre; a cópia temporária é zerada;
    ''' credencial ausente devolve <b>vazio</b> em vez de lançar, para o
    ''' <c>Pronto()</c> recusar limpo; e a chave é rejeitada se contiver quebra
    ''' de linha, que é como se injeta cabeçalho HTTP.
    '''
    ''' <b>Não</b> garante que a chave não fique em memória gerenciada: ela vira
    ''' uma <c>String</c>, que o .NET não deixa zerar e o coletor move. Dizer
    ''' "a credencial nunca fica na memória" seria falso. O que se pode dizer é
    ''' que ela não vira campo, não vira arquivo e não vira log.
    ''' </summary>
    Public NotInheritable Class CredencialDoWindows

        ''' <summary>O alvo padrão, e o que o <c>cmdkey</c> acima grava.</summary>
        Public Const AlvoPadrao As String = "Iris/OpenRouter"

        Private Const CRED_TYPE_GENERIC As UInteger = 1UI

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure CREDENTIAL
            Public Flags As UInteger
            Public Type As UInteger
            Public TargetName As IntPtr
            Public Comment As IntPtr
            Public LastWritten As Runtime.InteropServices.ComTypes.FILETIME
            Public CredentialBlobSize As UInteger
            Public CredentialBlob As IntPtr
            Public Persist As UInteger
            Public AttributeCount As UInteger
            Public Attributes As IntPtr
            Public TargetAlias As IntPtr
            Public UserName As IntPtr
        End Structure

        <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True,
                   EntryPoint:="CredReadW")>
        Private Shared Function CredRead(alvo As String, tipo As UInteger, reservado As UInteger,
                                         ByRef credencial As IntPtr) As Boolean
        End Function

        <DllImport("advapi32.dll", EntryPoint:="CredFree")>
        Private Shared Sub CredFree(buffer As IntPtr)
        End Sub

        ''' <summary>
        ''' <b>Os caracteres lidos do cofre viram chave — ou nada.</b>
        '''
        ''' Separada da leitura nativa para poder ser <b>testada</b>: a parte
        ''' que fala com o Windows depende da máquina, e esta não depende de
        ''' nada. Sem a separação, as regras abaixo eram promessas de
        ''' comentário — e uma delas era falsa.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O NUL, E O QUE ESTAVA ERRADO</b>
        '''
        ''' Havia um <c>TrimEnd(ChrW(0))</c> <b>antes</b> da verificação, então
        ''' NUL no fim era aceito — o comentário prometia rejeitar todo NUL e o
        ''' código rejeitava só os do meio.
        '''
        ''' Agora: um único NUL <b>final</b> é terminador e sai; qualquer NUL
        ''' que sobre depois disso reprova a chave inteira.
        ''' </summary>
        Friend Shared Function Interpretar(bruto As Char()) As String
            If bruto Is Nothing OrElse bruto.Length = 0 Then Return ""

            Dim fim = bruto.Length
            If bruto(fim - 1) = ChrW(0) Then fim -= 1
            If fim = 0 Then Return ""

            Dim chave = New String(bruto, 0, fim)

            ' QUEBRA DE LINHA NAO. E assim que se injeta cabecalho HTTP, e uma
            ' chave com CR ou LF dentro nao e uma chave. NUL tambem nao: depois
            ' de tirar o terminador, o que sobrou tem de ser texto.
            For Each ch In chave
                If ch = ChrW(13) OrElse ch = ChrW(10) OrElse ch = ChrW(0) Then Return ""
            Next

            ' Espaco em volta e erro de colagem, e nao parte do segredo.
            Return chave.Trim()
        End Function

        ''' <summary>
        ''' A função que o provedor chama <b>na hora</b> de montar o cabeçalho.
        '''
        ''' Devolve uma <c>Func</c> e não a chave: guardar a chave numa variável
        ''' capturada no arranque a manteria viva por toda a execução, e entre
        ''' uma chamada e outra ela pode ter sido revogada.
        ''' </summary>
        Public Shared Function Leitor(Optional alvo As String = AlvoPadrao) As Func(Of String)
            Return Function() Ler(alvo)
        End Function

        ''' <summary>
        ''' A chave, ou <b>vazio</b>. Nunca lança, e nunca diz por quê.
        '''
        ''' O motivo não sai daqui de propósito: a mensagem de erro do Windows
        ''' para um alvo de credencial é o tipo de texto que acaba em log, e o
        ''' nome do alvo já diz o suficiente para quem está configurando.
        ''' </summary>
        Public Shared Function Ler(Optional alvo As String = AlvoPadrao) As String
            If String.IsNullOrWhiteSpace(alvo) Then Return ""

            Dim ponteiro As IntPtr = IntPtr.Zero
            Try
                If Not CredRead(alvo, CRED_TYPE_GENERIC, 0UI, ponteiro) Then Return ""
                If ponteiro = IntPtr.Zero Then Return ""

                Dim c = Marshal.PtrToStructure(Of CREDENTIAL)(ponteiro)
                If c.CredentialBlob = IntPtr.Zero OrElse c.CredentialBlobSize = 0UI Then Return ""

                ' O TAMANHO VEM EM BYTES, e a chave e UTF-16: um numero impar
                ' quer dizer que o que esta la nao e o que se espera — provavel
                ' credencial gravada por outra ferramenta, em outra codificacao.
                ' Ler assim mesmo produziria uma chave silenciosamente truncada.
                If c.CredentialBlobSize Mod 2UI <> 0UI Then Return ""

                ' Teto de sanidade. Chave de API tem dezenas de caracteres; um
                ' blob de megabytes e outra coisa, e alocar para descobrir o que
                ' e ja seria ter alocado.
                If c.CredentialBlobSize > 8192UI Then Return ""

                Dim caracteres = CInt(c.CredentialBlobSize \ 2UI)
                Dim buffer(caracteres - 1) As Char
                Try
                    Marshal.Copy(c.CredentialBlob, buffer, 0, caracteres)
                    Return Interpretar(buffer)
                Finally
                    Array.Clear(buffer, 0, buffer.Length)
                End Try

            Catch
                ' Inclusive DllNotFoundException fora do Windows. Falha fechada:
                ' sem credencial, o Pronto() recusa e nada sai.
                Return ""
            Finally
                If ponteiro <> IntPtr.Zero Then CredFree(ponteiro)
            End Try
        End Function

    End Class

End Namespace

Imports System.Linq
Imports System.Runtime.InteropServices
Imports Iris.App
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace Global.Iris.Tests

    ''' <summary>
    ''' <b>A leitura da credencial do Gerenciador de Credenciais.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ELE EXISTE</b>
    '''
    ''' O Codex apontou que esta classe não tinha prova nenhuma, e que uma das
    ''' regras que ela promete no comentário estava <b>errada no código</b>:
    ''' havia um <c>TrimEnd(ChrW(0))</c> antes da verificação, então NUL no fim
    ''' passava. Regra que só existe em comentário não é regra.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DOIS NÍVEIS, E POR QUÊ</b>
    '''
    ''' A <b>interpretação</b> é pura e testada sem tocar em nada. O
    ''' <b>caminho nativo</b> é testado gravando uma credencial num alvo
    ''' aleatório <c>Iris.Tests/{GUID}</c>, lendo, e apagando no <c>Finally</c>.
    '''
    ''' <b>A credencial real do usuário não é tocada.</b> Nem lida, nem
    ''' escrita, nem apagada — um teste que mexesse nela poderia derrubar a
    ''' ativação de quem está usando o programa.
    ''' </summary>
    <TestClass>
    Public Class CredencialDoWindowsTests

        ' ==================================================================
        ' A interpretação, pura

        ''' <summary>Controle: uma chave comum atravessa inteira.</summary>
        <TestMethod>
        Public Sub Controle_uma_chave_comum_ATRAVESSA()
            Assert.AreEqual("sk-or-v1-abc123", Interpretar("sk-or-v1-abc123"))
        End Sub

        ''' <summary>
        ''' <b>O NUL final é terminador e sai; o do meio reprova.</b>
        '''
        ''' É o defeito que o Codex achou. O blob do Windows pode vir com
        ''' terminador, e "terminador" e "conteúdo com NUL" não se distinguem
        ''' olhando só o fim da cadeia — por isso um sai e o resto reprova.
        ''' </summary>
        <TestMethod>
        Public Sub O_NUL_final_SAI_e_o_do_meio_REPROVA()
            Assert.AreEqual("abc", Interpretar("abc" & ChrW(0)))
            Assert.AreEqual("", Interpretar("ab" & ChrW(0) & "c"),
                            "NUL no meio nao e terminador: e conteudo que nao devia estar la")
            Assert.AreEqual("", Interpretar("abc" & ChrW(0) & ChrW(0)),
                            "dois NUL no fim: o segundo sobra, e sobra reprova")
        End Sub

        ''' <summary>
        ''' <b>Quebra de linha reprova — é assim que se injeta cabeçalho.</b>
        '''
        ''' Uma chave com CR ou LF dentro produziria um <c>Authorization</c>
        ''' partido em dois e um cabeçalho a mais na requisição.
        ''' </summary>
        <TestMethod>
        Public Sub Quebra_de_linha_REPROVA()
            For Each ruim In {"abc" & ChrW(13) & "def",
                              "abc" & ChrW(10) & "def",
                              "abc" & ChrW(13) & ChrW(10) & "X-Injetado: sim"}
                Assert.AreEqual("", Interpretar(ruim), ruim.Replace(ChrW(13), "\r").
                                                            Replace(ChrW(10), "\n"))
            Next
        End Sub

        ''' <summary>
        ''' <b>Espaço em volta é erro de colagem, e sai.</b>
        '''
        ''' Copiar a chave do painel do provedor traz espaço ou quebra atrás com
        ''' frequência. Mandar " sk-…" produziria um 401 que ninguém entende.
        ''' </summary>
        <TestMethod>
        Public Sub Espaco_em_volta_SAI()
            Assert.AreEqual("sk-abc", Interpretar("  sk-abc  "))
        End Sub

        ''' <summary><b>Vazio, só NUL, ou nada: chave nenhuma.</b></summary>
        <TestMethod>
        Public Sub Vazio_e_chave_NENHUMA()
            Assert.AreEqual("", Interpretar(""))
            Assert.AreEqual("", Interpretar(ChrW(0)))
            Assert.AreEqual("", Interpretar("   "))
            Assert.AreEqual("", CredencialDoWindows.Interpretar(Nothing))
        End Sub

        ' ==================================================================
        ' O caminho nativo, contra o cofre de verdade

        ''' <summary>
        ''' <b>Grava, lê e apaga — num alvo que só este teste usa.</b>
        '''
        ''' Prova o que a parte pura não alcança: o P/Invoke, o tamanho do blob
        ''' em bytes, e a conversão UTF-16. Sem isto, a leitura real só seria
        ''' exercitada na primeira chamada de verdade.
        ''' </summary>
        <TestMethod>
        Public Sub Grava_le_e_apaga_no_cofre_de_VERDADE()
            Dim alvo = "Iris.Tests/" & Guid.NewGuid().ToString("N")
            Const segredo = "sk-de-teste-" & "abcdef0123456789"

            Try
                Assert.IsTrue(Gravar(alvo, segredo), "nao consegui gravar a credencial de teste")

                Assert.AreEqual(segredo, CredencialDoWindows.Ler(alvo))

                ' E o leitor tardio le a MESMA coisa — ele existe para ser
                ' chamado na hora de cada envio, e nao no arranque.
                Assert.AreEqual(segredo, CredencialDoWindows.Leitor(alvo)())
            Finally
                Apagar(alvo)
            End Try

            ' E depois de apagada, vazio — e nao excecao. E isto que faz o
            ' Pronto() do provedor recusar limpo em vez de explodir.
            Assert.AreEqual("", CredencialDoWindows.Ler(alvo))
        End Sub

        ''' <summary>
        ''' <b>Alvo que não existe devolve vazio, e não lança.</b>
        '''
        ''' O caso normal antes de o usuário guardar a chave.
        ''' </summary>
        <TestMethod>
        Public Sub Alvo_INEXISTENTE_devolve_vazio()
            Assert.AreEqual("", CredencialDoWindows.Ler(
                "Iris.Tests/nao-existe-" & Guid.NewGuid().ToString("N")))
            Assert.AreEqual("", CredencialDoWindows.Ler(""))
            Assert.AreEqual("", CredencialDoWindows.Ler("   "))
        End Sub

        ' ==================================================================

        Private Shared Function Interpretar(s As String) As String
            Return CredencialDoWindows.Interpretar(s.ToCharArray())
        End Function

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure CREDENTIAL
            Public Flags As UInteger
            Public Type As UInteger
            Public TargetName As IntPtr
            Public Comment As IntPtr
            Public LastWritten As ComTypes.FILETIME
            Public CredentialBlobSize As UInteger
            Public CredentialBlob As IntPtr
            Public Persist As UInteger
            Public AttributeCount As UInteger
            Public Attributes As IntPtr
            Public TargetAlias As IntPtr
            Public UserName As IntPtr
        End Structure

        <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True,
                   EntryPoint:="CredWriteW")>
        Private Shared Function CredWrite(ByRef c As CREDENTIAL, flags As UInteger) As Boolean
        End Function

        <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True,
                   EntryPoint:="CredDeleteW")>
        Private Shared Function CredDelete(alvo As String, tipo As UInteger,
                                           reservado As UInteger) As Boolean
        End Function

        Private Shared Function Gravar(alvo As String, segredo As String) As Boolean
            Dim bytes = Text.Encoding.Unicode.GetBytes(segredo)
            Dim blob = Marshal.AllocHGlobal(bytes.Length)
            Dim nome = Marshal.StringToHGlobalUni(alvo)
            Dim usuario = Marshal.StringToHGlobalUni("iris-teste")
            Try
                Marshal.Copy(bytes, 0, blob, bytes.Length)
                Dim c As New CREDENTIAL With {
                    .Type = 1UI,
                    .TargetName = nome,
                    .CredentialBlobSize = CUInt(bytes.Length),
                    .CredentialBlob = blob,
                    .Persist = 2UI,
                    .UserName = usuario}
                Return CredWrite(c, 0UI)
            Finally
                Marshal.FreeHGlobal(blob)
                Marshal.FreeHGlobal(nome)
                Marshal.FreeHGlobal(usuario)
            End Try
        End Function

        Private Shared Sub Apagar(alvo As String)
            Try
                CredDelete(alvo, 1UI, 0UI)
            Catch
                ' O teste ja terminou; nao ha o que fazer aqui.
            End Try
        End Sub

    End Class

End Namespace

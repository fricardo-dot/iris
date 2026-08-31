Imports System.IO
Imports System.Linq
Imports Iris.Integration
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>AS DUAS AÇÕES QUE LIMPAM A FILA SEM IA.</b>
'''
''' Newsletter, notificação automática e "obrigado, recebido" entram na fila
''' como pendências de vinte dias. Uma fila com lixo é uma fila que se aprende
''' a ignorar — inclusive nas linhas em que ela acertou —, e as duas ações daqui
''' resolvem a maior parte disso <b>antes</b> de existir classificador.
'''
''' O controle negativo é <see cref="Falha_ao_gravar_NAO_e_sucesso"/>: dizer que
''' dispensou sem ter dispensado deixa a linha na fila e o dono achando que
''' resolveu — e quem foi enganado não tenta de novo.
''' </summary>
<TestClass>
Public Class DispensasDaFilaTests

    Private Shared Function Pasta() As String
        Return Path.Combine(Path.GetTempPath(), "iris-dispensas-" & Guid.NewGuid().ToString("N"))
    End Function

    Private Shared Sub Comigo(corpo As Action(Of DispensasDaFila, String))
        Dim pasta = DispensasDaFilaTests.Pasta()
        Try
            corpo(New DispensasDaFila(pasta), pasta)
        Finally
            If Directory.Exists(pasta) Then Directory.Delete(pasta, recursive:=True)
        End Try
    End Sub

    <TestMethod>
    Public Sub Conversa_dispensada_volta_na_leitura()
        Comigo(Sub(d, pasta)
                   Assert.AreEqual(0, d.Conversas().Count, "nasce vazio")

                   Assert.IsTrue(d.DispensarConversa("conversa-1"))
                   Assert.IsTrue(d.DispensarConversa("conversa-2"))

                   CollectionAssert.AreEquivalent({"conversa-1", "conversa-2"},
                                                  d.Conversas().ToArray())
               End Sub)
    End Sub

    ''' <summary>
    ''' Clicar duas vezes não escreve duas linhas — e responde sucesso, porque a
    ''' conversa <b>está</b> dispensada, que é o que o dono queria.
    ''' </summary>
    <TestMethod>
    Public Sub Dispensar_a_mesma_conversa_duas_vezes_e_inofensivo()
        Comigo(Sub(d, pasta)
                   d.DispensarConversa("c1")
                   Assert.IsTrue(d.DispensarConversa("c1"))

                   Assert.AreEqual(1, d.Conversas().Count)
               End Sub)
    End Sub

    ''' <summary>
    ''' O remetente ignorado volta como <see cref="MinhasIdentidades"/> — o mesmo
    ''' tipo das identidades do dono, e portanto o mesmo casamento de endereço:
    ''' caixa, nome de exibição e X.500 tratados igual. Um casamento escrito de
    ''' novo divergiria em algum caso, e a regra funcionaria para uns remetentes
    ''' e não para outros.
    ''' </summary>
    <TestMethod>
    Public Sub Remetente_ignorado_casa_como_identidade()
        Comigo(Sub(d, pasta)
                   Assert.IsTrue(d.IgnorarRemetente("noreply@boletim.com"))

                   Dim ignorados = d.Remetentes()
                   Assert.AreEqual(Direcao.Minha, ignorados.DirecaoDe("NOREPLY@BOLETIM.COM"),
                       "a caixa das letras mudou o casamento")
                   Assert.AreEqual(Direcao.Minha,
                       ignorados.DirecaoDe("Boletim <noreply@boletim.com>"),
                       "o nome de exibicao mudou o casamento")
                   Assert.AreEqual(Direcao.DoOutro, ignorados.DirecaoDe("gente@empresa.com"))
               End Sub)
    End Sub

    ''' <summary>
    ''' Os dois arquivos têm vidas diferentes: conversa dispensada é fato
    ''' pontual, remetente ignorado é regra permanente. Misturá-los faria a
    ''' limpeza de um apagar o outro.
    ''' </summary>
    <TestMethod>
    Public Sub Sao_DOIS_arquivos()
        Comigo(Sub(d, pasta)
                   d.DispensarConversa("c1")
                   d.IgnorarRemetente("noreply@boletim.com")

                   Assert.AreNotEqual(d.CaminhoDasConversas, d.CaminhoDosRemetentes)
                   Assert.AreEqual(1, d.Conversas().Count)
                   Assert.AreEqual(1, d.Remetentes().Quantas)

                   ' Apagar um nao leva o outro.
                   File.Delete(d.CaminhoDasConversas)
                   Assert.AreEqual(0, d.Conversas().Count)
                   Assert.AreEqual(1, d.Remetentes().Quantas)
               End Sub)
    End Sub

    ''' <summary>
    ''' O cabeçalho de comentários não vira conversa dispensada — seria uma
    ''' dispensa fantasma que ninguém pediu.
    ''' </summary>
    <TestMethod>
    Public Sub O_cabecalho_nao_vira_dispensa()
        Comigo(Sub(d, pasta)
                   d.DispensarConversa("c1")

                   Dim linhas = File.ReadAllLines(d.CaminhoDasConversas)
                   Assert.IsTrue(linhas.Length > 1, "o cabecalho tem de estar la")
                   Assert.AreEqual(1, d.Conversas().Count, "o cabecalho virou dispensa")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo.</b> Dizer que dispensou sem ter dispensado deixa
    ''' a linha na fila e o dono achando que resolveu — e quem foi enganado não
    ''' tenta de novo. Aqui o caminho é impossível de gravar.
    ''' </summary>
    <TestMethod>
    Public Sub Falha_ao_gravar_NAO_e_sucesso()
        ' Um caminho que nao da para criar: um arquivo no lugar da pasta.
        Dim atrapalho = Path.Combine(Path.GetTempPath(),
                                     "iris-atrapalho-" & Guid.NewGuid().ToString("N"))
        Try
            File.WriteAllText(atrapalho, "sou um arquivo, e nao uma pasta")
            Dim d As New DispensasDaFila(atrapalho)

            Assert.IsFalse(d.DispensarConversa("c1"),
                "disse que dispensou sobre uma pasta que nao existe")
            Assert.AreEqual(0, d.Conversas().Count)
        Finally
            If File.Exists(atrapalho) Then File.Delete(atrapalho)
        End Try
    End Sub

    ''' <summary>Valor vazio não vira dispensa de coisa nenhuma.</summary>
    <TestMethod>
    Public Sub Vazio_nao_dispensa_nada()
        Comigo(Sub(d, pasta)
                   Assert.IsFalse(d.DispensarConversa(""))
                   Assert.IsFalse(d.DispensarConversa("   "))
                   Assert.IsFalse(d.IgnorarRemetente(Nothing))
                   Assert.AreEqual(0, d.Conversas().Count)
               End Sub)
    End Sub

    ''' <summary>
    ''' Arquivo ausente vale como nada dispensado, e a fila mostra <i>mais</i> do
    ''' que deveria. É o lado certo de errar: linha a mais o dono descarta de
    ''' novo; linha a menos ele nunca vê, e não sabe que não viu.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_arquivo_nada_esta_dispensado()
        Comigo(Sub(d, pasta)
                   Assert.AreEqual(0, d.Conversas().Count)
                   Assert.IsTrue(d.Remetentes().Vazio)
               End Sub)
    End Sub

End Class

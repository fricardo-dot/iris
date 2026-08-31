Imports System.IO
Imports System.Linq
Imports System.Text
Imports Iris.Assist
Imports Iris.Integration
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>AS REGRAS QUE O DONO ESCREVE — Fase 6.</b>
'''
''' O arquivo é dele: ele abre, lê tudo de uma vez, corrige e apaga. O que este
''' arquivo prende é o que o leitor <b>não</b> pode fazer com o que ele
''' escreveu — porque as regras mudam o que a caixa mostra, e uma regra que o
''' programa reinterpreta sozinho é uma regra que ele não pode auditar.
'''
''' <b>O teto não é aplicado aqui</b>, e isso é uma decisão:
''' <see cref="Onze_regras_voltam_TODAS_as_onze"/>. Cortar na décima
''' classificaria a caixa com dez das onze sem dizer qual ficou de fora.
''' </summary>
' NAO PARALELIZAR: cada teste escreve num arquivo proprio, mas o
' Semear() sem caminho tocaria o %LOCALAPPDATA% do dono.
<TestClass>
<DoNotParallelize>
Public Class RegrasDoDonoTests

    Private _pasta As String

    <TestInitialize>
    Public Sub Antes()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-regras-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_pasta)
    End Sub

    <TestCleanup>
    Public Sub Depois()
        Try
            Directory.Delete(_pasta, recursive:=True)
        Catch
        End Try
    End Sub

    Private Function Arquivo() As String
        Return Path.Combine(_pasta, "regras.txt")
    End Function

    Private Sub Escrever(ParamArray linhas As String())
        File.WriteAllLines(Arquivo(), linhas, Encoding.UTF8)
    End Sub

    ' ==================================================================

    <TestMethod>
    Public Sub Arquivo_que_nao_existe_e_nenhuma_regra()
        Assert.AreEqual(0, New RegrasDoDono(Arquivo()).Ler().Count)
    End Sub

    <TestMethod>
    Public Sub Uma_regra_por_linha_na_ordem_do_dono()
        Escrever("clientes reclamando de atraso",
                 "e-mails sobre boleto")

        Dim lidas = New RegrasDoDono(Arquivo()).Ler()

        Assert.AreEqual(2, lidas.Count)
        Assert.AreEqual("clientes reclamando de atraso", lidas(0))
        Assert.AreEqual("e-mails sobre boleto", lidas(1))
    End Sub

    ''' <summary>
    ''' O cabeçalho semeado é todo comentário, e os exemplos dentro dele estão
    ''' comentados. Um leitor que engolisse <c>#</c> classificaria a caixa do
    ''' dono com três regras que ele nunca escreveu — e ele descobriria pelo
    ''' resultado, que é o pior jeito de descobrir.
    ''' </summary>
    <TestMethod>
    Public Sub O_arquivo_recem_semeado_nao_tem_regra_NENHUMA()
        Dim regras = New RegrasDoDono(Arquivo())

        Assert.IsTrue(regras.Semear())
        Assert.AreEqual(0, regras.Ler().Count)
    End Sub

    <TestMethod>
    Public Sub Semear_NAO_pisa_no_que_ele_escreveu()
        Escrever("a minha regra")

        Assert.IsFalse(New RegrasDoDono(Arquivo()).Semear())
        Assert.AreEqual("a minha regra", New RegrasDoDono(Arquivo()).Ler().Single())
    End Sub

    <TestMethod>
    Public Sub Linha_em_branco_e_espaco_nas_pontas_nao_viram_regra()
        Escrever("  uma regra com espaco  ", "", "   ", "outra")

        Dim lidas = New RegrasDoDono(Arquivo()).Ler()

        Assert.AreEqual(2, lidas.Count)
        Assert.AreEqual("uma regra com espaco", lidas(0))
    End Sub

    ''' <summary>
    ''' <b>O teto mora no lote, e não aqui.</b> A leitura devolve as onze; quem
    ''' recusa é <see cref="LoteDeClassificacao.Preparar"/>, que é quem tem como
    ''' dizer o número a quem chamou. Se este leitor cortasse na décima, a
    ''' décima primeira sumiria em silêncio e o dono veria uma caixa
    ''' classificada por dez regras achando que eram onze.
    ''' </summary>
    <TestMethod>
    Public Sub Onze_regras_voltam_TODAS_as_onze()
        Escrever(Enumerable.Range(1, RegrasDoDono.Maximo + 1).
                 Select(Function(i) "regra " & i).ToArray())

        Assert.AreEqual(RegrasDoDono.Maximo + 1, New RegrasDoDono(Arquivo()).Ler().Count)
        Assert.IsNull(LoteDeClassificacao.Preparar(
            {New Iris.Model.ItemKey("a", "s")},
            New RegrasDoDono(Arquivo()).Ler()))
    End Sub

    ''' <summary>
    ''' <b>Um teto só.</b> Dois números iguais escritos em dois lugares
    ''' divergem, e a divergência apareceria como uma classificação que
    ''' simplesmente não acontece: o dono escreve a décima primeira regra, o
    ''' leitor a entrega, o lote a recusa, e nada na tela explica por quê.
    ''' </summary>
    <TestMethod>
    Public Sub O_teto_e_o_MESMO_dos_dois_lados()
        Assert.AreEqual(LoteDeClassificacao.MaximoDeRegras, RegrasDoDono.Maximo)
    End Sub

    <TestMethod>
    Public Sub Pasta_que_nao_existe_nao_estoura_a_leitura()
        Dim fundo = Path.Combine(_pasta, "nao", "existe", "regras.txt")

        Assert.AreEqual(0, New RegrasDoDono(fundo).Ler().Count)
    End Sub

End Class

Imports Iris.Core
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' O gate de conteudo protegido, e o caso que ele existe para pegar.
'''
''' Antes de ProtectionState existir, Permission era lido por um helper
''' que convertia QUALQUER excecao em zero. "Nao consegui determinar"
''' virava "nao e protegida", e o corpo era lido logo em seguida.
'''
''' O teste que importa e o do Unknown. Um teste que so verificasse
''' Restricted passaria com o codigo antigo tambem.
''' </summary>
<TestClass>
Public Class ProtectionPolicyTests

    <TestMethod>
    Public Sub So_Unprotected_libera_o_corpo()
        Assert.IsTrue(ProtectionPolicy.CanReadBody(ProtectionState.Unprotected))
        Assert.IsFalse(ProtectionPolicy.CanReadBody(ProtectionState.Restricted))
    End Sub

    ''' <summary>
    ''' O caso que o desenho antigo deixava passar: falha de leitura de
    ''' Permission NAO pode virar permissao de leitura.
    ''' </summary>
    <TestMethod>
    Public Sub Unknown_bloqueia_igual_a_Restricted()
        Assert.IsFalse(ProtectionPolicy.CanReadBody(ProtectionState.Unknown),
                       "gate que falha ABERTO e pior que gate nenhum")
        Assert.IsFalse(ProtectionPolicy.CanSendToAi(ProtectionState.Unknown))
    End Sub

    ''' <summary>
    ''' Unknown e o PRIMEIRO valor do enum, entao um DTO que ninguem
    ''' preencheu ja nasce bloqueado. Se alguem reordenar o enum e
    ''' Unprotected virar zero, o default passa a LIBERAR — em silencio.
    ''' </summary>
    <TestMethod>
    Public Sub O_default_do_enum_bloqueia()
        Dim naoPreenchido As ProtectionState = Nothing
        Assert.AreEqual(ProtectionState.Unknown, naoPreenchido)
        Assert.IsFalse(ProtectionPolicy.CanReadBody(naoPreenchido))

        Dim detalhe As New MessageDetail()
        Assert.IsFalse(ProtectionPolicy.CanReadBody(detalhe.Protection),
                       "MessageDetail recem-criado tem de estar bloqueado")
    End Sub

    <TestMethod>
    Public Sub Cada_bloqueio_se_explica()
        Assert.AreNotEqual("", ProtectionPolicy.DescribeBlock(ProtectionState.Restricted))
        Assert.AreNotEqual("", ProtectionPolicy.DescribeBlock(ProtectionState.Unknown))
        Assert.AreNotEqual(ProtectionPolicy.DescribeBlock(ProtectionState.Restricted),
                           ProtectionPolicy.DescribeBlock(ProtectionState.Unknown),
                           "os dois bloqueios tem causas diferentes e mensagens diferentes")
        Assert.AreEqual("", ProtectionPolicy.DescribeBlock(ProtectionState.Unprotected))
    End Sub

End Class

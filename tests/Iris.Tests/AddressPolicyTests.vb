Imports System.Collections.Generic
Imports Iris.Core
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' A regra que decide se um endereço serve para o usuário CONFERIR antes
''' de um envio.
'''
''' Ganhou testes próprios porque estava enterrada como método privado
''' dentro da camada COM, onde nada a alcançava — e é uma regra que decide
''' se uma mensagem sai ou não sai.
'''
''' Roda sem Outlook e sem WPF: é lógica de string pura.
''' </summary>
<TestClass>
Public Class AddressPolicyTests

    <TestMethod>
    Public Sub Endereco_SMTP_comum_passa()
        Assert.IsTrue(AddressPolicy.IsUsableSmtp("fulano@empresa.com"))
        Assert.IsTrue(AddressPolicy.IsUsableSmtp("nome.sobrenome@sub.empresa.com.br"))
        Assert.IsTrue(AddressPolicy.IsUsableSmtp("  espacos@empresa.com  "))
    End Sub

    ''' <summary>
    ''' O caso que motivou a regra. O Outlook resolve nomes internos para
    ''' X.500 e considera resolvido. Para quem lê a tela de confirmação,
    ''' não é: ninguém reconhece um colega nessa string.
    ''' </summary>
    <TestMethod>
    Public Sub Endereco_Exchange_legado_nao_passa()
        Assert.IsFalse(AddressPolicy.IsUsableSmtp(
            "/O=EMPRESA/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=FULANO"))
        Assert.IsFalse(AddressPolicy.IsUsableSmtp("/o=empresa/cn=alguem"))
    End Sub

    <TestMethod>
    Public Sub Nao_passa_o_que_nem_endereco_e()
        Assert.IsFalse(AddressPolicy.IsUsableSmtp(Nothing))
        Assert.IsFalse(AddressPolicy.IsUsableSmtp(""))
        Assert.IsFalse(AddressPolicy.IsUsableSmtp("   "))
        Assert.IsFalse(AddressPolicy.IsUsableSmtp("Fulano de Tal"), "sem arroba")
        Assert.IsFalse(AddressPolicy.IsUsableSmtp("@empresa.com"), "sem nada antes da arroba")
        Assert.IsFalse(AddressPolicy.IsUsableSmtp("fulano@"), "sem domínio")
        Assert.IsFalse(AddressPolicy.IsUsableSmtp("a@b@c.com"), "duas arrobas")
        Assert.IsFalse(AddressPolicy.IsUsableSmtp("fulano@empresa"), "domínio sem ponto")
        Assert.IsFalse(AddressPolicy.IsUsableSmtp("fulano@.com"), "domínio começando com ponto")
        Assert.IsFalse(AddressPolicy.IsUsableSmtp("fulano@empresa."), "domínio terminando em ponto")
    End Sub

    <TestMethod>
    Public Sub Todos_conferiveis_exige_todos()
        Dim bons = Lista(Bom("a@x.com"), Bom("b@x.com"))
        Assert.IsTrue(AddressPolicy.AllRecipientsUsable(bons))

        ' Um único ruim reprova o conjunto: mandar para nove certos e um
        ' errado continua sendo mandar para o errado.
        Dim comUmRuim = Lista(Bom("a@x.com"), Legado("b"))
        Assert.IsFalse(AddressPolicy.AllRecipientsUsable(comUmRuim))
    End Sub

    ''' <summary>
    ''' Resolvido pelo Outlook E endereço conferível são coisas diferentes.
    ''' O caso perigoso é o que diz "resolvido" e entrega /O=...
    ''' </summary>
    <TestMethod>
    Public Sub Resolvido_com_endereco_ilegivel_nao_conta_como_conferivel()
        ' Nao chamar de "legado" nem de "lista": VB e case-insensitive e o
        ' nome eclipsaria as funcoes Legado() e Lista() desta classe.
        Dim so_x500 = Lista(Legado("fulano"))
        Assert.IsTrue(so_x500(0).Resolved, "o cenário é justamente este: o Outlook diz que resolveu")
        Assert.IsFalse(AddressPolicy.AllRecipientsUsable(so_x500))
        Assert.AreEqual(1, AddressPolicy.Unusable(so_x500).Count)
    End Sub

    <TestMethod>
    Public Sub Nao_resolvido_nao_conta_mesmo_com_endereco_bonito()
        Dim pendentes = Lista(New RecipientInfo With {
            .DisplayName = "Fulano", .Address = "fulano@empresa.com",
            .Kind = RecipientKind.To, .Resolved = False})

        Assert.IsFalse(AddressPolicy.AllRecipientsUsable(pendentes))
    End Sub

    ''' <summary>
    ''' Lista vazia não é "todos conferíveis": não há para quem mandar.
    ''' Devolver True aqui liberaria o envio de uma mensagem sem
    ''' destinatário nenhum.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_destinatario_nao_e_conferivel()
        Assert.IsFalse(AddressPolicy.AllRecipientsUsable(New List(Of RecipientInfo)()))
        Assert.IsFalse(AddressPolicy.AllRecipientsUsable(Nothing))
    End Sub

    <TestMethod>
    Public Sub Os_reprovados_sao_nomeados()
        Dim mistos = Lista(Bom("a@x.com"), Legado("bruno"), Bom("c@x.com"))
        Dim ruins = AddressPolicy.Unusable(mistos)

        Assert.AreEqual(1, ruins.Count)
        Assert.AreEqual("bruno", ruins(0).DisplayName,
            "A mensagem de erro precisa dizer QUAL destinatário, senão o usuário caça no escuro.")
    End Sub

    ' ----------------------------------------------------------------

    Private Shared Function Lista(ParamArray itens As RecipientInfo()) As List(Of RecipientInfo)
        Return New List(Of RecipientInfo)(itens)
    End Function

    Private Shared Function Bom(endereco As String) As RecipientInfo
        Return New RecipientInfo With {
            .DisplayName = endereco, .Address = endereco,
            .Kind = RecipientKind.To, .Resolved = True}
    End Function

    Private Shared Function Legado(nome As String) As RecipientInfo
        Return New RecipientInfo With {
            .DisplayName = nome,
            .Address = "/O=EMPRESA/OU=GRUPO/CN=RECIPIENTS/CN=" & nome,
            .Kind = RecipientKind.To, .Resolved = True}
    End Function

End Class

Imports Iris.App.ViewModels
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A marcação sai; a interpretação não entra.</b>
'''
''' ------------------------------------------------------------------
''' O modelo escreve em Markdown por hábito, e a faixa mostra a resposta num
''' <c>TextBlock</c> que não interpreta nada — então <c>**Marta:**</c>
''' aparecia com os asteriscos na tela.
'''
''' A saída óbvia seria renderizar o Markdown. Estes testes existem para
''' guardar a saída <b>escolhida</b>: só se apaga caractere, e o resultado
''' continua sendo uma string passiva.
''' </summary>
<TestClass>
Public Class TextoDoModeloTests

    ''' <summary>O caso que o usuário viu na tela.</summary>
    <TestMethod>
    Public Sub Os_asteriscos_do_modelo_SOMEM()
        Dim bruto = "* **Marta:** Ficará responsável pelo orçamento." & vbLf &
                    "* **Túlio:** Fará o levantamento."

        Dim limpo = TextoDoModelo.Limpar(bruto)

        Assert.IsFalse(limpo.Contains("*"), limpo)
        StringAssert.Contains(limpo, "• Marta: Ficará responsável pelo orçamento.")
        StringAssert.Contains(limpo, "• Túlio: Fará o levantamento.")
    End Sub

    ''' <summary>
    ''' <b>A lista é tratada antes da ênfase.</b>
    '''
    ''' Uma linha <c>* **Marta:** …</c> começa com um asterisco de lista seguido
    ''' de um par de ênfase. Tirar a ênfase primeiro deixaria <c>* Marta: …</c>,
    ''' e o marcador de lista viraria texto solto no começo da frase.
    ''' </summary>
    <TestMethod>
    Public Sub A_lista_vem_antes_da_enfase()
        Assert.AreEqual("• Marta", TextoDoModelo.Limpar("* **Marta**"))
    End Sub

    ''' <summary>Títulos, crases e ênfase simples também saem.</summary>
    <TestMethod>
    Public Sub Titulo_crase_e_enfase_simples_saem()
        Assert.AreEqual("Resumo", TextoDoModelo.Limpar("## Resumo"))
        Assert.AreEqual("o campo prazo", TextoDoModelo.Limpar("o campo `prazo`"))
        Assert.AreEqual("bem urgente", TextoDoModelo.Limpar("bem *urgente*"))
    End Sub

    ''' <summary>
    ''' <b>Asterisco solto continua asterisco.</b>
    '''
    ''' Ele pode ser do texto: "custa 3 * 4 reais", uma nota de rodapé, um
    ''' campo obrigatório. Apagar todo asterisco mudaria o que o modelo disse
    ''' em vez de limpar como ele disse.
    ''' </summary>
    <TestMethod>
    Public Sub Asterisco_SOLTO_fica()
        Assert.AreEqual("3 * 4", TextoDoModelo.Limpar("3 * 4"))
        Assert.AreEqual("campo obrigatório *", TextoDoModelo.Limpar("campo obrigatório *"))
    End Sub

    ''' <summary>
    ''' <b>Um par que atravessa linha não come o meio.</b>
    '''
    ''' Dois asteriscos soltos em parágrafos diferentes não são um par. Sem a
    ''' âncora de linha na expressão, tudo entre eles sumiria — inclusive
    ''' conteúdo do resumo.
    ''' </summary>
    <TestMethod>
    Public Sub Par_que_atravessa_LINHA_nao_e_par()
        Dim r = TextoDoModelo.Limpar("nota * um" & vbLf & "outra * coisa")

        StringAssert.Contains(r, "um")
        StringAssert.Contains(r, "outra")
    End Sub

    ''' <summary>
    ''' <b>Isca: a limpeza nunca ACRESCENTA conteúdo.</b>
    '''
    ''' O controle que guarda a fronteira inteira. Se um dia alguém trocar
    ''' "apagar marcador" por "renderizar", vai ser tentador fazer a função
    ''' produzir marcação própria — e é aí que texto de terceiro vira árvore
    ''' visual. A única coisa que entra e não estava é o marcador de lista.
    ''' </summary>
    <TestMethod>
    Public Sub A_limpeza_nao_ACRESCENTA_nada()
        Dim bruto = "# T" & vbLf & "* **a** `b` *c*" & vbLf & vbLf & vbLf & "fim"

        Dim limpo = TextoDoModelo.Limpar(bruto)

        For Each ch In limpo
            If ch = "•"c OrElse ch = " "c OrElse ch = vbCr OrElse ch = vbLf Then Continue For
            Assert.IsTrue(bruto.Contains(ch),
                $"o caractere '{ch}' apareceu do nada")
        Next
    End Sub

    ''' <summary>Quebras do Windows não escapam da âncora de linha.</summary>
    <TestMethod>
    Public Sub Quebra_do_Windows_tambem_conta()
        Assert.AreEqual("• um" & Environment.NewLine & "• dois",
                        TextoDoModelo.Limpar("* um" & vbCrLf & "* dois"))
    End Sub

    <TestMethod>
    Public Sub Vazio_e_nulo_nao_explodem()
        Assert.AreEqual("", TextoDoModelo.Limpar(Nothing))
        Assert.AreEqual("", TextoDoModelo.Limpar(""))
        Assert.AreEqual("", TextoDoModelo.Limpar("   " & vbLf & "  "))
    End Sub

End Class

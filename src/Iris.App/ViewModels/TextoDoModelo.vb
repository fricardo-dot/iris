Imports System.Text
Imports System.Text.RegularExpressions

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>Tira a marcação da resposta do modelo — sem interpretá-la.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>REMOVER NÃO É RENDERIZAR, E A DIFERENÇA É A GARANTIA INTEIRA</b>
    '''
    ''' O modelo escreve em Markdown por hábito, e a faixa mostra a resposta
    ''' num <c>TextBlock</c> que não interpreta nada — então
    ''' <c>**Marta:**</c> aparecia com os asteriscos na tela.
    '''
    ''' A saída óbvia seria renderizar o Markdown. <b>Não é o que isto faz.</b>
    ''' A resposta veio de um lugar que leu o e-mail, que por sua vez veio de
    ''' fora; interpretar marcação transformaria texto de terceiro em árvore
    ''' visual, e o dia em que alguém acrescentasse link ou imagem à gramática
    ''' — porque "já que interpretamos negrito" — o e-mail teria ganhado um
    ''' jeito de fazer o Iris buscar coisa na rede.
    '''
    ''' Aqui só se <b>apaga</b> caractere. A saída continua sendo uma string
    ''' passiva, e o <c>TextBlock</c> continua sendo um <c>TextBlock</c>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E O QUE SAI DAQUI É O QUE VAI PARA TODO LUGAR</b>
    '''
    ''' O texto limpo é o que aparece na tela, o que o botão Copiar leva, e o
    ''' que a redação escreve no rascunho. Mostrar uma coisa e copiar outra
    ''' seria pior que os asteriscos.
    ''' </summary>
    Public Module TextoDoModelo

        ''' <summary>Marcador de lista no começo da linha: <c>*</c>, <c>-</c>, <c>+</c>.</summary>
        Private ReadOnly Lista As New Regex("^([ \t]*)[*+\-][ \t]+",
                                            RegexOptions.Multiline Or RegexOptions.Compiled)

        ''' <summary>Marcador de título: <c>#</c> a <c>######</c> no começo da linha.</summary>
        Private ReadOnly Titulo As New Regex("^[ \t]*#{1,6}[ \t]+",
                                             RegexOptions.Multiline Or RegexOptions.Compiled)

        ''' <summary>
        ''' Ênfase forte. <c>[^\n]</c> e não <c>.</c>: um par que atravessasse
        ''' linha comeria o texto entre dois asteriscos soltos de parágrafos
        ''' diferentes.
        ''' </summary>
        Private ReadOnly Forte As New Regex("\*\*([^\n]+?)\*\*", RegexOptions.Compiled)

        ''' <summary>Ênfase simples, depois da forte — senão <c>**x**</c> vira <c>*x*</c>.</summary>
        Private ReadOnly Simples As New Regex("\*([^*\n]+?)\*", RegexOptions.Compiled)

        ''' <summary>Código entre crases.</summary>
        Private ReadOnly Crase As New Regex("`+([^`\n]*?)`+", RegexOptions.Compiled)

        ''' <summary>Três ou mais linhas em branco viram duas.</summary>
        Private ReadOnly Vazias As New Regex("\n{3,}", RegexOptions.Compiled)

        ''' <summary>
        ''' Limpa a marcação. <b>Só apaga e substitui</b> — nunca acrescenta
        ''' conteúdo, e a única coisa que entra que não estava é o
        ''' <c>•</c> no lugar do marcador de lista.
        ''' </summary>
        Public Function Limpar(bruto As String) As String
            If String.IsNullOrEmpty(bruto) Then Return ""

            ' Quebras normalizadas antes de qualquer ancora de linha: com
            ' vbCrLf, o "^" da Multiline casa depois do \n e o \r sobra no
            ' comeco, fazendo o marcador de lista escapar da substituicao.
            Dim t = bruto.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)

            ' LISTA ANTES DA ENFASE. Uma linha "* **Marta:** ..." comeca com
            ' um asterisco de lista seguido de um par de enfase; tirar a enfase
            ' primeiro deixaria "* Marta: ..." e o marcador de lista viraria
            ' texto solto no comeco da frase.
            t = Lista.Replace(t, "$1• ")
            t = Titulo.Replace(t, "")

            t = Forte.Replace(t, "$1")
            t = Simples.Replace(t, "$1")
            t = Crase.Replace(t, "$1")

            t = Vazias.Replace(t, vbLf & vbLf)

            ' Espaco no fim da linha nao aparece, mas conta na largura da
            ' quebra automatica.
            Dim linhas = t.Split(CChar(vbLf))
            For i = 0 To linhas.Length - 1
                linhas(i) = linhas(i).TrimEnd()
            Next

            Return String.Join(Environment.NewLine, linhas).Trim()
        End Function

    End Module

End Namespace

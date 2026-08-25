Imports System.Collections.Generic
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Iris.App.ViewModels
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' A faixa da IA <b>renderiza o texto</b> — não só resolve o binding.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ISTO EXISTE</b>
'''
''' O <c>BindingsDaJanelaTests</c> prova que <c>Assistente.Aviso</c> resolve
''' numa propriedade que existe, e que o binding está no XAML. Isso é bem menos
''' do que "o aviso aparece": binding correto não detecta faixa com altura
''' zero, controle colapsado, ou dois elementos disputando a mesma linha do
''' <c>Grid</c>.
'''
''' E aqui isso importa mais que na Fase 2: <b>o aviso e o resultado ocupam a
''' mesma <c>Grid.Row</c></b>. Se os dois ficassem visíveis ao mesmo tempo, um
''' cobriria o outro — e o que sumiria seria justamente a frase que diz que
''' pode ter saído conteúdo.
'''
''' Nada aqui abre o Outlook nem mostra janela: <c>Measure</c> e
''' <c>Arrange</c> fora do vídeo.
''' </summary>
<TestClass>
Public Class FaixaDaIaRenderizaTests

    ''' <summary>
    ''' Monta a faixa com a <b>mesma estrutura</b> do <c>MainWindow.xaml</c>:
    ''' duas bordas, mesma linha, visibilidade por binding.
    '''
    ''' Reconstruir em vez de instanciar a janela inteira: ela exige um
    ''' <c>MainViewModel</c>, que exige um broker, que exige o Outlook. O preço
    ''' é este teste não pegar mudanças no XAML — e é por isso que ele anda
    ''' junto do <c>BindingsDaJanelaTests</c>, que lê o XAML de verdade. Um
    ''' cobre o que o outro não alcança.
    ''' </summary>
    Private Shared Function Montar(aviso As String, temAlgoADizer As Boolean,
                                   resultado As String) As Grid
        Dim linha As New Grid()
        linha.RowDefinitions.Add(New RowDefinition With {.Height = GridLength.Auto})

        Dim dados As New FaixaFalsa(aviso, temAlgoADizer, resultado)

        linha.Children.Add(Borda(New TextBlock(), TextBlock.TextProperty, "Aviso",
                                 "TemAlgoADizer", dados))
        linha.Children.Add(Borda(New TextBlock(), TextBlock.TextProperty, "Resultado",
                                 "TemResultado", dados))

        ' DUAS passadas, e nao uma. A primeira mede antes de os bindings de
        ' Visibility terem sido aplicados, entao ela ve os dois filhos visiveis
        ' e o Grid sai com altura de quem nao tem nada a mostrar. Medir de novo
        ' depois do UpdateLayout e o que faz o teste falar do que a tela mostra,
        ' e nao de um estado intermediario.
        For passada = 1 To 2
            linha.Measure(New Size(900, 600))
            linha.Arrange(New Rect(0, 0, 900, linha.DesiredSize.Height))
            linha.UpdateLayout()
        Next
        Return linha
    End Function

    Private Shared Function Borda(alvo As TextBlock, prop As DependencyProperty,
                                  caminho As String, visivelSe As String,
                                  dados As Object) As Border
        alvo.TextWrapping = TextWrapping.Wrap
        alvo.FontSize = 12
        alvo.SetBinding(prop, New Data.Binding(caminho))

        Dim b As New Border With {.Padding = New Thickness(16, 8, 16, 8), .Child = alvo}
        b.SetBinding(UIElement.VisibilityProperty,
                     New Data.Binding(visivelSe) With {
                         .Converter = New BooleanToVisibilityConverter()})
        b.DataContext = dados
        Grid.SetRow(b, 0)
        Return b
    End Function

    Private NotInheritable Class FaixaFalsa
        Public ReadOnly Property Aviso As String
        Public ReadOnly Property TemAlgoADizer As Boolean
        Public ReadOnly Property Resultado As String
        Public ReadOnly Property TemResultado As Boolean

        Public Sub New(aviso As String, temAlgoADizer As Boolean, resultado As String)
            Me.Aviso = aviso
            Me.TemAlgoADizer = temAlgoADizer
            Me.Resultado = resultado
            TemResultado = resultado.Length > 0
        End Sub
    End Class

    ' ==================================================================

    ''' <summary>
    ''' O motivo de a IA estar desligada <b>aparece na tela</b>, e com altura.
    ''' </summary>
    <STATestMethod>
    Public Sub O_motivo_da_IA_desligada_APARECE()
        Const aviso = "A IA externa não está habilitada."
        Dim faixa = Montar(aviso, True, "")

        StringAssert.Contains(TextoVisivel(faixa), aviso)
        Assert.IsTrue(faixa.ActualHeight > 0, "faixa com texto nao pode ter altura zero")
    End Sub

    ''' <summary>
    ''' <b>Sem nada a dizer, a faixa não ocupa espaço.</b>
    '''
    ''' O contraponto, e sem ele o teste acima passaria numa faixa que mostra
    ''' tudo sempre — inclusive vazia. Faixa vazia é ruído que ensina o usuário
    ''' a ignorar aquele lugar da tela.
    ''' </summary>
    <STATestMethod>
    Public Sub Sem_nada_a_dizer_a_faixa_nao_ocupa_espaco()
        Assert.AreEqual(0.0, Montar("", False, "").ActualHeight, 0.01)
    End Sub

    ''' <summary>
    ''' <b>O aviso e o resultado não se cobrem.</b>
    '''
    ''' Eles ocupam a mesma <c>Grid.Row</c>, e num <c>Grid</c> isso significa
    ''' <b>empilhados</b>. Com os dois visíveis, o de baixo cobriria o de cima —
    ''' e o que sumiria seria justamente a frase que diz que pode ter saído
    ''' conteúdo.
    '''
    ''' O ViewModel evita isso: com resultado, o aviso fica vazio. Este teste é
    ''' a prova de que a <b>consequência visual</b> é a esperada — se algum dia
    ''' os dois ficarem visíveis juntos, ele acusa.
    ''' </summary>
    <STATestMethod>
    Public Sub Aviso_e_resultado_juntos_SE_COBREM()
        Dim faixa = Montar("A operação não terminou.", True, "resumo do modelo")

        Dim caixas = Retangulos(faixa).ToList()
        Assert.AreEqual(2, caixas.Count, "os dois estao visiveis — e o teste e sobre isso")
        Assert.IsTrue(caixas(0).IntersectsWith(caixas(1)),
            "eles ocupam a mesma linha do Grid: se um dia os dois aparecerem " &
            "juntos, um cobre o outro, e isso tem de ser uma decisao e nao um acidente")
    End Sub

    ''' <summary>
    ''' E o caso que a produção produz — resultado sem aviso — mostra <b>só</b>
    ''' o resultado.
    ''' </summary>
    <STATestMethod>
    Public Sub Com_resultado_e_sem_aviso_so_o_resultado_aparece()
        Dim faixa = Montar("", False, "resumo do modelo")

        Dim texto = TextoVisivel(faixa)
        StringAssert.Contains(texto, "resumo do modelo")
        Assert.AreEqual(1, Retangulos(faixa).Count(), "so uma borda visivel")
    End Sub

    ''' <summary>
    ''' Controle: a leitura da árvore visual <b>realmente</b> encontra o texto.
    '''
    ''' Sem isto, um <c>TextoVisivel</c> que devolvesse sempre vazio faria o
    ''' contraponto passar e o principal falhar de um jeito que eu poderia
    ''' "consertar" afrouxando a asserção.
    ''' </summary>
    <STATestMethod>
    Public Sub Controle_a_leitura_da_arvore_encontra_texto_plantado()
        StringAssert.Contains(TextoVisivel(Montar("ISCA-VISIVEL-77", True, "")),
                              "ISCA-VISIVEL-77")
    End Sub

    ' ==================================================================

    Private Shared Function TextoVisivel(raiz As DependencyObject) As String
        Dim partes As New List(Of String)()
        Colher(raiz, partes)
        Return String.Join(" ", partes)
    End Function

    Private Shared Sub Colher(no As DependencyObject, partes As List(Of String))
        Dim ui = TryCast(no, UIElement)
        If ui IsNot Nothing AndAlso ui.Visibility <> Visibility.Visible Then Return

        Dim tb = TryCast(no, TextBlock)
        If tb IsNot Nothing AndAlso Not String.IsNullOrEmpty(tb.Text) Then partes.Add(tb.Text)

        For i = 0 To VisualTreeHelper.GetChildrenCount(no) - 1
            Colher(VisualTreeHelper.GetChild(no, i), partes)
        Next
    End Sub

    ''' <summary>Onde cada borda visível ficou, depois do arranjo.</summary>
    Private Shared Iterator Function Retangulos(raiz As Grid) As IEnumerable(Of Rect)
        For Each filho As UIElement In raiz.Children
            If filho.Visibility <> Visibility.Visible Then Continue For
            Dim fe = CType(filho, FrameworkElement)
            Dim canto = fe.TranslatePoint(New Point(0, 0), raiz)
            Yield New Rect(canto, New Size(fe.ActualWidth, fe.ActualHeight))
        Next
    End Function

End Class

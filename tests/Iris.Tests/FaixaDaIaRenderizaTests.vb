Imports System.Collections.Generic
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' A faixa da IA <b>de verdade</b>, medida e arranjada fora do vídeo.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE A FAIXA VIROU UM <c>UserControl</c></b>
'''
''' A primeira versão deste arquivo reconstruía uma <i>imitação</i> da faixa —
''' duas bordas montadas à mão, com os mesmos bindings. E por isso não pegou o
''' defeito mais óbvio que havia: <b>os botões estavam dentro da borda de
''' aviso</b>, cuja visibilidade é <c>TemAlgoADizer</c>. Com a IA funcionando —
''' sem aviso e sem ambiguidades — essa borda colapsa e leva os botões junto.
''' Ou seja, eles sumiam exatamente quando estavam habilitados.
'''
''' Uma imitação prova o que quem a escreveu já sabia. Aqui o controle é
''' instanciado de verdade, com o XAML que a janela usa.
'''
''' Nada disto abre o Outlook nem mostra janela: o <c>UserControl</c> só tem
''' bindings, e o teste faz <c>Measure</c> e <c>Arrange</c>.
''' </summary>
' NAO PARALELIZAR: esta classe carrega XAML de verdade.
'
' Application.LoadComponent le o baml do pacote de recursos do assembly, e
' System.IO.Packaging.PackagePart NAO e thread-safe --
' CleanUpRequestedStreamsList estoura NullReferenceException quando duas
' threads STA carregam o mesmo pacote ao mesmo tempo. Com
' Parallelize(MethodLevel) no assembly, os oito metodos daqui corriam juntos.
'
' O sintoma era intermitente: a mesma suite deu 5 falhas, depois 7, depois 0,
' sem mudanca nenhuma no codigo. Teste que as vezes passa nao prova nada, e
' pior: gasta a confianca do numero verde que ele mesmo produz.
'
' AdversarioPontaAPontaTests e FaixaDoAcervoRenderizaTests ja tinham o
' atributo. Esta era a unica que carregava XAML sem ele.
'
' (Comentario ACIMA dos atributos: entre <Attr> e a declaracao, o VB exige
' continuacao de linha e o compilador reclama de "attribute specifier is not
' a complete statement".)
<TestClass>
<DoNotParallelize>
Public Class FaixaDaIaRenderizaTests

    ''' <summary>
    ''' <b>Veste a faixa com os dicionarios de tema de verdade.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE PRECISA, E POR QUE DEPOIS DO CONSTRUTOR</b>
    '''
    ''' Em produção os estilos moram em <c>Application.Resources</c>. No
    ''' processo de teste não existe <c>Application</c>, então nada disso é
    ''' alcançável — e sem vestir, a faixa renderiza com o visual padrão do
    ''' WPF, que <b>não é o que o usuário vê</b>. Um teste de renderização
    ''' sobre um visual que não existe em lugar nenhum não é teste de nada.
    '''
    ''' Funciona depois do construtor porque a faixa usa <c>DynamicResource</c>:
    ''' ele resolve quando o dicionário aparece. Com <c>StaticResource</c> a
    ''' faixa nem carregaria aqui — e foi por isso que ela passou a usar
    ''' <c>DynamicResource</c> em tudo.
    ''' </summary>
    Private Shared Sub Vestir(faixa As FrameworkElement)
        ' O MESMO arquivo que o Application.xaml carrega. Uma lista propria
        ' aqui podia divergir da do programa, e o teste passaria a provar uma
        ' aparencia que nao existe em lugar nenhum.
        faixa.Resources.MergedDictionaries.Add(New ResourceDictionary() With {
            .Source = New Uri("pack://application:,,,/Iris;component/Themes/Tema.xaml")})
    End Sub

    ''' <summary>
    ''' A faixa real, com um contexto de dados que imita o ViewModel.
    '''
    ''' <b>Duas passadas</b> de layout: a primeira mede antes de os bindings de
    ''' <c>Visibility</c> terem sido aplicados, e o controle sai com altura de
    ''' quem não tem nada a mostrar.
    ''' </summary>
    Private Shared Function Montar(aviso As String, temAlgoADizer As Boolean,
                                   resultado As String,
                                   Optional podePedir As Boolean = False) As FrameworkElement
        Dim faixa As New Iris.App.Views.FaixaDaIa()
        Vestir(faixa)
        faixa.DataContext = New FaixaFalsa(aviso, temAlgoADizer, resultado, podePedir)

        For passada = 1 To 2
            faixa.Measure(New Size(900, 600))
            faixa.Arrange(New Rect(0, 0, 900, faixa.DesiredSize.Height))
            faixa.UpdateLayout()
        Next
        Return faixa
    End Function

    ''' <summary>
    ''' O contexto de dados. <b>Não</b> é o ViewModel: ele exigiria um
    ''' transmissor, que exige um diário, que exige um banco.
    '''
    ''' O que este teste cobra é a <b>estrutura visual</b>; que o ViewModel
    ''' produza estes valores é assunto do <c>AssistenteViewModelTests</c>, e
    ''' que o XAML da janela use estes nomes é do
    ''' <c>BindingsDaJanelaTests</c>. Três testes, três perguntas.
    ''' </summary>
    Private NotInheritable Class FaixaFalsa
        Public ReadOnly Property Aviso As String
        Public ReadOnly Property TemAlgoADizer As Boolean
        Public ReadOnly Property Resultado As String
        Public ReadOnly Property TemResultado As Boolean
        Public ReadOnly Property ResumirCommand As Input.ICommand
        Public ReadOnly Property RedigirCommand As Input.ICommand
        Public ReadOnly Property DesfazerCommand As Input.ICommand
        Public ReadOnly Property CancelarCommand As Input.ICommand
        Public ReadOnly Property Reconciliacao As Object

        Public Sub New(aviso As String, temAlgoADizer As Boolean, resultado As String,
                       podePedir As Boolean)
            Me.Aviso = aviso
            Me.TemAlgoADizer = temAlgoADizer
            Me.Resultado = resultado
            TemResultado = resultado.Length > 0
            ResumirCommand = New ComandoFalso(podePedir)
            RedigirCommand = New ComandoFalso(podePedir)
            DesfazerCommand = New ComandoFalso(False)
            CancelarCommand = New ComandoFalso(False)
            Reconciliacao = New With {.Aviso = ""}
        End Sub
    End Class

    Private NotInheritable Class ComandoFalso
        Implements Input.ICommand

        Private ReadOnly _pode As Boolean
        Public Sub New(pode As Boolean)
            _pode = pode
        End Sub

        Public Custom Event CanExecuteChanged As EventHandler _
                            Implements Input.ICommand.CanExecuteChanged
            AddHandler(value As EventHandler)
            End AddHandler
            RemoveHandler(value As EventHandler)
            End RemoveHandler
            RaiseEvent(sender As Object, e As EventArgs)
            End RaiseEvent
        End Event

        Public Function CanExecute(parameter As Object) As Boolean _
                                   Implements Input.ICommand.CanExecute
            Return _pode
        End Function

        Public Sub Execute(parameter As Object) Implements Input.ICommand.Execute
        End Sub
    End Class

    ''' <summary>
    ''' Roda o corpo numa thread <b>STA</b> com <c>Dispatcher</c> de verdade.
    '''
    ''' Os outros testes deste arquivo medem e arranjam sem tocar em
    ''' <c>IsEnabled</c>, e por isso passam em MTA. Ler <c>Button.IsEnabled</c>
    ''' constrói o <c>InputManager</c>, que exige STA.
    '''
    ''' O <c>Dispatcher</c> não é decoração: ele instala o
    ''' <c>SynchronizationContext</c> que traz as continuações do <c>Await</c> de
    ''' volta para <b>esta</b> thread. Sem isso, o <c>CanExecuteChanged</c> seria
    ''' levantado numa thread do pool, e o <c>Button</c> — que é
    ''' <c>DispatcherObject</c> — recusaria a visita.
    ''' </summary>
    Friend Shared Sub NaSTA(corpo As Func(Of Global.System.Threading.Tasks.Task))
        Dim erro As Exception = Nothing
        Dim t As New Global.System.Threading.Thread(
            Sub()
                Dim d = Global.System.Windows.Threading.Dispatcher.CurrentDispatcher
                d.BeginInvoke(
                    Async Sub()
                        Try
                            Await corpo()
                        Catch ex As Exception
                            erro = ex
                        Finally
                            d.InvokeShutdown()
                        End Try
                    End Sub)
                Global.System.Windows.Threading.Dispatcher.Run()
            End Sub)
        t.SetApartmentState(Global.System.Threading.ApartmentState.STA)
        t.Start()
        Assert.IsTrue(t.Join(TimeSpan.FromSeconds(30)), "a thread STA nao terminou")
        If erro IsNot Nothing Then Throw erro
    End Sub

    ' ==================================================================
    ' O botão de verdade, ligado ao ViewModel de verdade

    ''' <summary>
    ''' <b>O botão "Desfazer" da faixa real acompanha o estado do rascunho.</b>
    '''
    ''' Os outros testes deste arquivo usam <see cref="FaixaFalsa"/>, porque o
    ''' que eles cobram é estrutura visual. Este é diferente: o que se quer
    ''' provar é que a <b>notificação</b> do ViewModel chega ao
    ''' <c>Button.IsEnabled</c> — e para isso os dois lados têm de ser os de
    ''' verdade.
    '''
    ''' O defeito que ele fecha: <c>PodeDesfazer</c> passou a recusar quando o
    ''' usuário digita por cima da redação, e o <c>RelayCommand</c> não se
    ''' reconsulta sozinho. Sem alguém avisar, o botão continuaria habilitado
    ''' mostrando um estado que já não existe. Perguntar
    ''' <c>DesfazerCommand.CanExecute</c> não pegaria isso: a resposta estaria
    ''' certa e o botão errado.
    ''' </summary>
    <TestMethod>
    Public Sub O_botao_DESFAZER_real_cai_quando_o_usuario_digita()
        NaSTA(Async Function() As Global.System.Threading.Tasks.Task
        Dim rascunho As New AssistenteViewModelTests.RascunhoFalso() With {.Texto = "o que eu escrevi antes"}
        Dim provedor As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "resposta redigida pela IA"}
        Dim vm = AssistenteViewModelTests.Montar(
            AssistenteViewModelTests.Ativacao(), provedor,
            AssistenteViewModelTests.Pronta(), Nothing, rascunho)

        Dim faixa As New Iris.App.Views.FaixaDaIa()
        faixa.DataContext = vm
        For passada = 1 To 2
            faixa.Measure(New Size(900, 600))
            faixa.Arrange(New Rect(0, 0, 900, faixa.DesiredSize.Height))
            faixa.UpdateLayout()
        Next

        Dim desfazer = Botoes(faixa).Single(Function(b) CStr(b.Content) = "Desfazer")
        Assert.IsFalse(desfazer.IsEnabled, "antes da redacao nao ha o que desfazer")

        Await vm.RedigirCommand.ExecuteAsync(Nothing)
        faixa.UpdateLayout()
        Assert.IsTrue(desfazer.IsEnabled, "depois da redacao o botao tem de estar de pe")

        ' O usuário digita por cima da redação.
        rascunho.Texto = "resposta redigida pela IA, com o meu final"
        faixa.UpdateLayout()

        Assert.IsFalse(desfazer.IsEnabled,
            "o botao continuou habilitado mostrando um estado que ja nao existe")
                  End Function)
    End Sub

    ''' <summary>
    ''' Controle negativo: sem digitar, o mesmo botão real continua habilitado.
    '''
    ''' Sem ele, uma faixa cujo botão nunca habilitasse — ou um comando que
    ''' recusasse sempre — passaria no teste de cima.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_o_botao_DESFAZER_real_fica_de_pe_sem_edicao()
        NaSTA(Async Function() As Global.System.Threading.Tasks.Task
        Dim rascunho As New AssistenteViewModelTests.RascunhoFalso() With {.Texto = "o que eu escrevi antes"}
        Dim provedor As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "resposta redigida pela IA"}
        Dim vm = AssistenteViewModelTests.Montar(
            AssistenteViewModelTests.Ativacao(), provedor,
            AssistenteViewModelTests.Pronta(), Nothing, rascunho)

        Dim faixa As New Iris.App.Views.FaixaDaIa()
        faixa.DataContext = vm
        For passada = 1 To 2
            faixa.Measure(New Size(900, 600))
            faixa.Arrange(New Rect(0, 0, 900, faixa.DesiredSize.Height))
            faixa.UpdateLayout()
        Next

        Await vm.RedigirCommand.ExecuteAsync(Nothing)
        faixa.UpdateLayout()

        Dim desfazer = Botoes(faixa).Single(Function(b) CStr(b.Content) = "Desfazer")
        Assert.IsTrue(desfazer.IsEnabled)
                  End Function)
    End Sub

    ' ==================================================================
    ' Os botões

    ''' <summary>
    ''' <b>Os botões aparecem quando a IA está FUNCIONANDO.</b>
    '''
    ''' Este é o teste que a imitação não fazia, e é o que teria pego o defeito:
    ''' os botões moravam dentro da borda de aviso, e com a IA funcionando —
    ''' sem aviso e sem ambiguidades — a borda colapsa e leva os botões junto.
    ''' </summary>
    <STATestMethod>
    Public Sub Os_botoes_aparecem_com_a_IA_FUNCIONANDO()
        Dim faixa = Montar(aviso:="", temAlgoADizer:=False, resultado:="", podePedir:=True)

        Dim rotulos = Botoes(faixa).Select(Function(b) CStr(b.Content)).ToList()
        CollectionAssert.Contains(rotulos, "Resumir",
            "o botao sumiu exatamente quando estava habilitado")
        Assert.IsTrue(faixa.ActualHeight > 0)
    End Sub

    ''' <summary>
    ''' E aparecem <b>também</b> com a IA desligada — desabilitados.
    '''
    ''' Um botão que some esconderia a funcionalidade e o motivo dela estar
    ''' desligada, e o motivo é o que o usuário precisa ler no lugar onde
    ''' procuraria a ação.
    ''' </summary>
    <STATestMethod>
    Public Sub Os_botoes_aparecem_DESABILITADOS_com_a_IA_desligada()
        Dim faixa = Montar("A IA externa não está habilitada.", True, "", podePedir:=False)

        Dim resumir = Botoes(faixa).First(Function(b) CStr(b.Content) = "Resumir")
        Assert.IsFalse(resumir.IsEnabled, "visivel, e desabilitado")
    End Sub

    ' ==================================================================
    ' O aviso e o resultado

    <STATestMethod>
    Public Sub O_motivo_da_IA_desligada_APARECE()
        Const aviso = "A IA externa não está habilitada."

        StringAssert.Contains(TextoVisivel(Montar(aviso, True, "")), aviso)
    End Sub

    ''' <summary>
    ''' <b>Sem nada a dizer, o aviso não aparece.</b>
    '''
    ''' O contraponto: sem ele, o teste acima passaria numa faixa que mostra
    ''' tudo sempre — inclusive vazia. Faixa vazia é ruído que ensina o usuário
    ''' a ignorar aquele lugar da tela.
    ''' </summary>
    <STATestMethod>
    Public Sub Sem_nada_a_dizer_o_aviso_nao_aparece()
        Dim texto = TextoVisivel(Montar("ISCA-QUE-NAO-DEVIA-APARECER", False, ""))

        Assert.IsFalse(texto.Contains("ISCA-QUE-NAO-DEVIA-APARECER"),
                       "texto de aviso com TemAlgoADizer falso nao pode aparecer")
    End Sub

    ''' <summary>
    ''' <b>O aviso e o resultado não se cobrem.</b>
    '''
    ''' Eles moravam na mesma <c>Grid.Row</c> — num <c>Grid</c> isso significa
    ''' empilhados, e o de baixo cobriria o de cima. O que sumiria seria
    ''' justamente a frase que diz que pode ter saído conteúdo.
    '''
    ''' Agora cada um tem a sua linha, e este teste é o que acusa se alguém os
    ''' juntar de novo.
    ''' </summary>
    <STATestMethod>
    Public Sub Aviso_e_resultado_NAO_se_cobrem()
        Dim faixa = Montar("A operação não terminou.", True, "resumo do modelo")

        Dim texto = TextoVisivel(faixa)
        StringAssert.Contains(texto, "A operação não terminou.")
        StringAssert.Contains(texto, "resumo do modelo")

        Dim caixas = Visiveis(faixa).ToList()
        Assert.AreEqual(3, caixas.Count, "botoes, aviso e resultado")

        For i = 0 To caixas.Count - 2
            For j = i + 1 To caixas.Count - 1
                ' ENCOSTAR nao e cobrir: faixas empilhadas dividem a borda, e
                ' IntersectsWith conta isso como intersecao. O que importa e
                ' area sobreposta de verdade.
                Dim comum = Rect.Intersect(caixas(i), caixas(j))
                Assert.IsTrue(comum.IsEmpty OrElse comum.Height < 0.5,
                    "duas faixas se cobrindo. " &
                    String.Join(" | ", caixas.Select(Function(c) c.ToString())))
            Next
        Next
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

    Private Shared Iterator Function Descendentes(no As DependencyObject) _
                                                  As IEnumerable(Of DependencyObject)
        For i = 0 To VisualTreeHelper.GetChildrenCount(no) - 1
            Dim filho = VisualTreeHelper.GetChild(no, i)
            Yield filho
            For Each neto In Descendentes(filho)
                Yield neto
            Next
        Next
    End Function

    Private Shared Function Botoes(raiz As DependencyObject) As IEnumerable(Of Button)
        Return Descendentes(raiz).OfType(Of Button)().
               Where(Function(b) b.Visibility = Visibility.Visible)
    End Function

    ''' <summary>
    ''' Onde cada <b>faixa</b> ficou, depois do arranjo.
    '''
    ''' Só os filhos diretos do <c>Grid</c> raiz, e não todo <c>Border</c> da
    ''' árvore: um <c>Button</c> tem borda própria no template, o
    ''' <c>UserControl</c> tem a dele, e varrer a árvore inteira devolvia oito
    ''' retângulos aninhados — que se sobrepõem por construção, e não porque
    ''' alguma faixa cobre outra.
    ''' </summary>
    Private Shared Iterator Function Visiveis(raiz As FrameworkElement) As IEnumerable(Of Rect)
        Dim grade = Descendentes(raiz).OfType(Of Grid)().FirstOrDefault()
        If grade Is Nothing Then Return

        For Each filho As UIElement In grade.Children
            If filho.Visibility <> Visibility.Visible Then Continue For
            Dim fe = CType(filho, FrameworkElement)
            If fe.ActualHeight <= 0 Then Continue For
            Dim canto = fe.TranslatePoint(New Point(0, 0), raiz)
            Yield New Rect(canto, New Size(fe.ActualWidth, fe.ActualHeight))
        Next
    End Function

    ' ==================================================================
    ' LEGIBILIDADE
    '
    ' O primeiro resumo de verdade chegou certo e ILEGIVEL: a faixa nao
    ' consumia token nenhum, entao herdava preto sobre fundo escuro. O texto
    ' estava la, na arvore visual, e os testes todos verdes -- porque nenhum
    ' deles olhava para a COR.

    ''' <summary>
    ''' <b>A resposta da IA é legível sobre o fundo da faixa.</b>
    '''
    ''' Não confere igualdade com um token: confere <b>contraste</b>. Amarrar
    ''' o teste ao brush deixaria trocar a paleta por uma ilegível sem nada
    ''' falhar — e o que importa aqui não é qual cor é, é dar para ler.
    ''' </summary>
    <STATestMethod>
    Public Sub A_resposta_da_IA_e_LEGIVEL()
        Dim faixa = Montar("", False, "resumo do modelo")

        Dim alvo = Descendentes(faixa).OfType(Of TextBlock)().
                   FirstOrDefault(Function(t) t.Text = "resumo do modelo")
        Assert.IsNotNull(alvo, "nao achei o texto da resposta")

        Dim frente = CorDe(alvo.Foreground)
        Dim fundo = CorDe(FundoAtras(alvo))

        Dim razao = Contraste(frente, fundo)
        Assert.IsTrue(razao >= 4.5,
            $"contraste {razao:F2}:1 entre {frente} e {fundo} -- abaixo de 4.5:1 " &
            "isto e o que o usuario chamou de 'ilegivel nesta cor preta'")
    End Sub

    ''' <summary>
    ''' <b>Controle negativo: o teste acima reprova mesmo.</b>
    '''
    ''' Sem ele, um <c>Contraste</c> que devolvesse sempre 21 faria o principal
    ''' passar para qualquer paleta — inclusive a que quebrou.
    ''' </summary>
    <STATestMethod>
    Public Sub Controle_preto_sobre_o_fundo_escuro_REPROVA()
        Dim escuro = Color.FromRgb(&H17, &H1A, &H21)
        Dim preto = Color.FromRgb(0, 0, 0)

        Assert.IsTrue(Contraste(preto, escuro) < 4.5,
                      "preto sobre #171A21 tinha de reprovar")
        Assert.IsTrue(Contraste(Color.FromRgb(&HE6, &HEA, &HF0), escuro) >= 4.5,
                      "e o texto claro dos tokens tinha de passar")
    End Sub

    ''' <summary>A primeira cor de fundo opaca subindo a árvore.</summary>
    Private Shared Function FundoAtras(qual As DependencyObject) As Brush
        Dim atual = qual
        While atual IsNot Nothing
            Dim b = TryCast(atual, Border)
            If b IsNot Nothing AndAlso TypeOf b.Background Is SolidColorBrush Then
                Return b.Background
            End If
            Dim p = TryCast(atual, Panel)
            If p IsNot Nothing AndAlso TypeOf p.Background Is SolidColorBrush Then
                Return p.Background
            End If
            atual = Media.VisualTreeHelper.GetParent(atual)
        End While
        Return Nothing
    End Function

    Private Shared Function CorDe(b As Brush) As Color
        Dim s = TryCast(b, SolidColorBrush)
        ' Sem brush, o WPF pinta texto preto e nao pinta fundo nenhum. Devolver
        ' preto nos dois casos e o pior caso HONESTO: e exatamente o que estava
        ' acontecendo na tela quando o defeito apareceu.
        If s Is Nothing Then Return Color.FromRgb(0, 0, 0)
        Return s.Color
    End Function

    ''' <summary>Razão de contraste da WCAG 2.1.</summary>
    Private Shared Function Contraste(a As Color, b As Color) As Double
        Dim la = Luminancia(a), lb = Luminancia(b)
        Dim claro = Math.Max(la, lb), escuro = Math.Min(la, lb)
        Return (claro + 0.05) / (escuro + 0.05)
    End Function

    Private Shared Function Luminancia(c As Color) As Double
        Return 0.2126 * Canal(c.R) + 0.7152 * Canal(c.G) + 0.0722 * Canal(c.B)
    End Function

    Private Shared Function Canal(v As Byte) As Double
        Dim x = v / 255.0
        If x <= 0.03928 Then Return x / 12.92
        Return Math.Pow((x + 0.055) / 1.055, 2.4)
    End Function

End Class

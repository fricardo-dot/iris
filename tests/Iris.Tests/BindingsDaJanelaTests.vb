Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports Iris.App.ViewModels
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Os caminhos de <c>Binding</c> do XAML resolvem contra os ViewModels.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ISTO EXISTE</b>
'''
''' Binding com caminho errado no WPF <b>falha em silêncio</b>. A propriedade
''' não existe, o controle fica vazio, nada é lançado, nada aparece no log — e
''' a suíte continua verde porque nenhum teste toca em XAML.
'''
''' É "verde mas quebrado" na forma mais pura, e é o formato de erro que esta
''' fase inteira persegue. A faixa do acervo é o caso concreto: se
''' <c>Acervo.Ressalva</c> virasse <c>Acervo.Ressalvas</c> num refactor, a
''' ressalva que a §23 obriga a mostrar simplesmente <b>sumiria da tela</b>,
''' e o produto voltaria a exibir o acervo como se fosse o estado corrente da
''' caixa — sem nenhum sinal de que algo quebrou.
''' </summary>
<TestClass>
Public Class BindingsDaJanelaTests

    ''' <summary>
    ''' Raízes conhecidas: o prefixo do caminho e o tipo em que ele começa.
    '''
    ''' Só as raízes que este teste sabe resolver. Um <c>Binding</c> dentro de
    ''' <c>DataTemplate</c> tem outro DataContext e não é verificável assim —
    ''' por isso a lista é explícita em vez de "tudo o que aparecer".
    ''' </summary>
    Private Shared Function Raizes() As Dictionary(Of String, Type)
        Return New Dictionary(Of String, Type) From {
            {"Acervo.", GetType(AcervoViewModel)},
            {"Connection.", GetType(ConnectionViewModel)},
            {"Composer.", GetType(ComposerViewModel)},
            {"Detail.", GetType(MessageDetailViewModel)},
            {"Messages.", GetType(MessageListViewModel)},
            {"Folders.", GetType(FolderTreeViewModel)}}
    End Function

    <TestMethod>
    Public Sub Todo_binding_conhecido_resolve_no_ViewModel()
        Dim xaml = LerXaml()
        Dim conhecidas = Raizes()
        Dim quebrados As New List(Of String)()
        Dim conferidos = 0

        For Each caminho In CaminhosDeBinding(xaml)
            Dim raiz = conhecidas.Keys.FirstOrDefault(Function(k) caminho.StartsWith(k, StringComparison.Ordinal))
            If raiz Is Nothing Then Continue For

            conferidos += 1
            Dim membro = caminho.Substring(raiz.Length)
            ' So o primeiro segmento: "A.B.C" resolve A, e B/C dependem do tipo
            ' de A, que este teste nao persegue.
            Dim primeiro = membro.Split("."c)(0)
            If primeiro.Length = 0 Then Continue For

            If conhecidas(raiz).GetProperty(primeiro,
                    BindingFlags.Public Or BindingFlags.Instance) Is Nothing Then
                quebrados.Add($"{caminho}  (nao existe em {conhecidas(raiz).Name})")
            End If
        Next

        Assert.IsTrue(conferidos > 5,
            $"so {conferidos} bindings conferidos — o teste nao esta encontrando o XAML")
        Assert.AreEqual(0, quebrados.Count,
            "binding com caminho errado falha em SILENCIO no WPF: " &
            Environment.NewLine & String.Join(Environment.NewLine, quebrados))
    End Sub

    ''' <summary>
    ''' A faixa do acervo existe, e mostra a RESSALVA.
    '''
    ''' Não basta o binding resolver: a §23 obriga a ressalva a aparecer junto
    ''' do acervo, e um refactor que removesse o <c>TextBlock</c> passaria no
    ''' teste acima sem problema nenhum — não haveria binding quebrado, haveria
    ''' binding ausente.
    ''' </summary>
    <TestMethod>
    Public Sub A_janela_mostra_a_ressalva_do_acervo()
        Dim xaml = LerXaml()
        StringAssert.Contains(xaml, "Acervo.Ressalva",
            "a ressalva da §23 tem de estar na janela, nao so no ViewModel")
        StringAssert.Contains(xaml, "AcervoIndisponivel",
            "cache que nao abre tem de aparecer — vazio silencioso e " &
            "indistinguivel de 'nao ha nada guardado'")
    End Sub

    ''' <summary>
    ''' Controle: a busca por caminho quebrado REALMENTE acusa.
    '''
    ''' Sem isto, um extrator que não achasse binding nenhum faria o teste
    ''' principal passar para sempre.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_um_caminho_inventado_seria_acusado()
        Assert.IsNull(GetType(AcervoViewModel).GetProperty("RessalvaQueNaoExiste",
                          BindingFlags.Public Or BindingFlags.Instance),
                      "a propriedade inventada nao pode existir de verdade")

        Dim achados = CaminhosDeBinding(
            "<TextBlock Text=""{Binding Acervo.RessalvaQueNaoExiste}"" />").ToList()
        CollectionAssert.Contains(achados, "Acervo.RessalvaQueNaoExiste",
            "o extrator nao encontra nem um caminho plantado — ele nao extrai nada")
    End Sub

    ' ==================================================================

    Private Shared Iterator Function CaminhosDeBinding(xaml As String) As IEnumerable(Of String)
        ' {Binding Caminho} e {Binding Path=Caminho} e {Binding Caminho, ...}
        For Each m As Match In Regex.Matches(xaml, "\{Binding\s+(?:Path=)?([A-Za-z_][\w.]*)")
            Yield m.Groups(1).Value
        Next
    End Function

    Private Shared _xaml As String

    Private Shared Function LerXaml() As String
        If _xaml IsNot Nothing Then Return _xaml
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
            d = d.Parent
        End While
        Assert.IsNotNull(d, "nao achei a raiz do repositorio")
        Dim caminho = Path.Combine(d.FullName, "src", "Iris.App", "MainWindow.xaml")
        Assert.IsTrue(File.Exists(caminho), "MainWindow.xaml nao encontrado em " & caminho)
        _xaml = File.ReadAllText(caminho)
        Return _xaml
    End Function

End Class

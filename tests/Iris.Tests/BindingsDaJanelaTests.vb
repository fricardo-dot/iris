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
            {"Busca.", GetType(BuscaViewModel)},
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
    ''' <b>A varredura tem botao na janela, e ele nao some.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' A faixa do acervo era visivel so quando havia ressalva
    ''' (<c>TemAlgoADizer</c>). Com o botao dentro dela, ele sumiria justamente
    ''' quando nao ha nada a ressalvar — que e quando alguem quer varrer.
    '''
    ''' E o mesmo defeito que a faixa da IA ja teve, com o mesmo custo: botao
    ''' que some esconde a funcionalidade E o motivo de ela estar
    ''' indisponivel. A visibilidade passou a ser "existe acervo".
    ''' </summary>
    <TestMethod>
    Public Sub A_varredura_tem_BOTAO_na_janela()
        Dim xaml = LerXaml()

        StringAssert.Contains(xaml, "Acervo.VarrerCommand",
            "sem botao, o SweepRunner continua sendo codigo que ninguem chama")
        StringAssert.Contains(xaml, "Acervo.Varrendo",
            "varrer bloqueia e demora: sem sinal, 'lendo' e 'travou' sao iguais")

        Assert.IsFalse(xaml.Contains("Acervo.TemAlgoADizer"),
            "a faixa do acervo voltou a sumir quando nao ha ressalva, e leva o " &
            "botao de varrer junto")
    End Sub

    ''' <summary>
    ''' <b>A faixa do acervo e a da IA não moram na mesma linha da grade.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O DEFEITO QUE ESTE TESTE FECHA</b>
    '''
    ''' As duas ficavam em <c>Grid.Row="2"</c>. Num <c>Grid</c> isso significa
    ''' <b>empilhadas</b>, e a da IA — declarada depois, e com fundo próprio —
    ''' pintava por cima. <b>A faixa do acervo nunca foi vista na tela</b>, e a
    ''' pendência da Fase 2 que dizia exatamente isso não era falta de dado: era
    ''' uma linha de grade faltando.
    '''
    ''' Ninguém notou porque as duas tinham visibilidade condicional: a do
    ''' acervo só aparecia havendo ressalva, e a da IA cobria justamente quando
    ''' aparecia. Duas condições escondendo uma sobreposição.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ESTE TESTE LÊ TEXTO, E NÃO PIXEL</b>
    '''
    ''' O jeito honesto seria instanciar a <c>MainWindow</c> e medir os
    ''' retângulos, como o <c>Aviso_e_resultado_NAO_se_cobrem</c> faz dentro da
    ''' faixa da IA. A janela exige broker, cache e sessão, e montar tudo isso
    ''' para conferir um número de linha custaria mais do que vale.
    '''
    ''' Então ele lê o XAML e cobra a propriedade que causou o defeito: as duas
    ''' faixas, que aparecem <b>ao mesmo tempo</b>, declaram linhas diferentes.
    ''' Não prova ausência de sobreposição em geral — prova que esta não voltou.
    ''' </summary>
    <TestMethod>
    Public Sub O_acervo_e_a_IA_nao_dividem_a_LINHA_da_grade()
        Dim xaml = LerXaml()

        Dim daIa = Text.RegularExpressions.Regex.Match(
            xaml, "<local:FaixaDaIa\s+Grid\.Row=""(\d+)""")
        Assert.IsTrue(daIa.Success, "nao achei a faixa da IA na janela")

        Dim doAcervo = Text.RegularExpressions.Regex.Match(
            xaml, "<Border Grid\.Row=""(\d+)""\s*?
\s*Visibility=""\{Binding Acervo,")
        Assert.IsTrue(doAcervo.Success, "nao achei a faixa do acervo na janela")

        Assert.AreNotEqual(doAcervo.Groups(1).Value, daIa.Groups(1).Value,
            "as duas faixas voltaram para a mesma linha da grade, e a de baixo " &
            "cobre a de cima -- foi assim que a faixa do acervo passou a fase " &
            "inteira sem nunca ter sido vista")
    End Sub

    ''' <summary>
    ''' <b>A janela hospeda a faixa da IA, com o contexto certo.</b>
    '''
    ''' A faixa é um <c>UserControl</c> próprio — foi extraída para que o teste
    ''' de renderização pudesse instanciar a <b>faixa de verdade</b> em vez de
    ''' uma imitação. O que a janela precisa fazer é hospedá-la e dar a ela o
    ''' <c>DataContext</c> certo; sem isso, todos os bindings de dentro
    ''' resolveriam contra o <c>MainViewModel</c> e ficariam vazios em silêncio.
    ''' </summary>
    <TestMethod>
    Public Sub A_janela_hospeda_a_faixa_da_IA()
        Dim xaml = LerXaml()
        StringAssert.Contains(xaml, "<local:FaixaDaIa",
            "a faixa da IA tem de estar na janela")
        StringAssert.Contains(xaml, "DataContext=" & Q & "{Binding Assistente}" & Q,
            "sem o contexto certo, os bindings de dentro resolvem contra o " &
            "MainViewModel e ficam vazios em silencio")
    End Sub

    ''' <summary>
    ''' <b>Todo binding da faixa resolve no <c>AssistenteViewModel</c>.</b>
    '''
    ''' O <c>DataContext</c> da faixa é o assistente, então os caminhos são
    ''' diretos — <c>Aviso</c>, e não <c>Assistente.Aviso</c>. Um refactor que
    ''' renomeasse uma propriedade deixaria o binding perfeito e o texto
    ''' invisível: binding com caminho errado no WPF <b>falha em silêncio</b>.
    ''' </summary>
    <TestMethod>
    Public Sub Todo_binding_da_faixa_resolve_no_AssistenteViewModel()
        Dim quebrados As New List(Of String)()

        For Each caminho In CaminhosDeBinding(LerFaixa())
            Dim partes = caminho.Split("."c)
            Dim alvo As Type = GetType(AssistenteViewModel)

            For Each membro In partes
                If alvo Is Nothing Then Exit For
                Dim p = alvo.GetProperty(membro)
                If p Is Nothing Then
                    quebrados.Add(caminho)
                    Exit For
                End If
                alvo = p.PropertyType
            Next
        Next

        Assert.AreEqual(0, quebrados.Count,
            "caminho que nao resolve fica vazio em silencio: " &
            String.Join(", ", quebrados))
    End Sub

    ''' <summary>
    ''' <b>A faixa mostra a situação da IA — e o que ficou sem desfecho.</b>
    '''
    ''' Binding ausente não é binding quebrado, e passaria pelo teste de cima sem
    ''' reclamar. O que a §28.2 obriga a mostrar é o motivo de a IA não estar
    ''' habilitada; o que a §29.6 obriga é o número de envios que ficaram
    ''' ambíguos numa execução anterior.
    '''
    ''' "Pode ter saído conteúdo desta caixa e ninguém sabe" não pode viver só no
    ''' banco.
    ''' </summary>
    <TestMethod>
    Public Sub A_faixa_mostra_a_situacao_da_IA()
        Dim xaml = LerFaixa()
        StringAssert.Contains(xaml, "{Binding Aviso}",
            "o motivo de a IA nao estar habilitada tem de aparecer")
        StringAssert.Contains(xaml, "{Binding Reconciliacao.Aviso}",
            "envios sem desfecho conhecido nao podem viver so no banco")
        StringAssert.Contains(xaml, "{Binding Resultado}",
            "e a resposta do modelo tem de ter onde aparecer")
    End Sub

    ''' <summary>
    ''' <b>A ação existe na faixa.</b>
    '''
    ''' Sem os botões, o 3.5 seria uma tela de status: os comandos existiriam no
    ''' ViewModel e ninguém os alcançaria, e nem uma ativação futura tornaria a
    ''' funcionalidade utilizável.
    ''' </summary>
    <TestMethod>
    Public Sub A_acao_da_IA_existe_na_faixa()
        Dim xaml = LerFaixa()
        For Each comando In {"ResumirCommand", "RedigirCommand",
                             "DesfazerCommand", "CancelarCommand"}
            StringAssert.Contains(xaml, "{Binding " & comando & "}",
                comando & " nao esta na faixa — o comando existiria sem ninguem alcancar")
        Next
    End Sub

    ''' <summary>
    ''' <b>A resposta do modelo aparece num <c>TextBlock</c>.</b>
    '''
    ''' Não num controle que interprete Markdown, HTML ou link: ela vem de um
    ''' lugar que leu o e-mail, que por sua vez veio de fora. A barreira da §29.5
    ''' é estrutural, e este teste é onde ela fica presa ao XAML.
    ''' </summary>
    <TestMethod>
    Public Sub A_resposta_do_modelo_aparece_em_TEXTBLOCK()
        Dim xaml = LerFaixa()
        Dim i = xaml.IndexOf("{Binding Resultado}", StringComparison.Ordinal)
        Assert.IsTrue(i > 0, "o binding tem de existir")

        Dim antes = xaml.Substring(0, i)
        Dim elemento = antes.Substring(antes.LastIndexOf("<"c))

        StringAssert.StartsWith(elemento, "<TextBlock",
            "a resposta do modelo nao pode ir para um controle que INTERPRETE: " &
            "ela e dado, e dado que veio de fora")
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

    ''' <summary>Aspas duplas, sem duplicar aspas dentro do literal.</summary>
    Private Const Q As String = """"

    ''' <summary>O XAML da faixa da IA, que é um <c>UserControl</c> próprio.</summary>
    Private Shared Function LerFaixa() As String
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
            d = d.Parent
        End While
        Assert.IsNotNull(d, "nao achei a raiz do repositorio")
        Dim caminho = Path.Combine(d.FullName, "src", "Iris.App", "Views", "FaixaDaIa.xaml")
        Assert.IsTrue(File.Exists(caminho), "FaixaDaIa.xaml nao encontrado em " & caminho)
        Return File.ReadAllText(caminho)
    End Function

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

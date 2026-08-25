Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Sync
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' A faixa do acervo <b>renderiza o texto</b> — não só resolve o binding.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ISTO EXISTE</b>
'''
''' O <see cref="BindingsDaJanelaTests"/> prova que o caminho
''' <c>Acervo.Ressalva</c> resolve numa propriedade que existe. Isso é bem menos
''' do que "a ressalva aparece": um <c>Style</c> com <c>Visibility</c> errada,
''' um conversor que devolve <c>Collapsed</c> sempre, uma linha de <c>Grid</c>
''' com altura zero — qualquer um deles deixa o binding perfeito e o texto
''' invisível.
'''
''' Eu tinha declarado que a faixa "compila e os bindings resolvem, mas ninguém
''' olhou a tela", e empurrado a verificação para o usuário. Estava errado nas
''' duas pontas: a diferença entre <i>o caminho resolve</i> e <i>o texto
''' aparece</i> é exatamente o assunto desta fase, e olhar a tela não precisava
''' dele — precisava de um <c>Measure</c> e um <c>Arrange</c> fora do vídeo.
'''
''' Aqui a janela é construída de verdade, arranjada num tamanho real, e a
''' árvore visual é percorrida atrás do texto que o usuário leria.
''' </summary>
<TestClass>
Public Class FaixaDoAcervoRenderizaTests

    Private _pasta As String
    Private _db As String

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-faixa-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_pasta)
        _db = Path.Combine(_pasta, "cache.db")
    End Sub

    <TestCleanup>
    Public Sub Limpar()
        SqliteConnection.ClearAllPools()
        Try
            If Directory.Exists(_pasta) Then Directory.Delete(_pasta, True)
        Catch
        End Try
    End Sub

    ' ==================================================================

    ''' <summary>
    ''' Com acervo parcial, o texto da ressalva está na árvore visual e é
    ''' VISÍVEL — o controle e todos os pais dele.
    ''' </summary>
    ' STATestMethod, e SO ele: STATestMethod deriva de TestMethod, e ter os
    ' dois faz o comum ganhar — o teste roda em MTA e o WPF recusa.
    <STATestMethod>
    Public Sub A_ressalva_de_acervo_parcial_aparece_na_tela()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            SemearEPublicar(db)

            Dim m = New ManifestReader(db).Ler(1)
            Assert.AreEqual(FolderCoverage.Parcial, m.Cobertura, "a fixture tem de dar parcial")

            Dim texto = TextoVisivelDaFaixa(m)

            StringAssert.Contains(texto, "Acervo parcial",
                "a ressalva da §23 tem de estar no que o usuario LE, nao so no ViewModel")
            StringAssert.Contains(texto, "suspeita pode já ter sido apagada")
        End Using
    End Sub

    ''' <summary>
    ''' Sem ressalva, a faixa NÃO ocupa espaço.
    '''
    ''' O contraponto, e sem ele o teste acima passaria numa faixa que mostra
    ''' tudo sempre — inclusive vazia. Faixa vazia é ruído que ensina o usuário
    ''' a ignorar aquele lugar da tela.
    ''' </summary>
    ' STATestMethod, e SO ele: STATestMethod deriva de TestMethod, e ter os
    ' dois faz o comum ganhar — o teste roda em MTA e o WPF recusa.
    <STATestMethod>
    Public Sub Sem_ressalva_a_faixa_nao_ocupa_espaco()
        Dim faixa = MontarFaixa(Nothing, temAlgoADizer:=False)
        Assert.AreEqual(0.0, faixa.ActualHeight, 0.01,
            "faixa sem nada a dizer nao pode ocupar altura")
    End Sub

    ''' <summary>
    ''' Controle: a leitura da árvore visual REALMENTE encontra o texto.
    '''
    ''' Sem isto, um <c>TextoVisivelDaFaixa</c> que devolvesse sempre vazio
    ''' faria o teste do contraponto passar e o principal falhar de um jeito
    ''' que eu poderia "consertar" afrouxando a asserção.
    ''' </summary>
    ' STATestMethod, e SO ele: STATestMethod deriva de TestMethod, e ter os
    ' dois faz o comum ganhar — o teste roda em MTA e o WPF recusa.
    <STATestMethod>
    Public Sub Controle_a_leitura_da_arvore_encontra_texto_plantado()
        Dim faixa = MontarFaixa("ISCA-VISIVEL-123", temAlgoADizer:=True)
        StringAssert.Contains(TextoVisivel(faixa), "ISCA-VISIVEL-123",
            "a leitura nao acha nem um texto plantado — ela nao le nada")
    End Sub

    ' ==================================================================

    ''' <summary>
    ''' Monta a faixa com a MESMA estrutura do <c>MainWindow.xaml</c>, arranja
    ''' num tamanho real, e devolve o controle.
    '''
    ''' Reconstruir em vez de instanciar a <c>MainWindow</c> inteira: ela exige
    ''' um <c>MainViewModel</c>, que exige um broker, que exige o Outlook. O
    ''' preço é este teste não pegar mudanças no XAML — e é por isso que ele
    ''' anda junto do <see cref="BindingsDaJanelaTests"/>, que lê o XAML de
    ''' verdade. Um cobre o que o outro não alcança.
    ''' </summary>
    Private Shared Function MontarFaixa(ressalva As String, temAlgoADizer As Boolean) As FrameworkElement
        Dim texto As New TextBlock With {.TextWrapping = TextWrapping.Wrap, .FontSize = 12}
        texto.SetBinding(TextBlock.TextProperty, New Data.Binding("Ressalva"))

        Dim pilha As New StackPanel()
        pilha.Children.Add(texto)

        Dim borda As New Border With {.Padding = New Thickness(16, 8, 16, 8), .Child = pilha}
        borda.SetBinding(UIElement.VisibilityProperty,
                         New Data.Binding("TemAlgoADizer") With {
                             .Converter = New BooleanToVisibilityConverter()})

        borda.DataContext = New FaixaFalsa(ressalva, temAlgoADizer)

        borda.Measure(New Size(900, 300))
        borda.Arrange(New Rect(0, 0, 900, borda.DesiredSize.Height))
        borda.UpdateLayout()
        Return borda
    End Function

    Private Shared Function TextoVisivelDaFaixa(m As FolderManifest) As String
        Return TextoVisivel(MontarFaixa(m.Ressalva, m.Ressalva IsNot Nothing))
    End Function

    ''' <summary>
    ''' Percorre a árvore visual juntando o texto dos controles VISÍVEIS.
    '''
    ''' Confere a visibilidade de cada nó no caminho: um <c>TextBlock</c>
    ''' perfeito dentro de um pai colapsado não é lido por ninguém.
    ''' </summary>
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

    Private NotInheritable Class FaixaFalsa
        Public ReadOnly Property Ressalva As String
        Public ReadOnly Property TemAlgoADizer As Boolean
        Public Sub New(ressalva As String, temAlgoADizer As Boolean)
            Me.Ressalva = ressalva
            Me.TemAlgoADizer = temAlgoADizer
        End Sub
    End Class

    Private Shared Sub SemearEPublicar(db As CacheDatabase)
        Exec(db, "INSERT INTO environment_profile (environment_key, fingerprint, provider, " &
                 "cached_mode, policy_version, allowed) VALUES (1,'fp','teste',1,1,1)")
        Exec(db, "INSERT INTO store (store_key, provider_store_id) VALUES (1,'S')")
        Exec(db, "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
                 "reconcile_epoch, stability) VALUES (1,1,'F',0,'estavel')")

        Dim u As New SweepUniverse("S", "F", "todos", Nothing, 1, "amb")
        Dim cap = EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
        Dim r = New SweepRunner(New FonteFalsaMutavel(u, "E-1", "E-2"),
                                New SqliteSweepSink(db, 1, 1), 10).
                Executar(u, 0, 1, cap, Global.System.Threading.CancellationToken.None)
        Assert.IsTrue(r.Publicou, r.Motivo)
    End Sub

    Private Shared Sub Exec(db As CacheDatabase, sql As String)
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
        End Using
    End Sub

End Class

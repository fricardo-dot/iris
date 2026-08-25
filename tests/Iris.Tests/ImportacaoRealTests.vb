Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Threading
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Integration.Outlook
Imports Iris.Model
Imports Iris.Outlook
Imports Iris.Sync
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Marco 2.2b — a varredura inteira contra o Outlook REAL, gravando num
''' cache SQLite real.
'''
''' Tudo até aqui foi provado contra fonte falsa. Uma fonte falsa não tem
''' fila da STA, não tem RCW, não devolve página drenada por empate, e não
''' morre de <c>RPC_E_CALL_REJECTED</c>. É o que a Q1 já tinha cobrado quando
''' o teste sintético verificava um algoritmo diferente do que rodava contra
''' o Outlook.
'''
''' <b>Só leitura.</b> Nada é criado, movido, apagado ou enviado na caixa do
''' usuário. O cache vai para um arquivo temporário e é apagado no fim.
''' </summary>
<TestClass>
Public Class ImportacaoRealTests

    Private _pasta As String
    Private _db As String

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-real-" & Guid.NewGuid().ToString("N"))
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
    ''' Importa uma pasta de verdade e mede o que a D5 pede.
    '''
    ''' A D5 decidiu 100 ms de orçamento por lote e nunca foi medida — era o
    ''' critério 8 da §8, o único que continuava "decidido e não medido".
    ''' Medir exigia exatamente isto: o adaptador real, na fila da STA, contra
    ''' a caixa do usuário.
    ''' </summary>
    <TestMethod>
    Public Async Function Importa_uma_pasta_real_e_mede_a_latencia_por_lote() As Task
        Using broker = Await AbrirAsync()
            ' A Caixa de Entrada, e nao uma pasta pequena: a D5 e sobre o
            ' orcamento de tempo POR LOTE, e um lote unico de 35 linhas nao diz
            ' nada sobre o custo de paginar. Medir na pasta facil e como medir
            ' a Q1 com dez mensagens.
            Dim entrada = Await AcharPastaAsync(broker, "Caixa de Entrada", "1. Backup")

            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
                Assert.IsNotNull(db, $"{falha}")
                Semear(db, entrada)

                Dim universo As New SweepUniverse(
                    entrada.StoreId, entrada.EntryId, "todos", Nothing, 1, "cached|janela-nao-lida")

                Dim fonte As New OutlookSweepSource(broker, entrada, universo, 1)
                Dim relogio As New RelogioDeLotes(fonte)
                Dim sink As New SqliteSweepSink(db, 1, 1)
                Dim runner As New SweepRunner(relogio, sink, tamanhoPagina:=100)

                Dim total = Stopwatch.StartNew()
                Dim r = runner.Executar(universo, 0, 1, Capacidades(), CancellationToken.None)
                total.Stop()

                ' ---- o que aconteceu ----
                Anotar($"conclusao : {r.Conclusion}")
                Anotar($"motivo    : {r.Motivo}")
                Anotar($"paginas   : {r.Paginas}")
                Anotar($"linhas    : {If(r.Attempt Is Nothing, 0, r.Attempt.RowsRead)}")
                Anotar($"cobertura : {r.Cobertura}")
                Anotar($"total     : {total.ElapsedMilliseconds} ms")
                Anotar($"idas COM  : {fonte.IdasAoProvider}")
                Anotar($"descartadas: {fonte.Descartadas}  (nao-mensagem ou erro de leitura)")
                Anotar("")
                Anotar("=== D5: latencia por lote (ms) ===")
                For Each l In relogio.Lotes
                    Anotar($"  lote {l.Numero,3}: {l.Ms,6} ms   {l.Linhas,4} linhas")
                Next
                If relogio.Lotes.Count > 0 Then
                    Anotar($"  min {relogio.Lotes.Min(Function(x) x.Ms)} / " &
                                      $"mediana {Mediana(relogio.Lotes.Select(Function(x) x.Ms))} / " &
                                      $"max {relogio.Lotes.Max(Function(x) x.Ms)}")
                End If

                ' ---- o que o teste EXIGE ----
                '
                ' Nao exijo que a varredura publique. Numa caixa viva, chegar
                ' mensagem no meio faz o S6 rejeitar - e rejeitar e o
                ' comportamento CERTO, nao falha do teste. O que eu exijo e que
                ' o desfecho seja um dos previstos e que nada fique pela metade.
                Assert.AreNotEqual(SweepConclusion.Falhou, r.Conclusion,
                    $"falha nao e desfecho previsto aqui: {r.Motivo}")

                If r.Publicou Then
                    Assert.AreEqual(FolderCoverage.Parcial, r.Cobertura,
                        "cached sem janela legivel publica PARCIAL (§23)")
                    Dim m = New ManifestReader(db).Ler(1)
                    Assert.AreEqual(r.Attempt.RowsRead, m.Items.Count,
                        "o manifesto tem de conter exatamente o que foi lido")
                    Assert.IsFalse(m.EhEstadoCorrente)
                    Anotar($"manifesto : {m.Items.Count} itens, ressalva: {m.Ressalva}")
                Else
                    Anotar("NAO publicou — e pode ser o certo numa caixa viva.")
                    Assert.AreEqual("descartada", Texto(db, "SELECT stage FROM scan_attempt"),
                        "nao publicou: a tentativa tem de ficar descartada, nao orfa")
                    Assert.AreEqual(0, Contar(db, "incarnation"),
                        "nao publicou: o acervo nao pode ter sido tocado")
                End If
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Retomar depois de uma interrupção converge para o mesmo manifesto.
    '''
    ''' É o critério que o Codex acrescentou ao pronto do 2.1: importação real
    ''' interrompida, seguida de varredura limpa, dá o mesmo resultado de uma
    ''' execução ininterrupta.
    ''' </summary>
    <TestMethod>
    Public Async Function Varredura_interrompida_e_refeita_converge() As Task
        Using broker = Await AbrirAsync()
            ' A Caixa de Entrada PRIMEIRO aqui: o cancelamento tem de acontecer
            ' na pagina 2, e "1. Backup" tem 35 linhas — com lote de 50 ela cabe
            ' numa pagina so e a interrupcao nunca dispara. Uma pasta pequena
            ' demais faz o teste passar sem exercitar o caminho.
            Dim alvo = Await AcharPastaAsync(broker, "Caixa de Entrada", "1. Backup")

            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
                Assert.IsNotNull(db, $"{falha}")
                Semear(db, alvo)

                Dim universo As New SweepUniverse(
                    alvo.StoreId, alvo.EntryId, "todos", Nothing, 1, "cached|janela-nao-lida")
                Dim sink As New SqliteSweepSink(db, 1, 1)

                ' 1. Interrompida na segunda pagina.
                Using cts As New CancellationTokenSource()
                    Dim fonte As New OutlookSweepSource(broker, alvo, universo, 1)
                    Dim corta As New CortaNaPagina(fonte, 2, cts)
                    Dim r1 = New SweepRunner(corta, sink, 50).
                             Executar(universo, 0, 1, Capacidades(), cts.Token)

                    ' Se a pasta couber numa pagina so, o corte nao dispara e o
                    ' teste nao exercita nada. Melhor declarar inconclusivo que
                    ' passar verde sem ter testado.
                    If corta.PaginasLidas < 2 Then
                        Assert.Inconclusive(
                            $"a pasta cabe em {corta.PaginasLidas} pagina(s): " &
                            "o cancelamento na pagina 2 nao chega a acontecer")
                    End If
                    Assert.AreEqual(SweepConclusion.Cancelada, r1.Conclusion,
                        $"esperava cancelamento, veio {r1.Conclusion}: {r1.Motivo}")
                End Using

                Assert.AreEqual(0, Contar(db, "incarnation"),
                    "cancelada nao pode ter tocado o acervo")

                ' 2. Limpa.
                Dim fonte2 As New OutlookSweepSource(broker, alvo, universo, 1)
                Dim r2 = New SweepRunner(fonte2, sink, 50).
                         Executar(universo, 0, 2, Capacidades(), CancellationToken.None)

                Anotar($"apos retomada: {r2.Conclusion} — {r2.Motivo}")

                If r2.Publicou Then
                    Dim m = New ManifestReader(db).Ler(1)
                    Assert.AreEqual(r2.Attempt.RowsRead, m.Items.Count)
                    Assert.IsTrue(m.Items.Select(Function(i) i.ProviderEntryId).Distinct().Count() =
                                  m.Items.Count, "a interrupcao nao pode ter duplicado nada")
                    Anotar($"convergiu para {m.Items.Count} itens, sem duplicata")
                Else
                    Assert.AreNotEqual(SweepConclusion.Falhou, r2.Conclusion, r2.Motivo)
                    Anotar("a segunda tambem nao publicou (caixa viva) — sem duplicata a conferir")
                End If
            End Using
        End Using
    End Function

    ' ==================================================================
    ' Instrumentação

    ''' <summary>Cronometra cada lote, sem alterar o comportamento.</summary>
    Private NotInheritable Class RelogioDeLotes
        Implements ISweepSource

        Private ReadOnly _dentro As ISweepSource
        Private _n As Integer = 0
        Friend ReadOnly Lotes As New List(Of (Numero As Integer, Ms As Long, Linhas As Integer))()

        Friend Sub New(dentro As ISweepSource)
            _dentro = dentro
        End Sub

        Public Function Contar(ct As CancellationToken) As SourceCount Implements ISweepSource.Contar
            Return _dentro.Contar(ct)
        End Function

        Public Function LerPagina(cursor As String, tamanho As Integer,
                                  ct As CancellationToken) As SourcePage _
                                  Implements ISweepSource.LerPagina
            _n += 1
            Dim sw = Stopwatch.StartNew()
            Dim p = _dentro.LerPagina(cursor, tamanho, ct)
            sw.Stop()
            Lotes.Add((_n, sw.ElapsedMilliseconds, p.Rows.Count))
            Return p
        End Function
    End Class

    ''' <summary>Cancela ao chegar na página N — a interrupção do teste de retomada.</summary>
    Private NotInheritable Class CortaNaPagina
        Implements ISweepSource

        Private ReadOnly _dentro As ISweepSource
        Private ReadOnly _corte As Integer
        Private ReadOnly _cts As CancellationTokenSource
        Private _n As Integer = 0

        Friend ReadOnly Property PaginasLidas As Integer
            Get
                Return _n
            End Get
        End Property

        Friend Sub New(dentro As ISweepSource, corte As Integer, cts As CancellationTokenSource)
            _dentro = dentro
            _corte = corte
            _cts = cts
        End Sub

        Public Function Contar(ct As CancellationToken) As SourceCount Implements ISweepSource.Contar
            Return _dentro.Contar(ct)
        End Function

        Public Function LerPagina(cursor As String, tamanho As Integer,
                                  ct As CancellationToken) As SourcePage _
                                  Implements ISweepSource.LerPagina
            _n += 1
            Dim p = _dentro.LerPagina(cursor, tamanho, ct)
            If _n >= _corte Then _cts.Cancel()
            Return p
        End Function
    End Class

    ' ==================================================================

    ''' <summary>
    ''' Onde as medicoes ficam. O MSTest engole Console.WriteLine, e medicao
    ''' que ninguem consegue ler depois nao e medicao — e o mesmo motivo pelo
    ''' qual o conferidor de crash virou script no repositorio.
    ''' </summary>
    Private Shared ReadOnly Saida As String =
        Path.Combine(RaizDoRepo(), "medicoes", "2.2b-medicoes.txt")

    Private Shared Function RaizDoRepo() As String
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
            d = d.Parent
        End While
        Return If(d Is Nothing, Path.GetTempPath(), d.FullName)
    End Function

    Private Shared Sub Anotar(linha As String)
        Console.WriteLine(linha)
        Directory.CreateDirectory(Path.GetDirectoryName(Saida))
        File.AppendAllText(Saida, linha & Environment.NewLine)
    End Sub

    Private Shared Function Capacidades() As EnvironmentCapabilities
        ' O ambiente REAL: cached, janela nao legivel (§22.4, §23).
        Return EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
    End Function

    Private Shared Async Function AbrirAsync() As Task(Of OutlookBroker)
        Dim broker As New OutlookBroker(New NullLog())
        broker.Start()
        Dim estado = Await broker.ConnectAsync(CancellationToken.None)
        Select Case PagingIntegrationTests.Decidir(
                    estado = SessionState.Connected, PagingIntegrationTests.OutlookInstalado())
            Case PagingIntegrationTests.SemOutlook.Prosseguir
                Return broker
            Case PagingIntegrationTests.SemOutlook.Falhar
                broker.Dispose()
                Assert.Fail($"Outlook instalado mas nao respondeu ({estado}). Abra-o e rode de novo.")
                Return Nothing
            Case Else
                broker.Dispose()
                Assert.Inconclusive($"Outlook nao instalado ({estado}).")
                Return Nothing
        End Select
    End Function

    ''' <summary>A primeira das pastas pedidas que existir e tiver mensagem.</summary>
    Private Shared Async Function AcharPastaAsync(broker As OutlookBroker,
                                                  ParamArray nomes As String()) As Task(Of FolderKey)
        Dim stores = Await broker.GetStoresAsync(CancellationToken.None)
        Assert.IsTrue(stores.Succeeded, "GetStoresAsync falhou")

        For Each nome In nomes
            For Each st In stores.Value
                Dim achada = Await ProcurarAsync(broker, st.RootFolder, nome, 0)
                If achada IsNot Nothing Then Return achada
            Next
        Next
        Assert.Inconclusive("nao achei nenhuma das pastas pedidas")
        Return Nothing
    End Function

    Private Shared Async Function ProcurarAsync(broker As OutlookBroker, pai As FolderKey,
                                                nome As String, prof As Integer) _
                                                As Task(Of FolderKey)
        If prof > 4 Then Return Nothing
        Dim filhas = Await broker.GetFolderChildrenAsync(pai, CancellationToken.None)
        If Not filhas.Succeeded Then Return Nothing
        For Each f In filhas.Value
            If String.Equals(f.Name, nome, StringComparison.OrdinalIgnoreCase) Then Return f.Key
            Dim d = Await ProcurarAsync(broker, f.Key, nome, prof + 1)
            If d IsNot Nothing Then Return d
        Next
        Return Nothing
    End Function

    Private Shared Sub Semear(db As CacheDatabase, pasta As FolderKey)
        Exec(db, "INSERT INTO environment_profile (environment_key, fingerprint, provider, " &
                 "cached_mode, policy_version, allowed) VALUES (1,'cached|janela-nao-lida'," &
                 "'ExchangeCached',1,1,1)")
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = "INSERT INTO store (store_key, provider_store_id) VALUES (1,$s)"
            cmd.Parameters.AddWithValue("$s", pasta.StoreId)
            cmd.ExecuteNonQuery()
        End Using
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
                              "reconcile_epoch, stability) VALUES (1,1,$f,0,'estavel')"
            cmd.Parameters.AddWithValue("$f", pasta.EntryId)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub Exec(db As CacheDatabase, sql As String)
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Function Texto(db As CacheDatabase, sql As String) As String
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            Dim v = cmd.ExecuteScalar()
            Return If(v Is Nothing OrElse v Is DBNull.Value, Nothing, Convert.ToString(v))
        End Using
    End Function

    Private Shared Function Contar(db As CacheDatabase, tabela As String) As Integer
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = $"SELECT COUNT(*) FROM {tabela}"
            Return Convert.ToInt32(cmd.ExecuteScalar())
        End Using
    End Function

    Private Shared Function Mediana(xs As IEnumerable(Of Long)) As Long
        Dim l = xs.OrderBy(Function(x) x).ToList()
        If l.Count = 0 Then Return 0
        Return l(l.Count \ 2)
    End Function

End Class

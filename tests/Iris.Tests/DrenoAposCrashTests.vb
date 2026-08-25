Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Marco 2.4 — a condição vinculante da §26.2, executada.
'''
''' <blockquote>
''' Para a integração ser considerada pronta, um teste tem de <b>matar o
''' processo entre <c>Receber</c> e <c>MarcarDrenada</c></b> e demonstrar que,
''' na reabertura, a UI converge sem perda.
''' </blockquote>
'''
''' É o que fecha o critério 9 — a parte que o 2.1 e o 2.2 não puderam provar
''' porque não havia consumidor.
'''
''' O consumidor usado aqui é o <see cref="AcervoService"/> de verdade, o mesmo
''' que o ViewModel usa. Se fosse uma imitação, a prova seria sobre a imitação
''' — o erro que a Q1 cobrou quando o teste sintético verificava um algoritmo
''' diferente do que rodava contra o Outlook.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class DrenoAposCrashTests

    Private _pasta As String
    Private _db As String
    Private Const FolderKey As Long = 1
    Private Const TotalLinhas As Integer = 6

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-dreno-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_pasta)
        _db = Path.Combine(_pasta, "cache.db")

        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"preparar: {falha}")
            Exec(db, "INSERT INTO environment_profile (environment_key, fingerprint, provider, " &
                     "cached_mode, policy_version, allowed) VALUES (1,'fp','teste',1,1,1)")
            Exec(db, "INSERT INTO store (store_key, provider_store_id) VALUES (1,'S')")
            Exec(db, "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
                     "reconcile_epoch, stability) VALUES (1,1,'F',0,'estavel')")
        End Using
        SqliteConnection.ClearAllPools()
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
    ''' <b>O teste que fecha o critério 9.</b>
    '''
    ''' O processo morre depois de o consumidor receber e antes de a dívida ser
    ''' marcada. Na reabertura, a mesma geração é entregue outra vez — entrega
    ''' <i>ao menos uma vez</i> — e o acervo converge para o mesmo conteúdo.
    '''
    ''' Convergir aqui não é "ficar igual por sorte": o
    ''' <see cref="AcervoService"/> é idempotente pela forma mais simples que
    ''' existe — não acumula nada, relê o manifesto inteiro. Receber duas vezes
    ''' e receber uma vez dão o mesmo estado porque o segundo substitui, não
    ''' soma.
    ''' </summary>
    <TestMethod>
    Public Sub Morrer_entre_Receber_e_MarcarDrenada_reentrega_e_converge()
        ' 1. Varre, publica, e MORRE no meio do dreno.
        Dim r1 = RodarHarness(CrashInjection.DepoisDeReceberAntesDeMarcarDrenada, drenar:=True)
        Assert.AreNotEqual(0, r1.ExitCode, "o harness deveria ter morrido")
        StringAssert.Contains(r1.Stdout, "resultado=Publicada",
            "tem de ter publicado antes de morrer no dreno")

        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")

            ' 2. O disco diz que a UI NAO recebeu — e ela recebeu.
            Dim dreno As New PublicationDrain(db)
            Assert.AreEqual(1, dreno.Pendentes().Count,
                "morreu antes de marcar: a divida continua pendente, que e o certo")

            ' 3. A reabertura entrega DE NOVO, e o acervo converge.
            Dim acervo As New AcervoService(db, FolderKey)
            Assert.AreEqual(TotalLinhas, acervo.Atual.Items.Count,
                "o acervo ja esta correto so de abrir — a publicacao sobreviveu")

            Assert.AreEqual(1, dreno.Drenar(acervo))
            Assert.AreEqual(1, acervo.Recebidas, "a geracao foi entregue OUTRA vez")
            Assert.AreEqual(TotalLinhas, acervo.Atual.Items.Count,
                "e o acervo converge para o mesmo conteudo — nada duplicou")
            Assert.AreEqual(0, dreno.Pendentes().Count, "agora sim, drenada")
        End Using
    End Sub

    ''' <summary>
    ''' Sem crash, o dreno entrega uma vez só — o contraponto.
    '''
    ''' Sem ele, um dreno que reentregasse sempre passaria no teste acima e
    ''' ninguém notaria que a entrega virou "muitas vezes" em vez de "ao menos
    ''' uma".
    ''' </summary>
    <TestMethod>
    Public Sub Sem_crash_entrega_uma_vez_so()
        Dim r = RodarHarness("nenhum", drenar:=True)
        Assert.AreEqual(0, r.ExitCode, r.Stderr)
        StringAssert.Contains(r.Stdout, "drenadas=1 recebidas=1",
            "sem crash, uma entrega e uma so")
        StringAssert.Contains(r.Stdout, $"itens={TotalLinhas}")

        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.AreEqual(0, New PublicationDrain(db).Pendentes().Count)
        End Using
    End Sub

    ''' <summary>
    ''' O acervo que o consumidor entrega carrega a ressalva — a §23 chegando
    ''' até onde a UI lê.
    ''' </summary>
    <TestMethod>
    Public Sub O_acervo_entregue_carrega_a_ressalva()
        Dim r = RodarHarness("nenhum", drenar:=True)
        Assert.AreEqual(0, r.ExitCode, r.Stderr)

        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim acervo As New AcervoService(db, FolderKey)
            Assert.IsFalse(acervo.Atual.EhEstadoCorrente,
                "o acervo nunca pode se apresentar como o estado corrente da caixa")
            Assert.IsNotNull(acervo.Atual.Ressalva)
        End Using
    End Sub

    ''' <summary>
    ''' Receber a mesma geração N vezes dá o mesmo acervo — a idempotência que
    ''' a entrega ao-menos-uma-vez cobra do consumidor.
    ''' </summary>
    <TestMethod>
    Public Sub Receber_a_mesma_geracao_varias_vezes_da_o_mesmo_acervo()
        Dim r = RodarHarness("nenhum", drenar:=False)
        Assert.AreEqual(0, r.ExitCode, r.Stderr)

        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim acervo As New AcervoService(db, FolderKey)
            Dim geracao = acervo.Atual.GenerationKey.Value

            For i = 1 To 5
                acervo.Receber(geracao)
            Next

            Assert.AreEqual(5, acervo.Recebidas)
            Assert.AreEqual(TotalLinhas, acervo.Atual.Items.Count,
                "cinco entregas da mesma geracao nao acumulam — o consumidor RELE, nao soma")
        End Using
    End Sub

    ' ==================================================================

    Private Structure Resultado
        Public ExitCode As Integer
        Public Stdout As String
        Public Stderr As String
    End Structure

    Private Function RodarHarness(ponto As String, drenar As Boolean) As Resultado
        Dim exe = LocalizarHarness()
        Dim psi As New ProcessStartInfo(exe) With {
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .UseShellExecute = False,
            .CreateNoWindow = True}
        For Each a In {_db, FolderKey.ToString(), ponto, "kill", "0", "",
                       If(drenar, "drenar", "")}
            psi.ArgumentList.Add(a)
        Next

        Using p = Process.Start(psi)
            Dim o = p.StandardOutput.ReadToEnd()
            Dim e = p.StandardError.ReadToEnd()
            If Not p.WaitForExit(60000) Then
                p.Kill(True)
                Assert.Fail("o harness travou")
            End If
            If e.Contains("Unhandled exception") OrElse e.Contains("abrir falhou") Then
                Assert.Fail("o harness falhou em vez de morrer no ponto pedido:" &
                            Environment.NewLine & e)
            End If
            Return New Resultado With {.ExitCode = p.ExitCode, .Stdout = o, .Stderr = e}
        End Using
    End Function

    Private Shared _exe As String

    Private Shared Function LocalizarHarness() As String
        If _exe IsNot Nothing Then Return _exe
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
            d = d.Parent
        End While
        Assert.IsNotNull(d, "nao achei a raiz do repositorio")
        Dim raiz = Path.Combine(d.FullName, "tools", "Iris.CrashHarness", "bin")
        Dim achado = Directory.GetFiles(raiz, "Iris.CrashHarness.exe", SearchOption.AllDirectories).
                     OrderByDescending(Function(f) File.GetLastWriteTimeUtc(f)).FirstOrDefault()
        Assert.IsNotNull(achado, "Iris.CrashHarness.exe nao encontrado")
        _exe = achado
        Return _exe
    End Function

    Private Shared Sub Exec(db As CacheDatabase, sql As String)
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
        End Using
    End Sub

End Class

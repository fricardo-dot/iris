Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports Iris.Cache
Imports Iris.Core
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Critério 9 do 2.0: gravar as linhas, avançar o checkpoint e publicar
''' precisam sobreviver a morrer no meio.
'''
''' Estes testes matam um processo DE VERDADE — <c>TerminateProcess</c>, via
''' o harness em <c>tools/Iris.CrashHarness</c> — e depois reabrem o arquivo.
''' A distinção importa e não é preciosismo: injetar exceção prova
''' ATOMICIDADE (a transação desfaz), mas o processo continua vivo e o SQLite
''' roda o rollback com ordem. Num crash ninguém desfaz nada e ninguém fecha
''' nada; quem recupera é o WAL, na abertura seguinte. Um teste que só injeta
''' exceção e se declara "teste de crash" exercita o caminho de encerramento
''' limpo — é o mesmo erro da §16.5 com o <c>Restrict</c>.
'''
''' O que NÃO está provado aqui: falta de energia. TerminateProcess mata o
''' processo, mas o Windows segue vivo e o que já foi entregue ao sistema de
''' arquivos continua lá.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class CrashTests

    Private _pasta As String
    Private _db As String
    Private Const FolderKey As Long = 1

    ' O harness escreve 3 paginas de 2 linhas.
    Private Const TotalLinhas As Integer = 6

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-crash-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_pasta)
        _db = Path.Combine(_pasta, "cache.db")

        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"preparar: {falha}")
            Exec(db.Connection, "INSERT INTO environment_profile (environment_key, fingerprint, " &
                                "provider, cached_mode, policy_version, allowed) " &
                                "VALUES (1,'fp','teste',1,1,1)")
            Exec(db.Connection, "INSERT INTO store (store_key, provider_store_id) VALUES (1,'S')")
            Exec(db.Connection, "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
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
    ' Os quatro pontos onde morrer

    ''' <summary>
    ''' Morrer DENTRO da transação da página: nada, nem meia linha.
    ''' </summary>
    <TestMethod>
    Public Sub Morrer_dentro_da_pagina_nao_deixa_nada_pela_metade()
        Dim r = RodarHarness(CrashInjection.DentroDaPaginaAntesDoCommit)
        Assert.AreNotEqual(0, r.ExitCode, "o harness deveria ter morrido")

        Comparar(Sub(c)
                     Assert.AreEqual(0, Contar(c, "scan_stage"), "nenhuma linha encenada")
                     Assert.AreEqual(0, Contar(c, "incarnation"), "nenhuma encarnacao")
                     Assert.AreEqual(0, Contar(c, "metadata_observation"))
                     Assert.AreEqual("aberta", Texto(c, "SELECT stage FROM scan_attempt"))
                     Assert.IsNull(Valor(c, "SELECT cursor FROM scan_attempt"),
                                   "o checkpoint NAO pode ter avancado")
                 End Sub)
    End Sub

    ''' <summary>
    ''' Morrer DEPOIS do commit da página: linhas e checkpoint concordam.
    '''
    ''' É o invariante que importa. Checkpoint à frente das linhas significa
    ''' que a retomada pula mensagens — perda silenciosa. Linhas à frente do
    ''' checkpoint só é inofensivo porque a gravação é idempotente.
    ''' </summary>
    <TestMethod>
    Public Sub Morrer_depois_do_commit_deixa_linhas_e_checkpoint_de_acordo()
        Dim r = RodarHarness(CrashInjection.DepoisDoCommitDaPagina)
        Assert.AreNotEqual(0, r.ExitCode)

        Comparar(Sub(c)
                     Assert.AreEqual(2, Contar(c, "scan_stage"), "so a pagina 1")
                     ' ZERO encarnacoes, e e o desenho: a pagina so ENCENA. O
                     ' acervo e materializado a partir de scan_stage na
                     ' transacao da PUBLICACAO. Antes, gravar a pagina escrevia
                     ' direto em incarnation/metadata/association, e uma
                     ' tentativa rejeitada depois alterava o manifesto da UI.
                     Assert.AreEqual(0, Contar(c, "incarnation"),
                        "a pagina encena; o acervo so e tocado ao publicar")
                     Assert.AreEqual(0, Contar(c, "association"))
                     Assert.AreEqual("cursor-1", Texto(c, "SELECT cursor FROM scan_attempt"))
                     Assert.AreEqual("varrendo", Texto(c, "SELECT stage FROM scan_attempt"))
                     Assert.AreEqual(2L, Convert.ToInt64(Valor(c, "SELECT rows_read FROM scan_attempt")))
                     ' Nada publicado: a varredura nem terminou.
                     Assert.AreEqual(0, Contar(c, "generation"))
                     Assert.IsNull(Valor(c, "SELECT published_generation_key FROM folder"))
                 End Sub)
    End Sub

    ''' <summary>
    ''' Morrer DENTRO da publicação: as páginas ficam, a publicação não
    ''' acontece pela metade. Nada de geração sem dívida, nem dívida sem
    ''' geração, nem cabeça apontando para o vazio.
    ''' </summary>
    <TestMethod>
    Public Sub Morrer_dentro_da_publicacao_nao_deixa_geracao_parcial()
        Dim r = RodarHarness(CrashInjection.DentroDaPublicacaoAntesDoCommit)
        Assert.AreNotEqual(0, r.ExitCode)

        Comparar(Sub(c)
                     Assert.AreEqual(TotalLinhas, Contar(c, "scan_stage"), "as 3 paginas ficaram")
                     Assert.AreEqual(0, Contar(c, "incarnation"),
                        "morreu antes de publicar: o acervo nao foi tocado")
                     Assert.AreEqual(0, Contar(c, "generation"), "nenhuma geracao")
                     Assert.AreEqual(0, Contar(c, "publication_log"), "nenhuma divida")
                     Assert.IsNull(Valor(c, "SELECT published_generation_key FROM folder"),
                                   "a cabeca nao pode ter avancado")
                     Assert.AreEqual("varrendo", Texto(c, "SELECT stage FROM scan_attempt"))
                 End Sub)
    End Sub

    ''' <summary>
    ''' Morrer DEPOIS do commit da publicação — o caso que justifica o desenho.
    '''
    ''' A UI nunca recebeu aviso nenhum: o processo morreu antes de qualquer
    ''' evento poder ser entregue. Mas a dívida está NO DISCO, não drenada, e
    ''' a próxima abertura a encontra. É por isso que publicar é uma linha e
    ''' não um evento: evento perdido não deixa rastro.
    ''' </summary>
    <TestMethod>
    Public Sub Morrer_depois_de_publicar_deixa_a_divida_registrada_para_a_UI()
        Dim r = RodarHarness(CrashInjection.DepoisDoCommitDaPublicacao)
        Assert.AreNotEqual(0, r.ExitCode)

        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Dim c = db.Connection
            Assert.AreEqual(1, Contar(c, "generation"))
            Assert.AreEqual("publicada", Texto(c, "SELECT stage FROM scan_attempt"))

            Dim g = Convert.ToInt64(Valor(c, "SELECT published_generation_key FROM folder"))
            Assert.AreEqual(Convert.ToInt64(Valor(c, "SELECT generation_key FROM generation")), g,
                            "a cabeca aponta para a geracao publicada")

            Dim w As New CacheWriter(db)
            Dim pendentes = w.PublicacoesPendentes()
            Assert.AreEqual(1, pendentes.Count, "a UI tem exatamente uma divida a drenar")
            Assert.AreEqual(g, pendentes(0))

            ' E drenar e idempotente: drenou, some.
            w.MarcarDrenada(g)
            Assert.AreEqual(0, w.PublicacoesPendentes().Count)
        End Using
    End Sub

    ' ==================================================================
    ' Retomada

    ''' <summary>
    ''' Morre na página 1, retoma, termina: o resultado é o mesmo de uma
    ''' execução que não morreu. Nada perdido, nada duplicado.
    '''
    ''' É o critério que o Codex acrescentou ao pronto do 2.1: importação
    ''' interrompida seguida de varredura limpa converge para o mesmo
    ''' manifesto de uma execução ininterrupta.
    ''' </summary>
    <TestMethod>
    Public Sub Retomada_apos_crash_converge_para_o_mesmo_manifesto()
        Dim r1 = RodarHarness(CrashInjection.DepoisDoCommitDaPagina)
        Assert.AreNotEqual(0, r1.ExitCode)
        Dim tentativa = TentativaDe(r1.Stdout)

        Dim r2 = RodarHarness("nenhum", retomar:=tentativa)
        Assert.AreEqual(0, r2.ExitCode, r2.Stderr)

        Dim comCrash = Manifesto()

        ' Agora o mesmo trabalho, do zero, sem morrer nenhuma vez.
        Dim limpo = ManifestoDeExecucaoLimpa()

        CollectionAssert.AreEqual(limpo.ToList(), comCrash.ToList(),
            "retomada apos crash tem de dar o MESMO manifesto de uma execucao ininterrupta")
        Assert.AreEqual(TotalLinhas, comCrash.Count)
    End Sub

    ''' <summary>
    ''' A página reexecutada não duplica. Sem isto, retomar contaria a
    ''' página 1 duas vezes e o S6 rejeitaria a varredura inteira por um
    ''' sintoma sem relação nenhuma com a causa.
    ''' </summary>
    <TestMethod>
    Public Sub Retomada_nao_duplica_nem_infla_a_contagem()
        Dim r1 = RodarHarness(CrashInjection.DepoisDoCommitDaPagina)
        Dim tentativa = TentativaDe(r1.Stdout)
        RodarHarness("nenhum", retomar:=tentativa)

        Comparar(Sub(c)
                     Assert.AreEqual(TotalLinhas, Contar(c, "scan_stage"))
                     Assert.AreEqual(TotalLinhas, Contar(c, "incarnation"))
                     Assert.AreEqual(TotalLinhas, Contar(c, "item"))
                     Assert.AreEqual(TotalLinhas, Contar(c, "metadata_observation"))
                     Assert.AreEqual(CLng(TotalLinhas),
                        Convert.ToInt64(Valor(c, "SELECT rows_read FROM generation")))
                     Assert.AreEqual(CLng(TotalLinhas),
                        Convert.ToInt64(Valor(c, "SELECT distinct_keys FROM generation")))
                 End Sub)
    End Sub

    ' ==================================================================
    ' CONTROLE NEGATIVO

    ''' <summary>
    ''' O controle negativo, e é ele que dá sentido a todos os anteriores.
    '''
    ''' Um teste que só confirma "depois do crash está tudo consistente" passa
    ''' igualzinho num escritor correto e num que não grava nada. Aqui eu ligo
    ''' o defeito — avançar o checkpoint numa transação própria ANTES de
    ''' gravar as linhas, que é o desenho ingênuo "salva o progresso primeiro"
    ''' — e cobro que o mesmo cenário PERCA a página 1.
    '''
    ''' Perda silenciosa: sem erro, sem log, e invisível para qualquer
    ''' contagem que confie no cursor.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_negativo_checkpoint_antes_das_linhas_PERDE_mensagens()
        ' --- com o defeito ligado ---
        Dim r1 = RodarHarness(CrashInjection.DentroDaPaginaAntesDoCommit, defeito:="checkpoint-antes")
        Assert.AreNotEqual(0, r1.ExitCode)
        Dim tentativa = TentativaDe(r1.Stdout)

        Comparar(Sub(c)
                     Assert.AreEqual(0, Contar(c, "scan_stage"), "as linhas nao foram gravadas")
                     Assert.AreEqual("cursor-1", Texto(c, "SELECT cursor FROM scan_attempt"),
                        "mas o checkpoint AVANCOU — e e exatamente esse o defeito")
                 End Sub)

        RodarHarness("nenhum", retomar:=tentativa)
        Dim comDefeito = Manifesto()

        Assert.AreEqual(TotalLinhas - 2, comDefeito.Count,
            "a retomada comeca na pagina 2 e a pagina 1 se perde para sempre")
        Assert.IsFalse(comDefeito.Contains("E-1-1"), "E-1-1 sumiu sem deixar rastro")

        ' --- o MESMO cenario, sem o defeito ---
        Limpar() : Preparar()
        Dim r2 = RodarHarness(CrashInjection.DentroDaPaginaAntesDoCommit)
        Dim t2 = TentativaDe(r2.Stdout)
        RodarHarness("nenhum", retomar:=t2)
        Dim semDefeito = Manifesto()

        Assert.AreEqual(TotalLinhas, semDefeito.Count, "sem o defeito nao se perde nada")
        Assert.IsTrue(semDefeito.Contains("E-1-1"))
    End Sub

    ' ==================================================================
    ' Critério 10: geração velha não sobrescreve geração nova

    ''' <summary>
    ''' Duas varreduras da mesma pasta. A que ABRIU primeiro termina por
    ''' último, e tem de ser RECUSADA.
    '''
    ''' A armadilha do critério 10 está na chave usada para ordenar.
    ''' <c>generation_key</c> é atribuída no INSERT, que só acontece ao
    ''' PUBLICAR — então a varredura velha que termina tarde recebe a chave
    ''' MAIOR, e um teste de monotonicidade sobre ela aprovaria exatamente o
    ''' caso que deveria barrar. <c>attempt_key</c> é atribuída ao ABRIR, e é
    ''' essa a ordem que diz qual leitura é mais antiga.
    ''' </summary>
    <TestMethod>
    Public Sub Geracao_velha_terminando_tarde_e_RECUSADA()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Dim w As New CacheWriter(db)

            Dim velha = w.AbrirTentativa(FolderKey, 1, "U", 0, 1)
            Dim nova = w.AbrirTentativa(FolderKey, 1, "U", 0, 2)
            Assert.IsTrue(nova > velha)

            w.GravarPagina(nova, FolderKey, 1, Linhas("N-1", "N-2"), "c1")
            Assert.AreEqual(PublishOutcome.Publicada, w.Publicar(nova, FolderKey, "completa", 2, 2))
            Dim cabecaBoa = w.CabecaPublicada(FolderKey)

            ' A velha so agora termina.
            w.GravarPagina(velha, FolderKey, 1, Linhas("V-1"), "c1")
            Assert.AreEqual(PublishOutcome.RecusadaPorOrdem,
                            w.Publicar(velha, FolderKey, "completa", 1, 1))

            Assert.AreEqual(cabecaBoa, w.CabecaPublicada(FolderKey),
                            "a cabeca nao pode ter recuado")
            Assert.AreEqual("descartada", w.EstagioDa(velha))
            Assert.AreEqual("ordem", Texto(db.Connection,
                $"SELECT rejection FROM scan_attempt WHERE attempt_key={velha}"))
        End Using
    End Sub

    ''' <summary>
    ''' Controle negativo do critério 10: se a ordem fosse por
    ''' <c>generation_key</c>, a velha passaria.
    '''
    ''' Não é hipótese — é medido aqui. A velha, se publicasse, receberia uma
    ''' generation_key MAIOR que a da nova. Qualquer guarda escrita sobre essa
    ''' chave aprovaria a sobrescrita.
    ''' </summary>
    <TestMethod>
    Public Sub Ordenar_por_generation_key_aprovaria_a_velha()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim w As New CacheWriter(db)
            Dim velha = w.AbrirTentativa(FolderKey, 1, "U", 0, 1)
            Dim nova = w.AbrirTentativa(FolderKey, 1, "U", 0, 2)

            w.GravarPagina(nova, FolderKey, 1, Linhas("N-1"), "c1")
            Dim gNova As Long = 0
            w.Publicar(nova, FolderKey, "completa", 1, 1, gNova)

            ' O proximo rowid que a velha receberia se publicasse.
            Dim proximo = Convert.ToInt64(Valor(db.Connection,
                "SELECT COALESCE(MAX(generation_key),0)+1 FROM generation"))

            Assert.IsTrue(proximo > gNova,
                "a velha receberia a chave MAIOR — por isso a ordem nao pode vir daqui")
            Assert.IsTrue(velha < nova,
                "e a ordem de ABERTURA, que e a correta, diz o contrario")
        End Using
    End Sub

    ''' <summary>
    ''' Época: o universo mudou embaixo da varredura, e ela é descartada.
    ''' </summary>
    <TestMethod>
    Public Sub Universo_invalidado_no_meio_recusa_a_publicacao()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim w As New CacheWriter(db)
            Dim t = w.AbrirTentativa(FolderKey, 1, "U", 0, 1)
            w.GravarPagina(t, FolderKey, 1, Linhas("A-1"), "c1")

            w.InvalidarUniverso(FolderKey)

            Assert.AreEqual(PublishOutcome.RecusadaPorEpoca, w.Publicar(t, FolderKey, "completa", 1, 1))
            Assert.IsNull(w.CabecaPublicada(FolderKey))
            Assert.AreEqual(0, Contar(db.Connection, "generation"))
            Assert.AreEqual(0, Contar(db.Connection, "publication_log"),
                            "sem geracao nao pode haver divida para a UI")
        End Using
    End Sub

    ''' <summary>Publicar duas vezes a mesma tentativa não cria duas gerações.</summary>
    <TestMethod>
    Public Sub Publicar_duas_vezes_e_recusado()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim w As New CacheWriter(db)
            Dim t = w.AbrirTentativa(FolderKey, 1, "U", 0, 1)
            w.GravarPagina(t, FolderKey, 1, Linhas("A-1"), "c1")

            Assert.AreEqual(PublishOutcome.Publicada, w.Publicar(t, FolderKey, "completa", 1, 1))
            Assert.AreEqual(PublishOutcome.RecusadaPorEstado, w.Publicar(t, FolderKey, "completa", 1, 1))
            Assert.AreEqual(1, Contar(db.Connection, "generation"))
            Assert.AreEqual(1, Contar(db.Connection, "publication_log"))
        End Using
    End Sub

    ' ==================================================================
    ' Infraestrutura

    Private Structure Resultado
        Public ExitCode As Integer
        Public Stdout As String
        Public Stderr As String
    End Structure

    Private Function RodarHarness(ponto As String,
                                  Optional retomar As Long = 0,
                                  Optional defeito As String = "") As Resultado
        Dim exe = LocalizarHarness()
        Dim psi As New ProcessStartInfo(exe) With {
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .UseShellExecute = False,
            .CreateNoWindow = True}
        For Each a In {_db, FolderKey.ToString(), ponto, "kill", retomar.ToString(), defeito}
            psi.ArgumentList.Add(a)
        Next

        Using p = Process.Start(psi)
            Dim o = p.StandardOutput.ReadToEnd()
            Dim e = p.StandardError.ReadToEnd()
            If Not p.WaitForExit(60000) Then
                p.Kill(True)
                Assert.Fail("o harness travou")
            End If

            ' Sem isto, "saiu com codigo != 0" cobre tanto "morreu no ponto
            ' pedido" quanto "explodiu antes de comecar" — e foi assim que um
            ' schema quebrado passou por quatro testes de crash se dizendo
            ' verificado. Excecao nao tratada NAO e o crash que eu pedi.
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
        Assert.IsNotNull(d, "nao achei a raiz do repositorio a partir de " & AppContext.BaseDirectory)

        Dim raiz = Path.Combine(d.FullName, "tools", "Iris.CrashHarness", "bin")
        Assert.IsTrue(Directory.Exists(raiz),
            "o harness de crash nao foi compilado. Rode: dotnet build Iris.slnx")

        Dim achado = Directory.GetFiles(raiz, "Iris.CrashHarness.exe", SearchOption.AllDirectories).
                     OrderByDescending(Function(f) File.GetLastWriteTimeUtc(f)).FirstOrDefault()
        Assert.IsNotNull(achado, "Iris.CrashHarness.exe nao encontrado sob " & raiz)

        _exe = achado
        Return _exe
    End Function

    Private Shared Function TentativaDe(stdout As String) As Long
        Dim linha = stdout.Split(ControlChars.Lf).FirstOrDefault(
            Function(l) l.StartsWith("tentativa=", StringComparison.Ordinal))
        Assert.IsNotNull(linha, "o harness nao informou a tentativa. stdout: " & stdout)
        Return Long.Parse(linha.Trim().Substring("tentativa=".Length))
    End Function

    ''' <summary>Os EntryIDs guardados, ordenados — o "manifesto".</summary>
    Private Function Manifesto() As List(Of String)
        Dim r As New List(Of String)()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText = "SELECT provider_entry_id FROM incarnation ORDER BY provider_entry_id"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        r.Add(rd.GetString(0))
                    End While
                End Using
            End Using
        End Using
        SqliteConnection.ClearAllPools()
        Return r
    End Function

    Private Function ManifestoDeExecucaoLimpa() As List(Of String)
        Dim outro = Path.Combine(_pasta, "limpo.db")
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(outro, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Exec(db.Connection, "INSERT INTO environment_profile (environment_key, fingerprint, " &
                                "provider, cached_mode, policy_version, allowed) " &
                                "VALUES (1,'fp','teste',1,1,1)")
            Exec(db.Connection, "INSERT INTO store (store_key, provider_store_id) VALUES (1,'S')")
            Exec(db.Connection, "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
                                "reconcile_epoch, stability) VALUES (1,1,'F',0,'estavel')")
        End Using
        SqliteConnection.ClearAllPools()

        Dim guardado = _db
        _db = outro
        Try
            Dim r = RodarHarness("nenhum")
            Assert.AreEqual(0, r.ExitCode, r.Stderr)
            Return Manifesto()
        Finally
            _db = guardado
        End Try
    End Function

    Private Sub Comparar(verificar As Action(Of SqliteConnection))
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"reabrir apos o crash: {falha}")
            verificar(db.Connection)
        End Using
        SqliteConnection.ClearAllPools()
    End Sub

    Private Shared Function Linhas(ParamArray ids As String()) As IReadOnlyList(Of StagedRow)
        Return ids.Select(Function(x) New StagedRow With {
            .ProviderEntryId = x, .Subject = "s", .MessageClass = "IPM.Note"}).ToList()
    End Function

    Private Shared Sub Exec(c As SqliteConnection, sql As String)
        Using cmd = c.CreateCommand()
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Function Valor(c As SqliteConnection, sql As String) As Object
        Using cmd = c.CreateCommand()
            cmd.CommandText = sql
            Dim v = cmd.ExecuteScalar()
            Return If(v Is Nothing OrElse v Is DBNull.Value, Nothing, v)
        End Using
    End Function

    Private Shared Function Texto(c As SqliteConnection, sql As String) As String
        Dim v = Valor(c, sql)
        Return If(v Is Nothing, Nothing, Convert.ToString(v))
    End Function

    Private Shared Function Contar(c As SqliteConnection, tabela As String) As Integer
        Return Convert.ToInt32(Valor(c, $"SELECT COUNT(*) FROM {tabela}"))
    End Function

End Class

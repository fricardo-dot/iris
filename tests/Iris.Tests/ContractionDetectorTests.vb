Imports System.IO
Imports System.Linq
Imports System.Threading
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Sync
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Marco 2.3 — o detector de contração.
'''
''' É o que sobrou da Q8 depois de dois resultados negativos: a janela de
''' sincronização não é legível (§22.4) e a contagem do servidor não vem por
''' <c>PropertyAccessor</c> (§22.11). Sem fonte externa do universo, a única
''' referência disponível é o que o próprio Iris já guardou.
'''
''' <b>Ele invalida, nunca autoriza.</b> Os testes abaixo cobram as duas
''' coisas: que ele acuse quando encolhe, e que ele <b>não</b> transforme
''' histórico em cobertura.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class ContractionDetectorTests

    Private _pasta As String
    Private _db As String

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-contr-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_pasta)
        _db = Path.Combine(_pasta, "cache.db")

        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
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

    Private Shared Function Universo() As SweepUniverse
        Return New SweepUniverse("S", "F", "todos", Nothing, 1, "amb")
    End Function

    Private Shared Function Real() As EnvironmentCapabilities
        Return EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
    End Function

    Private Shared Sub Varrer(db As CacheDatabase, numero As Integer, ParamArray chaves As String())
        Dim r = New SweepRunner(New FonteFalsaMutavel(Universo(), chaves),
                                New SqliteSweepSink(db, 1, 1), 10).
                Executar(Universo(), 0, numero, Real(), CancellationToken.None)
        Assert.IsTrue(r.Publicou, $"varredura {numero} deveria publicar: {r.Motivo}")
    End Sub

    ' ==================================================================

    <TestMethod>
    Public Sub Primeira_geracao_nao_tem_com_o_que_comparar()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Varrer(db, 1, "a", "b", "c")
            Dim r = New ContractionDetector(db).Comparar(1)
            Assert.AreEqual(ContractionVerdict.SemReferencia, r.Verdict)
            Assert.IsNull(r.Aviso, "sem referencia nao ha o que avisar")
        End Using
    End Sub

    <TestMethod>
    Public Sub Alcance_estavel_nao_acusa()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Varrer(db, 1, "a", "b", "c")
            Varrer(db, 2, "a", "b", "c")
            Dim r = New ContractionDetector(db).Comparar(1)
            Assert.AreEqual(ContractionVerdict.Estavel, r.Verdict)
            Assert.IsNull(r.Aviso)
        End Using
    End Sub

    ''' <summary>
    ''' Alcance que ENCOLHE é acusado, e o aviso diz o que importa: as
    ''' mensagens que sumiram não foram necessariamente apagadas.
    ''' </summary>
    <TestMethod>
    Public Sub Alcance_que_encolhe_e_ACUSADO()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Varrer(db, 1, "a", "b", "c", "d")
            Varrer(db, 2, "a", "b")

            Dim r = New ContractionDetector(db).Comparar(1)

            Assert.AreEqual(ContractionVerdict.Encolheu, r.Verdict)
            Assert.AreEqual(4, r.AlcanceAntes)
            Assert.AreEqual(2, r.AlcanceAgora)
            CollectionAssert.AreEquivalent({"c", "d"}, r.Sumiram.ToList())
            StringAssert.Contains(r.Aviso, "nao foram necessariamente apagadas".
                                            Replace("nao", "não"))
        End Using
    End Sub

    ''' <summary>
    ''' <b>Encolhimento COMPENSADO por correio novo.</b>
    '''
    ''' Sai um, entra outro: a contagem não se mexe. Um detector que comparasse
    ''' contagens passaria batido — foi um dos buracos que a §22.11 listou, e é
    ''' o motivo de este comparar CONJUNTOS.
    ''' </summary>
    <TestMethod>
    Public Sub Encolhimento_compensado_por_correio_novo_e_ACUSADO()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Varrer(db, 1, "a", "b", "c")
            Varrer(db, 2, "a", "b", "z")   ' "c" sumiu, "z" chegou: 3 e 3

            Dim r = New ContractionDetector(db).Comparar(1)

            Assert.AreEqual(ContractionVerdict.Encolheu, r.Verdict,
                "comparar CONTAGENS passaria batido: 3 antes, 3 agora")
            Assert.AreEqual(3, r.AlcanceAntes)
            Assert.AreEqual(3, r.AlcanceAgora)
            CollectionAssert.AreEquivalent({"c"}, r.Sumiram.ToList())
            Assert.AreEqual(1, r.Chegaram)
        End Using
    End Sub

    ''' <summary>
    ''' Só correio novo chegando NÃO é contração.
    '''
    ''' O contraponto: sem ele, um detector que acusasse qualquer diferença
    ''' passaria no teste acima e avisaria o usuário toda vez que chegasse
    ''' e-mail.
    ''' </summary>
    <TestMethod>
    Public Sub So_correio_novo_NAO_e_contracao()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Varrer(db, 1, "a", "b")
            Varrer(db, 2, "a", "b", "novo")

            Dim r = New ContractionDetector(db).Comparar(1)

            Assert.AreEqual(ContractionVerdict.Estavel, r.Verdict,
                "chegar e-mail nao e o Iris enxergar menos")
            Assert.AreEqual(1, r.Chegaram)
            Assert.AreEqual(0, r.Sumiram.Count)
        End Using
    End Sub

    ''' <summary>
    ''' <b>O buraco que ele NÃO fecha, e que precisa estar escrito como teste.</b>
    '''
    ''' Se a janela sempre foi pequena, não há contração para detectar: o Iris
    ''' nunca viu o que falta. É o estado da caixa do usuário hoje — 1.013
    ''' itens alcançáveis numa Caixa de Entrada que o servidor diz ter 17.728.
    '''
    ''' O detector diz <c>Estavel</c>, e está certo: nada encolheu. Mas
    ''' <c>Estavel</c> não é <c>Completa</c>, e confundir os dois seria
    ''' exatamente transformar a derrota da Q8 em recurso.
    ''' </summary>
    <TestMethod>
    Public Sub Janela_sempre_pequena_nao_tem_contracao_a_detectar()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            ' O Iris SEMPRE viu so "a" e "b". Que existam outros mil no
            ' servidor, ele nao tem como saber.
            Varrer(db, 1, "a", "b")
            Varrer(db, 2, "a", "b")

            Dim r = New ContractionDetector(db).Comparar(1)
            Assert.AreEqual(ContractionVerdict.Estavel, r.Verdict)

            ' E estavel NAO virou cobertura completa.
            Dim m = New ManifestReader(db).Ler(1)
            Assert.AreEqual(FolderCoverage.Parcial, m.Cobertura,
                "historico estavel NAO promove cobertura: isso seria transformar a " &
                "derrota da Q8 em recurso")
            Assert.IsFalse(m.EhEstadoCorrente)
        End Using
    End Sub

    ''' <summary>
    ''' O detector lê o que cada geração VIU, não o estado corrente da
    ''' associação.
    '''
    ''' Se lesse a associação, a resposta seria sempre a de agora — a geração
    ''' seguinte já sobrescreveu — e ele nunca acusaria nada. É o tipo de
    ''' detector que passa em todo teste de "não acusa falso positivo" e falha
    ''' o único que importa.
    ''' </summary>
    <TestMethod>
    Public Sub Le_o_que_a_geracao_VIU_e_nao_o_estado_corrente()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Varrer(db, 1, "a", "b", "c")
            Varrer(db, 2, "a")

            ' O estado corrente diz que b e c estao SUSPEITOS, e continuam na
            ' tabela association. Um detector que olhasse ali veria tres itens
            ' nas duas geracoes.
            Dim m = New ManifestReader(db).Ler(1)
            Assert.AreEqual(3, m.Items.Count, "o acervo mantem os tres")
            Assert.AreEqual(2, m.Items.Where(Function(i) i.Presence = PresenceState.Suspeito).Count())

            Dim r = New ContractionDetector(db).Comparar(1)
            Assert.AreEqual(ContractionVerdict.Encolheu, r.Verdict,
                "olhar a associacao daria 3 e 3, e o detector nunca acusaria")
            Assert.AreEqual(3, r.AlcanceAntes)
            Assert.AreEqual(1, r.AlcanceAgora)
        End Using
    End Sub

    ' ==================================================================
    ' O detector tem um CONSUMIDOR

    ''' <summary>
    ''' A contração aparece na ressalva do manifesto — que é o que a UI mostra.
    '''
    ''' Sem isto o detector calculava um veredito que ninguém lia, e a §25
    ''' afirmava que "encolheu → as conclusões anteriores caem" quando o que
    ''' acontecia era "o detector consegue devolver um diagnóstico, se alguém
    ''' o chamar".
    '''
    ''' E o efeito é <b>aviso</b>, não invalidação: não há conclusão a
    ''' invalidar, porque em cached a cobertura já é sempre parcial e ausência
    ''' já é proibida (§23).
    ''' </summary>
    <TestMethod>
    Public Sub A_contracao_aparece_na_ressalva_do_manifesto()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Varrer(db, 1, "a", "b", "c", "d")
            Varrer(db, 2, "a", "b")

            Dim m = New ManifestReader(db).Ler(1)

            Assert.IsNotNull(m.Contracao, "o manifesto tem de carregar o veredito")
            Assert.AreEqual(ContractionVerdict.Encolheu, m.Contracao.Verdict)
            StringAssert.Contains(m.Ressalva, "passou a enxergar menos",
                "a contracao tem de chegar na ressalva, que e o que a UI mostra")
            StringAssert.Contains(m.Ressalva, "Acervo parcial",
                "e a ressalva de cobertura continua junto")
        End Using
    End Sub

    ''' <summary>
    ''' E sem contração a ressalva não ganha ruído: só o que já havia.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_contracao_a_ressalva_nao_ganha_ruido()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Varrer(db, 1, "a", "b")
            Varrer(db, 2, "a", "b", "novo")

            Dim m = New ManifestReader(db).Ler(1)

            Assert.AreEqual(ContractionVerdict.Estavel, m.Contracao.Verdict)
            Assert.IsFalse(m.Ressalva.Contains("passou a enxergar menos"),
                "chegar e-mail nao pode virar aviso de contracao")
        End Using
    End Sub

    ' ==================================================================

    Private Shared Sub Exec(db As CacheDatabase, sql As String)
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
        End Using
    End Sub

End Class

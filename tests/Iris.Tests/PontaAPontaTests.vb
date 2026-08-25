Imports System.Collections.Generic
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
''' Orquestração real contra persistência real, sem COM.
'''
''' Os testes do <see cref="SweepRunnerTests"/> provam a orquestração contra um
''' destino falso; os do <see cref="CacheDatabaseTests"/> provam a persistência
''' contra o banco. Nenhum dos dois prova que as duas metades <b>se
''' encaixam</b> — e é justamente na junção que as suposições de cada lado
''' aparecem: o que o runner chama de página, o que o writer chama de
''' tentativa, quem decide a época.
'''
''' Aqui roda o <see cref="SweepRunner"/> de verdade sobre o
''' <see cref="SqliteSweepSink"/> de verdade, num arquivo SQLite de verdade.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class PontaAPontaTests

    Private _pasta As String
    Private _db As String
    Private Const FolderKey As Long = 1
    Private Const EnvKey As Long = 1

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-e2e-" & Guid.NewGuid().ToString("N"))
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

    Private Shared Function Universo() As SweepUniverse
        Return New SweepUniverse("S", "F", "todos", Nothing, 1, "amb-1")
    End Function

    ''' <summary>O ambiente REAL do usuário: cached, janela não legível (§23).</summary>
    Private Shared Function Real() As EnvironmentCapabilities
        Return EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>O teste que fecha o 2.2a.</b> Varre, encena, publica, e a UI drena.
    '''
    ''' E roda com as capacidades REAIS do usuário — cached sem janela legível
    ''' —, então prova a §23 na prática: o produto opera, e opera declarando
    ''' cobertura parcial.
    ''' </summary>
    <TestMethod>
    Public Sub Varre_publica_e_a_UI_drena()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")

            Dim fonte As New FonteFalsaMutavel(Universo(), "E-1", "E-2", "E-3", "E-4", "E-5")
            Dim sink As New SqliteSweepSink(db, FolderKey, EnvKey)
            Dim runner As New SweepRunner(fonte, sink, tamanhoPagina:=2)

            Dim r = runner.Executar(Universo(), 0, 1, Real(), CancellationToken.None)

            Assert.IsTrue(r.Publicou, $"deveria publicar. motivo: {r.Motivo}")
            Assert.AreEqual(FolderCoverage.Parcial, r.Cobertura,
                "cached sem janela legivel publica PARCIAL (§23)")

            ' --- o que ficou no banco ---
            Assert.AreEqual(5, Contar(db, "incarnation"))
            Assert.AreEqual(5, Contar(db, "metadata_observation"))
            Assert.AreEqual(5, Contar(db, "association"))
            Assert.AreEqual(1, Contar(db, "generation"))
            Assert.AreEqual("publicada", Texto(db, "SELECT stage FROM scan_attempt"))

            ' O ALCANCE ficou registrado como parcial, e a geracao aponta para ele.
            Assert.AreEqual("parcial",
                Texto(db, "SELECT coverage FROM coverage_observation"),
                "o alcance parcial tem de estar no banco, nao so no resultado em memoria")
            Assert.AreEqual("completa",
                Texto(db, "SELECT coverage_kind FROM generation"),
                "o TIPO da varredura foi completo — ela percorreu a pasta inteira")
            Assert.IsNotNull(Valor(db, "SELECT coverage_key FROM generation"),
                "a geracao tem de apontar para a observacao de alcance")

            ' --- a UI drena ---
            Dim drain As New PublicationDrain(db)
            Assert.AreEqual(1, drain.Pendentes().Count, "a divida tem de estar pendente")

            Dim ui As New ConsumidorFalso()
            Assert.AreEqual(1, drain.Drenar(ui))
            Assert.AreEqual(1, ui.Recebidas.Count)
            Assert.AreEqual(0, drain.Pendentes().Count, "drenou, some")
        End Using
    End Sub

    ''' <summary>
    ''' Varredura rejeitada não deixa tentativa órfã no banco.
    '''
    ''' Sem isto, cada rejeição deixaria uma linha <c>varrendo</c> que a
    ''' retomada seguinte encontraria e trataria como trabalho a continuar.
    ''' </summary>
    <TestMethod>
    Public Sub Varredura_rejeitada_descarta_a_tentativa_no_banco()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim fonte As New FonteFalsaMutavel(Universo(), "E-1", "E-2", "E-3", "E-4")
            fonte.Agenda(1) = Sub(x) x.Remover("E-4")

            Dim runner As New SweepRunner(fonte, New SqliteSweepSink(db, FolderKey, EnvKey), 2)
            Dim r = runner.Executar(Universo(), 0, 1, Real(), CancellationToken.None)

            Assert.IsFalse(r.Publicou)
            Assert.AreEqual("descartada", Texto(db, "SELECT stage FROM scan_attempt"))
            Assert.IsNotNull(Texto(db, "SELECT rejection FROM scan_attempt"),
                "o motivo tem de ficar no banco para quem for investigar")
            Assert.AreEqual(0, Contar(db, "generation"))
            Assert.AreEqual(0, Contar(db, "publication_log"))
        End Using
    End Sub

    ''' <summary>
    ''' Duas varreduras seguidas: a segunda não duplica nada, e a cabeça avança.
    ''' </summary>
    <TestMethod>
    Public Sub Segunda_varredura_nao_duplica_e_a_cabeca_avanca()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim sink As New SqliteSweepSink(db, FolderKey, EnvKey)

            Dim r1 = New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1", "E-2", "E-3"),
                                     sink, 2).Executar(Universo(), 0, 1, Real(), CancellationToken.None)
            Assert.IsTrue(r1.Publicou, r1.Motivo)
            Dim cabeca1 = Convert.ToInt64(Valor(db, "SELECT published_generation_key FROM folder"))

            Dim r2 = New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1", "E-2", "E-3"),
                                     sink, 2).Executar(Universo(), 0, 2, Real(), CancellationToken.None)
            Assert.IsTrue(r2.Publicou, r2.Motivo)
            Dim cabeca2 = Convert.ToInt64(Valor(db, "SELECT published_generation_key FROM folder"))

            Assert.AreEqual(3, Contar(db, "incarnation"), "a segunda varredura nao duplica")
            Assert.AreEqual(3, Contar(db, "item"))
            Assert.AreEqual(2, Contar(db, "generation"))
            Assert.AreNotEqual(cabeca1, cabeca2, "a cabeca tem de avancar")
            Assert.AreEqual(2, Contar(db, "publication_log"))
        End Using
    End Sub

    ''' <summary>
    ''' Consumidor que lança INTERROMPE o dreno e deixa tudo pendente.
    '''
    ''' Engolir a exceção e seguir marcaria como drenada uma geração que a UI
    ''' não recebeu — perder em silêncio o que este desenho existe para não
    ''' perder.
    ''' </summary>
    <TestMethod>
    Public Sub Consumidor_que_lanca_nao_marca_drenada()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim sink As New SqliteSweepSink(db, FolderKey, EnvKey)
            Dim runner As New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1"), sink, 2)
            Dim publicou = runner.Executar(Universo(), 0, 1, Real(), CancellationToken.None)
            Assert.IsTrue(publicou.Publicou, publicou.Motivo)

            Dim drain As New PublicationDrain(db)
            Dim ui As New ConsumidorFalso() With {.Lancar = True}

            Assert.ThrowsException(Of InvalidOperationException)(Sub() drain.Drenar(ui))
            Assert.AreEqual(1, drain.Pendentes().Count,
                "a divida continua pendente: a UI nao recebeu")
        End Using
    End Sub

    ''' <summary>
    ''' A entrega é AO MENOS UMA VEZ: drenar de novo sem marcar repete.
    '''
    ''' É a consequência da ordem escolhida — consumir antes de marcar. Está
    ''' aqui como teste porque é obrigação transferida ao consumidor, e
    ''' obrigação transferida em comentário é obrigação esquecida.
    ''' </summary>
    <TestMethod>
    Public Sub A_entrega_e_ao_menos_uma_vez()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim sink As New SqliteSweepSink(db, FolderKey, EnvKey)
            Dim runner As New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1"), sink, 2)
            Dim publicou = runner.Executar(Universo(), 0, 1, Real(), CancellationToken.None)
            Assert.IsTrue(publicou.Publicou, publicou.Motivo)

            Dim drain As New PublicationDrain(db)

            ' Uma UI que recebe e morre antes de a marcacao acontecer.
            Dim primeira As New ConsumidorFalso() With {.LancarDepoisDeReceber = True}
            Try
                drain.Drenar(primeira)
            Catch ex As InvalidOperationException
            End Try
            Assert.AreEqual(1, primeira.Recebidas.Count)

            ' Na volta, a mesma geracao e entregue OUTRA VEZ.
            Dim segunda As New ConsumidorFalso()
            Assert.AreEqual(1, drain.Drenar(segunda))
            Assert.AreEqual(primeira.Recebidas(0), segunda.Recebidas(0),
                "a MESMA geracao foi entregue duas vezes — por isso o consumidor tem de ser idempotente")
        End Using
    End Sub

    ' ==================================================================
    ' O manifesto: nao da para pegar os itens sem pegar a cobertura

    ''' <summary>
    ''' O manifesto entrega itens E cobertura no mesmo objeto, e a ressalva
    ''' escrita — a §23 virando estrutura em vez de recomendacao.
    ''' </summary>
    <TestMethod>
    Public Sub O_manifesto_declara_que_e_ACERVO_e_nao_estado_corrente()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim sink As New SqliteSweepSink(db, FolderKey, EnvKey)
            Dim runner As New SweepRunner(
                New FonteFalsaMutavel(Universo(), "E-1", "E-2", "E-3"), sink, 2)
            Assert.IsTrue(runner.Executar(Universo(), 0, 1, Real(), CancellationToken.None).Publicou)

            Dim m = New ManifestReader(db).Ler(FolderKey)

            Assert.AreEqual(3, m.Items.Count)
            Assert.AreEqual(FolderCoverage.Parcial, m.Cobertura)
            Assert.IsFalse(m.EhEstadoCorrente,
                "cobertura parcial nao pode ser exibida como o estado corrente da pasta")
            StringAssert.Contains(m.Ressalva, "parcial")
            Assert.IsNotNull(m.GenerationKey)
        End Using
    End Sub

    ''' <summary>
    ''' Pasta nunca varrida devolve manifesto vazio com ressalva propria — nao
    ''' devolve nulo nem lista vazia sem explicacao.
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_nunca_varrida_diz_que_nunca_foi_varrida()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim m = New ManifestReader(db).Ler(FolderKey)
            Assert.AreEqual(0, m.Items.Count)
            Assert.IsNull(m.GenerationKey)
            Assert.IsFalse(m.EhEstadoCorrente)
            StringAssert.Contains(m.Ressalva, "nao foi varrida".Replace("nao", "não"))
        End Using
    End Sub

    ''' <summary>
    ''' Linhas de uma tentativa REJEITADA não aparecem no manifesto — e agora
    ''' nem chegam ao acervo.
    '''
    ''' A primeira versão deste teste afirmava o contrário: que o
    ''' <c>GravarPagina</c> escrevia em <c>incarnation</c> antes da publicação
    ''' e que a fronteira do <c>generation_key</c> segurava o vazamento. A
    ''' fronteira segurava — mas só para itens NOVOS, e era esse o buraco.
    ''' </summary>
    <TestMethod>
    Public Sub Tentativa_rejeitada_nao_toca_o_acervo()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim fonte As New FonteFalsaMutavel(Universo(), "E-1", "E-2", "E-3", "E-4")
            fonte.Agenda(1) = Sub(x) x.Remover("E-4")
            Dim runner As New SweepRunner(fonte, New SqliteSweepSink(db, FolderKey, EnvKey), 2)
            Assert.IsFalse(runner.Executar(Universo(), 0, 1, Real(), CancellationToken.None).Publicou)

            Assert.AreEqual(0, Contar(db, "incarnation"),
                "a pagina so ENCENA: o acervo nao pode ter sido tocado")
            Assert.AreEqual(0, Contar(db, "association"))
            Assert.IsTrue(Contar(db, "scan_stage") > 0, "mas as linhas ficaram encenadas")

            Dim m = New ManifestReader(db).Ler(FolderKey)
            Assert.AreEqual(0, m.Items.Count)
            Assert.IsNull(m.GenerationKey)
        End Using
    End Sub

    ''' <summary>
    ''' <b>O teste que faltava, e que o buraco exigia.</b>
    '''
    ''' Uma tentativa rejeitada não pode alterar o que JÁ ESTAVA PUBLICADO.
    '''
    ''' Era exatamente aqui que o defeito vivia, e o meu teste anterior não
    ''' chegava perto: ele só usava itens novos, cujas associações ainda não
    ''' pertencem a geração nenhuma, então a fronteira do
    ''' <c>generation_key</c> as escondia. Com um item <b>preexistente</b>, o
    ''' <c>GravarPagina</c> substituía o metadado publicado e devolvia a
    ''' associação para <c>presente</c> — e rejeitar depois não desfazia nada.
    ''' Uma varredura recusada alterava o manifesto que a UI mostra.
    ''' </summary>
    <TestMethod>
    Public Sub Tentativa_rejeitada_nao_altera_item_JA_PUBLICADO()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim sink As New SqliteSweepSink(db, FolderKey, EnvKey)

            ' 1. Uma geracao publicada, com assunto conhecido.
            Dim r1 = New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1", "E-2"), sink, 2).
                     Executar(Universo(), 0, 1, Real(), CancellationToken.None)
            Assert.IsTrue(r1.Publicou, r1.Motivo)

            Dim antes = New ManifestReader(db).Ler(FolderKey)
            Assert.AreEqual(2, antes.Items.Count)
            Dim assuntoAntes = antes.Items.First(Function(i) i.ProviderEntryId = "E-1").Subject
            Dim geracaoAntes = antes.GenerationKey

            ' 2. Uma segunda tentativa que le os MESMOS itens e e REJEITADA.
            Dim fonte2 As New FonteFalsaMutavel(Universo(), "E-1", "E-2", "E-3")
            fonte2.Agenda(1) = Sub(x) x.Remover("E-3")
            Dim r2 = New SweepRunner(fonte2, sink, 2).
                     Executar(Universo(), 0, 2, Real(), CancellationToken.None)
            Assert.IsFalse(r2.Publicou, "a segunda tem de ser rejeitada")

            ' 3. O manifesto tem de estar INTACTO.
            Dim depois = New ManifestReader(db).Ler(FolderKey)
            Assert.AreEqual(2, depois.Items.Count, "a rejeitada nao pode acrescentar item")
            Assert.AreEqual(geracaoAntes, depois.GenerationKey, "a cabeca nao pode ter mudado")
            Assert.AreEqual(assuntoAntes,
                depois.Items.First(Function(i) i.ProviderEntryId = "E-1").Subject,
                "a rejeitada nao pode ter substituido o metadado publicado")
            Assert.IsFalse(depois.Items.Any(Function(i) i.ProviderEntryId = "E-3"),
                "E-3 foi lido pela tentativa rejeitada e nao pode aparecer")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Não vistos viram SUSPEITOS</b> — o efeito que era emitido e nunca
    ''' executado.
    '''
    ''' O <c>SweepModel</c> emitia <c>MarcarNaoVistosComoSuspeitos</c> e ninguém
    ''' executava: o sink nem tinha a operação. Depois de uma geração com A e B
    ''' e outra só com A, o B continuava <c>presente</c> e o manifesto o
    ''' devolvia como se estivesse lá.
    '''
    ''' E suspeito é onde para. Geração que passa NÃO promove suspeita a
    ''' ausência — nem por contagem, nem por tempo. É a §23.
    ''' </summary>
    <TestMethod>
    Public Sub Item_que_sumiu_na_segunda_geracao_vira_SUSPEITO_e_nao_ausente()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim sink As New SqliteSweepSink(db, FolderKey, EnvKey)

            Dim r1 = New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1", "E-2"), sink, 2).
                     Executar(Universo(), 0, 1, Real(), CancellationToken.None)
            Assert.IsTrue(r1.Publicou, r1.Motivo)

            ' Segunda geracao: E-2 sumiu.
            Dim r2 = New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1"), sink, 2).
                     Executar(Universo(), 0, 2, Real(), CancellationToken.None)
            Assert.IsTrue(r2.Publicou, r2.Motivo)

            Dim m = New ManifestReader(db).Ler(FolderKey)
            Assert.AreEqual(2, m.Items.Count, "o item que sumiu nao some do acervo")

            Dim e1 = m.Items.First(Function(i) i.ProviderEntryId = "E-1")
            Dim e2 = m.Items.First(Function(i) i.ProviderEntryId = "E-2")

            Assert.AreEqual(PresenceState.Presente, e1.Presence)
            Assert.AreEqual(PresenceState.Suspeito, e2.Presence,
                "nao visto na geracao publicada tem de virar SUSPEITO")
            Assert.AreNotEqual(PresenceState.AusenteDaPasta, e2.Presence,
                "e NUNCA ausente: cached nao conclui ausencia (§23)")
        End Using
    End Sub

    ''' <summary>
    ''' E suspeito volta a presente quando o item reaparece — visto é visto,
    ''' venha de onde vier.
    ''' </summary>
    <TestMethod>
    Public Sub Suspeito_volta_a_presente_quando_o_item_reaparece()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim sink As New SqliteSweepSink(db, FolderKey, EnvKey)

            Dim a1 = New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1", "E-2"), sink, 2).
                     Executar(Universo(), 0, 1, Real(), CancellationToken.None)
            Dim a2 = New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1"), sink, 2).
                     Executar(Universo(), 0, 2, Real(), CancellationToken.None)
            Assert.IsTrue(a1.Publicou AndAlso a2.Publicou)

            Dim suspeito = New ManifestReader(db).Ler(FolderKey).Items.
                First(Function(i) i.ProviderEntryId = "E-2")
            Assert.AreEqual(PresenceState.Suspeito, suspeito.Presence)

            ' O item volta.
            Dim a3 = New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1", "E-2"), sink, 2).
                     Executar(Universo(), 0, 3, Real(), CancellationToken.None)
            Assert.IsTrue(a3.Publicou, a3.Motivo)

            Dim voltou = New ManifestReader(db).Ler(FolderKey).Items.
                First(Function(i) i.ProviderEntryId = "E-2")
            Assert.AreEqual(PresenceState.Presente, voltou.Presence, "o item voltou")
        End Using
    End Sub
    ''' <summary>
    ''' Fila travada fica VISÍVEL.
    '''
    ''' Um consumidor que falha sempre trava o dreno na cabeça da fila, e as
    ''' gerações seguintes nunca chegam. <b>O bloqueio é deliberado</b> — marcar
    ''' como drenada uma geração que a UI não recebeu seria perder em silêncio,
    ''' e pular a cabeça entregaria as seguintes fora de ordem.
    '''
    ''' O que não pode é o bloqueio ser invisível. Depois de N falhas, quem
    ''' perguntar descobre onde travou e por quê.
    ''' </summary>
    <TestMethod>
    Public Sub Fila_travada_fica_VISIVEL()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(_db, CacheSchema.Intended(), falha)
            Dim sink As New SqliteSweepSink(db, FolderKey, EnvKey)
            Dim r = New SweepRunner(New FonteFalsaMutavel(Universo(), "E-1"), sink, 2).
                    Executar(Universo(), 0, 1, Real(), CancellationToken.None)
            Assert.IsTrue(r.Publicou, r.Motivo)

            Dim drain As New PublicationDrain(db)
            Dim ui As New ConsumidorFalso() With {.Lancar = True}

            Assert.IsNull(drain.TravadoEm(), "ainda nao tentou nenhuma vez")

            For i = 1 To 3
                Try
                    drain.Drenar(ui)
                Catch ex As InvalidOperationException
                End Try
            Next

            Dim travada = drain.TravadoEm()
            Assert.IsNotNull(travada, "depois de 3 falhas a fila tem de se declarar travada")
            StringAssert.Contains(drain.UltimoErro(travada.Value), "UI falhou",
                "o motivo tem de estar no banco, senao ninguem descobre por que parou")
            Assert.AreEqual(1, drain.Pendentes().Count, "e continua pendente, que e o certo")
        End Using
    End Sub

    ' ==================================================================

    Private Class ConsumidorFalso
        Implements IPublicationConsumer

        Friend ReadOnly Recebidas As New List(Of Long)()
        Friend Property Lancar As Boolean = False
        Friend Property LancarDepoisDeReceber As Boolean = False

        Public Sub Receber(geracao As Long) Implements IPublicationConsumer.Receber
            If Lancar Then Throw New InvalidOperationException("UI falhou")
            Recebidas.Add(geracao)
            If LancarDepoisDeReceber Then Throw New InvalidOperationException("UI morreu depois de receber")
        End Sub
    End Class

    Private Shared Sub Exec(db As CacheDatabase, sql As String)
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Function Valor(db As CacheDatabase, sql As String) As Object
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            Dim v = cmd.ExecuteScalar()
            Return If(v Is Nothing OrElse v Is DBNull.Value, Nothing, v)
        End Using
    End Function

    Private Shared Function Texto(db As CacheDatabase, sql As String) As String
        Dim v = Valor(db, sql)
        Return If(v Is Nothing, Nothing, Convert.ToString(v))
    End Function

    Private Shared Function Contar(db As CacheDatabase, tabela As String) As Integer
        Return Convert.ToInt32(Valor(db, $"SELECT COUNT(*) FROM {tabela}"))
    End Function

End Class

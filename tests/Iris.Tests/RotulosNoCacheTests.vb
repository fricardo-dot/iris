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
''' <b>ONDE OS RÓTULOS MORAM — Fase 5.</b>
'''
''' ------------------------------------------------------------------
''' <b>O RÓTULO É UMA OBSERVAÇÃO, E HERDA O QUE ISSO SIGNIFICA</b>
'''
''' Ele pende da encarnação e da geração. Geração nova invalida a anterior de
''' graça — e é essa herança que este arquivo prende: sem ela, um rótulo de uma
''' varredura de junho apareceria como atual em agosto, e ninguém teria como
''' saber a idade dele olhando a tela.
'''
''' ------------------------------------------------------------------
''' <b>O CONTROLE NEGATIVO</b>
'''
''' <see cref="Rotulo_de_geracao_ANTERIOR_nao_aparece_como_atual"/>. Sem ele, um
''' leitor que ignorasse a geração passaria em todos os outros testes daqui — e
''' seria exatamente o que faz a fila mostrar classificação velha com cara de
''' nova.
''' </summary>
' NAO PARALELIZAR: cada teste abre um SQLite proprio.
<TestClass>
<DoNotParallelize>
Public Class RotulosNoCacheTests

    Private Shared ReadOnly Quando As New DateTimeOffset(2026, 8, 31, 12, 0, 0, TimeSpan.Zero)

    <TestMethod>
    Public Sub O_rotulo_gravado_volta_na_leitura()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, "f-1", {"a", "b"}, rodada:=1)
                   Dim geracao = GeracaoPublicada(db, pasta)

                   Dim quantos = New RotulosNoCache(db).Gravar(
                       pasta, geracao, "ativacao-1", Quando,
                       New Dictionary(Of String, String) From {
                           {"f-1-a", "fyi"}, {"f-1-b", "precisa_de_mim"}},
                       New Dictionary(Of String, Double) From {{"f-1-a", 0.8}})

                   Assert.AreEqual(2, quantos)

                   Dim lidos = New RotulosNoCache(db).Publicados(pasta)
                   Assert.AreEqual("fyi", lidos("f-1-a"))
                   Assert.AreEqual("precisa_de_mim", lidos("f-1-b"))
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo.</b> O rótulo é da geração em que foi observado;
    ''' uma varredura nova republica a pasta, e a classificação antiga deixa de
    ''' ser a publicada — fica no banco, e não aparece.
    '''
    ''' Sem isso, a fila mostraria uma classificação de junho como se fosse de
    ''' hoje, e nada na tela diria a idade dela.
    ''' </summary>
    <TestMethod>
    Public Sub Rotulo_de_geracao_ANTERIOR_nao_aparece_como_atual()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, "f-1", {"a"}, rodada:=1)
                   Dim primeira = GeracaoPublicada(db, pasta)

                   Dim oCache As New RotulosNoCache(db)
                   oCache.Gravar(
                       pasta, primeira, "ativacao-1", Quando,
                       New Dictionary(Of String, String) From {{"f-1-a", "fyi"}}, Nothing)
                   Assert.AreEqual(1, New RotulosNoCache(db).Publicados(pasta).Count,
                       "o preparo do teste esta errado")

                   ' Varre de novo: a geracao publicada passa a ser outra.
                   Varrer(db, "f-1", {"a"}, rodada:=2, existente:=pasta)
                   Assert.AreNotEqual(primeira, GeracaoPublicada(db, pasta))

                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(pasta).Count,
                       "um rotulo de geracao vencida apareceu como atual")
               End Sub)
    End Sub

    ''' <summary>
    ''' Reclassificar a mesma geração é <b>correção</b>, e não um segundo fato:
    ''' substitui em vez de acrescentar. Sem isso a leitura teria de escolher
    ''' entre duas linhas, e a escolha não tem critério.
    ''' </summary>
    <TestMethod>
    Public Sub Reclassificar_a_MESMA_geracao_substitui()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, "f-1", {"a"}, rodada:=1)
                   Dim geracao = GeracaoPublicada(db, pasta)
                   Dim rotulos As New RotulosNoCache(db)

                   rotulos.Gravar(pasta, geracao, "ativacao-1", Quando,
                       New Dictionary(Of String, String) From {{"f-1-a", "fyi"}}, Nothing)
                   rotulos.Gravar(pasta, geracao, "ativacao-1", Quando,
                       New Dictionary(Of String, String) From {{"f-1-a", "promocao"}}, Nothing)

                   Dim lidos = rotulos.Publicados(pasta)
                   Assert.AreEqual(1, lidos.Count, "acrescentou em vez de substituir")
                   Assert.AreEqual("promocao", lidos("f-1-a"))
               End Sub)
    End Sub

    ''' <summary>
    ''' Item que não está na pasta é ignorado — ele saiu entre a classificação e
    ''' a gravação, e insistir criaria encarnação para uma mensagem que não está
    ''' mais lá.
    ''' </summary>
    <TestMethod>
    Public Sub Item_que_nao_esta_na_pasta_e_ignorado()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, "f-1", {"a"}, rodada:=1)
                   Dim geracao = GeracaoPublicada(db, pasta)

                   Dim quantos = New RotulosNoCache(db).Gravar(
                       pasta, geracao, "ativacao-1", Quando,
                       New Dictionary(Of String, String) From {
                           {"f-1-a", "fyi"}, {"f-1-SUMIU", "promocao"}}, Nothing)

                   Assert.AreEqual(1, quantos)
                   Assert.AreEqual(1, New RotulosNoCache(db).Publicados(pasta).Count)
               End Sub)
    End Sub

    ''' <summary>
    ''' Confiança que não veio vira zero — e zero é gravável, ao contrário de
    ''' nulo. "Não sei o quanto" e "tenho certeza" não podem ser o mesmo valor,
    ''' e o zero é o lado que não afirma.
    ''' </summary>
    <TestMethod>
    Public Sub Confianca_ausente_e_gravada_como_ZERO()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, "f-1", {"a"}, rodada:=1)

                   Dim oCache As New RotulosNoCache(db)
                   oCache.Gravar(
                       pasta, GeracaoPublicada(db, pasta), "ativacao-1", Quando,
                       New Dictionary(Of String, String) From {{"f-1-a", "fyi"}}, Nothing)

                   Using cmd = db.Connection.CreateCommand()
                       cmd.CommandText = "SELECT confidence FROM label_observation"
                       Assert.AreEqual(0.0, Convert.ToDouble(cmd.ExecuteScalar()), 0.0001)
                   End Using
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>A ativação fica junto.</b> Sem ela, um rótulo gravado sob uma
    ''' autorização vencida seria indistinguível de um recente, e o dono não
    ''' teria como saber sob que regra a classificação foi feita.
    ''' </summary>
    <TestMethod>
    Public Sub A_ATIVACAO_fica_gravada_com_o_rotulo()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, "f-1", {"a"}, rodada:=1)

                   Dim oCache As New RotulosNoCache(db)
                   oCache.Gravar(
                       pasta, GeracaoPublicada(db, pasta), "ativacao-de-agosto", Quando,
                       New Dictionary(Of String, String) From {{"f-1-a", "fyi"}}, Nothing)

                   Using cmd = db.Connection.CreateCommand()
                       cmd.CommandText = "SELECT activation_id FROM label_observation"
                       Assert.AreEqual("ativacao-de-agosto", CStr(cmd.ExecuteScalar()))
                   End Using
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O cache não guarda justificativa.</b> "Por que este rótulo" cita o
    ''' corpo, e o D1 diz que o cache guarda metadado. Este teste olha a forma da
    ''' tabela: nenhuma coluna de texto livre onde uma frase caiba.
    ''' </summary>
    <TestMethod>
    Public Sub A_tabela_dos_rotulos_NAO_tem_onde_guardar_corpo()
        Comigo(Sub(db)
                   Dim colunas As New List(Of String)()
                   Using cmd = db.Connection.CreateCommand()
                       cmd.CommandText = "PRAGMA table_info(label_observation)"
                       Using rd = cmd.ExecuteReader()
                           While rd.Read()
                               colunas.Add(rd.GetString(1))
                           End While
                       End Using
                   End Using

                   CollectionAssert.AreEquivalent(
                       {"label_key", "incarnation_key", "generation_key", "label",
                        "confidence", "activation_id", "observed_at"},
                       colunas.ToArray(),
                       "a tabela ganhou uma coluna: se ela guarda texto livre, o " &
                       "corpo da mensagem cabe nela")
               End Sub)
    End Sub

    ''' <summary>
    ''' Pasta sem geração publicada devolve vazio — e vazio aqui quer dizer "não
    ''' há rótulo publicado", que é diferente de "não há rótulo".
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_sem_geracao_publicada_devolve_vazio()
        Comigo(Sub(db)
                   Dim pasta = New ResolvedorDoAcervo(db).Pasta("store-1", "f-9", "Nunca varrida")
                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(pasta).Count)
               End Sub)
    End Sub

    ' ==================================================================
    ' O ANDAIME

    Private Shared Function Impressao() As EnvironmentFingerprint
        Return New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing)
    End Function

    ''' <summary>Varre e publica uma pasta, pela varredura de verdade.</summary>
    Private Shared Function Varrer(db As CacheDatabase, entryId As String,
                                   sufixos As String(), rodada As Integer,
                                   Optional existente As Long? = Nothing) As Long
        Dim resolvedor As New ResolvedorDoAcervo(db)
        Dim chave = If(existente.HasValue, existente.Value,
                       resolvedor.Pasta("store-1", entryId, "Pasta " & entryId))
        Dim amb = resolvedor.Ambiente(Impressao())

        Dim universo As New SweepUniverse("store-1", entryId, "f", Nothing, rodada, "amb-1")
        Dim fonte As New FonteDeLinhas(universo, sufixos.Select(
            Function(s) New SourceRow With {
                .Key = $"{entryId}-{s}",
                .Subject = "assunto " & s,
                .SenderName = "quem",
                .ReceivedAt = Quando.ToString("o"),
                .MessageClass = "IPM.Note"}))

        Dim sink As New SqliteSweepSink(db, chave, amb.Chave)
        Dim r = New SweepRunner(fonte, sink, 50).
                Executar(universo, 0, rodada, EnvironmentPolicy.Capacidades(Impressao()),
                         CancellationToken.None)
        Assert.IsTrue(r.Publicou, $"controle: a varredura tinha de publicar. motivo: {r.Motivo}")
        Return chave
    End Function

    Private Shared Function GeracaoPublicada(db As CacheDatabase, pasta As Long) As Long
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = "SELECT published_generation_key FROM folder WHERE folder_key = $f"
            cmd.Parameters.AddWithValue("$f", pasta)
            Return Convert.ToInt64(cmd.ExecuteScalar())
        End Using
    End Function

    Private Shared Sub Comigo(corpo As Action(Of CacheDatabase))
        Dim caminho = Path.Combine(Path.GetTempPath(),
                                   "iris-rotulos-" & Guid.NewGuid().ToString("N") & ".db")
        Try
            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(caminho, CacheSchema.Intended(), falha)
                Assert.IsNotNull(db, $"{falha}")
                corpo(db)
            End Using
        Finally
            SqliteConnection.ClearAllPools()
            For Each sufixo In {"", "-wal", "-shm"}
                If File.Exists(caminho & sufixo) Then File.Delete(caminho & sufixo)
            Next
        End Try
    End Sub

End Class

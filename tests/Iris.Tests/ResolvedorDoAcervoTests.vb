Imports System.IO
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Model
Imports Iris.Sync
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A ponte que faltava entre o Outlook e o cache.</b>
'''
''' ------------------------------------------------------------------
''' A Fase 2 entregou o <c>SweepRunner</c>, o <c>OutlookSweepSource</c> e o
''' <c>SqliteSweepSink</c> — e nada que os ligasse. O sink pede
''' <c>folderKey</c> e <c>environmentKey</c> como <c>Long</c>; o Outlook fala
''' em <c>(StoreId, EntryId)</c>, que são strings. Nenhum código de produção
''' fazia a travessia, e por isso <b>o aplicativo nunca varreu</b>: os testes
''' semeavam <c>1, 1, 1</c> na mão e seguiam.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class ResolvedorDoAcervoTests

    Private _pasta As String
    Private _caminho As String

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-resolv-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_pasta)
        _caminho = Path.Combine(_pasta, "cache.db")
    End Sub

    <TestCleanup>
    Public Sub Limpar()
        SqliteConnection.ClearAllPools()
        Try
            If Directory.Exists(_pasta) Then Directory.Delete(_pasta, True)
        Catch
        End Try
    End Sub

    Private Function Abrir() As CacheDatabase
        Dim falha As OpenFailure = Nothing
        Dim db = CacheDatabase.Open(_caminho, CacheSchema.Intended(), falha)
        Assert.IsNotNull(db, $"{falha}")
        Return db
    End Function

    Private Shared Function Impressao() As EnvironmentFingerprint
        Return New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing)
    End Function

    ' ==================================================================
    ' A pasta

    <TestMethod>
    Public Sub A_mesma_pasta_devolve_a_MESMA_chave()
        Using db = Abrir()
            Dim r As New ResolvedorDoAcervo(db)

            Dim a = r.Pasta("store-1", "entry-1", "Caixa de Entrada")
            Dim b = r.Pasta("store-1", "entry-1", "Caixa de Entrada")

            Assert.AreEqual(a, b)
            Assert.AreEqual(1, Contar(db, "store"))
            Assert.AreEqual(1, Contar(db, "folder"))
        End Using
    End Sub

    ''' <summary>
    ''' <b>Pastas diferentes do mesmo store compartilham o store.</b>
    '''
    ''' O controle negativo do teste acima: sem ele, um resolvedor que
    ''' devolvesse sempre a mesma chave passaria.
    ''' </summary>
    <TestMethod>
    Public Sub Pastas_DIFERENTES_tem_chaves_diferentes()
        Using db = Abrir()
            Dim r As New ResolvedorDoAcervo(db)

            Dim a = r.Pasta("store-1", "entry-1", "Caixa de Entrada")
            Dim b = r.Pasta("store-1", "entry-2", "Iris-Teste")

            Assert.AreNotEqual(a, b)
            Assert.AreEqual(1, Contar(db, "store"), "e o store e o mesmo")
            Assert.AreEqual(2, Contar(db, "folder"))
        End Using
    End Sub

    ''' <summary>
    ''' <b>O mesmo EntryId em stores diferentes é outra pasta.</b>
    '''
    ''' O índice único é <c>(store_key, provider_entry_id)</c>, e não o EntryId
    ''' sozinho. Resolver só pelo EntryId juntaria duas caixas numa chave — e o
    ''' que o usuário veria é o acervo de uma conta debaixo do nome da outra.
    ''' </summary>
    <TestMethod>
    Public Sub O_mesmo_EntryId_em_OUTRO_store_e_outra_pasta()
        Using db = Abrir()
            Dim r As New ResolvedorDoAcervo(db)

            Dim a = r.Pasta("store-1", "entry-igual", "Caixa de Entrada")
            Dim b = r.Pasta("store-2", "entry-igual", "Caixa de Entrada")

            Assert.AreNotEqual(a, b)
            Assert.AreEqual(2, Contar(db, "store"))
        End Using
    End Sub

    ''' <summary>
    ''' <b>Reencontrar a pasta NÃO apaga o estado de sincronização.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <c>reconcile_epoch</c>, <c>published_generation_key</c> e
    ''' <c>stability</c> são o resultado da varredura. Um resolvedor que
    ''' fizesse "criar ou atualizar" ingênuo os reescreveria a cada clique na
    ''' pasta, e o trabalho da varredura anterior sumiria sem ninguém pedir —
    ''' com a faixa do acervo continuando a mostrar número, agora errado.
    ''' </summary>
    <TestMethod>
    Public Sub Reencontrar_a_pasta_NAO_apaga_a_sincronizacao()
        Using db = Abrir()
            Dim r As New ResolvedorDoAcervo(db)
            Dim k = r.Pasta("store-1", "entry-1", "Caixa de Entrada")

            Executar(db, $"UPDATE folder SET reconcile_epoch = 7, stability = 'instavel' " &
                         $"WHERE folder_key = {k}")

            Dim k2 = r.Pasta("store-1", "entry-1", "Caixa de Entrada renomeada")

            Assert.AreEqual(k, k2)
            Assert.AreEqual(7L, Numero(db, $"SELECT reconcile_epoch FROM folder WHERE folder_key = {k}"))
            Assert.AreEqual("instavel", Texto(db, $"SELECT stability FROM folder WHERE folder_key = {k}"))
            Assert.AreEqual("Caixa de Entrada renomeada",
                            Texto(db, $"SELECT name FROM folder WHERE folder_key = {k}"),
                            "o nome e a unica coisa que acompanha o Outlook")
        End Using
    End Sub

    <TestMethod>
    Public Sub Identificador_vazio_e_RECUSADO()
        Using db = Abrir()
            Dim r As New ResolvedorDoAcervo(db)
            Assert.ThrowsException(Of ArgumentException)(Sub() r.Pasta("", "e", "n"))
            Assert.ThrowsException(Of ArgumentException)(Sub() r.Pasta("s", "  ", "n"))
        End Using
    End Sub

    ' ==================================================================
    ' O ambiente

    ''' <summary>
    ''' <b>O ambiente nasce NÃO autorizado.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' É o teste que guarda a decisão inteira. O gate D2 diz que a allowlist
    ''' do ambiente é <b>dado</b>, e não constante no código, para que
    ''' "ambiente não medido" possa recusar operar. Se o resolvedor gravasse
    ''' <c>allowed = 1</c> para o que ele mesmo detectou, o Iris estaria
    ''' medindo e aprovando a própria medição, e o D2 viraria decoração.
    ''' </summary>
    <TestMethod>
    Public Sub O_ambiente_nasce_NAO_autorizado()
        Using db = Abrir()
            Dim p = New ResolvedorDoAcervo(db).Ambiente(Impressao())

            Assert.IsTrue(p.Novo)
            Assert.IsFalse(p.Permitido,
                "o programa nao pode autorizar o ambiente que ele mesmo mediu")
            Assert.AreEqual(0L, Numero(db, "SELECT allowed FROM environment_profile"))
        End Using
    End Sub

    ''' <summary>
    ''' <b>E a autorização DURA.</b>
    '''
    ''' O outro lado da mesma moeda: se reencontrar o perfil reescrevesse
    ''' <c>allowed = 0</c>, a cerimônia valeria até o próximo start do
    ''' programa — e o usuário autorizaria o mesmo ambiente todo dia sem
    ''' entender por quê.
    ''' </summary>
    <TestMethod>
    Public Sub Reencontrar_o_ambiente_NAO_rebaixa_a_autorizacao()
        Using db = Abrir()
            Dim r As New ResolvedorDoAcervo(db)
            Dim antes = r.Ambiente(Impressao())

            ' A cerimonia autoriza.
            Executar(db, $"UPDATE environment_profile SET allowed = 1 " &
                         $"WHERE environment_key = {antes.Chave}")

            Dim depois = r.Ambiente(Impressao())

            Assert.AreEqual(antes.Chave, depois.Chave)
            Assert.IsFalse(depois.Novo)
            Assert.IsTrue(depois.Permitido, "a autorizacao foi rebaixada na releitura")
            Assert.AreEqual(1, Contar(db, "environment_profile"))
        End Using
    End Sub

    ''' <summary>
    ''' <b>Ambiente diferente é OUTRO perfil, e ele nasce não autorizado.</b>
    '''
    ''' Autorizar Exchange em cache não autoriza um PST. A impressão digital é
    ''' a identidade, e a §18.4 já tinha decidido que "Exchange em cache" não é
    ''' um ambiente e sim uma família.
    ''' </summary>
    <TestMethod>
    Public Sub Autorizar_um_ambiente_nao_autoriza_OUTRO()
        Using db = Abrir()
            Dim r As New ResolvedorDoAcervo(db)
            Dim cached = r.Ambiente(Impressao())
            Executar(db, $"UPDATE environment_profile SET allowed = 1 " &
                         $"WHERE environment_key = {cached.Chave}")

            Dim outro = r.Ambiente(
                New EnvironmentFingerprint(ProviderKind.PstLocal, False, Nothing))

            Assert.AreNotEqual(cached.Chave, outro.Chave)
            Assert.IsFalse(outro.Permitido)
        End Using
    End Sub

    ' ==================================================================
    ' A medição

    ''' <summary>
    ''' <b>O que o Outlook reporta vira impressão digital.</b>
    '''
    ''' Sem COM: o broker já devolve <c>IsCachedExchange</c> e
    ''' <c>ExchangeStoreType</c> no <see cref="StoreInfo"/>.
    ''' </summary>
    <TestMethod>
    Public Sub O_store_do_Outlook_vira_impressao_digital()
        Dim cached = AmbienteMedido.De(New StoreInfo() With {
            .ExchangeStoreType = "3", .IsCachedExchange = True})
        Assert.AreEqual(ProviderKind.ExchangeCached, cached.Provider)
        Assert.IsTrue(cached.CachedMode)

        Dim online = AmbienteMedido.De(New StoreInfo() With {
            .ExchangeStoreType = "3", .IsCachedExchange = False})
        Assert.AreEqual(ProviderKind.ExchangeOnline, online.Provider)

        Dim pst = AmbienteMedido.De(New StoreInfo() With {.ExchangeStoreType = "4"})
        Assert.AreEqual(ProviderKind.PstLocal, pst.Provider)
    End Sub

    ''' <summary>
    ''' <b>O que não se reconhece vira Desconhecido — e desconhecido recusa.</b>
    '''
    ''' O controle negativo da medição. Um <c>Case Else</c> que chutasse
    ''' <c>ExchangeCached</c> faria o Iris tratar um provedor nunca medido como
    ''' um que ele mediu, que é exatamente o que a §19.3 proíbe.
    ''' </summary>
    <TestMethod>
    Public Sub Store_que_nao_se_reconhece_vira_DESCONHECIDO()
        For Each esquisito In {"", "9", "banana", Nothing}
            Assert.AreEqual(ProviderKind.Desconhecido,
                AmbienteMedido.De(New StoreInfo() With {.ExchangeStoreType = esquisito}).Provider,
                $"'{esquisito}'")
        Next
        Assert.AreEqual(ProviderKind.Desconhecido, AmbienteMedido.De(Nothing).Provider)
    End Sub

    ''' <summary>
    ''' <b>A janela de sincronização sai NULA, e sai declaradamente.</b>
    '''
    ''' A §22.3 mediu que <c>Store</c> não a expõe; a §22.4 mediu que o
    ''' registro do perfil também não — o usuário moveu o cursor três vezes e
    ''' 294 valores não mudaram. Inventar um token aqui faria o ambiente
    ''' parecer completo, e <c>ExigeReconciliacao</c> nunca disparar sem que
    ''' ninguém soubesse por quê.
    ''' </summary>
    <TestMethod>
    Public Sub A_janela_de_sincronizacao_sai_NULA()
        Dim f = AmbienteMedido.De(New StoreInfo() With {
            .ExchangeStoreType = "3", .IsCachedExchange = True})

        Assert.IsNull(f.WindowToken)
        StringAssert.Contains(f.Value(), "janela-nao-lida")
    End Sub

    ' ==================================================================

    Private Shared Function Contar(db As CacheDatabase, tabela As String) As Integer
        Return CInt(Numero(db, $"SELECT COUNT(*) FROM {tabela}"))
    End Function

    Private Shared Function Numero(db As CacheDatabase, sql As String) As Long
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            Return Convert.ToInt64(cmd.ExecuteScalar())
        End Using
    End Function

    Private Shared Function Texto(db As CacheDatabase, sql As String) As String
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            Dim v = cmd.ExecuteScalar()
            Return If(v Is Nothing OrElse v Is DBNull.Value, Nothing, Convert.ToString(v))
        End Using
    End Function

    Private Shared Sub Executar(db As CacheDatabase, sql As String)
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
        End Using
    End Sub

End Class

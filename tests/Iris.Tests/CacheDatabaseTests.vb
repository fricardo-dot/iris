Imports System.IO
Imports System.Linq
Imports Iris.Cache
Imports Iris.Core
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' O banco de verdade, contra o modelo.
'''
''' O <see cref="SchemaGate"/> prova que o schema que eu ESCREVI respeita os
''' invariantes. Ele não prova que o arquivo no disco é aquele schema — um
''' banco de versão antiga, ou editado à mão, passa no gate e não
''' corresponde a nada. A divergência só apareceria muito depois, num INSERT
''' que falha ou, pior, num SELECT que devolve menos do que deveria.
'''
''' É o mesmo padrão que a Q1 cobrou com a coluna <c>Permission</c> e a
''' §16.5 com o <c>Restrict</c>: <b>"não lançou" não é "funciona"</b>.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class CacheDatabaseTests

    Private _pasta As String

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-cache-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_pasta)
    End Sub

    <TestCleanup>
    Public Sub Limpar()
        SqliteConnection.ClearAllPools()
        Try
            If Directory.Exists(_pasta) Then Directory.Delete(_pasta, True)
        Catch
        End Try
    End Sub

    Private Function Caminho(Optional nome As String = "cache.db") As String
        Return Path.Combine(_pasta, nome)
    End Function

    ' ==================================================================

    <TestMethod>
    Public Sub Cria_e_reabre_sem_divergencia()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"falhou ao criar: {falha}")
            Assert.IsNull(falha)
        End Using

        ' Reabrir compara o arquivo REAL com o modelo. Se o DDL gerado nao
        ' corresponder ao modelo, e aqui que aparece.
        Using db2 = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db2, $"falhou ao reabrir: {falha}")
        End Using
    End Sub

    ''' <summary>
    ''' FK no SQLite vem DESLIGADA. Um schema com FK declarada e desligada dá
    ''' a impressão de integridade sem tê-la.
    ''' </summary>
    <TestMethod>
    Public Sub Foreign_keys_ficam_ligadas()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText = "PRAGMA foreign_keys"
                Assert.AreEqual(1, Convert.ToInt32(cmd.ExecuteScalar()))
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' E elas funcionam de verdade — declarar não é aplicar.
    ''' </summary>
    <TestMethod>
    Public Sub FK_realmente_recusa_referencia_orfa()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Using cmd = db.Connection.CreateCommand()
                ' folder aponta para um store que nao existe.
                cmd.CommandText =
                    "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
                    "reconcile_epoch, stability) VALUES (1, 999, 'X', 0, 'estavel')"
                Assert.ThrowsException(Of SqliteException)(Sub() cmd.ExecuteNonQuery())
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Toda FK aponta para uma coluna QUE EXISTE.
    '''
    ''' Este teste não consulta o modelo — só o banco — e é por isso que ele
    ''' vale. O SQLite aceita <c>REFERENCES pai (coluna_que_nao_existe)</c> no
    ''' CREATE TABLE sem reclamar; a falha só aparece no primeiro INSERT, como
    ''' "foreign key mismatch", longe de onde o erro foi escrito.
    '''
    ''' Foi exatamente o que aconteceu: o gerador supunha que a chave primária
    ''' de <c>x</c> se chama <c>x_key</c>, e <c>environment_profile</c> tem
    ''' <c>environment_key</c>. Os outros dez testes desta classe passaram por
    ''' cima do schema quebrado — quem acusou foi o primeiro INSERT real, no
    ''' teste de crash.
    ''' </summary>
    <TestMethod>
    Public Sub Toda_FK_aponta_para_coluna_que_existe()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Dim c = db.Connection

            Dim tabelas As New List(Of String)()
            Using cmd = c.CreateCommand()
                cmd.CommandText = "SELECT name FROM sqlite_schema WHERE type='table' " &
                                  "AND name NOT LIKE 'sqlite_%'"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        tabelas.Add(rd.GetString(0))
                    End While
                End Using
            End Using
            Assert.IsTrue(tabelas.Count > 10, "esperava o schema inteiro")

            Dim problemas As New List(Of String)()
            Dim verificadas = 0
            For Each t In tabelas
                For Each fk In ForeignKeys(c, t)
                    verificadas += 1
                    Dim colunas = ColunasDe(c, fk.Pai)
                    If colunas.Count = 0 Then
                        problemas.Add($"{t}.{fk.De} -> tabela {fk.Pai} nao existe")
                    ElseIf fk.Para IsNot Nothing AndAlso
                           Not colunas.Contains(fk.Para, StringComparer.OrdinalIgnoreCase) Then
                        problemas.Add($"{t}.{fk.De} -> {fk.Pai}.{fk.Para}: coluna alvo nao existe")
                    End If
                Next
            Next

            Assert.IsTrue(verificadas > 15, $"so {verificadas} FKs verificadas — teste vazio nao prova nada")
            Assert.AreEqual(0, problemas.Count, String.Join(" | ", problemas))
        End Using
    End Sub

    Private Shared Function ForeignKeys(c As SqliteConnection, tabela As String) _
            As List(Of (De As String, Pai As String, Para As String))
        Dim r As New List(Of (De As String, Pai As String, Para As String))()
        Using cmd = c.CreateCommand()
            cmd.CommandText = $"PRAGMA foreign_key_list({tabela})"
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    r.Add((rd.GetString(3), rd.GetString(2),
                           If(rd.IsDBNull(4), Nothing, rd.GetString(4))))
                End While
            End Using
        End Using
        Return r
    End Function

    Private Shared Function ColunasDe(c As SqliteConnection, tabela As String) As List(Of String)
        Dim r As New List(Of String)()
        Using cmd = c.CreateCommand()
            cmd.CommandText = $"PRAGMA table_info({tabela})"
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    r.Add(rd.GetString(1))
                End While
            End Using
        End Using
        Return r
    End Function

    ''' <summary>
    ''' Os CHECK do modelo viram CHECK no banco. Sem isto, 'presence' aceita
    ''' qualquer string e o enum vira sugestão.
    ''' </summary>
    <TestMethod>
    Public Sub CHECK_recusa_estado_de_presenca_invalido()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Executar(db, "INSERT INTO store (store_key, provider_store_id) VALUES (1, 'S')")
            Executar(db, "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
                         "reconcile_epoch, stability) VALUES (1, 1, 'F', 0, 'estavel')")
            Executar(db, "INSERT INTO item (item_key, created_at) VALUES (1, '2026-01-01')")

            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText =
                    "INSERT INTO association (association_key, item_key, folder_key, " &
                    "presence, observability, version) VALUES (1, 1, 1, 'inventado', 'observavel', 0)"
                Assert.ThrowsException(Of SqliteException)(Sub() cmd.ExecuteNonQuery())
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Idempotência: a mesma pasta duas vezes não entra. Sem isto, uma
    ''' retomada ou uma segunda importação duplicam, e o S6 passa a rejeitar
    ''' TODA varredura por chave repetida — sintoma longe da causa.
    ''' </summary>
    <TestMethod>
    Public Sub Unico_impede_pasta_duplicada()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Executar(db, "INSERT INTO store (store_key, provider_store_id) VALUES (1, 'S')")
            Executar(db, "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
                         "reconcile_epoch, stability) VALUES (1, 1, 'F', 0, 'estavel')")
            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText = "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
                                  "reconcile_epoch, stability) VALUES (2, 1, 'F', 0, 'estavel')"
                Assert.ThrowsException(Of SqliteException)(Sub() cmd.ExecuteNonQuery())
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' S5 na prática: apagar um item não pode levar o trabalho do usuário.
    ''' O modelo diz RESTRICT, e o banco tem de recusar.
    ''' </summary>
    <TestMethod>
    Public Sub Apagar_item_com_estado_do_usuario_e_RECUSADO()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Executar(db, "INSERT INTO item (item_key, created_at) VALUES (1, '2026-01-01')")
            Executar(db, "INSERT INTO user_state (user_state_key, item_key, triaged) VALUES (1, 1, 1)")

            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText = "DELETE FROM item WHERE item_key = 1"
                Assert.ThrowsException(Of SqliteException)(Sub() cmd.ExecuteNonQuery(),
                    "apagar item com estado do usuario tem de ser RECUSADO, nao cascatear")
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Versão incompatível FALHA FECHADO. Migrar sem saber de onde para onde
    ''' é pior que recusar: o dado já está no disco, e um palpite errado o
    ''' corrompe em silêncio.
    ''' </summary>
    <TestMethod>
    Public Sub Versao_incompativel_falha_fechado()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Executar(db, "PRAGMA user_version = 999")
        End Using
        SqliteConnection.ClearAllPools()

        Dim db2 = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
        Assert.IsNull(db2, "deveria recusar abrir")
        Assert.IsNotNull(falha)
        Assert.AreEqual("versao", falha.Reason)
    End Sub

    ''' <summary>
    ''' O controle negativo da INTROSPECÇÃO, e é o que dá sentido a ela.
    '''
    ''' Um banco adulterado — coluna a mais — tem de ser recusado. Sem este
    ''' teste, o introspector poderia sempre devolver "sem divergência" e
    ''' ninguém notaria.
    ''' </summary>
    <TestMethod>
    Public Sub Banco_adulterado_e_recusado()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Executar(db, "ALTER TABLE item ADD COLUMN intruso TEXT")
        End Using
        SqliteConnection.ClearAllPools()

        Dim db2 = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
        Assert.IsNull(db2, "coluna extra deveria ser recusada")
        Assert.AreEqual("divergencia", falha.Reason)
        StringAssert.Contains(falha.Detail, "intruso")
    End Sub

    ''' <summary>
    ''' E o outro lado: tabela FALTANDO também é divergência.
    ''' </summary>
    <TestMethod>
    Public Sub Tabela_faltando_e_recusada()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Executar(db, "PRAGMA foreign_keys = OFF")
            Executar(db, "DROP TABLE publication_log")
        End Using
        SqliteConnection.ClearAllPools()

        Dim db2 = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
        Assert.IsNull(db2)
        StringAssert.Contains(falha.Detail, "publication_log")
    End Sub

    ''' <summary>
    ''' Modelo que viola invariante nem cria banco. O gate roda ANTES.
    ''' </summary>
    <TestMethod>
    Public Sub Modelo_que_viola_invariante_nao_cria_banco()
        Dim s = CacheSchema.Intended()
        Dim semLog = New CacheSchema(s.Tables.Where(Function(t) t.Name <> "publication_log"))

        Dim falha As OpenFailure = Nothing
        Dim db = CacheDatabase.Open(Caminho("nunca.db"), semLog, falha)
        Assert.IsNull(db)
        Assert.AreEqual("gate", falha.Reason)
        Assert.IsFalse(File.Exists(Caminho("nunca.db")) AndAlso New FileInfo(Caminho("nunca.db")).Length > 0,
            "nao deveria ter criado arquivo com conteudo")
    End Sub

    ''' <summary>
    ''' WAL: crash no meio de uma transação não corrompe o banco.
    ''' </summary>
    <TestMethod>
    Public Sub Journal_e_WAL()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText = "PRAGMA journal_mode"
                Assert.AreEqual("wal", Convert.ToString(cmd.ExecuteScalar()).ToLowerInvariant())
            End Using
        End Using
    End Sub

    Private Shared Sub Executar(db As CacheDatabase, sql As String)
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' ==================================================================
    ' A MIGRACAO
    '
    ' Enquanto o arquivo guardava so metadado do Outlook, apagar nao custava
    ' nada. Depois que o DIARIO DO EGRESS passou a morar dentro dele, apagar
    ' virou destruir o registro do que saiu desta maquina -- que nao se
    ' reconstroi de lugar nenhum.

    ''' <summary>
    ''' <b>Um banco na versão 1 sobe para a 2 sem perder linha.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' O teste finge a versão 1 do jeito mais próximo do real que dá: cria o
    ''' banco de hoje, <b>derruba</b> a coluna nova e volta o
    ''' <c>user_version</c>. Depois grava uma linha no diário e reabre.
    '''
    ''' O que ele cobra é a linha ainda estar lá. Antes disto, a única saída
    ''' para quem tinha um banco antigo era apagar o arquivo — e junto com ele o
    ''' registro do que já tinha saído da máquina.
    ''' </summary>
    <TestMethod>
    Public Sub Banco_da_versao_1_MIGRA_sem_perder_o_diario()
        Dim falha As OpenFailure = Nothing

        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            ' Volta ao formato da versao 1: DROP COLUMN e o user_version.
            Executar(db, "ALTER TABLE disclosure_log DROP COLUMN http_status")
            Executar(db, "INSERT INTO disclosure_log (request_id, seq, capability_id, " &
                         " stage, activation_id, activation_version, operation, provider, " &
                         " endpoint, model, payload_hash, payload_bytes, message_count, " &
                         " intended_at, note, gate_reason) VALUES " &
                         "('req-1', 1, 'cap-1', 'Ambigua', 'ativacao-1', 1, 'Resumir', " &
                         " 'openrouter', 'https://x.invalido/v1', 'modelo-1', 'hash-1', " &
                         " 3856, 1, '2026-08-26T00:00:00.0000000+00:00', " &
                         " 'ProvedorRecusou', 'NaoDecidido')")
            Executar(db, "PRAGMA user_version = 1")
        End Using
        SqliteConnection.ClearAllPools()

        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"a migracao devia ter deixado abrir: {falha}")

            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText = "SELECT payload_bytes, note, http_status " &
                                  "FROM disclosure_log WHERE request_id = 'req-1'"
                Using r = cmd.ExecuteReader()
                    Assert.IsTrue(r.Read(), "A LINHA DO DIARIO SUMIU NA MIGRACAO")
                    Assert.AreEqual(3856, r.GetInt32(0))
                    Assert.AreEqual("ProvedorRecusou", r.GetString(1))
                    Assert.IsTrue(r.IsDBNull(2), "linha velha nao tem codigo, e isso e certo")
                End Using
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' <b>E o banco migrado aceita o código, com a mesma guarda do novo.</b>
    '''
    ''' Sem isto, uma migração que criasse a coluna sem o <c>CHECK</c> passaria:
    ''' o <see cref="SchemaIntrospector"/> compara nome, tipo, nulidade e chave
    ''' — <b>não</b> compara restrição. O banco migrado ficaria com uma guarda a
    ''' menos que o criado do zero, e ninguém saberia.
    ''' </summary>
    <TestMethod>
    Public Sub O_banco_MIGRADO_fica_com_a_mesma_guarda_do_novo()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Executar(db, "ALTER TABLE disclosure_log DROP COLUMN http_status")
            Executar(db, "PRAGMA user_version = 1")
        End Using
        SqliteConnection.ClearAllPools()

        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Executar(db, "INSERT INTO disclosure_log (request_id, seq, capability_id, " &
                         " stage, activation_id, activation_version, operation, provider, " &
                         " endpoint, model, payload_bytes, message_count, intended_at, " &
                         " note, gate_reason, http_status) VALUES " &
                         "('req-2', 2, 'cap-2', 'Ambigua', 'a', 1, 'Resumir', 'p', " &
                         " 'https://x.invalido/v1', 'm', 1, 1, " &
                         " '2026-08-26T00:00:00.0000000+00:00', 'ProvedorRecusou', " &
                         " 'NaoDecidido', 404)")

            Assert.ThrowsException(Of SqliteException)(
                Sub()
                    Executar(db, "INSERT INTO disclosure_log (request_id, seq, capability_id, " &
                                 " stage, activation_id, activation_version, operation, provider, " &
                                 " endpoint, model, payload_bytes, message_count, intended_at, " &
                                 " note, gate_reason, http_status) VALUES " &
                                 "('req-3', 3, 'cap-3', 'Ambigua', 'a', 1, 'Resumir', 'p', " &
                                 " 'https://x.invalido/v1', 'm', 1, 1, " &
                                 " '2026-08-26T00:00:00.0000000+00:00', 'ProvedorRecusou', " &
                                 " 'NaoDecidido', 99999)")
                End Sub, "o CHECK tinha de vir junto com a coluna migrada")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Versão sem passo conhecido continua falhando fechado.</b>
    '''
    ''' É o controle negativo da migração, e o que impede a tabela de passos de
    ''' virar "migre o que aparecer". Migrar sem saber de onde para onde
    ''' continua sendo pior que recusar; o que mudou foi só que os caminhos
    ''' <b>listados</b> deixaram de ser recusados.
    ''' </summary>
    <TestMethod>
    Public Sub Versao_desconhecida_continua_RECUSANDO()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Executar(db, "PRAGMA user_version = 97")
        End Using
        SqliteConnection.ClearAllPools()

        Dim db2 = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
        Assert.IsNull(db2, "97 nao tem passo conhecido; abrir seria adivinhar")
        Assert.IsNotNull(falha)
        StringAssert.Contains(falha.ToString(), "97")
    End Sub

    ''' <summary>
    ''' <b>E banco mais novo que o programa também recusa.</b>
    '''
    ''' Para trás não há caminho: um arquivo na versão 3 foi escrito por uma
    ''' versão que sabia coisas que esta não sabe. Sem esta guarda, o laço da
    ''' migração simplesmente não daria nenhuma volta e o arquivo seguiria para
    ''' a introspecção como se estivesse em ordem.
    ''' </summary>
    <TestMethod>
    Public Sub Banco_MAIS_NOVO_que_o_programa_recusa()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Executar(db, $"PRAGMA user_version = {SqliteDdl.SchemaVersion + 1}")
        End Using
        SqliteConnection.ClearAllPools()

        Assert.IsNull(CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha))
        Assert.IsNotNull(falha)
    End Sub

End Class

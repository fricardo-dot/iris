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

    ''' <summary>
    ''' <b>Devolve um banco de hoje ao formato da versao 1.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' Fingir versao antiga por SUBTRACAO tem um custo: cada migracao nova
    ''' precisa ser desfeita aqui tambem. Quando a 2 -> 3 entrou, os tres
    ''' testes de migracao quebraram com "duplicate column name: discarded" --
    ''' eles criavam um banco v3, tiravam so a coluna da migracao 1 -> 2, e
    ''' mandavam a escada subir por cima do que ja existia.
    '''
    ''' Por isso mora num lugar so. A alternativa seria guardar um arquivo
    ''' .db de verdade na versao 1, que prova mais e envelhece pior.
    ''' </summary>
    Private Shared Sub FingirVersao1(db As CacheDatabase)
        Executar(db, "DROP TABLE label_observation")
        Executar(db, "ALTER TABLE metadata_observation DROP COLUMN conversation_id")
        Executar(db, "ALTER TABLE metadata_observation DROP COLUMN conversation_index")
        Executar(db, "ALTER TABLE metadata_observation DROP COLUMN sender_address")
        Executar(db, "ALTER TABLE scan_stage DROP COLUMN conversation_id")
        Executar(db, "ALTER TABLE scan_stage DROP COLUMN conversation_index")
        Executar(db, "ALTER TABLE scan_stage DROP COLUMN sender_address")
        Executar(db, "ALTER TABLE disclosure_log DROP COLUMN http_status")
        Executar(db, "ALTER TABLE generation DROP COLUMN discarded")
        Executar(db, "PRAGMA user_version = 1")
    End Sub

    ''' <summary>
    ''' Devolve um banco de hoje ao formato da <b>versão 2</b>: só a coluna da
    ''' migração 2→3 sai.
    '''
    ''' Existe para o degrau 2→3 ser exercitado <b>sozinho</b>. Subir sempre de
    ''' 1 até 3 prova a escada inteira e esconde qual degrau quebrou — e o
    ''' degrau novo é justamente o que ninguém rodou ainda.
    ''' </summary>
    Private Shared Sub FingirVersao2(db As CacheDatabase)
        Executar(db, "DROP TABLE label_observation")
        Executar(db, "ALTER TABLE metadata_observation DROP COLUMN conversation_id")
        Executar(db, "ALTER TABLE metadata_observation DROP COLUMN conversation_index")
        Executar(db, "ALTER TABLE metadata_observation DROP COLUMN sender_address")
        Executar(db, "ALTER TABLE scan_stage DROP COLUMN conversation_id")
        Executar(db, "ALTER TABLE scan_stage DROP COLUMN conversation_index")
        Executar(db, "ALTER TABLE scan_stage DROP COLUMN sender_address")
        Executar(db, "ALTER TABLE generation DROP COLUMN discarded")
        Executar(db, "PRAGMA user_version = 2")
    End Sub

    ''' <summary>
    ''' Devolve um banco de hoje ao formato da <b>versão 3</b>: só as três
    ''' colunas da migração 3→4 saem.
    '''
    ''' Mesmo motivo do 2→3: o degrau novo é o que ninguém rodou ainda, e
    ''' subir a escada inteira esconde qual degrau quebrou.
    ''' </summary>
    Private Shared Sub FingirVersao3(db As CacheDatabase)
        Executar(db, "DROP TABLE label_observation")
        Executar(db, "ALTER TABLE metadata_observation DROP COLUMN conversation_id")
        Executar(db, "ALTER TABLE metadata_observation DROP COLUMN conversation_index")
        Executar(db, "ALTER TABLE metadata_observation DROP COLUMN sender_address")
        Executar(db, "ALTER TABLE scan_stage DROP COLUMN conversation_id")
        Executar(db, "ALTER TABLE scan_stage DROP COLUMN conversation_index")
        Executar(db, "ALTER TABLE scan_stage DROP COLUMN sender_address")
        Executar(db, "PRAGMA user_version = 3")
    End Sub

    ''' <summary>
    ''' Devolve um banco de hoje ao formato da <b>versão 4</b>: a tabela dos
    ''' rótulos sai inteira.
    ''' </summary>
    Private Shared Sub FingirVersao4(db As CacheDatabase)
        Executar(db, "DROP TABLE label_observation")
        Executar(db, "PRAGMA user_version = 4")
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
            FingirVersao1(db)
            Executar(db, "INSERT INTO disclosure_log (request_id, seq, capability_id, " &
                         " stage, activation_id, activation_version, operation, provider, " &
                         " endpoint, model, payload_hash, payload_bytes, message_count, " &
                         " intended_at, note, gate_reason) VALUES " &
                         "('req-1', 1, 'cap-1', 'Ambigua', 'ativacao-1', 1, 'Resumir', " &
                         " 'openrouter', 'https://x.invalido/v1', 'modelo-1', 'hash-1', " &
                         " 3856, 1, '2026-08-26T00:00:00.0000000+00:00', " &
                         " 'ProvedorRecusou', 'NaoDecidido')")
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
            FingirVersao1(db)
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

    ''' <summary>
    ''' <b>Depois de uma instância migrar, a seguinte abre o mesmo arquivo.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTE TESTE COBRE, E O QUE ELE NÃO COBRE</b>
    '''
    ''' Cobre a metade visível do defeito: uma segunda abertura <b>não</b> pode
    ''' tentar acrescentar de novo uma coluna que já existe, nem recusar um
    ''' banco que a primeira deixou correto.
    '''
    ''' <b>Não</b> cobre a corrida de verdade — duas aberturas <i>ao mesmo
    ''' tempo</i>, com a versão lida antes da trava de escrita. Eu escrevi esse
    ''' teste, com <c>Barrier</c> e duas <c>Task</c>, e ele <b>passou com o
    ''' defeito presente</b>: a janela entre ler a versão e escrever é curta
    ''' demais, e as duas aberturas quase nunca se cruzam nela. Um teste que
    ''' passa com o defeito no lugar não é cobertura — é uma linha verde que
    ''' faz alguém acreditar que a corrida está fechada.
    '''
    ''' Então ele foi removido, e fica escrito aqui: o <c>BEGIN IMMEDIATE</c>
    ''' com releitura da versão <b>dentro</b> dele se apoia no travamento do
    ''' SQLite e na leitura do código, e não em teste.
    ''' </summary>
    <TestMethod>
    Public Sub Depois_de_migrar_a_abertura_seguinte_NAO_tenta_de_novo()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            FingirVersao1(db)
        End Using
        SqliteConnection.ClearAllPools()

        Using primeira = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(primeira, $"a primeira devia migrar: {falha}")
        End Using
        SqliteConnection.ClearAllPools()

        Using segunda = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(segunda,
                $"a segunda recusou um banco que a primeira deixou correto: {falha}")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Migração que não produz o schema esperado é DESFEITA.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' Este é o defeito que a revisão pegou: o passo era escolhido pelo número
    ''' declarado no <c>user_version</c>, e não pela forma real do arquivo. Um
    ''' banco marcado como 1 mas com <b>outra</b> coisa faltando levava o
    ''' <c>ALTER TABLE</c>, era promovido a 2, e só então era recusado pela
    ''' introspecção — e na abertura seguinte já não havia caminho nenhum,
    ''' porque a própria tentativa tinha apagado a versão de origem.
    '''
    ''' Agora a verificação mora <b>dentro</b> da transação. O que este teste
    ''' cobra não é a recusa: é o arquivo continuar <b>na versão 1</b> depois
    ''' dela, com o diário intacto — quer dizer, ainda recuperável.
    ''' </summary>
    <TestMethod>
    Public Sub Migracao_que_nao_bate_com_o_modelo_DESFAZ_tudo()
        Dim falha As OpenFailure = Nothing
        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            FingirVersao1(db)
            ' E ALGO MAIS quebrado, que a migracao 1 -> 2 nao conserta.
            Executar(db, "ALTER TABLE disclosure_log DROP COLUMN gate_reason")
        End Using
        SqliteConnection.ClearAllPools()

        Assert.IsNull(CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha))
        Assert.IsNotNull(falha)
        SqliteConnection.ClearAllPools()

        ' O ARQUIVO TEM DE ESTAR INTACTO.
        Using conn As New SqliteConnection($"Data Source={Caminho()}")
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "PRAGMA user_version"
                Assert.AreEqual(1, Convert.ToInt32(cmd.ExecuteScalar()),
                    "a tentativa promoveu a versao e nao voltou atras")
            End Using
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM pragma_table_info('disclosure_log') " &
                                  "WHERE name = 'http_status'"
                Assert.AreEqual(0, Convert.ToInt32(cmd.ExecuteScalar()),
                    "a coluna ficou de uma migracao que devia ter sido desfeita")
            End Using
        End Using
        SqliteConnection.ClearAllPools()
    End Sub

    ''' <summary>
    ''' <b>A tabela de migracoes e fechada de verdade, e nao so no tipo.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' Guardar um <c>Dictionary</c> numa variavel tipada
    ''' <c>IReadOnlyDictionary</c> nao fecha nada: um <c>DirectCast</c> de volta
    ''' devolve a colecao <b>viva</b>, e um <c>Clear()</c> ali desliga a
    ''' migracao inteira — quer dizer, transforma "seu banco sobe de versao" em
    ''' "seu banco e recusado", que e o caminho que termina em apagar o diario.
    '''
    ''' Este projeto ja levou exatamente este golpe uma vez, em
    ''' <c>ActivationRecord.Congelar</c>, que devolvia <c>ToList()</c> tipado
    ''' como <c>IReadOnlyList</c> — e um <c>TryCast</c> reabria a lista de
    ''' operacoes autorizadas.
    ''' </summary>
    <TestMethod>
    Public Sub A_tabela_de_migracoes_NAO_da_para_reabrir()
        Assert.IsNull(TryCast(SqliteDdl.Migracoes,
                              Dictionary(Of Integer, IReadOnlyList(Of String))),
                      "da para reabrir o dicionario e esvazia-lo")

        For Each par In SqliteDdl.Migracoes
            Assert.IsNull(TryCast(par.Value, String()),
                          $"da para reabrir a lista do passo {par.Key} e reescreve-la")
            Assert.IsNull(TryCast(par.Value, List(Of String)),
                          $"da para reabrir a lista do passo {par.Key}")
        Next
    End Sub

    ''' <summary>
    ''' <b>A migração 2→3, sozinha, e a linha antiga sobrevive com NULO.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' Os outros testes de migração sobem de 1 até 3 e provam a escada
    ''' inteira. Este isola o degrau novo — o único que nenhum banco de
    ''' verdade jamais subiu — e cobra o que a coluna significa numa geração
    ''' anterior a ela: <b>nulo, e não zero</b>.
    '''
    ''' Zero afirmaria que aquela varredura não descartou nada. Nulo diz que
    ''' ela não contou. A diferença é exatamente a que existe entre as
    ''' gerações 2 e 3 da Caixa de Entrada do usuário: mesmas 1123 linhas
    ''' lidas, e só a segunda sabe que 12 foram recusadas.
    ''' </summary>
    <TestMethod>
    Public Sub A_migracao_2_para_3_preserva_a_geracao_antiga()
        Dim falha As OpenFailure = Nothing

        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            Semear(db)
            Executar(db,
                "INSERT INTO generation (generation_key, folder_key, attempt_key, " &
                "  coverage_kind, universe_fingerprint, rows_read, count_before, " &
                "  count_after, discarded, distinct_keys, reconcile_epoch, published_at) " &
                "VALUES (7, 1, 1, 'completa', 'u', 1123, 1135, 1135, NULL, 1123, 0, " &
                "        '2026-08-27T00:00:00.0000000+00:00')")
            FingirVersao2(db)
        End Using
        SqliteConnection.ClearAllPools()

        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"a migracao 2 -> 3 devia ter deixado abrir: {falha}")

            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText = "SELECT rows_read, count_before, discarded " &
                                  "FROM generation WHERE generation_key = 7"
                Using r = cmd.ExecuteReader()
                    Assert.IsTrue(r.Read(), "A GERACAO SUMIU NA MIGRACAO")
                    Assert.AreEqual(1123, r.GetInt32(0))
                    Assert.AreEqual(1135, r.GetInt32(1))
                    Assert.IsTrue(r.IsDBNull(2),
                        "geracao anterior a coluna tem de ficar NULA -- zero " &
                        "afirmaria que ela nao descartou nada, e ela nao contou")
                End Using
            End Using
        End Using
    End Sub
    ''' <summary>
    ''' <b>O degrau 3 → 4, sozinho.</b>
    '''
    ''' <c>FingirVersao3</c> existia e <b>ninguém o chamava</b>: a migração
    ''' 3→4 só era percorrida como parte da escada que começa na versão 1, e
    ''' o comentário do fixture prometia uma cobertura que não existia.
    ''' Escada inteira prova que o conjunto funciona e esconde qual degrau
    ''' quebrou — e o degrau novo é justamente o que ninguém rodou ainda.
    '''
    ''' Achado por revisão externa em 31/08/2026.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E O QUE ELE NÃO PROVA</b>
    '''
    ''' O fixture finge a versão 3 por <b>subtração</b>: parte do banco de
    ''' hoje e tira seis colunas. O resultado é um híbrido — tem tudo o
    ''' mais da versão 4 — então isto prova que <i>o degrau executa e
    ''' recria as colunas</i>, e não que um banco realmente produzido pela
    ''' versão 3 migra.
    '''
    ''' A alternativa seria guardar um arquivo <c>.db</c> de verdade na
    ''' versão 3, que prova mais e envelhece pior — é a mesma escolha que
    ''' <see cref="FingirVersao1"/> já documenta. Para seis colunas
    ''' <c>TEXT</c> anuláveis e sem índice, a diferença não morde; se elas
    ''' ganharem restrição ou <i>default</i>, morde.
    ''' </summary>
    <TestMethod>
    Public Sub A_migracao_3_para_4_recria_as_colunas_da_conversa()
        Dim falha As OpenFailure = Nothing

        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            FingirVersao3(db)
        End Using
        SqliteConnection.ClearAllPools()

        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"a migracao 3 -> 4 devia ter deixado abrir: {falha}")

            For Each tabela In {"metadata_observation", "scan_stage"}
                For Each coluna In {"conversation_id", "conversation_index", "sender_address"}
                    Assert.IsTrue(TemColuna(db, tabela, coluna),
                        $"{tabela}.{coluna} nao chegou pela migracao 3 -> 4")
                Next
            Next
        End Using
    End Sub

    ''' <summary>A tabela tem esta coluna? Por <c>PRAGMA table_info</c>.</summary>
    Private Shared Function TemColuna(db As CacheDatabase, tabela As String,
                                      coluna As String) As Boolean
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = $"PRAGMA table_info({tabela})"
            Using leitor = cmd.ExecuteReader()
                While leitor.Read()
                    If String.Equals(leitor.GetString(1), coluna, StringComparison.Ordinal) Then
                        Return True
                    End If
                End While
            End Using
        End Using
        Return False
    End Function

    ''' <summary>Uma pasta e uma tentativa, para a geração do teste ter a que se prender.</summary>
    Private Shared Sub Semear(db As CacheDatabase)
        Executar(db, "INSERT INTO environment_profile (environment_key, fingerprint, " &
                     "  provider, cached_mode, policy_version, allowed) " &
                     "VALUES (1, 'f', 'ExchangeCached', 1, 1, 1)")
        Executar(db, "INSERT INTO store (store_key, provider_store_id) VALUES (1, 's')")
        Executar(db, "INSERT INTO folder (folder_key, store_key, provider_entry_id, " &
                     "  reconcile_epoch, stability) VALUES (1, 1, 'e', 0, 'estavel')")
        Executar(db, "INSERT INTO scan_attempt (attempt_key, folder_key, environment_key, " &
                     "  universe_fingerprint, algorithm_version, reconcile_epoch, " &
                     "  attempt_number, stage, rows_read, started_at) " &
                     "VALUES (1, 1, 1, 'u', 1, 0, 1, 'publicada', 1123, " &
                     "        '2026-08-27T00:00:00.0000000+00:00')")
    End Sub

    ''' <summary>
    ''' <b>O degrau 4 → 5, sozinho.</b>
    '''
    ''' Ele CRIA uma tabela, e não acrescenta coluna — é o primeiro assim, e
    ''' por isso vale exercitá-lo separado: os degraus anteriores não provam
    ''' nada sobre <c>CREATE TABLE</c> dentro de uma migração.
    '''
    ''' Como os outros deste arquivo, o fixture finge a versão por subtração:
    ''' prova que o degrau executa e recria a tabela, não que um banco
    ''' realmente produzido pela versão 4 migra.
    ''' </summary>
    <TestMethod>
    Public Sub A_migracao_4_para_5_recria_a_tabela_dos_rotulos()
        Dim falha As OpenFailure = Nothing

        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"{falha}")
            FingirVersao4(db)
        End Using
        SqliteConnection.ClearAllPools()

        Using db = CacheDatabase.Open(Caminho(), CacheSchema.Intended(), falha)
            Assert.IsNotNull(db, $"a migracao 4 -> 5 devia ter deixado abrir: {falha}")

            For Each coluna In {"incarnation_key", "generation_key", "label",
                                "confidence", "activation_id", "observed_at"}
                Assert.IsTrue(TemColuna(db, "label_observation", coluna),
                    $"label_observation.{coluna} nao chegou pela migracao 4 -> 5")
            Next
        End Using
    End Sub

End Class

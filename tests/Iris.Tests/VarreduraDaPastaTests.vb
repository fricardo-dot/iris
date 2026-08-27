Imports System.Collections.Generic
Imports System.IO
Imports System.Threading
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Integration.Outlook
Imports Iris.Model
Imports Iris.Sync
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O aplicativo finalmente varre — e recusa quando tem de recusar.</b>
'''
''' ------------------------------------------------------------------
''' As três peças da varredura existiam desde a Fase 2 e nada as ligava. Só
''' os testes montavam a cadeia, semeando as chaves na mão; em produção
''' ninguém chamava o <c>SweepRunner</c>. O cache só tinha o que uma
''' importação manual tivesse posto nele.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class VarreduraDaPastaTests

    Private _pasta As String
    Private _caminho As String

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-varre-" & Guid.NewGuid().ToString("N"))
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

    Private Shared ReadOnly Alvo As New FolderKey("entry-1", "store-1")

    Private Shared Function Store() As StoreInfo
        ' O NOME, e nao o numero: o broker guarda ToString() de um enum, com
        ' ligacao antecipada. E este e o valor REAL medido na caixa do usuario.
        ' Enquanto este fixture dizia "3", ele carregava o meu chute -- e "3" e
        ' olNotExchange, quer dizer, PST local.
        '
        ' (Comentario AQUI, e nao dentro do With: em VB a continuacao implicita
        ' de { } nao aceita linha so de comentario, e o erro sai na anterior.)
        Return New StoreInfo() With {
            .StoreId = "store-1",
            .DisplayName = "Caixa",
            .ExchangeStoreType = "olPrimaryExchangeMailbox",
            .IsCachedExchange = True}
    End Function

    ''' <summary>Um broker que devolve uma página e acaba.</summary>
    Private Shared Function Broker() As FakeBroker
        Dim b As New FakeBroker()
        b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
            New MessagePage() With {
                .Items = New List(Of MailSummary) From {
                    New MailSummary() With {
                        .Key = New ItemKey("E-1", "store-1"), .Subject = "um"},
                    New MailSummary() With {
                        .Key = New ItemKey("E-2", "store-1"), .Subject = "dois"}},
                .TotalAtStart = 2,
                .NextCursor = Nothing})
        Return b
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>Ambiente não autorizado NÃO varre — e é o gate D2 em ação.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' É o teste que sustenta a decisão inteira. Sem esta recusa, "a allowlist
    ''' do ambiente é dado, e não constante no código" seria uma frase sem
    ''' consequência nenhuma: o perfil seria gravado com <c>allowed = 0</c> e a
    ''' varredura rodaria do mesmo jeito.
    '''
    ''' E repare no que a recusa <b>deixa pronto</b>: o perfil fica gravado,
    ''' com chave, para a cerimônia ter o que autorizar. Recusar sem registrar
    ''' o que foi medido deixaria o usuário sem nada para decidir.
    ''' </summary>
    <TestMethod>
    Public Sub Ambiente_NAO_autorizado_nao_varre()
        Using db = Abrir()
            Dim v As New VarreduraDaPasta(Broker(), db)

            Dim r = v.Executar(Alvo, "Caixa de Entrada", Store(), CancellationToken.None)

            Assert.AreEqual(RecusaDeVarredura.AmbienteNaoAutorizado, r.Recusa)
            Assert.IsFalse(r.Ok)
            Assert.AreEqual(0, Contar(db, "scan_attempt"), "varreu sem autorizacao")
            Assert.AreEqual(0, Contar(db, "folder"),
                "criou a pasta antes de saber se podia varrer")

            Assert.AreEqual(1, Contar(db, "environment_profile"),
                "o perfil TEM de ficar gravado: e o que a cerimonia autoriza")
            Assert.IsTrue(r.ChaveDoAmbiente > 0, "e a cerimonia precisa da chave")
            StringAssert.Contains(r.Ambiente, "ExchangeCached")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Depois da cerimônia, varre.</b>
    '''
    ''' O controle negativo do teste acima: sem ele, uma varredura que
    ''' recusasse <b>sempre</b> passaria — e a recusa pareceria correta.
    ''' </summary>
    <TestMethod>
    Public Sub Depois_de_AUTORIZADO_varre()
        Using db = Abrir()
            Dim v As New VarreduraDaPasta(Broker(), db)

            ' A primeira tentativa mede e recusa; e ela que da o que autorizar.
            Dim chave = v.Executar(Alvo, "Caixa", Store(), CancellationToken.None).ChaveDoAmbiente
            Autorizar(db, chave)

            Dim r = v.Executar(Alvo, "Caixa de Entrada", Store(), CancellationToken.None)

            Assert.AreEqual(RecusaDeVarredura.Nenhuma, r.Recusa)
            Assert.IsTrue(r.Ok, "a varredura nao produziu resultado")
            Assert.IsTrue(r.Pasta > 0, "a pasta tinha de ter sido resolvida")
            Assert.AreEqual(1, Contar(db, "folder"))
            Assert.AreEqual(1, Contar(db, "scan_attempt"))
            ' PUBLICOU, e nao so "chegou a varrer". Sem esta linha o teste
            ' passava com o fixture base, que nao declarava TotalAtStart: sem
            ' total nao ha S6, nada publica, e "Ok" continuava verdadeiro.
            Assert.AreEqual(1, Contar(db, "generation"))
        End Using
    End Sub

    ''' <summary>
    ''' <b>Varrer duas vezes não abre a mesma tentativa de novo.</b>
    '''
    ''' O número da tentativa entra no registro, e repetir um já usado tornaria
    ''' duas varreduras diferentes indistinguíveis no histórico — que é
    ''' justamente onde alguém vai olhar para entender por que o acervo mudou.
    ''' </summary>
    <TestMethod>
    Public Sub Varrer_duas_vezes_gera_tentativas_DIFERENTES()
        Using db = Abrir()
            Dim v As New VarreduraDaPasta(Broker(), db)
            Autorizar(db, v.Executar(Alvo, "Caixa", Store(), CancellationToken.None).ChaveDoAmbiente)

            v.Executar(Alvo, "Caixa", Store(), CancellationToken.None)
            v.Executar(Alvo, "Caixa", Store(), CancellationToken.None)

            Assert.AreEqual(2, Contar(db, "scan_attempt"))
            Assert.AreEqual(2, CInt(Numero(db, "SELECT COUNT(DISTINCT attempt_number) FROM scan_attempt")),
                "duas varreduras com o mesmo numero de tentativa")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Store que o Iris não reconhece não é autorizável por acidente.</b>
    '''
    ''' O ambiente vira <c>Desconhecido</c>, que é outro fingerprint — então
    ''' autorizar o Exchange em cache <b>não</b> autoriza o que não se
    ''' reconheceu. É a §19.3: ambiente não medido recusa.
    ''' </summary>
    <TestMethod>
    Public Sub Store_DESCONHECIDO_nao_pega_carona_na_autorizacao()
        Using db = Abrir()
            Dim v As New VarreduraDaPasta(Broker(), db)
            Autorizar(db, v.Executar(Alvo, "Caixa", Store(), CancellationToken.None).ChaveDoAmbiente)

            Dim estranho As New StoreInfo() With {
                .StoreId = "store-1", .ExchangeStoreType = "banana"}

            Dim r = v.Executar(Alvo, "Caixa", estranho, CancellationToken.None)

            Assert.AreEqual(RecusaDeVarredura.AmbienteNaoAutorizado, r.Recusa)
            Assert.AreEqual(2, Contar(db, "environment_profile"),
                "o desconhecido tem de ser OUTRO perfil")
        End Using
    End Sub

    <TestMethod>
    Public Sub Sem_pasta_e_sem_store_recusa_antes_de_tudo()
        Using db = Abrir()
            Dim v As New VarreduraDaPasta(Broker(), db)

            Assert.AreEqual(RecusaDeVarredura.SemPasta,
                v.Executar(Nothing, "x", Store(), CancellationToken.None).Recusa)
            Assert.AreEqual(RecusaDeVarredura.SemPasta,
                v.Executar(New FolderKey("", "store-1"), "x", Store(), CancellationToken.None).Recusa)
            Assert.AreEqual(RecusaDeVarredura.StoreDesconhecido,
                v.Executar(Alvo, "x", Nothing, CancellationToken.None).Recusa)

            Assert.AreEqual(0, Contar(db, "environment_profile"),
                "recusou antes de medir, entao nao ha o que gravar")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Fonte que recusa NÃO publica, e a tentativa é descartada.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' Este teste começou com a premissa errada: eu esperava exceção subindo e
    ''' virando <c>RecusaDeVarredura.Falhou</c>. Não é o que acontece, e o que
    ''' acontece é melhor — o <c>SweepRunner</c> trata recusa da fonte como
    ''' <b>desfecho</b> da varredura, descarta a tentativa e devolve o motivo.
    ''' Exceção fica para defeito de contrato.
    '''
    ''' O que importa provar é o que <b>não</b> aconteceu: nada foi publicado,
    ''' e o acervo não ganhou número vindo de uma leitura que falhou.
    ''' </summary>
    <TestMethod>
    Public Sub Fonte_que_RECUSA_nao_publica()
        Using db = Abrir()
            Dim b = Broker()
            Dim v As New VarreduraDaPasta(b, db)
            Autorizar(db, v.Executar(Alvo, "Caixa", Store(), CancellationToken.None).ChaveDoAmbiente)

            ' Sem resposta canonica, o duble devolve "fora da alcada".
            b.RespostaDaPagina = Nothing

            Dim r = v.Executar(Alvo, "Caixa", Store(), CancellationToken.None)

            Assert.AreNotEqual(SweepConclusion.Publicada, r.Varredura.Conclusion,
                "publicou sobre uma leitura que a fonte recusou")
            Assert.AreEqual("descartada",
                Texto(db, "SELECT stage FROM scan_attempt ORDER BY attempt_key DESC LIMIT 1"))
            Assert.AreEqual(0, Contar(db, "incarnation"),
                "gravou item de uma pagina que nunca veio")
        End Using
    End Sub

    ''' <summary>Texto de uma coluna, para as asserções acima.</summary>
    Private Shared Function Texto(db As CacheDatabase, sql As String) As String
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            Dim v = cmd.ExecuteScalar()
            Return If(v Is Nothing OrElse v Is DBNull.Value, Nothing, Convert.ToString(v))
        End Using
    End Function

    ' ==================================================================

    Private Shared Sub Autorizar(db As CacheDatabase, chave As Long)
        Assert.IsTrue(chave > 0, "nao havia perfil para autorizar")
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = "UPDATE environment_profile SET allowed = 1 WHERE environment_key = $k"
            cmd.Parameters.AddWithValue("$k", chave)
            Assert.AreEqual(1, cmd.ExecuteNonQuery())
        End Using
    End Sub

    Private Shared Function Contar(db As CacheDatabase, tabela As String) As Integer
        Return CInt(Numero(db, $"SELECT COUNT(*) FROM {tabela}"))
    End Function

    Private Shared Function Numero(db As CacheDatabase, sql As String) As Long
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = sql
            Return Convert.ToInt64(cmd.ExecuteScalar())
        End Using
    End Function

    ' ==================================================================
    ' O DESCARTE DECLARADO
    '
    ' A varredura da Caixa de Entrada leu 1123 e a pasta declarava 1135. A
    ' diferenca -- 12 linhas vistas e recusadas por nao serem mensagem -- era
    ' CALCULADA, usada pela guarda S6 para decidir se publicava, e jogada fora
    ' no mesmo instante. Quem olhasse o acervo depois via dois numeros e nao
    ' tinha onde procurar o terceiro.

    ''' <summary>
    ''' <b>O descarte declarado sobrevive à varredura.</b>
    ''' </summary>
    <TestMethod>
    Public Sub O_descarte_declarado_e_GUARDADO()
        Using db = Abrir()
            Dim b = Broker()
            ' Tres na pasta, uma delas recusada pela fonte: o Iris guarda duas
            ' e a contagem continua tres.
            b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
                New MessagePage() With {
                    .Items = New List(Of MailSummary) From {
                        New MailSummary() With {.Key = New ItemKey("E-1", "store-1")},
                        New MailSummary() With {.Key = New ItemKey("E-2", "store-1")}},
                    .SkippedCount = 1,
                    .TotalAtStart = 3,
                    .NextCursor = Nothing})

            Dim v As New VarreduraDaPasta(b, db)
            Autorizar(db, v.Executar(Alvo, "Caixa", Store(), CancellationToken.None).ChaveDoAmbiente)
            Dim r = v.Executar(Alvo, "Caixa", Store(), CancellationToken.None)

            Assert.IsTrue(r.Ok, "a conta 2 + 1 = 3 tinha de fechar e publicar")
            Assert.AreEqual(1L, Numero(db, "SELECT discarded FROM generation ORDER BY generation_key DESC LIMIT 1"))
            Assert.AreEqual(2L, Numero(db, "SELECT rows_read FROM generation ORDER BY generation_key DESC LIMIT 1"))
            Assert.AreEqual(3L, Numero(db, "SELECT count_before FROM generation ORDER BY generation_key DESC LIMIT 1"))
        End Using
    End Sub

    ''' <summary>
    ''' <b>E o número aparece na ressalva, com as palavras certas.</b>
    '''
    ''' O projeto já tinha tomado esta decisão do outro lado: a lista de
    ''' mensagens mostra o <c>SkippedCount</c>, com o comentário dizendo que sem
    ''' ele <i>"28 de 30 viraria mistério"</i>. A varredura fazia o mesmo e não
    ''' contava.
    ''' </summary>
    <TestMethod>
    Public Sub O_descarte_APARECE_na_ressalva()
        Using db = Abrir()
            Dim b = Broker()
            b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
                New MessagePage() With {
                    .Items = New List(Of MailSummary) From {
                        New MailSummary() With {.Key = New ItemKey("E-1", "store-1")},
                        New MailSummary() With {.Key = New ItemKey("E-2", "store-1")}},
                    .SkippedCount = 1,
                    .TotalAtStart = 3,
                    .NextCursor = Nothing})

            Dim v As New VarreduraDaPasta(b, db)
            Autorizar(db, v.Executar(Alvo, "Caixa", Store(), CancellationToken.None).ChaveDoAmbiente)
            Dim r = v.Executar(Alvo, "Caixa", Store(), CancellationToken.None)

            Dim m = New ManifestReader(db).Ler(r.Pasta)

            Assert.AreEqual(CType(1, Integer?), m.Descartadas)
            StringAssert.Contains(m.Ressalva, "recusada por não ser mensagem")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Zero não vira frase.</b>
    '''
    ''' "0 recusadas" é ruído numa faixa que já carrega a ressalva da §23.
    ''' E é o controle negativo do teste acima: sem ele, uma ressalva que
    ''' acrescentasse a frase sempre passaria.
    ''' </summary>
    <TestMethod>
    Public Sub ZERO_descartes_nao_vira_frase()
        Using db = Abrir()
            Dim v As New VarreduraDaPasta(Broker(), db)
            Autorizar(db, v.Executar(Alvo, "Caixa", Store(), CancellationToken.None).ChaveDoAmbiente)
            Dim r = v.Executar(Alvo, "Caixa", Store(), CancellationToken.None)

            Dim m = New ManifestReader(db).Ler(r.Pasta)

            Assert.AreEqual(CType(0, Integer?), m.Descartadas, "contou zero, e zero e um numero")
            Assert.IsFalse(m.Ressalva.Contains("recusad"), m.Ressalva)
        End Using
    End Sub

    ''' <summary>
    ''' <b>A guarda S6 continua mandando: descarte que não fecha a conta NÃO
    ''' publica.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' É o que dá sentido ao número guardado. O S6 exige que tudo o que a
    ''' pasta declarava tenha sido <b>lido ou explicitamente recusado</b> —
    ''' recusa declarada é mais forte que recusa silenciosa, porque obriga a
    ''' fonte a contar o que jogou fora.
    '''
    ''' Aqui a fonte diz que a pasta tinha 5, entrega 2 e declara 1 descarte:
    ''' faltam 2 que ninguém sabe onde estão. Publicar seria gravar um acervo
    ''' que perdeu mensagem sem dizer.
    ''' </summary>
    <TestMethod>
    Public Sub Conta_que_NAO_fecha_nao_publica()
        Using db = Abrir()
            Dim b = Broker()
            b.RespostaDaPagina = OperationResult(Of MessagePage).Ok(
                New MessagePage() With {
                    .Items = New List(Of MailSummary) From {
                        New MailSummary() With {.Key = New ItemKey("E-1", "store-1")},
                        New MailSummary() With {.Key = New ItemKey("E-2", "store-1")}},
                    .SkippedCount = 1,
                    .TotalAtStart = 5,
                    .NextCursor = Nothing})

            Dim v As New VarreduraDaPasta(b, db)
            Autorizar(db, v.Executar(Alvo, "Caixa", Store(), CancellationToken.None).ChaveDoAmbiente)
            Dim r = v.Executar(Alvo, "Caixa", Store(), CancellationToken.None)

            Assert.AreNotEqual(SweepConclusion.Publicada, r.Varredura.Conclusion,
                "2 lidas + 1 descartada = 3, e a pasta declarava 5")
            Assert.AreEqual(0, Contar(db, "generation"))
        End Using
    End Sub

    ''' <summary>
    ''' <b>O store medido tem de ser o da pasta varrida.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' O ambiente é medido a partir do <c>StoreInfo</c> recebido, e a
    ''' varredura acontece sobre <c>pasta.StoreId</c>. Sem conferir que são o
    ''' mesmo, a autorização dada a uma conta valeria para varrer <b>outra</b>
    ''' — basta a lista de stores de quem chama estar vencida.
    '''
    ''' A invariante não pode viver só na boa vontade de quem monta o par.
    ''' </summary>
    <TestMethod>
    Public Sub Store_de_OUTRA_conta_nao_autoriza_esta_pasta()
        Using db = Abrir()
            Dim v As New VarreduraDaPasta(Broker(), db)
            Autorizar(db, v.Executar(Alvo, "Caixa", Store(), CancellationToken.None).ChaveDoAmbiente)

            ' A pasta e do store-1; o StoreInfo e de outra conta.
            Dim deOutra As New StoreInfo() With {
                .StoreId = "store-OUTRO",
                .ExchangeStoreType = "olPrimaryExchangeMailbox",
                .IsCachedExchange = True}

            Dim r = v.Executar(Alvo, "Caixa", deOutra, CancellationToken.None)

            Assert.AreEqual(RecusaDeVarredura.StoreDesconhecido, r.Recusa,
                "a autorizacao de uma conta nao pode varrer a pasta de outra")
            Assert.AreEqual(0, Contar(db, "scan_attempt"))
        End Using
    End Sub

End Class

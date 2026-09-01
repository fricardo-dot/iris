Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Threading
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Model
Imports Iris.Sync
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O RÓTULO QUE SOBREVIVE A UMA VARREDURA.</b>
'''
''' ------------------------------------------------------------------
''' <b>O PROBLEMA</b>
'''
''' Toda republicação apagava a classificação inteira da pasta, inclusive a de
''' mensagens que não mudaram nada — uma varredura de manutenção jogava fora
''' todo o dinheiro gasto classificando.
'''
''' ------------------------------------------------------------------
''' <b>O CONTROLE NEGATIVO</b>
'''
''' <see cref="Mensagem_que_MUDOU_nao_herda_o_rotulo"/>. Sem ele, a herança
''' viraria "o rótulo é para sempre" — e um rótulo lido de um corpo que foi
''' reescrito é pior do que reclassificar, porque parece atual.
'''
''' E <see cref="Sem_saber_QUANDO_mudou_nao_se_herda"/>: metadado que não sabe
''' dizer quando a mensagem mudou não pode ser usado para afirmar que ela não
''' mudou.
''' </summary>
' NAO PARALELIZAR: cada teste abre um SQLite proprio.
<TestClass>
<DoNotParallelize>
Public Class HerancaDosRotulosTests

    Private Shared ReadOnly Quando As New DateTimeOffset(2026, 9, 1, 12, 0, 0, TimeSpan.Zero)

    ''' <summary>
    ''' <b>O caso que motivou tudo.</b> A varredura vê exatamente as mesmas
    ''' mensagens e republica; a classificação continua valendo.
    ''' </summary>
    <TestMethod>
    Public Sub Varredura_que_nao_mudou_nada_PRESERVA_o_rotulo()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, {Linha("a", tamanho:=100, mudouEm:="2026-08-01T10:00:00Z")})
                   Rotular(db, pasta, "a", "precisa_de_mim")

                   Varrer(db, {Linha("a", tamanho:=100, mudouEm:="2026-08-01T10:00:00Z")},
                          rodada:=2, existente:=pasta)

                   Dim lidos = New RotulosNoCache(db).Publicados(pasta)
                   Assert.AreEqual(1, lidos.Count, "a varredura apagou um rótulo válido")
                   Assert.AreEqual("precisa_de_mim", lidos("a").Rotulo)
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo.</b> Mensagem que mudou perde o rótulo — ele foi
    ''' lido de um texto que não está mais lá, e um rótulo assim é pior do que
    ''' nenhum, porque parece atual.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_que_MUDOU_nao_herda_o_rotulo()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, {Linha("a", tamanho:=100, mudouEm:="2026-08-01T10:00:00Z")})
                   Rotular(db, pasta, "a", "precisa_de_mim")

                   ' Mesmo tamanho, outra data: o corpo foi mexido.
                   Varrer(db, {Linha("a", tamanho:=100, mudouEm:="2026-08-30T09:00:00Z")},
                          rodada:=2, existente:=pasta)

                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(pasta).Count,
                       "herdou o rótulo de um corpo que mudou")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>Os dois critérios, e não um.</b> Uma edição que preserve a data —
    ''' relógio parado, importação — muda o tamanho.
    ''' </summary>
    <TestMethod>
    Public Sub Mesma_data_e_outro_TAMANHO_nao_herda()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, {Linha("a", tamanho:=100, mudouEm:="2026-08-01T10:00:00Z")})
                   Rotular(db, pasta, "a", "fyi")

                   Varrer(db, {Linha("a", tamanho:=250, mudouEm:="2026-08-01T10:00:00Z")},
                          rodada:=2, existente:=pasta)

                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(pasta).Count)
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O outro controle negativo.</b> Sem saber quando mudou, não se afirma
    ''' que não mudou. O preço é reclassificar; o preço do contrário é um rótulo
    ''' de um texto que não existe mais.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_saber_QUANDO_mudou_nao_se_herda()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, {Linha("a", tamanho:=100, mudouEm:=Nothing)})
                   Rotular(db, pasta, "a", "fyi")

                   Varrer(db, {Linha("a", tamanho:=100, mudouEm:=Nothing)},
                          rodada:=2, existente:=pasta)

                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(pasta).Count)
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>A observação carrega a data e a ativação originais.</b> O rótulo
    ''' <i>é</i> daquela hora e daquela autorização — carimbá-lo de hoje faria a
    ''' tela dizer que foi classificado agora, e apagaria a única prova de sob
    ''' que ativação o conteúdo saiu.
    ''' </summary>
    <TestMethod>
    Public Sub O_rotulo_herdado_guarda_a_ativacao_de_ORIGEM()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, {Linha("a", tamanho:=100, mudouEm:="2026-08-01T10:00:00Z")})
                   Rotular(db, pasta, "a", "fyi", ativacao:="ativacao-de-agosto")

                   Varrer(db, {Linha("a", tamanho:=100, mudouEm:="2026-08-01T10:00:00Z")},
                          rodada:=2, existente:=pasta)

                   Assert.AreEqual("ativacao-de-agosto",
                                   New RotulosNoCache(db).Publicados(pasta)("a").Ativacao)
               End Sub)
    End Sub

    ''' <summary>
    ''' Mensagem que saiu da pasta não volta pela herança: ela nem tem metadado
    ''' na geração nova, então não há o que comparar.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_que_SAIU_nao_volta_pela_heranca()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, {Linha("a", tamanho:=100, mudouEm:="2026-08-01T10:00:00Z"),
                                           Linha("b", tamanho:=100, mudouEm:="2026-08-01T10:00:00Z")})
                   Rotular(db, pasta, "a", "fyi")
                   Rotular(db, pasta, "b", "fyi")

                   ' "b" sumiu da pasta.
                   Varrer(db, {Linha("a", tamanho:=100, mudouEm:="2026-08-01T10:00:00Z")},
                          rodada:=2, existente:=pasta)

                   Dim lidos = New RotulosNoCache(db).Publicados(pasta)
                   Assert.AreEqual(1, lidos.Count)
                   Assert.IsTrue(lidos.ContainsKey("a"))
               End Sub)
    End Sub

    ' ==================================================================
    ' O ENTRYID REAPROVEITADO

    ''' <summary>
    ''' <b>Mesma chave, mesmo tamanho, mesma hora — e outro assunto.</b>
    '''
    ''' A herança casa pelo <c>EntryID</c>, e o Outlook não promete que ele nunca
    ''' se repita. Coincidir em tamanho <i>e</i> hora de modificação já era
    ''' improvável; o assunto fecha. Rótulo errado herdado é pior que rótulo
    ''' perdido — o perdido a próxima classificação repõe, o errado ninguém
    ''' revisita.
    ''' </summary>
    <TestMethod>
    Public Sub Chave_reaproveitada_com_outro_ASSUNTO_nao_herda()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, {Linha("a", 100, "2026-08-01T10:00:00Z")})
                   Rotular(db, pasta, "a", "fyi")

                   Varrer(db, {Linha("a", 100, "2026-08-01T10:00:00Z",
                                     titulo:="outra mensagem inteiramente")},
                          rodada:=2, existente:=pasta)

                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(pasta).Count,
                       "O ROTULO PASSOU PARA OUTRA MENSAGEM")
               End Sub)
    End Sub

    ''' <summary>O mesmo, pela data de recebimento.</summary>
    <TestMethod>
    Public Sub Chave_reaproveitada_com_outro_RECEBIMENTO_nao_herda()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, {Linha("a", 100, "2026-08-01T10:00:00Z")})
                   Rotular(db, pasta, "a", "fyi")

                   Varrer(db, {Linha("a", 100, "2026-08-01T10:00:00Z",
                                     recebidoEm:="2020-01-01T00:00:00.0000000+00:00")},
                          rodada:=2, existente:=pasta)

                   Assert.AreEqual(0, New RotulosNoCache(db).Publicados(pasta).Count,
                       "O ROTULO PASSOU PARA OUTRA MENSAGEM")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo, e ele é o ponto todo.</b>
    '''
    ''' As duas condições novas proíbem <i>discordar</i>; não exigem existir. Um
    ''' provedor que não entrega assunto faria a herança falhar sempre se a regra
    ''' fosse "tem de bater" — trocaria um risco remoto por uma perda certa, e a
    ''' pasta inteira seria reclassificada a cada varredura.
    ''' </summary>
    <TestMethod>
    Public Sub Assunto_DESCONHECIDO_dos_dois_lados_ainda_herda()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, {Linha("a", 100, "2026-08-01T10:00:00Z",
                                                 titulo:="")})
                   Rotular(db, pasta, "a", "fyi")

                   Varrer(db, {Linha("a", 100, "2026-08-01T10:00:00Z", titulo:="")},
                          rodada:=2, existente:=pasta)

                   Assert.AreEqual(1, New RotulosNoCache(db).Publicados(pasta).Count,
                       "a heranca virou refem de um campo que o provedor nao entrega")
               End Sub)
    End Sub

    ' ==================================================================
    ' O ANDAIME

    Private Shared Function Linha(chave As String, tamanho As Integer,
                                  mudouEm As String,
                                  Optional titulo As String = Nothing,
                                  Optional recebidoEm As String = Nothing) As SourceRow
        Return New SourceRow With {
            .Key = chave,
            .Subject = If(titulo, "assunto " & chave),
            .SenderName = "quem",
            .ReceivedAt = If(recebidoEm, Quando.ToString("o")),
            .LastModifiedAt = mudouEm,
            .SizeBytes = tamanho,
            .MessageClass = "IPM.Note"}
    End Function

    Private Shared Sub Rotular(db As CacheDatabase, pasta As Long, chave As String,
                               rotulo As String, Optional ativacao As String = "ativacao-1")
        Dim geracao As Long
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = "SELECT published_generation_key FROM folder WHERE folder_key = $f"
            cmd.Parameters.AddWithValue("$f", pasta)
            geracao = Convert.ToInt64(cmd.ExecuteScalar())
        End Using

        Dim cache = New RotulosNoCache(db)
        Dim r = cache.Gravar(pasta, geracao, ativacao, Quando,
                             New Dictionary(Of String, String) From {{chave, rotulo}},
                             New Dictionary(Of String, Double?)())
        Assert.IsTrue(r.Gravou AndAlso r.Entraram = 1,
                      "controle: o rótulo tinha de ser gravado")
    End Sub

    Private Shared Function Impressao() As EnvironmentFingerprint
        Return New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing)
    End Function

    Private Shared Function Varrer(db As CacheDatabase, linhas As SourceRow(),
                                   Optional rodada As Integer = 1,
                                   Optional existente As Long? = Nothing) As Long
        Dim resolvedor As New ResolvedorDoAcervo(db)
        Dim chave = If(existente.HasValue, existente.Value,
                       resolvedor.Pasta("store-1", "f-1", "Pasta f-1"))
        Dim amb = resolvedor.Ambiente(Impressao())

        Dim universo As New SweepUniverse("store-1", "f-1", "f", Nothing, rodada, "amb-1")
        Dim fonte As New FonteDeLinhas(universo, linhas)

        Dim sink As New SqliteSweepSink(db, chave, amb.Chave)
        Dim r = New SweepRunner(fonte, sink, 50).
                Executar(universo, 0, rodada, EnvironmentPolicy.Capacidades(Impressao()),
                         CancellationToken.None)
        Assert.IsTrue(r.Publicou, $"controle: a varredura tinha de publicar. motivo: {r.Motivo}")
        Return chave
    End Function

    Private Shared Sub Comigo(corpo As Action(Of CacheDatabase))
        Dim caminho = Path.Combine(Path.GetTempPath(),
                                   "iris-heranca-" & Guid.NewGuid().ToString("N") & ".db")
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

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Threading
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Model
Imports Iris.Sync
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A FILA LIDA DO ACERVO DE VERDADE.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ESTE ARQUIVO EXISTE, SE A LÓGICA JÁ ESTÁ TESTADA</b>
'''
''' <see cref="FilaDeRespostasTests"/> prova as regras sobre mensagens
''' fabricadas. Ele não prova que essas mensagens chegam: entre o cache e elas
''' há uma varredura, uma publicação, um dreno, um manifesto, e quatro estados
''' de presença. É a mesma distância que separou "as colunas da conversa
''' existem" de "as colunas da conversa vêm preenchidas" — um dia inteiro.
'''
''' Por isso aqui a semeadura passa pela <b>varredura de verdade</b>: fonte
''' falsa, <c>SweepRunner</c> real, <c>SqliteSweepSink</c> real, dreno real.
''' Semear a tabela na mão testaria SQL meu contra SQL meu.
'''
''' ------------------------------------------------------------------
''' <b>O CONTROLE NEGATIVO</b>
'''
''' <see cref="Pasta_de_enviados_conhecida_e_NUNCA_VARRIDA_recusa"/>. É o caso
''' que faz toda conversa já respondida parecer pendente, e ele não aparece em
''' teste de lógica pura: lá o "viu os enviados" é um parâmetro, e aqui é uma
''' pergunta ao banco.
''' </summary>
' NAO PARALELIZAR: cada teste abre um SQLite proprio, e a suite tem um teste
' que cobra este atributo de toda classe que toca o banco -- foi assim que a
' falha rara de 25/08/2026 apareceu.
<TestClass>
<DoNotParallelize>
Public Class FilaDoAcervoTests

    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 31, 12, 0, 0, TimeSpan.Zero)
    Private Shared ReadOnly Fuso As TimeZoneInfo = TimeZoneInfo.Utc

    Private Const Eu As String = "ricardo@empresa.com"
    Private Const Ela As String = "caroline@outra.com"

    Private Shared Function Identidades() As MinhasIdentidades
        Return New MinhasIdentidades({Eu})
    End Function

    ' ==================================================================

    <TestMethod>
    Public Sub A_fila_sai_do_acervo_com_as_duas_pontas()
        Comigo(Sub(db)
                   Dim entrada = Semear(db, "Caixa de Entrada", "f-entrada", {
                       ("c1", Ela, "pergunta antiga", 20),
                       ("c2", Ela, "outra pergunta", 5)})
                   Dim enviados = Semear(db, "Itens Enviados", "f-enviados", {
                       ("c1", Eu, "ja respondi", 19)})

                   Dim r = Montar(db, enviados)

                   Assert.AreEqual(MotivoDaFila.Respondida, r.Motivo)
                   Assert.AreEqual(2, r.Linhas.Count)

                   ' c1 terminou COMIGO: a espera e dela.
                   Assert.AreEqual(1, r.Deles().Count)
                   Assert.AreEqual("ja respondi", r.Deles()(0).Assunto)

                   ' c2 terminou com ELA: pode ser a minha vez.
                   Assert.AreEqual(1, r.Minhas().Count)
                   Assert.AreEqual("outra pergunta", r.Minhas()(0).Assunto)
                   Assert.AreEqual(5, r.Minhas()(0).Dias)
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo.</b> A pasta de enviados existe no acervo e nunca
    ''' foi varrida — e é exatamente aí que a fila mentiria mais: sem ver as
    ''' respostas do dono, toda conversa respondida vira pendência.
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_de_enviados_conhecida_e_NUNCA_VARRIDA_recusa()
        Comigo(Sub(db)
                   Semear(db, "Caixa de Entrada", "f-entrada", {("c1", Ela, "pergunta", 20)})
                   Dim enviados = SoRegistrar(db, "Itens Enviados", "f-enviados")

                   Dim r = Montar(db, enviados)

                   Assert.AreEqual(MotivoDaFila.SemOsEnviados, r.Motivo,
                       "montou a fila sem ter varrido os enviados: a conversa " &
                       "respondida apareceria como pendente")
                   Assert.AreEqual(0, r.Linhas.Count)
               End Sub)
    End Sub

    ''' <summary>
    ''' Sem saber <b>qual</b> pasta é a de enviados, a fila recusa igual. Não
    ''' saber é o mesmo que não ter varrido — nos dois casos as respostas do dono
    ''' estão fora do alcance.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_saber_qual_pasta_e_a_dos_enviados_recusa()
        Comigo(Sub(db)
                   Semear(db, "Caixa de Entrada", "f-entrada", {("c1", Ela, "pergunta", 20)})
                   Semear(db, "Itens Enviados", "f-enviados", {("c1", Eu, "resposta", 19)})

                   Assert.AreEqual(MotivoDaFila.SemOsEnviados, Montar(db, Nothing).Motivo)
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>Mensagem que saiu da pasta não vira pendência.</b> Uma linha sobre uma
    ''' mensagem ausente manda o dono abrir o que não existe — e a conversa passa
    ''' a ser decidida pela mensagem anterior, que é a verdade do acervo.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_AUSENTE_da_pasta_nao_entra()
        Comigo(Sub(db)
                   Dim entrada = Semear(db, "Caixa de Entrada", "f-entrada", {
                       ("c1", Ela, "a que ficou", 20),
                       ("c1", Ela, "a que sumiu", 2)})
                   Dim enviados = Semear(db, "Itens Enviados", "f-enviados", {
                       ("c9", Eu, "outra conversa", 30)})

                   ' A segunda mensagem some da pasta: varre de novo sem ela.
                   Revarrer(db, entrada, "f-entrada", {("c1", Ela, "a que ficou", 20)})

                   Dim r = Montar(db, enviados)
                   Dim c1 = r.Linhas.Single(Function(l) l.Conversa = "c1")

                   Assert.AreEqual("a que ficou", c1.Assunto,
                       "a fila apontou para uma mensagem que nao esta mais na pasta")
                   Assert.AreEqual(20, c1.Dias)
               End Sub)
    End Sub

    ''' <summary>
    ''' A chave da linha carrega o <b>store</b>, e não só o EntryID. Sem ele a
    ''' tela teria o identificador e não teria onde abri-lo — e o defeito
    ''' apareceria só em quem tem mais de uma caixa.
    ''' </summary>
    <TestMethod>
    Public Sub A_chave_da_linha_carrega_o_STORE()
        Comigo(Sub(db)
                   Semear(db, "Caixa de Entrada", "f-entrada", {("c1", Ela, "pergunta", 5)})
                   Dim enviados = Semear(db, "Itens Enviados", "f-enviados", {
                       ("c9", Eu, "outra", 30)})

                   Dim linha = Montar(db, enviados).Linhas.Single(Function(l) l.Conversa = "c1")

                   Assert.AreEqual("store-1", linha.Chave.StoreId,
                       "ItemKey sem store nao identifica mensagem fora de uma caixa so")
                   Assert.IsTrue(linha.Chave.EntryId.Length > 0)
               End Sub)
    End Sub

    ''' <summary>
    ''' Data escrita pelo cache é ISO 8601 com deslocamento, e é lida com a
    ''' cultura invariante. Ler com a cultura da máquina faria dia virar mês em
    ''' metade do calendário — sem erro, e errado só em alguns dias do mês.
    ''' </summary>
    <TestMethod>
    Public Sub A_data_e_lida_pela_cultura_INVARIANTE()
        Dim quando = New DateTimeOffset(2026, 3, 11, 8, 30, 0, TimeSpan.FromHours(-3))

        Dim lida = FilaDoAcervo.Instante(quando.ToString("o"))

        Assert.IsTrue(lida.HasValue)
        Assert.AreEqual(quando, lida.Value)
        Assert.AreEqual(3, lida.Value.Month, "o dia 11 virou o mes 11")

        Assert.IsFalse(FilaDoAcervo.Instante("nao e data").HasValue)
        Assert.IsFalse(FilaDoAcervo.Instante("").HasValue)
    End Sub

    ' ==================================================================
    ' O ANDAIME

    Private Shared Function Montar(db As CacheDatabase, enviados As Long?) As ResultadoDaFila
        Dim todas As New AcervoDeTodasAsPastas(db)
        Dim dreno As New PublicationDrain(db)

        ' A MESMA SEQUENCIA DA PRODUCAO: drena primeiro, e so carrega a mao se
        ' nada veio. Um andaime que fizesse diferente testaria outro programa.
        dreno.Drenar(todas)
        If todas.Recarregado = 0 Then todas.Recarregar()

        Return New FilaDoAcervo(todas).Montar(Identidades(), Agora, Fuso, enviados, Nothing)
    End Function

    Private Shared Function Impressao() As EnvironmentFingerprint
        Return New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing)
    End Function

    ''' <summary>
    ''' Semeia uma pasta varrida e publicada. Passa pela varredura de verdade —
    ''' semear a tabela na mão testaria SQL meu contra SQL meu.
    ''' </summary>
    Private Shared Function Semear(db As CacheDatabase, nome As String, entryId As String,
                                   linhas As IEnumerable(Of (Conversa As String,
                                                             Remetente As String,
                                                             Assunto As String,
                                                             DiasAtras As Integer))) As Long
        Dim chave = New ResolvedorDoAcervo(db).Pasta("store-1", entryId, nome)
        Varrer(db, chave, entryId, linhas, 1)
        Return chave
    End Function

    ''' <summary>Varre a mesma pasta de novo, com outro conjunto de linhas.</summary>
    Private Shared Sub Revarrer(db As CacheDatabase, chave As Long, entryId As String,
                                linhas As IEnumerable(Of (Conversa As String,
                                                          Remetente As String,
                                                          Assunto As String,
                                                          DiasAtras As Integer)))
        Varrer(db, chave, entryId, linhas, 2)
    End Sub

    Private Shared Sub Varrer(db As CacheDatabase, chave As Long, entryId As String,
                              linhas As IEnumerable(Of (Conversa As String,
                                                        Remetente As String,
                                                        Assunto As String,
                                                        DiasAtras As Integer)),
                              rodada As Integer)
        Dim resolvedor As New ResolvedorDoAcervo(db)
        Dim amb = resolvedor.Ambiente(Impressao())

        Dim universo As New SweepUniverse("store-1", entryId, "f", Nothing, rodada, "amb-1")
        Dim fonte As New FonteDeLinhas(universo, linhas.Select(
            Function(l) New SourceRow With {
                .Key = $"{entryId}-{l.Conversa}-{l.Assunto}",
                .Subject = l.Assunto,
                .SenderName = l.Remetente,
                .SenderAddress = l.Remetente,
                .ConversationId = l.Conversa,
                .ReceivedAt = Agora.AddDays(-l.DiasAtras).ToString("o"),
                .MessageClass = "IPM.Note"}))

        Dim sink As New SqliteSweepSink(db, chave, amb.Chave)
        Dim cap = EnvironmentPolicy.Capacidades(Impressao())
        Dim r = New SweepRunner(fonte, sink, 50).
                Executar(universo, 0, rodada, cap, CancellationToken.None)
        Assert.IsTrue(r.Publicou, $"controle: a semeadura tinha de publicar. motivo: {r.Motivo}")

        Dim servico As New AcervoService(db, chave)
        Dim dreno As New PublicationDrain(db)
        dreno.Drenar(servico)
    End Sub

    ''' <summary>Uma pasta conhecida que NUNCA foi varrida.</summary>
    Private Shared Function SoRegistrar(db As CacheDatabase, nome As String,
                                        entryId As String) As Long
        Return New ResolvedorDoAcervo(db).Pasta("store-1", entryId, nome)
    End Function

    Private Shared Sub Comigo(corpo As Action(Of CacheDatabase))
        Dim caminho = Path.Combine(Path.GetTempPath(),
                                   "iris-fila-" & Guid.NewGuid().ToString("N") & ".db")
        Try
            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(caminho, CacheSchema.Intended(), falha)
                Assert.IsNotNull(db, $"{falha}")
                corpo(db)
            End Using
        Finally
            Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools()
            For Each sufixo In {"", "-wal", "-shm"}
                If File.Exists(caminho & sufixo) Then File.Delete(caminho & sufixo)
            Next
        End Try
    End Sub

End Class

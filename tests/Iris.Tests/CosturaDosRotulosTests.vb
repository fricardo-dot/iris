Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Threading
Imports Iris.App.ViewModels
Imports Iris.Assist
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Model
Imports Iris.Sync
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A COSTURA: da classificação até a gaveta na tela.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ESTE ARQUIVO EXISTE</b>
'''
''' As dez fases foram testadas como ilhas, e cada ilha ficou verde. O que
''' ninguém exercitava era o caminho inteiro:
'''
''' <c>ClassificarUmaPasta</c> → <c>RotulosNoCache</c> →
''' <c>RotulosDoAcervo</c> → <c>CaixasSeparadas</c> → <c>CaixasViewModel</c>.
'''
''' Uma troca de chave, de pasta, de geração ou do nome textual do rótulo entre
''' duas dessas APIs deixa todos os testes de unidade verdes e produz gavetas
''' erradas ou vazias. Achado por revisão externa em 31/08/2026.
'''
''' ------------------------------------------------------------------
''' <b>O CONTROLE NEGATIVO</b>
'''
''' <see cref="Duas_pastas_que_DISCORDAM_nao_escolhem_uma"/>. Sem ele, a ponte
''' que junta as pastas faria a última vencer — e "a última" é a ordem em que o
''' acervo enumerou, que não quer dizer nada. A mensagem apareceria numa gaveta
''' que metade da evidência contradiz.
''' </summary>
' NAO PARALELIZAR: cada teste abre um SQLite proprio, e a passagem tem porta
' estatica de uma por vez.
<TestClass>
<DoNotParallelize>
Public Class CosturaDosRotulosTests

    Private Shared ReadOnly Quando As New DateTimeOffset(2026, 8, 31, 12, 0, 0, TimeSpan.Zero)

    ' ==================================================================
    ' O CAMINHO INTEIRO

    ''' <summary>
    ''' Classificar de verdade, ler de verdade, e a mensagem aparece na gaveta
    ''' do rótulo que o modelo devolveu — com o nome que a tela mostra.
    ''' </summary>
    <TestMethod>
    Public Sub Da_classificacao_ate_a_gaveta_na_tela()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, "f-1", {"pedido-do-cliente"})
                   Classificar(db, pasta, "precisa_de_mim")

                   Dim vm = Caixas(db)
                   vm.Atualizar()

                   Dim esperam = vm.Gavetas.Single(Function(g) g.Nome = "Esperam você")
                   Assert.AreEqual(1, esperam.Quantas,
                       "o rótulo gravado não chegou à gaveta")
                   Assert.AreEqual("pedido-do-cliente", esperam.Mensagens.Single().Assunto)

                   StringAssert.Contains(vm.Cobertura, "1 de 1")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>Sem saber se a mensagem mudou, varrer de novo apaga a classificação —
    ''' e a tela diz isso.</b>
    '''
    ''' O andaime daqui não preenche <c>last_modified_at</c> nem
    ''' <c>size_bytes</c>, então a herança de rótulos não consegue provar que a
    ''' mensagem continua a mesma e não herda nada. É o caso conservador, e é o
    ''' que este teste prende: a caixa dividida volta a dizer "não classificada"
    ''' em vez de continuar mostrando a gaveta de antes.
    '''
    ''' <b>Com o metadado preenchido o desfecho é outro</b>, e ele tem arquivo
    ''' próprio: <c>HerancaDosRotulosTests</c>.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_metadado_de_mudanca_a_varredura_devolve_para_as_nao_classificadas()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, "f-1", {"pedido"})
                   Classificar(db, pasta, "precisa_de_mim")

                   Varrer(db, "f-1", {"pedido"}, rodada:=2, existente:=pasta)

                   Dim vm = Caixas(db)
                   vm.Atualizar()

                   Assert.AreEqual(0, vm.Gavetas.Single(Function(g) g.Nome = "Esperam você").Quantas)
                   Assert.AreEqual(1, vm.Gavetas.Single(
                       Function(g) g.Nome = CaixasSeparadas.NomeDasNaoClassificadas).Quantas)
                   StringAssert.Contains(vm.Cobertura, "Nenhuma das 1")
               End Sub)
    End Sub

    ' ==================================================================
    ' DUAS PASTAS, UMA MENSAGEM

    ''' <summary>
    ''' <b>O controle negativo.</b> A mesma mensagem em duas pastas, com rótulos
    ''' diferentes. O cache guarda as duas observações — cada uma é verdadeira
    ''' na sua pasta —, e a ponte para <c>ItemKey</c> perde a pasta.
    '''
    ''' Escolher uma seria escolher pela ordem de enumeração. Ela vai para as não
    ''' classificadas, e o desacordo é contado.
    ''' </summary>
    <TestMethod>
    Public Sub Duas_pastas_que_DISCORDAM_nao_escolhem_uma()
        Comigo(Sub(db)
                   Dim a = Varrer(db, "f-1", {"o-mesmo-item"})
                   Dim b = Varrer(db, "f-2", {"o-mesmo-item"})
                   Classificar(db, a, "precisa_de_mim")
                   Classificar(db, b, "fyi")

                   Dim lido = New RotulosDoAcervo(Acervo(db), New RotulosNoCache(db)).Ler()

                   Assert.AreEqual(1, lido.Divergentes, "o desacordo sumiu")
                   Assert.AreEqual(0, lido.Rotulos.Count,
                       "escolheu um rótulo sem ter como escolher")
                   Assert.AreEqual(2, lido.PastasLidas)
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>E o contraponto</b>: duas pastas dizendo a mesma coisa são uma
    ''' informação só, repetida. Sem este teste, a regra do desacordo podia
    ''' virar "toda mensagem em duas pastas some", que é o defeito oposto e mais
    ''' comum.
    ''' </summary>
    <TestMethod>
    Public Sub Duas_pastas_que_CONCORDAM_valem_uma()
        Comigo(Sub(db)
                   Dim a = Varrer(db, "f-1", {"o-mesmo-item"})
                   Dim b = Varrer(db, "f-2", {"o-mesmo-item"})
                   Classificar(db, a, "fyi")
                   Classificar(db, b, "fyi")

                   Dim lido = New RotulosDoAcervo(Acervo(db), New RotulosNoCache(db)).Ler()

                   Assert.AreEqual(0, lido.Divergentes)
                   Assert.AreEqual("fyi", lido.Rotulos.Values.Single())
               End Sub)
    End Sub

    ' ==================================================================
    ' O CONTRATO QUE ATRAVESSA AS CAMADAS

    ''' <summary>
    ''' <b>Nulo e vazio sobrevivem do banco até a leitura.</b>
    '''
    ''' Ausência da chave é "ninguém respondeu sobre as regras"; lista vazia é
    ''' "respondeu, e nenhuma casou". A distinção nasce no cache e morria em
    ''' cada camada nova que a atravessava sem saber dela.
    ''' </summary>
    <TestMethod>
    Public Sub Nulo_e_vazio_das_regras_atravessam_a_ponte()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, "f-1", {"mudo", "falou"})
                   Dim geracao = GeracaoPublicada(db, pasta)

                   Dim cache = New RotulosNoCache(db)
                   cache.Gravar(pasta, geracao, "ativacao-1", Quando,
                                New Dictionary(Of String, String) From {
                                    {"mudo", "fyi"}, {"falou", "fyi"}},
                                New Dictionary(Of String, Double?)(),
                                New Dictionary(Of String, IReadOnlyList(Of String)) From {
                                    {"falou", Array.Empty(Of String)()}})

                   Dim lido = New RotulosDoAcervo(Acervo(db), cache).Ler()

                   Assert.IsFalse(lido.RegrasCasadas.ContainsKey(Chave("mudo")),
                       "silêncio virou resposta")
                   Assert.IsTrue(lido.RegrasCasadas.ContainsKey(Chave("falou")))
                   Assert.AreEqual(0, lido.RegrasCasadas(Chave("falou")).Count)
               End Sub)
    End Sub

    ''' <summary>
    ''' A regra do dono casada vira gaveta na tela, com a frase dele. É a Fase 6
    ''' chegando à Fase 7 pelo caminho real, e não pelo dublê.
    ''' </summary>
    <TestMethod>
    Public Sub A_regra_do_dono_vira_gaveta_pelo_caminho_real()
        Comigo(Sub(db)
                   Dim pasta = Varrer(db, "f-1", {"reclamacao"})
                   Dim geracao = GeracaoPublicada(db, pasta)

                   Dim cache = New RotulosNoCache(db)
                   cache.Gravar(pasta, geracao, "ativacao-1", Quando,
                                New Dictionary(Of String, String) From {{"reclamacao", "fyi"}},
                                New Dictionary(Of String, Double?)(),
                                New Dictionary(Of String, IReadOnlyList(Of String)) From {
                                    {"reclamacao", {"clientes reclamando"}}})

                   Dim vm = Caixas(db, {"clientes reclamando"})
                   vm.Atualizar()

                   Dim dele = vm.Gavetas.Single(Function(g) g.DoDono)
                   Assert.AreEqual("clientes reclamando", dele.Nome)
                   Assert.AreEqual(1, dele.Quantas)
                   Assert.AreEqual(0, vm.Gavetas.Single(Function(g) g.Nome = "Só para saber").Quantas)
               End Sub)
    End Sub

    ' ==================================================================
    ' O ANDAIME

    Private Shared Function Chave(id As String) As ItemKey
        Return New ItemKey(id, "store-1")
    End Function

    Private Shared Function Caixas(db As CacheDatabase,
                                   Optional regras As String() = Nothing) As CaixasViewModel
        Dim todas = Acervo(db)
        Dim lido = New RotulosDoAcervo(todas, New RotulosNoCache(db)).Ler()

        Return New CaixasViewModel(
            Function() New FilaDoAcervo(todas).Mensagens(),
            Function() lido.Rotulos,
            Function() lido.RegrasCasadas,
            Function() If(regras, Array.Empty(Of String)()))
    End Function

    ''' <summary>
    ''' Uma passagem de verdade, com o modelo obediente: devolve o rótulo pedido
    ''' e acerta o controle lido da instrução.
    ''' </summary>
    Private Shared Sub Classificar(db As CacheDatabase, pasta As Long, rotulo As String)
        Dim passagem As New ClassificarUmaPasta(Acervo(db), New RotulosNoCache(db))

        Dim r = passagem.Passar(
            pasta, Nothing, "ativacao-1", Quando,
            Function(pedidos, ct) pedidos.Select(
                Function(p) New MessagePart(p.Chave, "CK", "assunto", "de",
                                            {"para"}, "corpo", True, p.Ficha)).ToList(),
            Function(instrucao, partes, ct)
                Dim doControle = OControle(instrucao)
                Dim itens = partes.Select(
                    Function(p) "{""item_key"":""" & p.Ficha & """,""label"":""" &
                                If(p.Ficha = doControle.Ficha, doControle.Rotulo, rotulo) &
                                """}")
                Return "[" & String.Join(",", itens) & "]"
            End Function)

        Assert.AreEqual(MotivoDaClassificacao.Passou, r.Motivo,
            "controle: a passagem tinha de rodar")
        Assert.IsTrue(r.Classificados > 0, "controle: a passagem tinha de classificar")
    End Sub

    Private Shared Function OControle(instrucao As String) As (Ficha As String, Rotulo As String)
        Dim marca = "A mensagem de item_key "
        Dim i = instrucao.IndexOf(marca, StringComparison.Ordinal)
        Assert.IsTrue(i >= 0, "a instrução não anunciou o controle")

        Dim resto = instrucao.Substring(i + marca.Length)
        Dim ficha = resto.Substring(0, resto.IndexOf(" "c))

        Dim antes = "classifique-a como "
        Dim j = resto.IndexOf(antes, StringComparison.Ordinal)
        Dim rotulo = resto.Substring(j + antes.Length)
        rotulo = rotulo.Substring(0, rotulo.IndexOf(","c))

        Return (ficha, rotulo)
    End Function

    Private Shared Function Acervo(db As CacheDatabase) As AcervoDeTodasAsPastas
        Dim todas As New AcervoDeTodasAsPastas(db)
        Dim dreno As New PublicationDrain(db)
        dreno.Drenar(todas)
        If todas.Recarregado = 0 Then todas.Recarregar()
        Return todas
    End Function

    Private Shared Function GeracaoPublicada(db As CacheDatabase, pasta As Long) As Long
        Using cmd = db.Connection.CreateCommand()
            cmd.CommandText = "SELECT published_generation_key FROM folder WHERE folder_key = $f"
            cmd.Parameters.AddWithValue("$f", pasta)
            Return Convert.ToInt64(cmd.ExecuteScalar())
        End Using
    End Function

    Private Shared Function Impressao() As EnvironmentFingerprint
        Return New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing)
    End Function

    ''' <summary>
    ''' <b>A chave do item é a mesma nas duas pastas quando o sufixo é o mesmo</b>
    ''' — é isso que faz o teste do desacordo ser possível: um item que aparece
    ''' em duas pastas do mesmo store vira o mesmo <c>ItemKey</c>.
    ''' </summary>
    Private Shared Function Varrer(db As CacheDatabase, pastaId As String,
                                   sufixos As String(),
                                   Optional rodada As Integer = 1,
                                   Optional existente As Long? = Nothing) As Long
        Dim resolvedor As New ResolvedorDoAcervo(db)
        Dim chave = If(existente.HasValue, existente.Value,
                       resolvedor.Pasta("store-1", pastaId, "Pasta " & pastaId))
        Dim amb = resolvedor.Ambiente(Impressao())

        Dim universo As New SweepUniverse("store-1", pastaId, "f", Nothing, rodada, "amb-1")
        ' A CHAVE NAO LEVA O NOME DA PASTA. Um item que aparece em duas pastas do
        ' mesmo store tem de virar o MESMO ItemKey -- e a primeira versao deste
        ' andaime prefixava com a pasta, o que fazia o teste do desacordo nunca
        ' encontrar desacordo nenhum.
        '
        ' (Comentario AQUI, e nao dentro do inicializador com chaves: em VB ele
        ' quebra a continuacao implicita, e o erro sai nas linhas seguintes.)
        Dim fonte As New FonteDeLinhas(universo, sufixos.Select(
            Function(s) New SourceRow With {
                .Key = s,
                .Subject = s,
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

    Private Shared Sub Comigo(corpo As Action(Of CacheDatabase))
        Dim caminho = Path.Combine(Path.GetTempPath(),
                                   "iris-costura-" & Guid.NewGuid().ToString("N") & ".db")
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

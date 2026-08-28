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
''' <b>BUSCA TEXTUAL SOBRE O ACERVO.</b>
'''
''' ------------------------------------------------------------------
''' O <c>ESCOPO.md</c> dizia que a Fase 2 tinha entregue busca textual. Não
''' tinha. Estes testes existem junto com a entrega que faltava.
'''
''' O que eles cobram, em ordem de importância:
'''
'''   1. Que ela <b>ache</b> — controle positivo, antes de tudo. Uma busca que
'''      nunca acha nada passa em todos os testes de "não afirma demais".
'''   2. Que ela ache <b>sem acento e sem caixa</b>, porque a caixa é em
'''      português e o <c>LIKE</c> do SQLite não faria isso.
'''   3. Que zero achados <b>não</b> vire "não existe".
'''   4. Que pasta nunca varrida não seja confundida com pasta sem resultado.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class BuscaNoAcervoTests

    Private _pasta As String
    Private _caminho As String

    Private Shared ReadOnly EnvKey As Long = 1

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-busca-" & Guid.NewGuid().ToString("N"))
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

    ''' <summary>
    ''' Semeia uma pasta varrida e publicada, com as linhas dadas.
    '''
    ''' Passa pela varredura de verdade — fonte falsa, <c>SweepRunner</c> real,
    ''' <c>SqliteSweepSink</c> real. Semear a tabela na mão testaria SQL meu
    ''' contra SQL meu.
    ''' </summary>
    Private Shared Function Semear(db As CacheDatabase, nome As String,
                                   entryId As String,
                                   linhas As IEnumerable(Of (Assunto As String, Remetente As String))) As Long
        Dim resolvedor As New ResolvedorDoAcervo(db)
        Dim amb = resolvedor.Ambiente(Impressao())
        Dim chave = resolvedor.Pasta("store-1", entryId, nome)

        Dim universo As New SweepUniverse("store-1", entryId, "f", Nothing, 1, "amb-1")
        Dim fonte As New FonteDeLinhas(universo, linhas.Select(
            Function(l, i) New SourceRow With {
                .Key = $"{entryId}-{i}",
                .Subject = l.Assunto,
                .SenderName = l.Remetente,
                .ReceivedAt = New DateTimeOffset(2026, 8, 20, 9, i Mod 60, 0, TimeSpan.Zero).ToString("o"),
                .MessageClass = "IPM.Note"}))
        Dim sink As New SqliteSweepSink(db, chave, amb.Chave)
        Dim cap = EnvironmentPolicy.Capacidades(Impressao())
        Dim r = New SweepRunner(fonte, sink, 50).
                Executar(universo, 0, 1, cap, CancellationToken.None)
        Assert.IsTrue(r.Publicou, $"controle: a semeadura tinha de publicar. motivo: {r.Motivo}")

        ' A publicacao so vira acervo depois do dreno. Semear sem drenar
        ' deixaria a busca lendo um manifesto que a §26.2 diz que ninguem
        ' deveria estar vendo ainda.
        Dim servico As New AcervoService(db, chave)
        Dim dreno As New PublicationDrain(db)
        dreno.Drenar(servico)
        Return chave
    End Function

    ''' <summary>Uma pasta conhecida que NUNCA foi varrida.</summary>
    Private Shared Function SoRegistrar(db As CacheDatabase, nome As String, entryId As String) As Long
        Return New ResolvedorDoAcervo(db).Pasta("store-1", entryId, nome)
    End Function

    ''' <summary>
    ''' A busca como a producao a monta: sobre o acervo DRENADO.
    '''
    ''' Ate 28/08/2026 a tarde os testes construiam a busca com o banco, e
    ''' ela abria o ManifestReader sozinha -- o contorno da §26.2. Passar
    ''' pelo dreno aqui e o que faz o teste exercitar o caminho real.
    ''' </summary>
    Private Shared Function Buscar(db As CacheDatabase) As BuscaNoAcervo
        Dim todas As New AcervoDeTodasAsPastas(db)
        Dim dreno As New PublicationDrain(db)

        ' A MESMA SEQUENCIA DA PRODUCAO, e ela nao e acidental.
        '
        ' O acervo nasce VAZIO de proposito: ler no construtor seria ler na
        ' frente do dreno quando ha publicacao pendente de uma queda
        ' anterior. Entao drena primeiro -- cada entrega recarrega -- e so
        ' carrega a mao se nada veio, que e o caso da abertura normal.
        '
        ' Um helper que fizesse diferente do AcervoViewModel testaria outro
        ' programa.
        dreno.Drenar(todas)
        If todas.Recarregado = 0 Then todas.Recarregar()

        Return New BuscaNoAcervo(todas, dreno)
    End Function

    Private Shared ReadOnly Caixa As (String, String)() = {
        ("APROVAÇÃO AUDACCI: Cart. Brainmetyl 5ml", "Regulatório - Kate"),
        ("RES: Solicitação de informações sobre regularização", "Regulatório - Kate"),
        ("RE: Amostras de Aquaba TX 180", "Andre Bonini"),
        ("Contrato assinado", "Caroline Abreu"),
        ("almoço de sexta", "Aguinaldo")
    }

    ' ==================================================================

    ''' <summary>
    ''' <b>CONTROLE POSITIVO, e ele vem primeiro de propósito.</b>
    '''
    ''' Uma busca que nunca acha nada satisfaz todos os testes de "não afirma
    ''' ausência" que vêm depois. Sem esta linha, os outros não provam nada —
    ''' é a armadilha que o <c>CLAUDE.md</c> descreve com o compositor que
    ''' nunca envia.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_a_busca_ACHA()
        Using db = Abrir()
            Semear(db, "Caixa de Entrada", "entrada", Caixa)

            Dim r = Buscar(db).Procurar("contrato")

            Assert.AreEqual(1, r.Achados.Count, "controle: tinha de achar o contrato")
            Assert.AreEqual("Contrato assinado", r.Achados(0).Item.Subject)
            Assert.AreEqual("Caixa de Entrada", r.Achados(0).NomeDaPasta)
        End Using
    End Sub

    ''' <summary>
    ''' <b>Sem acento e sem caixa alta.</b>
    '''
    ''' É a razão de o casamento ser em memória: o <c>LIKE</c> do SQLite ignora
    ''' maiúsculas só em ASCII, e numa caixa em português isso faria
    ''' "Regulatório" e "regulatorio" serem palavras diferentes. Uma busca que
    ''' não acha o que o usuário está vendo na tela é pior que busca nenhuma.
    ''' </summary>
    <TestMethod>
    Public Sub Acha_sem_acento_e_sem_caixa()
        Using db = Abrir()
            Semear(db, "Caixa de Entrada", "entrada", Caixa)
            Dim busca = Buscar(db)

            Assert.AreEqual(2, busca.Procurar("REGULATORIO").Achados.Count,
                "maiuscula sem acento tinha de achar 'Regulatório - Kate'")
            Assert.AreEqual(2, busca.Procurar("regulatório").Achados.Count,
                "minuscula com acento tinha de achar o mesmo")
            Assert.AreEqual(1, busca.Procurar("APROVACAO").Achados.Count,
                "o assunto em maiuscula com til tinha de casar sem til")
            Assert.AreEqual(1, busca.Procurar("almoco").Achados.Count,
                "cedilha tambem")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Duas palavras exigem as duas, e podem cair em campos diferentes.</b>
    '''
    ''' Conjunção: "amostras aquaba" que devolvesse tudo o que tem "amostras"
    ''' seria ruído com cara de resultado.
    '''
    ''' E o alvo é assunto e remetente <b>juntos</b>: procurar "kate
    ''' regularizacao" tem de achar a mensagem cujo assunto tem uma palavra e
    ''' cujo remetente tem a outra. Exigir que caiam no mesmo campo faria a
    ''' busca falhar por um motivo que o usuário não tem como adivinhar.
    ''' </summary>
    <TestMethod>
    Public Sub Duas_palavras_sao_conjuncao_e_atravessam_os_campos()
        Using db = Abrir()
            Semear(db, "Caixa de Entrada", "entrada", Caixa)
            Dim busca = Buscar(db)

            Assert.AreEqual(1, busca.Procurar("amostras aquaba").Achados.Count)
            Assert.AreEqual(0, busca.Procurar("amostras contrato").Achados.Count,
                "conjuncao: nenhuma mensagem tem as duas")
            Assert.AreEqual(1, busca.Procurar("kate regularizacao").Achados.Count,
                "uma palavra no remetente e outra no assunto")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Zero achados NÃO é "não existe" — e a ressalva diz isso.</b>
    '''
    ''' A §23 proíbe concluir ausência, e busca é onde o usuário mais
    ''' interpreta silêncio como resposta.
    ''' </summary>
    <TestMethod>
    Public Sub Zero_achados_nao_afirma_ausencia()
        Using db = Abrir()
            Semear(db, "Caixa de Entrada", "entrada", Caixa)

            Dim r = Buscar(db).Procurar("palavraquenaoexiste")

            Assert.AreEqual(0, r.Achados.Count)
            StringAssert.Contains(r.Ressalva, "Nada no acervo observado")
            StringAssert.Contains(r.Ressalva, "não quer dizer que não exista")
            StringAssert.Contains(r.Ressalva, "acervo é parcial")

            ' A LIMITAÇÃO QUE MAIS SURPREENDE: o corpo não é procurável.
            StringAssert.Contains(r.Ressalva, "corpo da mensagem não")

            ' E o resultado carrega ONDE se procurou, no mesmo objeto.
            Assert.AreEqual(1, r.Consultadas.Count)
            Assert.AreEqual(5, r.TotalNoAcervo)
            Assert.IsTrue(r.AlgumaParcial, "em cached a cobertura e sempre parcial")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Pasta sem acervo publicado não é pasta sem resultado.</b>
    '''
    ''' Misturá-las faria o resultado dizer "procurei aqui e não achei" sobre
    ''' um lugar onde ninguém procurou. É a mesma distinção entre
    ''' <c>Nothing</c> e zero que o projeto já faz nas linhas descartadas.
    '''
    ''' <b>E o nome do estado mudou em 28/08/2026.</b> Eu chamava estas pastas
    ''' de "nunca varridas", e a revisão externa mostrou que é mais do que se
    ''' sabe: sem geração publicada cabe também a tentativa rejeitada pela S6,
    ''' a cancelada e a que falhou. O cache afirma que <b>não há acervo
    ''' publicado</b>, e é só isso que a frase pode dizer.
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_SEM_ACERVO_PUBLICADO_fica_SEPARADA_das_consultadas()
        Using db = Abrir()
            Semear(db, "Caixa de Entrada", "entrada", Caixa)
            SoRegistrar(db, "Itens Enviados", "enviados")

            Dim r = Buscar(db).Procurar("contrato")

            Assert.AreEqual(1, r.Consultadas.Count, "so a varrida foi consultada")
            Assert.AreEqual("Caixa de Entrada", r.Consultadas(0).Nome)
            Assert.AreEqual(1, r.SemAcervo.Count, "a nao varrida tem de aparecer, e a parte")
            Assert.AreEqual("Itens Enviados", r.SemAcervo(0).Nome)
            StringAssert.Contains(r.Ressalva, "não têm acervo publicado")
            StringAssert.Contains(r.Ressalva, "Itens Enviados")
        End Using
    End Sub

    ''' <summary>
    ''' <b>A busca atravessa pastas.</b>
    '''
    ''' E cada achado sabe de que pasta veio — sem isso o resultado seria uma
    ''' lista de assuntos sem lugar, e o usuário não teria como voltar à
    ''' mensagem.
    ''' </summary>
    <TestMethod>
    Public Sub Acha_em_mais_de_uma_pasta_e_diz_de_qual()
        Using db = Abrir()
            Semear(db, "Caixa de Entrada", "entrada", Caixa)
            Semear(db, "Itens Enviados", "enviados",
                   {("RES: Contrato assinado", "Eu mesmo")})

            Dim r = Buscar(db).Procurar("contrato")

            Assert.AreEqual(2, r.Achados.Count)
            CollectionAssert.AreEquivalent(
                {"Caixa de Entrada", "Itens Enviados"},
                r.Achados.Select(Function(a) a.NomeDaPasta).ToArray())
            Assert.AreEqual(2, r.Consultadas.Count)
        End Using
    End Sub

    ''' <summary>
    ''' <b>Termo vazio não devolve o acervo inteiro.</b>
    '''
    ''' Busca sem termo que lista tudo parece funcionalidade e é acidente: a
    ''' primeira abertura da tela despejaria mil linhas sem ninguém ter pedido.
    ''' </summary>
    <TestMethod>
    Public Sub Termo_vazio_nao_devolve_nada()
        Using db = Abrir()
            Semear(db, "Caixa de Entrada", "entrada", Caixa)
            Dim busca = Buscar(db)

            For Each vazio In {"", "   ", Nothing}
                Dim r = busca.Procurar(vazio)
                Assert.AreEqual(0, r.Achados.Count, $"'{vazio}' devolveu achados")
                StringAssert.Contains(r.Ressalva, "Digite alguma coisa")
                ' Mesmo vazio, ele diz ONDE procuraria.
                Assert.AreEqual(1, r.Consultadas.Count)
            Next
        End Using
    End Sub

    ''' <summary>
    ''' <b>Sem acervo nenhum, a busca diz isso — e não "não achei".</b>
    ''' </summary>
    <TestMethod>
    Public Sub Sem_pasta_varrida_a_ressalva_diz_que_nao_ha_onde_procurar()
        Using db = Abrir()
            SoRegistrar(db, "Caixa de Entrada", "entrada")

            Dim r = Buscar(db).Procurar("qualquer coisa")

            Assert.AreEqual(0, r.Achados.Count)
            Assert.AreEqual(0, r.Consultadas.Count)
            StringAssert.Contains(r.Ressalva, "Nenhuma pasta foi varrida")
        End Using
    End Sub

    ''' <summary>
    ''' <b>A BUSCA NÃO ENXERGA O QUE O DRENO AINDA NÃO ENTREGOU.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ESTE TESTE PROVAVA O CONTRÁRIO ATÉ 28/08/2026, À TARDE</b>
    '''
    ''' A versão anterior cobrava que a busca <i>avisasse</i> sobre a
    ''' publicação pendente — e passava porque a busca lia o
    ''' <c>ManifestReader</c> por conta própria. Ou seja: ela <b>via</b> a
    ''' geração nova e ao mesmo tempo dizia que a geração nova não tinha
    ''' chegado.
    '''
    ''' A revisão externa chamou isso pelo nome: o teste <b>cristalizava o
    ''' contorno</b> em vez de provar convergência. Consultar o estado do dreno
    ''' não é passar por ele.
    '''
    ''' Agora a busca lê o <see cref="AcervoDeTodasAsPastas"/>, que só muda
    ''' quando o dreno entrega. Então o que este teste prova é o oposto do que
    ''' provava: <b>publicou e não drenou ⇒ a busca não vê</b>, e o painel ao
    ''' lado também não. Ficar para trás junto é honesto; ficar na frente em
    ''' silêncio não era.
    ''' </summary>
    <TestMethod>
    Public Sub A_busca_NAO_ve_geracao_que_o_dreno_nao_entregou()
        Using db = Abrir()
            Dim chave = Semear(db, "Caixa de Entrada", "entrada", Caixa)

            ' A busca e montada AGORA, com o acervo ja drenado.
            Dim todas As New AcervoDeTodasAsPastas(db)
            Dim dreno As New PublicationDrain(db)
            dreno.Drenar(todas)
            If todas.Recarregado = 0 Then todas.Recarregar()
            If todas.Recarregado = 0 Then todas.Recarregar()
            Dim busca As New BuscaNoAcervo(todas, dreno)

            Assert.AreEqual(1, busca.Procurar("contrato").Achados.Count,
                "controle: a busca acha o que ja foi drenado")

            ' Uma segunda varredura publica UMA linha nova, e NINGUEM drena.
            Dim resolvedor As New ResolvedorDoAcervo(db)
            Dim amb = resolvedor.Ambiente(Impressao())
            Dim universo As New SweepUniverse("store-1", "entrada", "f", Nothing, 1, "amb-1")
            Dim fonte As New FonteDeLinhas(universo, {New SourceRow With {
                .Key = "entrada-9", .Subject = "Aditivo contratual novissimo",
                .SenderName = "Caroline Abreu",
                .ReceivedAt = New DateTimeOffset(2026, 8, 25, 9, 0, 0, TimeSpan.Zero).ToString("o"),
                .MessageClass = "IPM.Note"}})
            Dim r2 = New SweepRunner(fonte, New SqliteSweepSink(db, chave, amb.Chave), 50).
                     Executar(universo, 0, 2, EnvironmentPolicy.Capacidades(Impressao()),
                              CancellationToken.None)
            Assert.IsTrue(r2.Publicou, $"controle: a segunda varredura tinha de publicar. {r2.Motivo}")

            Dim r = busca.Procurar("novissimo")

            ' O CONSERTO, EM UMA LINHA.
            Assert.AreEqual(0, r.Achados.Count,
                "a busca enxergou uma geracao que o dreno ainda nao entregou -- " &
                "e o painel do acervo ao lado nao enxerga")

            ' E ELA AVISA -- COM A FRASE CERTA.
            '
            ' Esta assercao so cobrava a presenca de "painel do acervo", e por
            ' isso passou junto com uma ressalva que dizia o OPOSTO do estado:
            ' "a busca ja as enxerga; o painel pode estar atrasado", escrita
            ' quando a busca ainda contornava o dreno e nao revisada quando o
            ' contorno saiu. A revisao externa pegou.
            '
            ' Agora ela cobra o SENTIDO: que a ressalva diga que o retrato e o
            ' anterior, e que NAO afirme que a busca ja enxerga.
            '
            ' AS PROIBICOES VIERAM UMA POR VOLTA, E FORAM QUATRO.
            '
            ' A frase ja disse "nao foram entregues ao acervo" (falso: a
            ' publicacao materializa o acervo), "a busca ja as enxerga" (era
            ' verdade so enquanto a busca contornava o dreno), "na busca E no
            ' painel" (falso na entrega parcial) e "a busca mostra o retrato
            ' anterior" (falso com duas geracoes pendentes). Cada uma passou no
            ' teste da vez, porque o teste da vez so exercitava o estado em que
            ' ela era verdadeira.
            '
            ' O que ela pode afirmar e o estado da FILA, e e isso que se cobra
            ' aqui. Os dois estados que derrubaram as duas ultimas versoes tem
            ' teste proprio, logo abaixo.
            Assert.IsTrue(r.PublicacoesPendentes > 0,
                $"tinha de haver entrega pendente, achei {r.PublicacoesPendentes}")
            StringAssert.Contains(r.Ressalva, "retrato da última varredura")
            StringAssert.Contains(r.Ressalva, "painel do acervo")
            Assert.IsFalse(r.Ressalva.Contains("já as enxerga"), "volta 2")
            Assert.IsFalse(r.Ressalva.Contains("na busca e no painel"), "volta 3")
            Assert.IsFalse(r.Ressalva.Contains("a busca mostra é o retrato anterior"),
                "volta 4: a ressalva afirma o estado da BUSCA, e ela nao sabe " &
                "-- ver Com_DUAS_geracoes_pendentes_a_busca_ve_a_SEGUNDA_cedo")
            Assert.IsFalse(r.Ressalva.Contains("ainda não foram entregues"),
                "volta 5: Pendentes() e drained_at IS NULL, e nao 'nao recebeu'. " &
                "A entrega e ao menos uma vez, e o DrenoAposCrashTests cobre a " &
                "janela em que o consumidor recebeu e o disco ainda nao sabe")
        End Using
    End Sub

    ''' <summary>
    ''' <b>E depois de drenar, ela vê.</b>
    '''
    ''' O par do teste acima, e sem ele o outro passaria numa busca que
    ''' simplesmente nunca acha nada novo. É o mesmo controle positivo que o
    ''' resto deste arquivo tem, aplicado à convergência.
    ''' </summary>
    <TestMethod>
    Public Sub Depois_de_drenar_a_busca_VE()
        Using db = Abrir()
            Dim chave = Semear(db, "Caixa de Entrada", "entrada", Caixa)

            Dim todas As New AcervoDeTodasAsPastas(db)
            Dim dreno As New PublicationDrain(db)
            dreno.Drenar(todas)
            Dim busca As New BuscaNoAcervo(todas, dreno)

            Dim resolvedor As New ResolvedorDoAcervo(db)
            Dim amb = resolvedor.Ambiente(Impressao())
            Dim universo As New SweepUniverse("store-1", "entrada", "f", Nothing, 1, "amb-1")
            Dim fonte As New FonteDeLinhas(universo, {New SourceRow With {
                .Key = "entrada-9", .Subject = "Aditivo contratual novissimo",
                .SenderName = "Caroline Abreu",
                .ReceivedAt = New DateTimeOffset(2026, 8, 25, 9, 0, 0, TimeSpan.Zero).ToString("o"),
                .MessageClass = "IPM.Note"}})
            Dim r2 = New SweepRunner(fonte, New SqliteSweepSink(db, chave, amb.Chave), 50).
                     Executar(universo, 0, 2, EnvironmentPolicy.Capacidades(Impressao()),
                              CancellationToken.None)
            Assert.IsTrue(r2.Publicou, "controle: publicou")
            Assert.AreEqual(0, busca.Procurar("novissimo").Achados.Count, "controle: ainda nao drenou")

            dreno.Drenar(todas)

            Assert.AreEqual(1, busca.Procurar("novissimo").Achados.Count,
                "depois do dreno a busca tinha de enxergar a geracao nova")
        End Using
    End Sub

    ''' <summary>
    ''' <b>O consumidor composto entrega aos dois, e a falha de um trava.</b>
    '''
    ''' O <c>Drenar</c> entrega a UM consumidor, e agora há dois interessados.
    ''' Dois drenos seriam pior: cada um marcaria a geração como entregue por
    ''' conta própria, e o segundo nunca veria o que o primeiro drenou.
    '''
    ''' E se um falhar, a exceção tem de subir — a cabeça da fila trava de
    ''' propósito. Engolir a falha de um para agradar o outro marcaria a
    ''' geração como entregue a quem não a recebeu.
    ''' </summary>
    <TestMethod>
    Public Sub O_consumidor_composto_entrega_aos_dois_e_a_falha_SOBE()
        Dim a As New ContadorDeEntregas()
        Dim b As New ContadorDeEntregas()

        Dim composto As New ConsumidorComposto(a, b)
        composto.Receber(7)

        Assert.AreEqual(1, a.Recebidas, "o primeiro nao recebeu")
        Assert.AreEqual(1, b.Recebidas, "o segundo nao recebeu")

        b.Explodir = True
        Assert.ThrowsException(Of InvalidOperationException)(
            Sub() composto.Receber(8),
            "a falha de um consumidor foi engolida, e a geracao seria marcada " &
            "como entregue a quem nao a recebeu")
    End Sub

    ''' <summary>
    ''' <b>UMA ENTREGA QUE FALHA NO MEIO DEIXA O PRIMEIRO CONSUMIDOR À FRENTE.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTE TESTE PROVA, E O QUE ELE NÃO PROVA</b>
    '''
    ''' <b>Prova:</b> o <c>ConsumidorComposto</c> chama em sequência e sem
    ''' transação; a falha do segundo mantém a geração pendente <i>depois</i> de
    ''' o primeiro já ter recebido. É a dívida "o fan-out não é atômico", que
    ''' estava escrita no ESCOPO e não tinha cobertura nenhuma.
    '''
    ''' <b>Não prova:</b> que o painel de produção fique à frente da busca de
    ''' produção. Os dois membros aqui são contadores; o <c>AcervoService</c> e o
    ''' <c>AcervoDeTodasAsPastas</c> reais não entram. O teste anterior tinha o
    ''' nome <c>..._deixa_o_painel_a_FRENTE</c>, e a revisão externa apontou que
    ''' o nome prometia a integração e o corpo entregava a unidade. O nome mudou;
    ''' o corpo é o mesmo, e continua valendo pelo que é.
    '''
    ''' O estado de produção que de fato falsificou a ressalva tem teste próprio:
    ''' <c>Com_DUAS_geracoes_pendentes_a_busca_ve_a_SEGUNDA_cedo</c>, que usa o
    ''' acervo real.
    ''' </summary>
    <TestMethod>
    Public Sub Entrega_que_falha_no_meio_deixa_o_PRIMEIRO_consumidor_a_frente()
        Using db = Abrir()
            Dim chave = Semear(db, "Caixa de Entrada", "entrada", Caixa)

            Dim todas As New AcervoDeTodasAsPastas(db)
            Dim dreno As New PublicationDrain(db)
            dreno.Drenar(todas)
            If todas.Recarregado = 0 Then todas.Recarregar()
            If todas.Recarregado = 0 Then todas.Recarregar()
            Dim busca As New BuscaNoAcervo(todas, dreno)

            ' Uma varredura nova publica, e ninguem drena ainda.
            Dim resolvedor As New ResolvedorDoAcervo(db)
            Dim amb = resolvedor.Ambiente(Impressao())
            Dim universo As New SweepUniverse("store-1", "entrada", "f", Nothing, 1, "amb-1")
            Dim fonte As New FonteDeLinhas(universo, {New SourceRow With {
                .Key = "entrada-9", .Subject = "Aditivo contratual novissimo",
                .SenderName = "Caroline Abreu",
                .ReceivedAt = New DateTimeOffset(2026, 8, 25, 9, 0, 0, TimeSpan.Zero).ToString("o"),
                .MessageClass = "IPM.Note"}})
            Dim r2 = New SweepRunner(fonte, New SqliteSweepSink(db, chave, amb.Chave), 50).
                     Executar(universo, 0, 2, EnvironmentPolicy.Capacidades(Impressao()),
                              CancellationToken.None)
            Assert.IsTrue(r2.Publicou, $"controle: a varredura tinha de publicar. {r2.Motivo}")

            ' A ENTREGA PARCIAL: o primeiro recebe, o segundo falha.
            Dim painel As New ContadorDeEntregas()
            Dim quebrada As New ContadorDeEntregas() With {.Explodir = True}
            Assert.ThrowsException(Of InvalidOperationException)(
                Sub() dreno.Drenar(New ConsumidorComposto(painel, quebrada)))

            ' O ESTADO, medido e nao suposto.
            Assert.AreEqual(1, painel.Recebidas,
                "controle: o primeiro consumidor tinha de ter recebido")
            Assert.IsTrue(dreno.Pendentes().Count > 0,
                "controle: a geracao tinha de continuar pendente")

            Dim r = busca.Procurar("novissimo")
            Assert.AreEqual(0, r.Achados.Count, "a busca enxergou o que nao lhe foi entregue")
            Assert.IsTrue(r.PublicacoesPendentes > 0)

            ' O CONSERTO: a ressalva fala do estado da FILA, e nao do estado
            ' de nenhum dos dois lados.
            StringAssert.Contains(r.Ressalva, "retrato da última varredura")
            Assert.IsFalse(r.Ressalva.Contains("na busca e no painel"),
                "a ressalva poe os dois no mesmo retrato, e neste exato estado " &
                "o primeiro consumidor esta uma geracao a frente")
            Assert.IsFalse(r.Ressalva.Contains("nem a busca nem o painel"),
                "a ressalva afirma que nenhum dos dois enxerga, e um enxerga")
        End Using
    End Sub

    ''' <summary>
    ''' <b>COM DUAS GERAÇÕES PENDENTES, A BUSCA VÊ A SEGUNDA — antes de ela ser
    ''' entregue.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ESTE É O ESTADO QUE DERRUBOU A QUARTA VERSÃO DA RESSALVA</b>
    '''
    ''' <c>AcervoDeTodasAsPastas.Receber</c> ignora <i>qual</i> geração chegou e
    ''' relê o manifesto <b>corrente</b>. Com 10 e 11 pendentes e o manifesto já
    ''' apontando para 11, entregar a 10 faz a busca recarregar a 11. Se a
    ''' entrega da 11 falhar, a busca está enxergando exatamente a geração que a
    ''' ressalva jurava que ela não via.
    '''
    ''' A dívida estava escrita no <c>AcervoDeTodasAsPastas</c> desde a manhã,
    ''' com a observação de que "na prática a janela é curta". <b>Curta não é
    ''' inexistente</b> — e foi a revisão externa que ligou a dívida à frase.
    '''
    ''' Aqui o acervo é o <b>real</b>, e não um contador: o que se mede é a busca
    ''' de produção achando o que ninguém lhe entregou.
    '''
    ''' <b>Dois controles.</b> O de <i>causalidade</i> está no corpo: antes de
    ''' qualquer entrega a busca não acha nada, então o que a torna visível é a
    ''' entrega da <b>primeira</b> geração — sem essa linha o teste provaria o
    ''' estado final e não a causa, e a revisão externa cobrou isso. O
    ''' <i>negativo</i>: devolvendo <i>"O que a busca mostra é o retrato anterior
    ''' a elas"</i> à ressalva, a asserção do final cai.
    ''' </summary>
    <TestMethod>
    Public Sub Com_DUAS_geracoes_pendentes_a_busca_ve_a_SEGUNDA_cedo()
        Using db = Abrir()
            Dim chave = Semear(db, "Caixa de Entrada", "entrada", Caixa)

            Dim todas As New AcervoDeTodasAsPastas(db)
            Dim dreno As New PublicationDrain(db)
            dreno.Drenar(todas)
            If todas.Recarregado = 0 Then todas.Recarregar()
            If todas.Recarregado = 0 Then todas.Recarregar()
            Dim busca As New BuscaNoAcervo(todas, dreno)

            ' DUAS varreduras publicam, e ninguem drena entre elas.
            Varrer(db, chave, "entrada-9", "Aditivo contratual novissimo")
            Varrer(db, chave, "entrada-10", "Aditivo contratual rarissimo")
            Assert.AreEqual(2, dreno.Pendentes().Count,
                "controle: as duas geracoes tinham de estar pendentes")

            ' CONTROLE DE CAUSALIDADE, e ele foi cobrado pela revisao externa.
            '
            ' Sem esta linha o teste prova o ESTADO FINAL e nao a CAUSA: se um
            ' dia a publicacao passar a atualizar o acervo por outro caminho, a
            ' busca acharia "rarissimo" sem entrega nenhuma e o teste ficaria
            ' verde afirmando uma coisa que deixou de ser verdade.
            Assert.AreEqual(0, busca.Procurar("rarissimo").Achados.Count,
                "controle: antes de QUALQUER entrega a busca nao pode ver nada")

            ' So a PRIMEIRA e entregue; a segunda falha e continua pendente.
            Assert.ThrowsException(Of InvalidOperationException)(
                Sub() dreno.Drenar(New EntregaSoAPrimeira(todas)))
            Assert.AreEqual(1, dreno.Pendentes().Count,
                "controle: a segunda geracao tinha de continuar pendente")

            ' A DIVIDA, MEDIDA: a busca ja enxerga a geracao pendente.
            Assert.AreEqual(1, busca.Procurar("rarissimo").Achados.Count,
                "controle desta divida: se a busca NAO ve a segunda geracao, " &
                "entao o Receber deixou de reler o manifesto corrente -- e a " &
                "ressalva pode voltar a afirmar pela busca")

            ' ENTAO A RESSALVA NAO PODE DIZER QUE A BUSCA ESTA ATRAS.
            Dim r = busca.Procurar("rarissimo")
            Assert.IsTrue(r.PublicacoesPendentes > 0)
            StringAssert.Contains(r.Ressalva, "retrato da última varredura")
            Assert.IsFalse(r.Ressalva.Contains("a busca mostra é o retrato anterior"),
                "a ressalva afirma que a busca esta atras, e ela acabou de " &
                "achar a geracao pendente")
            Assert.IsFalse(r.Ressalva.Contains("esta busca não está enxergando"),
                "a ressalva afirma que a busca nao enxerga, e ela enxerga")
        End Using
    End Sub

    ''' <summary>
    ''' Uma varredura de uma linha só, publicando sem drenar.
    ''' </summary>
    Private Sub Varrer(db As CacheDatabase, chave As Long, id As String, assunto As String)
        Dim resolvedor As New ResolvedorDoAcervo(db)
        Dim amb = resolvedor.Ambiente(Impressao())
        Dim universo As New SweepUniverse("store-1", "entrada", "f", Nothing, 1, "amb-1")
        Dim fonte As New FonteDeLinhas(universo, {New SourceRow With {
            .Key = id, .Subject = assunto, .SenderName = "Caroline Abreu",
            .ReceivedAt = New DateTimeOffset(2026, 8, 25, 9, 0, 0, TimeSpan.Zero).ToString("o"),
            .MessageClass = "IPM.Note"}})
        Dim r = New SweepRunner(fonte, New SqliteSweepSink(db, chave, amb.Chave), 50).
                Executar(universo, 0, 2, EnvironmentPolicy.Capacidades(Impressao()),
                         CancellationToken.None)
        Assert.IsTrue(r.Publicou, $"controle: a varredura tinha de publicar. {r.Motivo}")
    End Sub

    ''' <summary>
    ''' Entrega a primeira geração ao alvo REAL e falha na segunda — que é a
    ''' sequência de uma queda no meio do laço do dreno.
    ''' </summary>
    Private NotInheritable Class EntregaSoAPrimeira
        Implements IPublicationConsumer

        Private ReadOnly _alvo As IPublicationConsumer
        Private _quantas As Integer

        Public Sub New(alvo As IPublicationConsumer)
            _alvo = alvo
        End Sub

        Public Sub Receber(geracao As Long) Implements IPublicationConsumer.Receber
            _quantas += 1
            If _quantas > 1 Then Throw New InvalidOperationException("entrega falhou na segunda")
            _alvo.Receber(geracao)
        End Sub
    End Class

    ''' <summary>Consumidor de teste que conta e, se pedirem, explode.</summary>
    Private NotInheritable Class ContadorDeEntregas
        Implements IPublicationConsumer

        Friend Recebidas As Integer
        Friend Property Explodir As Boolean

        Public Sub Receber(geracao As Long) Implements IPublicationConsumer.Receber
            If Explodir Then Throw New InvalidOperationException("consumidor falhou")
            Recebidas += 1
        End Sub
    End Class

    ''' <summary>
    ''' <b>REABRIR COM PUBLICAÇÃO PENDENTE — o caso que o dreno existe para
    ''' recuperar, e onde a primeira versão contornava.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE A REVISÃO EXTERNA ACHOU</b>
    '''
    ''' O <c>AcervoDeTodasAsPastas</c> lia o manifesto <b>no construtor</b>, com
    ''' a justificativa de que "na abertura não há entrega pendente". É falso:
    ''' uma queda entre publicar e marcar drenada deixa publicação pendente
    ''' <b>persistida no banco</b>. Na abertura seguinte, o construtor leria a
    ''' geração nova antes de o dreno entregá-la — o contorno de novo, reduzido
    ''' à abertura.
    '''
    ''' E os testes passavam pelo motivo errado: o <c>Semear</c> já drenava, e o
    ''' <c>Drenar</c> seguinte não tinha nada a entregar.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ESTE TESTE SIMULA A QUEDA</b>
    '''
    ''' Publica sem drenar, <b>fecha o banco</b>, reabre, e constrói o acervo do
    ''' zero — que é exatamente a sequência de uma reabertura depois de queda.
    '''
    ''' <b>Controle negativo:</b> devolvendo o <c>Recarregar()</c> ao construtor,
    ''' este teste falha — a busca enxerga a geração que ninguém entregou.
    ''' </summary>
    <TestMethod>
    Public Sub Reabrir_com_publicacao_pendente_NAO_mostra_a_geracao_nova()
        Dim chave As Long
        Using db = Abrir()
            chave = Semear(db, "Caixa de Entrada", "entrada", Caixa)

            ' Publica UMA linha nova e NINGUEM drena -- e ai o processo "morre".
            Dim resolvedor As New ResolvedorDoAcervo(db)
            Dim amb = resolvedor.Ambiente(Impressao())
            Dim universo As New SweepUniverse("store-1", "entrada", "f", Nothing, 1, "amb-1")
            Dim fonte As New FonteDeLinhas(universo, {New SourceRow With {
                .Key = "entrada-9", .Subject = "Aditivo pos-queda",
                .SenderName = "Caroline Abreu",
                .ReceivedAt = New DateTimeOffset(2026, 8, 25, 9, 0, 0, TimeSpan.Zero).ToString("o"),
                .MessageClass = "IPM.Note"}})
            Dim r2 = New SweepRunner(fonte, New SqliteSweepSink(db, chave, amb.Chave), 50).
                     Executar(universo, 0, 2, EnvironmentPolicy.Capacidades(Impressao()),
                              CancellationToken.None)
            Assert.IsTrue(r2.Publicou, $"controle: tinha de publicar. {r2.Motivo}")
        End Using

        ' O BANCO E REABERTO, e o acervo nasce do zero.
        SqliteConnection.ClearAllPools()
        Using db = Abrir()
            Dim dreno As New PublicationDrain(db)
            Assert.IsTrue(dreno.Pendentes().Count > 0,
                "controle: a publicacao pendente tinha de ter sobrevivido ao fechamento")

            Dim todas As New AcervoDeTodasAsPastas(db)
            Assert.AreEqual(0, todas.Recarregado,
                "o construtor leu o manifesto -- e nesta situacao isso e ler na " &
                "frente do dreno")

            Dim busca As New BuscaNoAcervo(todas, dreno)
            Assert.AreEqual(0, busca.Procurar("pos-queda").Achados.Count,
                "a busca enxergou uma geracao que ninguem entregou, logo apos reabrir")

            ' E depois de drenar, ela ve -- o par positivo.
            dreno.Drenar(todas)
            Assert.AreEqual(1, busca.Procurar("pos-queda").Achados.Count,
                "depois do dreno a geracao pendente tinha de aparecer")
        End Using
    End Sub

End Class

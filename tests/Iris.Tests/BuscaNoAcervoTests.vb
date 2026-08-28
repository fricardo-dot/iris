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

            Dim r = New BuscaNoAcervo(db).Procurar("contrato")

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
            Dim busca As New BuscaNoAcervo(db)

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
            Dim busca As New BuscaNoAcervo(db)

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

            Dim r = New BuscaNoAcervo(db).Procurar("palavraquenaoexiste")

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
    ''' <b>Pasta nunca varrida não é pasta sem resultado.</b>
    '''
    ''' Misturá-las faria o resultado dizer "procurei aqui e não achei" sobre
    ''' um lugar onde ninguém procurou. É a mesma distinção entre
    ''' <c>Nothing</c> e zero que o projeto já faz nas linhas descartadas.
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_nunca_varrida_fica_SEPARADA_das_consultadas()
        Using db = Abrir()
            Semear(db, "Caixa de Entrada", "entrada", Caixa)
            SoRegistrar(db, "Itens Enviados", "enviados")

            Dim r = New BuscaNoAcervo(db).Procurar("contrato")

            Assert.AreEqual(1, r.Consultadas.Count, "so a varrida foi consultada")
            Assert.AreEqual("Caixa de Entrada", r.Consultadas(0).Nome)
            Assert.AreEqual(1, r.SemAcervo.Count, "a nao varrida tem de aparecer, e a parte")
            Assert.AreEqual("Itens Enviados", r.SemAcervo(0).Nome)
            StringAssert.Contains(r.Ressalva, "nunca foram varridas")
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

            Dim r = New BuscaNoAcervo(db).Procurar("contrato")

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
            Dim busca As New BuscaNoAcervo(db)

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

            Dim r = New BuscaNoAcervo(db).Procurar("qualquer coisa")

            Assert.AreEqual(0, r.Achados.Count)
            Assert.AreEqual(0, r.Consultadas.Count)
            StringAssert.Contains(r.Ressalva, "Nenhuma pasta foi varrida")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Publicação não entregue aparece na ressalva.</b>
    '''
    ''' É a §26.2 na forma que a busca comporta. O <c>AcervoService</c> é de
    ''' uma pasta só, então uma busca entre pastas não cabe nele; o que ela
    ''' pode fazer é contar a fila e dizer. Sem isso, o dreno travado sumiria
    ''' atrás de uma lista que parece completa.
    ''' </summary>
    <TestMethod>
    Public Sub Publicacao_nao_entregue_aparece_na_ressalva()
        Using db = Abrir()
            Dim chave = Semear(db, "Caixa de Entrada", "entrada", Caixa)

            ' Uma segunda varredura publica de novo e NINGUEM drena.
            Dim resolvedor As New ResolvedorDoAcervo(db)
            Dim amb = resolvedor.Ambiente(Impressao())
            Dim universo As New SweepUniverse("store-1", "entrada", "f", Nothing, 1, "amb-1")
            Dim fonte As New FonteDeLinhas(universo, {New SourceRow With {
                .Key = "entrada-0", .Subject = "Contrato assinado",
                .SenderName = "Caroline Abreu",
                .ReceivedAt = New DateTimeOffset(2026, 8, 20, 9, 0, 0, TimeSpan.Zero).ToString("o"),
                .MessageClass = "IPM.Note"}})
            Dim r2 = New SweepRunner(fonte, New SqliteSweepSink(db, chave, amb.Chave), 50).
                     Executar(universo, 0, 2, EnvironmentPolicy.Capacidades(Impressao()),
                              CancellationToken.None)
            Assert.IsTrue(r2.Publicou, $"controle: a segunda varredura tinha de publicar. {r2.Motivo}")

            Dim r = New BuscaNoAcervo(db).Procurar("contrato")

            Assert.IsTrue(r.PublicacoesPendentes > 0,
                $"controle: tinha de haver publicacao pendente, achei {r.PublicacoesPendentes}")
            StringAssert.Contains(r.Ressalva, "não foram entregues")
        End Using
    End Sub

End Class

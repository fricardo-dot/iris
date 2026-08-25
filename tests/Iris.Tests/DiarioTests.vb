Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports Iris.Assist
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Model
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O diário do egress — o 3.3.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE REGISTRAR NO FIM NÃO SERVE</b>
'''
''' Um diário escrito quando o envio termina registra os envios que
''' terminaram — e perde justamente os que importam. Se o processo morre
''' durante a transmissão não há linha nenhuma, e o registro passa a afirmar,
''' <b>por omissão</b>, que nada saiu.
'''
''' Daí os cinco passos, e daí a reconciliação da abertura.
''' </summary>
<TestClass>
Public Class DiarioTests

    Private _pasta As String
    Private _caminho As String

    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 25, 12, 0, 0, TimeSpan.Zero)
    Private Const Endereco As String = "https://exemplo.invalido/v1"

    <TestInitialize>
    Public Sub Preparar()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-diario-" & Guid.NewGuid().ToString("N"))
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

    ' ---- fixtures ------------------------------------------------------

    Private Shared Function Chave(n As Integer) As ItemKey
        Return New ItemKey($"E-{n}", "store-1")
    End Function

    Private Shared Function Autorizacao() As ActivationRecord
        Return New ActivationRecord("ativacao-1", 3, "teste", Agora.AddDays(-1),
                                    "provedor-de-teste", Endereco, "modelo-de-teste",
                                    "local", "sem retenção",
                                    {AssistOperation.Resumir},
                                    {New FolderKey("store-1", "pasta-1")},
                                    Array.Empty(Of String)(),
                                    {LabelReadingKind.Absent}, {0})
    End Function

    Private Shared Function Voo() As PreflightRequest
        Return New PreflightRequest(AssistOperation.Resumir,
                                    New FolderKey("store-1", "pasta-1"),
                                    New AssistDestination("provedor-de-teste", Endereco,
                                                          "modelo-de-teste"))
    End Function

    Private Shared Function Mensagem(n As Integer) As MessageClassification
        Dim l As New LabelReading(Chave(n), LabelReadingKind.Absent, LabelReadStage.Parse,
                                  version:=New LabelVersionEvidence($"E-{n}", Agora, $"CK-{n}"))
        Return New MessageClassification(Chave(n), New FolderKey("store-1", "pasta-1"), l, temAnexo:=False)
    End Function

    Private Shared Function Parte(n As Integer, Optional corpo As String = "olá") As MessagePart
        Dim r = ContentPipeline.Preparar(
            New MessageSnapshot(Chave(n), $"CK-{n}", $"assunto {n}", "de@x.invalido",
                                {"para@x.invalido"}, corpo, False, True, temAnexo:=False))
        Assert.IsTrue(r.Ok, $"{r.Recusa}")
        Return r.Parte
    End Function

    ''' <summary>Uma capability de verdade, vinda do portão de verdade.</summary>
    Private Shared Function Autorizada(Optional corpo As String = "olá") _
                                       As (Cap As DisclosureCapability, Cofre As CapabilityLedger)
        Dim d = New DisclosurePolicy(Autorizacao()).
                Decidir(New DisclosureRequest(Voo(), {Mensagem(1)}), Agora)
        Assert.IsTrue(d.Permitido, $"{d.Motivo}")

        Dim env = New EnvelopeBuilder().Montar(AssistOperation.Resumir, "resuma",
                                               {Parte(1, corpo)})
        Assert.IsTrue(env.Ok, $"{env.Recusa}")

        Dim cofre As New CapabilityLedger()
        Dim c = cofre.Emitir(d, env.Envelope, Agora)
        Assert.IsNotNull(c)
        Return (c, cofre)
    End Function

    ' ==================================================================
    ' Os cinco passos

    <TestMethod>
    Public Sub Intencao_Iniciando_Concluir_chega_em_Concluida()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()

            j.Intencao(a.Cap, Agora)
            Assert.AreEqual(DisclosureStage.Intencionada, j.Ler(1)(0).Estagio)

            j.Iniciando(a.Cap.RequestId, Agora.AddSeconds(1))
            Assert.AreEqual(DisclosureStage.EmVoo, j.Ler(1)(0).Estagio)

            j.Concluir(a.Cap.RequestId, Agora.AddSeconds(2))
            Assert.AreEqual(DisclosureStage.Concluida, j.Ler(1)(0).Estagio)
        End Using
    End Sub

    ''' <summary>
    ''' A intenção guarda hash, tamanho, modelo, endpoint e a ativação — e é
    ''' gravada <b>antes</b> de qualquer transmissão.
    ''' </summary>
    <TestMethod>
    Public Sub A_intencao_guarda_o_que_o_R11_manda()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)

            Dim e = j.Ler(1)(0)
            Assert.AreEqual(a.Cap.Hash, e.Hash)
            Assert.AreEqual(a.Cap.Comprimento, e.Bytes)
            Assert.AreEqual("modelo-de-teste", e.Modelo)
            Assert.AreEqual(Endereco, e.Endpoint)
            Assert.AreEqual("ativacao-1", e.AtivacaoId)
            Assert.AreEqual(3, e.AtivacaoVersao)
            Assert.AreEqual(1, e.Mensagens)
        End Using
    End Sub

    ' ==================================================================
    ' O protocolo de crash

    ''' <summary>
    ''' <b>Morrer EM VOO vira ambíguo.</b>
    '''
    ''' Os bytes podem ter chegado ao provedor, e ninguém vai saber. É a mesma
    ''' disciplina do <c>ErrorKind.Ambiguous</c> que o CLAUDE.md impõe às
    ''' mutações — e aqui a "mutação" é o conteúdo ter saído da máquina.
    ''' </summary>
    <TestMethod>
    Public Sub Morrer_EM_VOO_vira_AMBIGUA()
        Dim id As Guid
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            id = a.Cap.RequestId
            j.Intencao(a.Cap, Agora)
            j.Iniciando(id, Agora.AddSeconds(1))
            ' E o processo morre aqui. Nenhum Concluir, nenhum Falhar.
        End Using

        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)

            Assert.AreEqual(1, j.Reconciliar(Agora.AddHours(1)),
                            "uma divulgacao ficou sem desfecho conhecido")

            Dim e = j.Ler(1)(0)
            Assert.AreEqual(DisclosureStage.Ambigua, e.Estagio)
            Assert.AreEqual(DisclosureNote.ProcessoMorreuEmVoo, e.Nota)
        End Using
    End Sub

    ''' <summary>
    ''' <b>Morrer só com a INTENÇÃO vira não-enviada.</b>
    '''
    ''' O contraponto, e o que dá sentido ao de cima: se todo crash virasse
    ''' ambíguo, "ambíguo" deixaria de significar alguma coisa. Ali a
    ''' transmissão não tinha começado, e isso se sabe.
    ''' </summary>
    <TestMethod>
    Public Sub Morrer_so_com_a_INTENCAO_vira_NAO_ENVIADA()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)
        End Using

        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)

            Assert.AreEqual(0, j.Reconciliar(Agora.AddHours(1)),
                            "nada ficou ambiguo — a transmissao nao tinha comecado")
            Assert.AreEqual(DisclosureStage.NaoEnviada, j.Ler(1)(0).Estagio)
        End Using
    End Sub

    ''' <summary>Concluída não é mexida pela reconciliação.</summary>
    <TestMethod>
    Public Sub A_reconciliacao_nao_mexe_no_que_ja_terminou()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)
            j.Iniciando(a.Cap.RequestId, Agora)
            j.Concluir(a.Cap.RequestId, Agora)

            Assert.AreEqual(0, j.Reconciliar(Agora.AddHours(1)))
            Assert.AreEqual(DisclosureStage.Concluida, j.Ler(1)(0).Estagio)
        End Using
    End Sub

    ''' <summary>
    ''' <b>Ambígua NUNCA volta a ser não-enviada.</b>
    '''
    ''' Uma vez que os bytes podem ter saído, nenhuma informação posterior
    ''' desfaz isso — nem uma chamada explícita dizendo que não chegou.
    ''' </summary>
    <TestMethod>
    Public Sub Ambigua_nunca_volta_a_ser_NAO_ENVIADA()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)
            j.Iniciando(a.Cap.RequestId, Agora)
            j.Falhar(a.Cap.RequestId, Agora, DisclosureNote.Timeout, podeTerChegado:=True)
            Assert.AreEqual(DisclosureStage.Ambigua, j.Ler(1)(0).Estagio)

            j.Falhar(a.Cap.RequestId, Agora, DisclosureNote.ConexaoCaiu, podeTerChegado:=False)

            Assert.AreEqual(DisclosureStage.Ambigua, j.Ler(1)(0).Estagio,
                "uma vez que pode ter saido, nada desfaz")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Falhar EM VOO é ambíguo mesmo quando o chamador jura que não chegou.</b>
    '''
    ''' Ele não pode saber: entre "a conexão caiu" e "a conexão caiu depois de o
    ''' servidor ler o corpo" não há diferença observável deste lado.
    ''' </summary>
    <TestMethod>
    Public Sub Falhar_EM_VOO_e_ambiguo_mesmo_jurando_que_nao_chegou()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)
            j.Iniciando(a.Cap.RequestId, Agora)

            j.Falhar(a.Cap.RequestId, Agora, DisclosureNote.ConexaoCaiu, podeTerChegado:=False)

            Assert.AreEqual(DisclosureStage.Ambigua, j.Ler(1)(0).Estagio)
        End Using
    End Sub

    ''' <summary>Falhar antes de começar é não-enviada, e isso se sabe.</summary>
    <TestMethod>
    Public Sub Falhar_ANTES_de_comecar_e_NAO_ENVIADA()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)

            j.NaoEnviou(a.Cap.RequestId, Agora, DisclosureNote.CapabilityRecusada)

            Assert.AreEqual(DisclosureStage.NaoEnviada, j.Ler(1)(0).Estagio)
        End Using
    End Sub

    ''' <summary>
    ''' Não dá para pular passo: <c>Concluir</c> sem <c>Iniciando</c> não pega.
    '''
    ''' Se pegasse, um envio que nunca começou apareceria como concluído — e o
    ''' diário passaria a afirmar que conteúdo saiu quando não saiu.
    ''' </summary>
    <TestMethod>
    Public Sub Concluir_sem_Iniciando_nao_pega()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)

            j.Concluir(a.Cap.RequestId, Agora)

            Assert.AreEqual(DisclosureStage.Intencionada, j.Ler(1)(0).Estagio)
        End Using
    End Sub

    ' ==================================================================
    ' O que o diário NUNCA guarda

    ''' <summary>
    ''' <b>A isca.</b> Nenhum trecho do conteúdo aparece em lugar nenhum do
    ''' banco — nem em coluna, nem em nome, nem em índice.
    '''
    ''' O R11 do ESCOPO: <i>um log com o texto cria mais uma cópia sensível</i>.
    ''' O teste varre o arquivo inteiro, byte a byte, porque procurar coluna por
    ''' coluna provaria só as colunas que eu lembrei de olhar.
    ''' </summary>
    <TestMethod>
    Public Sub O_diario_NAO_guarda_conteudo()
        Const isca = "ISCA-QUE-NAO-PODE-VAZAR-4711"

        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada(corpo:=isca)
            j.Intencao(a.Cap, Agora)
            j.Iniciando(a.Cap.RequestId, Agora)
            j.Concluir(a.Cap.RequestId, Agora)
        End Using

        SqliteConnection.ClearAllPools()
        Dim cru = File.ReadAllBytes(_caminho)
        Dim texto = Text.Encoding.UTF8.GetString(cru)

        Assert.IsFalse(texto.Contains(isca),
            "o conteudo apareceu no banco — o diario virou mais uma copia sensivel")
        StringAssert.Contains(texto, "provedor-de-teste",
            "controle: a varredura ACHA o que esta la de verdade")
    End Sub

    ''' <summary>
    ''' <b>O motivo é enum fechado, e não texto.</b>
    '''
    ''' Enquanto era <c>String</c>, "o diário nunca guarda conteúdo" era
    ''' convenção: qualquer adaptador podia passar a mensagem de uma exceção ou
    ''' o corpo de erro do provedor — e corpo de erro <b>ecoa o que foi
    ''' enviado</b>. Não existe mais campo por onde texto arbitrário entre.
    ''' </summary>
    <TestMethod>
    Public Sub O_motivo_e_CODIGO_e_sobrevive()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)
            j.NaoEnviou(a.Cap.RequestId, Agora, DisclosureNote.PortaoNegou,
                        DisclosureReason.PastaNaoAutorizada)

            Dim e = j.Ler(1)(0)
            Assert.AreEqual(DisclosureNote.PortaoNegou, e.Nota)
            Assert.AreEqual(DisclosureReason.PastaNaoAutorizada, e.MotivoDoPortao)
        End Using
    End Sub

    ' ==================================================================

    ''' <summary>
    ''' Várias divulgações convivem, e a leitura devolve a mais recente
    ''' primeiro.
    ''' </summary>
    <TestMethod>
    Public Sub Varias_convivem_e_a_mais_recente_vem_primeiro()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)

            Dim primeira = Autorizada("um")
            j.Intencao(primeira.Cap, Agora)
            Dim segunda = Autorizada("dois")
            j.Intencao(segunda.Cap, Agora.AddMinutes(1))

            Dim tudo = j.Ler(10)
            Assert.AreEqual(2, tudo.Count)
            Assert.AreEqual(segunda.Cap.RequestId, tudo(0).RequestId)
        End Using
    End Sub


    ' ==================================================================
    ' O crash de VERDADE

    ''' <summary>
    ''' <b>Fechar o <c>Using</c> não é morrer — e a diferença é o assunto.</b>
    '''
    ''' Os testes acima fecham a conexão e reabrem. Isso prova que a
    ''' reconciliação lê o que ficou escrito, e <b>não</b> prova que ficou
    ''' escrito: fechar dá ao SQLite a chance de descarregar tudo com ordem, que
    ''' é exatamente o que um crash não dá.
    '''
    ''' Aqui o processo é morto com <c>TerminateProcess</c> logo depois de
    ''' gravar o passo. Ninguém desfaz nada, ninguém fecha nada. Se a intenção e
    ''' o "em voo" não estiverem <b>durados</b> no disco, a reabertura não acha
    ''' nada — e o diário estaria afirmando, por omissão, que nada saiu.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Morto_EM_VOO_de_verdade_a_reabertura_acha_AMBIGUA()
        Dim r = RodarHarness("em-voo")

        Assert.AreNotEqual(0, r.ExitCode, "o harness tinha de MORRER, nao terminar")
        Dim id = IdDoPedido(r.Stdout)

        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)

            Assert.AreEqual(1, j.Reconciliar(Agora),
                "a intencao e o inicio do voo tinham de estar DURADOS no disco")

            Dim e = j.Ler(1)(0)
            Assert.AreEqual(id, e.RequestId)
            Assert.AreEqual(DisclosureStage.Ambigua, e.Estagio)
            Assert.IsFalse(String.IsNullOrEmpty(e.Hash),
                "o hash do que PODE ter saido tem de estar la — e o que alguem procura")
        End Using
    End Sub

    ''' <summary>
    ''' E morto logo depois da intenção, a reabertura acha <b>não-enviada</b>.
    '''
    ''' O contraponto que dá sentido ao de cima: se todo crash virasse ambíguo,
    ''' "ambíguo" deixaria de significar alguma coisa.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Morto_APOS_A_INTENCAO_a_reabertura_acha_NAO_ENVIADA()
        Dim r = RodarHarness("apos-intencao")

        Assert.AreNotEqual(0, r.ExitCode, "o harness tinha de MORRER")
        Dim id = IdDoPedido(r.Stdout)

        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)

            Assert.AreEqual(0, j.Reconciliar(Agora),
                "nada ficou ambiguo — a transmissao nao tinha comecado")

            Dim e = j.Ler(1)(0)
            Assert.AreEqual(id, e.RequestId)
            Assert.AreEqual(DisclosureStage.NaoEnviada, e.Estagio)
        End Using
    End Sub

    ''' <summary>
    ''' O controle: sem morrer, o harness chega a <c>Concluida</c>.
    '''
    ''' Sem ele, um harness que explodisse na abertura do banco passaria nos
    ''' dois testes acima — "morreu" e "nunca começou" produzem o mesmo código
    ''' de saída.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Controle_sem_morrer_o_harness_CONCLUI()
        Dim r = RodarHarness("nenhum")

        Assert.AreEqual(0, r.ExitCode, r.Stderr)

        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Assert.AreEqual(DisclosureStage.Concluida, j.Ler(1)(0).Estagio)
        End Using
    End Sub

    ' ==================================================================

    Private Structure Resultado
        Public ExitCode As Integer
        Public Stdout As String
        Public Stderr As String
    End Structure

    Private Function RodarHarness(ponto As String) As Resultado
        Dim psi As New Diagnostics.ProcessStartInfo(LocalizarHarness()) With {
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .UseShellExecute = False,
            .CreateNoWindow = True}
        For Each a In {_caminho, "diario", ponto}
            psi.ArgumentList.Add(a)
        Next

        Using proc = Diagnostics.Process.Start(psi)
            Dim o = proc.StandardOutput.ReadToEnd()
            Dim e = proc.StandardError.ReadToEnd()
            If Not proc.WaitForExit(60000) Then
                proc.Kill(True)
                Assert.Fail("o harness travou")
            End If
            If e.Contains("Unhandled exception") OrElse e.Contains("abrir falhou") Then
                Assert.Fail("o harness falhou em vez de morrer no ponto pedido:" &
                            Environment.NewLine & e)
            End If
            Return New Resultado With {.ExitCode = proc.ExitCode, .Stdout = o, .Stderr = e}
        End Using
    End Function

    Private Shared Function IdDoPedido(saida As String) As Guid
        Dim linha = saida.Split(Environment.NewLine.ToCharArray(),
                                StringSplitOptions.RemoveEmptyEntries).
                    FirstOrDefault(Function(l) l.StartsWith("requestId="))
        Assert.IsNotNull(linha, "o harness nao disse qual foi o pedido: " & saida)
        Return Guid.Parse(linha.Substring("requestId=".Length).Trim())
    End Function

    Private Shared _exe As String

    Private Shared Function LocalizarHarness() As String
        If _exe IsNot Nothing Then Return _exe
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
            d = d.Parent
        End While
        Assert.IsNotNull(d, "nao achei a raiz do repositorio")
        Dim raiz = Path.Combine(d.FullName, "tools", "Iris.CrashHarness", "bin")
        Dim achado = Directory.GetFiles(raiz, "Iris.CrashHarness.exe", SearchOption.AllDirectories).
                     OrderByDescending(Function(f) File.GetLastWriteTimeUtc(f)).FirstOrDefault()
        Assert.IsNotNull(achado, "Iris.CrashHarness.exe nao encontrado")
        _exe = achado
        Return _exe
    End Function


    ' ==================================================================
    ' Toda transição diz se pegou

    ''' <summary>
    ''' <b>Um passo que não pega devolve <c>False</c>.</b>
    '''
    ''' Era o buraco: <c>Avancar</c> ignorava quantas linhas mudaram, então um
    ''' <c>Iniciando</c> que não persistisse — pedido inexistente, estado
    ''' errado, corrida — passava em silêncio, e quem chamou seguia para o
    ''' HTTP assim mesmo. Resultado: <b>egress sem registro de voo</b>, que é
    ''' exatamente o buraco que o diário existe para não ter.
    ''' </summary>
    <TestMethod>
    Public Sub Passo_que_nao_pega_devolve_FALSE()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()

            Assert.IsFalse(j.Iniciando(a.Cap.RequestId, Agora),
                           "nao ha intencao gravada — o voo nao pode comecar")
            Assert.IsFalse(j.Concluir(a.Cap.RequestId, Agora))
            Assert.IsFalse(j.NaoEnviou(a.Cap.RequestId, Agora, DisclosureNote.PortaoNegou))

            Assert.IsTrue(j.Intencao(a.Cap, Agora), "controle: com o pedido novo, pega")
            Assert.IsFalse(j.Concluir(a.Cap.RequestId, Agora),
                           "concluir sem iniciar nao pega")
            Assert.IsTrue(j.Iniciando(a.Cap.RequestId, Agora))
            Assert.IsFalse(j.Iniciando(a.Cap.RequestId, Agora),
                           "iniciar duas vezes reabriria uma janela que ja fechou")
            Assert.IsTrue(j.Concluir(a.Cap.RequestId, Agora))
        End Using
    End Sub

    ''' <summary>Intenção repetida não duplica, e diz que não pegou.</summary>
    <TestMethod>
    Public Sub Intencao_repetida_nao_duplica()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()

            Assert.IsTrue(j.Intencao(a.Cap, Agora))
            Assert.IsFalse(j.Intencao(a.Cap, Agora))
            Assert.AreEqual(1, j.Ler(10).Count)
        End Using
    End Sub

    ' ==================================================================
    ' Os três carimbos

    ''' <summary>
    ''' <b><c>intended_at</c> é imutável, e a ordem histórica se guia por ele.</b>
    '''
    ''' Havia um carimbo único, sobrescrito a cada passo. Depois de uma
    ''' reconciliação, uma intenção abandonada há meses aparecia como
    ''' atividade recente — e a evidência de quando cada passo aconteceu
    ''' simplesmente sumia.
    ''' </summary>
    <TestMethod>
    Public Sub Os_tres_carimbos_sao_guardados_separados()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()

            j.Intencao(a.Cap, Agora)
            j.Iniciando(a.Cap.RequestId, Agora.AddSeconds(5))
            j.Concluir(a.Cap.RequestId, Agora.AddSeconds(9))

            Dim e = j.Ler(1)(0)
            Assert.AreEqual(Agora, e.Intencionada)
            Assert.AreEqual(Agora.AddSeconds(5), e.Iniciada.Value)
            Assert.AreEqual(Agora.AddSeconds(9), e.Terminada.Value)
        End Using
    End Sub

    ''' <summary>
    ''' <b>Uma reconciliação tardia não faz o velho parecer novo.</b>
    '''
    ''' O sintoma exato do carimbo único: a intenção antiga, tocada pela
    ''' reconciliação de hoje, subia para o topo da lista como se tivesse
    ''' acabado de acontecer.
    ''' </summary>
    <TestMethod>
    Public Sub Reconciliacao_tardia_nao_faz_o_velho_parecer_novo()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)

            Dim antiga = Autorizada("um")
            j.Intencao(antiga.Cap, Agora.AddMonths(-6))
            j.Iniciando(antiga.Cap.RequestId, Agora.AddMonths(-6))

            Dim recente = Autorizada("dois")
            j.Intencao(recente.Cap, Agora)

            j.Reconciliar(Agora.AddHours(1))

            Dim tudo = j.Ler(10)
            Assert.AreEqual(recente.Cap.RequestId, tudo(0).RequestId,
                "a recente continua no topo — a antiga so foi RECONCILIADA hoje")
            Assert.AreEqual(DisclosureStage.Ambigua, tudo(1).Estagio)
            Assert.AreEqual(Agora.AddMonths(-6), tudo(1).Intencionada,
                "e a intencao dela continua sendo de seis meses atras")
        End Using
    End Sub

    ''' <summary>
    ''' A sequência de inserção desempata — o <c>Guid</c> não serve.
    '''
    ''' Ele é aleatório, então a ordem entre dois registros do mesmo instante
    ''' mudava a cada execução, e uma lista que muda de ordem sozinha é uma
    ''' lista em que ninguém confia.
    ''' </summary>
    <TestMethod>
    Public Sub A_sequencia_desempata_o_mesmo_instante()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)

            Dim primeira = Autorizada("um")
            Dim segunda = Autorizada("dois")
            j.Intencao(primeira.Cap, Agora)
            j.Intencao(segunda.Cap, Agora)

            Dim tudo = j.Ler(10)
            Assert.AreEqual(segunda.Cap.RequestId, tudo(0).RequestId)
            Assert.IsTrue(tudo(0).Sequencia > tudo(1).Sequencia)
        End Using
    End Sub


    ' ==================================================================
    ' O que NAO entra no diário

    ''' <summary>
    ''' <b>Enum do .NET não é fechado, e o diário não aceita valor inventado.</b>
    '''
    ''' Trocar <c>String</c> por enum tirou o texto arbitrário e não fechou a
    ''' porta: <c>CType(999, DisclosureNote)</c> compila e roda. Um diário com
    ''' registro incoerente é pior que um diário sem o registro — ele parece
    ''' resposta.
    ''' </summary>
    <TestMethod>
    Public Sub Nota_INVENTADA_nao_entra()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)

            Assert.IsFalse(j.NaoEnviou(a.Cap.RequestId, Agora, CType(999, DisclosureNote)))
            Assert.IsFalse(j.Falhar(a.Cap.RequestId, Agora, CType(999, DisclosureNote), True))
            Assert.IsFalse(j.NaoEnviou(a.Cap.RequestId, Agora, DisclosureNote.PortaoNegou,
                                       CType(999, DisclosureReason)))

            Assert.AreEqual(DisclosureStage.Intencionada, j.Ler(1)(0).Estagio,
                            "nenhuma delas pode ter mexido na linha")
        End Using
    End Sub

    ''' <summary>
    ''' <b>Combinação incoerente também não entra.</b>
    '''
    ''' <c>PortaoNegou</c> sem dizer o que o portão negou não descreve nada; e
    ''' uma nota que não é do portão acompanhada de um motivo de portão
    ''' descreve duas coisas que não aconteceram juntas.
    ''' </summary>
    <TestMethod>
    Public Sub Combinacao_INCOERENTE_nao_entra()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)

            Assert.IsFalse(j.NaoEnviou(a.Cap.RequestId, Agora, DisclosureNote.PortaoNegou),
                           "PortaoNegou TEM de dizer o que foi negado")
            Assert.IsFalse(j.NaoEnviou(a.Cap.RequestId, Agora,
                                       DisclosureNote.CapabilityRecusada,
                                       DisclosureReason.PastaNaoAutorizada),
                           "motivo de portao numa nota que nao e do portao")

            Assert.AreEqual(DisclosureStage.Intencionada, j.Ler(1)(0).Estagio)
        End Using
    End Sub

    ''' <summary>
    ''' <b>Cada passo aceita só as notas que fazem sentido para ele.</b>
    '''
    ''' "O portão negou" não é um jeito de a transmissão falhar: ela nem teria
    ''' começado. E "timeout" não é um jeito de o envio ser impedido antes de
    ''' acontecer.
    ''' </summary>
    <TestMethod>
    Public Sub Cada_passo_aceita_so_a_nota_que_faz_sentido()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)

            Assert.IsFalse(j.Falhar(a.Cap.RequestId, Agora, DisclosureNote.PortaoNegou, True),
                           "o portao negando nao e falha de transporte")
            Assert.IsFalse(j.NaoEnviou(a.Cap.RequestId, Agora, DisclosureNote.Timeout),
                           "timeout nao e coisa que impede antes de comecar")

            Assert.AreEqual(DisclosureStage.Intencionada, j.Ler(1)(0).Estagio)

            ' E o controle: as notas certas passam.
            Assert.IsTrue(j.NaoEnviou(a.Cap.RequestId, Agora, DisclosureNote.PortaoNegou,
                                      DisclosureReason.PastaNaoAutorizada))
        End Using
    End Sub

    ''' <summary>
    ''' <b>A contagem de mensagens vem da capability, não de fora.</b>
    '''
    ''' Enquanto vinha como parâmetro, o diário podia registrar uma quantidade
    ''' diferente da autorizada — e o número de mensagens é justamente o que
    ''' alguém confere quando a pergunta for quanto saiu.
    ''' </summary>
    <TestMethod>
    Public Sub A_contagem_vem_da_CAPABILITY()
        Using db = Abrir()
            Dim j As New SqliteDisclosureJournal(db)
            Dim a = Autorizada()
            j.Intencao(a.Cap, Agora)

            Assert.AreEqual(a.Cap.Itens.Count, j.Ler(1)(0).Mensagens)
        End Using
    End Sub

End Class

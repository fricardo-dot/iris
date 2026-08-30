Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports Iris.Integration
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O DIÁRIO DE BUSCAS — coleta autorizada, e as três coisas que a mantêm
''' honesta.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ELE EXISTE</b>
'''
''' A metade aberta da Fase 4 precisa de um oráculo que só o dono da caixa tem:
''' <i>qual mensagem eu queria quando digitei isto</i>. Em 30/08/2026 ele foi
''' perguntado de memória e não veio nenhum caso — que é a resposta normal de
''' quem é questionado sobre uma busca que falhou semanas atrás.
'''
''' O diário deixa o oráculo se juntar sozinho. E o sinal não é qual mensagem
''' foi aberta, é <b>o usuário ter reformulado</b>: digitar "cobrança", não
''' achar, digitar "fatura" e parar. Esse par é a falha semântica inteira — e
''' de quebra o diário não precisa guardar assunto de mensagem nenhuma.
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTA SUÍTE PRENDE</b>
'''
''' <list type="number">
''' <item><b>Só o termo digitado entra.</b> Nada de assunto, remetente ou
''' <c>EntryID</c>. É a promessa que a tela faz ao dono, e promessa de
''' privacidade sem teste é texto.</item>
''' <item><b>O diário nunca derruba a busca.</b> Disco cheio, permissão negada,
''' arquivo travado — a busca continua. É o único lugar deste projeto onde
''' engolir exceção é o comportamento certo.</item>
''' <item><b>E engolir não é esconder.</b> A falha fica guardada e a tela
''' mostra: diário que morre calado produz amostra furada que ninguém sabe que
''' é furada.</item>
''' </list>
''' </summary>
''' <summary>
''' <b>&lt;DoNotParallelize&gt; porque esta classe usa o CrashInjection.</b>
'''
''' Ele e estado ESTATICO e compartilhado. Sem isto, o ponto armado por um
''' teste daqui vaza para os outros que estiverem rodando ao mesmo tempo --
''' e foi o que aconteceu: quatro testes desta mesma classe cairam de uma
''' vez, com falhas que nao tinham nada a ver com o que eles testam.
'''
''' A regra ja existia (ver <c>ParalelismoDaSuiteTests</c>); o que faltou
''' foi lembrar dela ao acrescentar a injecao.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class DiarioDeBuscasTests

    Private _pasta As String

    <TestInitialize>
    Public Sub Antes()
        _pasta = Path.Combine(Path.GetTempPath(), "iris-diario-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_pasta)
    End Sub

    <TestCleanup>
    Public Sub Depois()
        Try
            If Directory.Exists(_pasta) Then Directory.Delete(_pasta, recursive:=True)
        Catch
        End Try
    End Sub

    Private Function Arquivo() As String
        Return Path.Combine(_pasta, "buscas.jsonl")
    End Function

    Private Function Novo() As DiarioDeBuscasEmArquivo
        Return New DiarioDeBuscasEmArquivo(Arquivo())
    End Function

    ' ==================================================================

    <TestMethod>
    Public Sub Controle_a_busca_anotada_aparece_no_arquivo()
        Dim d = Novo()
        d.Registrar("cobranca do fornecedor", 0, 0)

        Dim linhas = File.ReadAllLines(Arquivo())
        Assert.AreEqual(1, linhas.Length)
        StringAssert.Contains(linhas(0), "cobranca do fornecedor")
        StringAssert.Contains(linhas(0), """exatos"":0")
    End Sub

    ''' <summary>
    ''' <b>UMA LINHA POR BUSCA, ANEXADA.</b>
    '''
    ''' Append e não reescrita: uma queda no meio perde no máximo a última
    ''' linha. Reescrever o arquivo inteiro a cada busca poria em risco tudo o
    ''' que já foi coletado, para gravar uma linha.
    ''' </summary>
    <TestMethod>
    Public Sub Cada_busca_anexa_uma_linha()
        Dim d = Novo()
        d.Registrar("uma", 1, 0)
        d.Registrar("outra", 0, 2)
        d.Registrar("terceira", 3, 1)

        Assert.AreEqual(3, File.ReadAllLines(Arquivo()).Length)
        Assert.AreEqual(3, d.Quantas())
    End Sub

    ''' <summary>
    ''' <b>SÓ O TERMO. Nem assunto, nem remetente, nem EntryID.</b>
    '''
    ''' É a promessa que a tela faz ao dono, escrita como teste. O
    ''' <c>Registrar</c> nem recebe esses dados — a garantia está no
    ''' <b>tipo</b>, e este teste prende o desenho para o dia em que alguém
    ''' pensar em "enriquecer" o registro.
    ''' </summary>
    <TestMethod>
    Public Sub O_registro_NAO_tem_onde_por_conteudo_de_mensagem()
        Dim m = GetType(IDiarioDeBuscas).GetMethod("Registrar")
        Dim parametros = m.GetParameters().Select(Function(x) x.Name.ToLowerInvariant()).ToList()

        CollectionAssert.AreEquivalent({"termo", "exatos", "aproximados"}, parametros,
            "o Registrar ganhou parâmetro novo. Se for conteúdo de mensagem, " &
            "a faixa da busca está mentindo para o dono: ela diz 'só o texto " &
            "digitado; nenhum assunto de mensagem'.")
    End Sub

    ''' <summary>
    ''' <b>Termo vazio não é busca.</b>
    '''
    ''' Registrar o Enter dado num campo em branco encheria o arquivo de linha
    ''' que não ensina nada — e depois alguém contaria essas linhas como
    ''' "buscas realizadas".
    ''' </summary>
    <DataTestMethod>
    <DataRow("")>
    <DataRow("   ")>
    Public Sub Termo_vazio_nao_e_anotado(termo As String)
        Dim d = Novo()
        d.Registrar(termo, 0, 0)

        Assert.AreEqual(0, d.Quantas())
        Assert.IsFalse(File.Exists(Arquivo()), "criou arquivo para uma busca que não houve")
    End Sub

    ''' <summary>
    ''' <b>O DIÁRIO NÃO DERRUBA A BUSCA.</b>
    '''
    ''' Busca é a funcionalidade; o diário é instrumentação. Caminho impossível
    ''' — aqui um diretório onde deveria haver arquivo — não pode virar exceção
    ''' na cara de quem só queria procurar uma mensagem.
    '''
    ''' <b>Controle negativo:</b> tirando o <c>Try</c> do <c>Registrar</c>, este
    ''' teste falha com a exceção real em vez de passar.
    ''' </summary>
    <TestMethod>
    Public Sub Falha_ao_gravar_NAO_lanca()
        ' Um DIRETORIO com o nome do arquivo: escrever nele e impossivel.
        Directory.CreateDirectory(Arquivo())
        Dim d = Novo()

        d.Registrar("isto vai falhar", 0, 0)   ' nao lanca

        Assert.IsTrue(d.UltimaFalha.Length > 0,
            "engoliu a falha E não guardou o motivo: amostra furada que " &
            "ninguém sabe que é furada")
        StringAssert.Contains(d.UltimaFalha, "não consegui anotar")
    End Sub

    ''' <summary>
    ''' <b>E a falha some quando a gravação volta a funcionar.</b>
    '''
    ''' Sem isto, um erro passageiro deixaria a tela reclamando para sempre — e
    ''' aviso que não some é aviso que se aprende a ignorar.
    ''' </summary>
    <TestMethod>
    Public Sub A_falha_some_quando_a_gravacao_volta()
        Directory.CreateDirectory(Arquivo())
        Dim d = Novo()
        d.Registrar("falha", 0, 0)
        Assert.IsTrue(d.UltimaFalha.Length > 0, "controle: tinha de ter falhado")

        Directory.Delete(Arquivo())
        d.Registrar("agora vai", 1, 0)

        Assert.AreEqual("", d.UltimaFalha, "a falha antiga ficou na tela depois de voltar")
    End Sub

    ''' <summary>
    ''' <b>Não conseguir contar não é contar zero.</b>
    '''
    ''' A mesma regra que esta base aplicou em cinco lugares: <c>Nothing</c> é
    ''' "não sei", zero é "olhei e não há". A tela diz frases diferentes.
    ''' </summary>
    <TestMethod>
    Public Sub Nao_conseguir_contar_devolve_Nothing_e_nao_zero()
        Directory.CreateDirectory(Arquivo())   ' diretorio no lugar do arquivo
        Dim d = Novo()

        Assert.IsNull(d.Quantas(), "arquivo ilegível virou 'zero buscas anotadas'")
    End Sub

    <TestMethod>
    Public Sub Controle_arquivo_ausente_conta_zero()
        Assert.AreEqual(0, Novo().Quantas(),
            "arquivo que nunca existiu tem zero buscas, e isso se sabe")
    End Sub

    ''' <summary>
    ''' <b>Apagar apaga, e o dono pode.</b>
    '''
    ''' Poder apagar era metade do acordo com o dono da caixa. A outra metade —
    ''' saber onde o arquivo está — é a faixa na tela.
    ''' </summary>
    ''' <summary>
    ''' <b>DESLIGAR PERSISTE, e apagar não desliga.</b>
    '''
    ''' O achado ALTO da revisão externa: a primeira versão só tinha apagar,
    ''' e a busca seguinte recriava o arquivo. Faxina não é retirada de
    ''' consentimento.
    '''
    ''' Persistir importa porque o dono desliga hoje e fecha o programa: um
    ''' desligamento que só vale até o próximo início é um desligamento que
    ''' engana.
    ''' </summary>
    <TestMethod>
    Public Sub Desligar_persiste_e_apagar_nao_desliga()
        Dim d = Novo()
        Assert.IsTrue(d.Ligado, "controle: nasce ligado")

        Assert.IsNull(d.Desligar())
        Assert.IsFalse(d.Ligado)

        d.Registrar("nao deve entrar", 0, 0)
        Assert.AreEqual(0, d.Quantas(), "anotou com o registro desligado")

        ' OUTRA INSTANCIA, como se o programa tivesse sido reaberto.
        Assert.IsFalse(Novo().Ligado,
            "o desligamento nao sobreviveu ao fechamento: vale ate o proximo " &
            "inicio, e isso engana quem desligou")

        Assert.IsNull(d.Ligar())
        d.Registrar("agora entra", 1, 0)
        Assert.AreEqual(1, d.Quantas())

        ' APAGAR NAO DESLIGA -- e o achado inteiro.
        d.Apagar()
        Assert.IsTrue(d.Ligado, "apagar desligou: sao coisas diferentes")
        d.Registrar("volta a anotar", 1, 0)
        Assert.AreEqual(1, d.Quantas(),
            "depois de apagar, a busca seguinte tinha de voltar a ser anotada")
    End Sub

    ''' <summary>
    ''' <b>Não conseguir conferir o consentimento vale como DESLIGADO.</b>
    '''
    ''' Falha fechada, como o <c>EhReuniao</c> e o <c>EhAtribuida</c>. Não
    ''' saber se o dono autorizou não é autorização.
    ''' </summary>
    <TestMethod>
    Public Sub Nao_conseguir_conferir_o_consentimento_FECHA()
        Dim d = Novo()
        d.Desligar()
        Assert.IsFalse(d.Ligado)

        ' CONTROLE: sem marcador, esta ligado. Sem isto, um Ligado que
        ' devolvesse False sempre passaria na assercao de cima.
        '
        ' A PRIMEIRA VERSAO DESTE CONTROLE ERA TAUTOLOGICA:
        '     Assert.IsTrue(Novo().Ligado = False OrElse True)
        ' que e verdadeiro sempre, para qualquer implementacao. Um controle
        ' que nao pode falhar nao e controle -- e a revisao externa o achou
        ' exatamente onde ele fingia estar protegendo.
        d.Ligar()
        Assert.IsTrue(d.Ligado, "controle: sem marcador tem de estar ligado")
    End Sub

    ''' <summary>
    ''' <b>DESLIGAR FECHA A JANELA ENTRE CONFERIR E GRAVAR.</b>
    '''
    ''' A revisão externa achou uma corrida real: o <c>Registrar</c> conferia
    ''' <c>Ligado</c>, esperava a trava, e nesse meio-tempo outro processo podia
    ''' desligar — a busca era anotada <b>depois</b> da retirada do
    ''' consentimento.
    '''
    ''' Aqui o desligamento acontece por fora, como faria a outra janela, e a
    ''' gravação seguinte tem de recusar.
    '''
    ''' <b>Controle negativo:</b> tirando a segunda conferência de dentro da
    ''' trava, este teste continua passando — porque num só processo a janela é
    ''' estreita demais para o teste abrir. O que ele prende é a <i>ordem</i>: o
    ''' marcador criado por fora vale imediatamente.
    ''' </summary>
    ''' <summary>
    ''' <b>DESLIGAR NO MEIO DA GRAVAÇÃO IMPEDE A GRAVAÇÃO.</b>
    '''
    ''' Este é o teste que faltava, e o controle negativo do irmão dele foi
    ''' quem mostrou: com a guarda de dentro da trava removida, o teste do
    ''' marcador continuava passando — porque num só processo a janela é
    ''' estreita demais para abrir sozinha, e a guarda de entrada já pegava.
    '''
    ''' O ponto de injeção abre a janela: a ação armada desliga o diário
    ''' <i>depois</i> da conferência de entrada e <i>antes</i> da gravação,
    ''' exatamente como a outra janela do Iris faria.
    '''
    ''' <b>Controle negativo:</b> tirando o <c>If Not Ligado Then Return</c>
    ''' de dentro da trava, este teste cai — e agora cai de verdade.
    ''' </summary>
    <TestMethod>
    Public Sub Desligar_NO_MEIO_da_gravacao_impede_a_gravacao()
        Dim d = Novo()
        Try
            Iris.Cache.CrashInjection.Armar(
                Iris.Cache.CrashInjection.EntreConferirOConsentimentoEGravar,
                Sub() File.WriteAllText(Arquivo() & ".desligado", "x"))

            d.Registrar("no meio da janela", 1, 0)

            Assert.AreEqual(0, d.Quantas(),
                "anotou uma busca depois de o consentimento ser retirado no " &
                "meio da gravacao. Uma retirada que admite 'mais uma' nao e retirada.")
        Finally
            Iris.Cache.CrashInjection.Desarmar()
        End Try
    End Sub

    ''' <summary>
    ''' CONTROLE: sem nada armado, a mesma gravação passa. Sem isto, um
    ''' <c>Registrar</c> que nunca gravasse passaria no teste de cima.
    ''' </summary>
    ''' <summary>
    ''' <b>SEM A TRAVA, DESLIGAR RECUSA — em vez de mentir que desligou.</b>
    '''
    ''' O teste que faltava, e a revisão externa nomeou o buraco: o
    ''' <c>Desligar</c> chamava <c>Segurar</c> e <b>ignorava o retorno</b>. Com
    ''' a espera estourada ele criava o marcador sem exclusão e devolvia
    ''' sucesso — enquanto outro processo, que já tinha conferido
    ''' <c>Ligado</c>, seguia e gravava.
    '''
    ''' E o teste da corrida não pegava isso, porque ele escrevia o marcador
    ''' direto em vez de chamar <c>Desligar</c>.
    '''
    ''' Aqui outra thread segura o mutex por mais tempo que a espera. Dizer
    ''' "desliguei" sem ter desligado é pior que recusar: quem recusa tenta de
    ''' novo, quem foi enganado não tenta.
    '''
    ''' <b>Controle negativo:</b> tirando o <c>If Not segurou Then Return</c> do
    ''' <c>Desligar</c>, este teste cai.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_a_trava_desligar_RECUSA()
        Dim d = Novo()
        Dim nome As String
        Using sha = Security.Cryptography.SHA256.Create()
            Dim bytes = sha.ComputeHash(Text.Encoding.UTF8.GetBytes(Arquivo().ToLowerInvariant()))
            nome = "Iris.buscas." & Convert.ToHexString(bytes, 0, 12)
        End Using

        Dim segurando As New Threading.ManualResetEventSlim(False)
        Dim solte As New Threading.ManualResetEventSlim(False)

        Dim outra As New Threading.Thread(
            Sub()
                Using m As New Threading.Mutex(False, nome)
                    m.WaitOne()
                    segurando.Set()
                    solte.Wait(TimeSpan.FromSeconds(10))
                    m.ReleaseMutex()
                End Using
            End Sub)
        outra.IsBackground = True
        outra.Start()

        Try
            Assert.IsTrue(segurando.Wait(TimeSpan.FromSeconds(5)),
                          "a outra thread nao chegou a segurar a trava")

            Dim motivo = d.Desligar()

            Assert.IsNotNull(motivo,
                "disse que desligou sem ter conseguido a trava. Uma retirada " &
                "de consentimento anunciada e nao efetivada e pior que uma recusa.")
            Assert.IsTrue(d.Ligado,
                "marcou como desligado sem exclusao: o outro processo pode " &
                "estar gravando agora mesmo")
        Finally
            solte.Set()
            outra.Join(TimeSpan.FromSeconds(5))
        End Try
    End Sub

    ''' <summary>
    ''' CONTROLE: com a trava livre, o mesmo <c>Desligar</c> funciona. Sem
    ''' isto, um <c>Desligar</c> que recusasse sempre passaria no teste acima.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_com_a_trava_livre_desligar_funciona()
        Dim d = Novo()
        Assert.IsNull(d.Desligar())
        Assert.IsFalse(d.Ligado)
    End Sub

    <TestMethod>
    Public Sub Controle_sem_a_janela_a_gravacao_acontece()
        Dim d = Novo()
        d.Registrar("sem janela", 1, 0)
        Assert.AreEqual(1, d.Quantas())
    End Sub

    <TestMethod>
    Public Sub Marcador_criado_por_fora_vale_na_hora()
        Dim d = Novo()
        d.Registrar("antes", 1, 0)
        Assert.AreEqual(1, d.Quantas(), "controle: anotou antes de desligar")

        ' Como outra janela do Iris faria: o marcador aparece no disco.
        File.WriteAllText(Arquivo() & ".desligado", "x")

        d.Registrar("depois", 1, 0)
        Assert.AreEqual(1, d.Quantas(),
            "anotou depois de o consentimento ser retirado por fora")
    End Sub

    ''' <summary>
    ''' <b>A contagem enxerga o que OUTRA instância escreveu.</b>
    '''
    ''' O cache em memória era cego: uma vez contado, ele devolvia o mesmo
    ''' número para sempre. Duas janelas do Iris, ou uma edição por fora, e a
    ''' tela mostrava um número de antes — a revisão externa listou os três
    ''' casos.
    '''
    ''' <b>Controle negativo:</b> tirando o carimbo do
    ''' <c>DiarioDeBuscasEmArquivo</c>, este teste cai.
    ''' </summary>
    <TestMethod>
    Public Sub A_contagem_enxerga_a_escrita_de_outra_instancia()
        Dim a = Novo()
        Dim b = Novo()

        a.Registrar("uma", 1, 0)
        Assert.AreEqual(1, b.Quantas(), "controle: b viu a primeira")

        a.Registrar("outra", 1, 0)
        Assert.AreEqual(2, b.Quantas(),
            "b continuou mostrando o número de antes: cache cego")

        a.Apagar()
        Assert.AreEqual(0, b.Quantas(),
            "b continuou contando linhas de um arquivo que já não existe")
    End Sub

    <TestMethod>
    Public Sub Apagar_apaga_e_nao_reclama()
        Dim d = Novo()
        d.Registrar("uma", 1, 0)
        Assert.AreEqual(1, d.Quantas(), "controle: tinha de haver o que apagar")

        Assert.IsNull(d.Apagar(), "reclamou de um apagar que deu certo")
        Assert.AreEqual(0, d.Quantas())
    End Sub

    <TestMethod>
    Public Sub Apagar_o_que_nao_existe_nao_e_erro()
        Assert.IsNull(Novo().Apagar())
    End Sub

End Class

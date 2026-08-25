Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports Iris.Assist
Imports Iris.Integration
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace Global.Iris.Tests

    ''' <summary>
    ''' <b>A cerimônia de ativação, lida de um arquivo.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTE ARQUIVO COBRA</b>
    '''
    ''' Que o parse seja <b>estrito</b> e a recusa <b>por inteiro</b>. Este é o
    ''' único ponto do Iris em que um arquivo de texto decide se conteúdo da
    ''' caixa pode sair da máquina — e é onde "não entendi este pedaço, o resto
    ''' vale" custaria mais caro.
    '''
    ''' Cada teste estraga <b>uma</b> coisa a partir de um arquivo bom, e o
    ''' controle é justamente esse arquivo bom carregando. Sem ele, um
    ''' carregador que recusasse tudo passaria em quase todo o resto.
    ''' </summary>
    <TestClass>
    <DoNotParallelize>
    Public Class ActivationLoaderTests

        Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 25, 12, 0, 0, TimeSpan.Zero)

        Private ReadOnly _pastas As New List(Of String)()

        <TestCleanup>
        Public Sub Limpar()
            For Each p In _pastas
                Try
                    If Directory.Exists(p) Then Directory.Delete(p, True)
                Catch
                End Try
            Next
            _pastas.Clear()
        End Sub

        ''' <summary>Um arquivo com o conteúdo dado, em pasta própria.</summary>
        Private Function Arquivo(conteudo As String) As String
            Dim pasta = Path.Combine(Path.GetTempPath(), "iris-ativ-" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(pasta)
            _pastas.Add(pasta)
            Dim caminho = Path.Combine(pasta, "ativacao.json")
            File.WriteAllText(caminho, conteudo, New UTF8Encoding(False))
            Return caminho
        End Function

        ''' <summary>
        ''' A ativação boa — a que o usuário vai escrever de verdade.
        '''
        ''' Todo teste adversarial parte desta e estraga um campo. Ela sozinha é
        ''' o controle do arquivo inteiro.
        ''' </summary>
        Private Shared Function Boa() As String
            Return "{" &
                """id"": ""ativacao-2026-08""," &
                """versao"": 1," &
                """autoridade"": ""Ricardo, dono da caixa""," &
                """politicaCorporativaVerificada"": false," &
                """quando"": ""2026-08-25T00:00:00Z""," &
                """ate"": ""2026-09-24T00:00:00Z""," &
                """provedor"": ""openrouter""," &
                """endpoint"": ""https://openrouter.ai/api/v1/chat/completions""," &
                """modelo"": ""anthropic/claude-sonnet-4""," &
                """regiao"": ""não imposta pelo pedido""," &
                """retencaoAceita"": ""retenção zero exigida no pedido""," &
                """exigirRetencaoZero"": true," &
                """provedoresPermitidos"": [""anthropic""]," &
                """operacoes"": [""Resumir"", ""Redigir""]," &
                """pastas"": [{""storeId"": ""S-1"", ""entryId"": ""F-teste""}]," &
                """rotulos"": []," &
                """leituras"": [""Absent""]," &
                """contentBits"": [0]" &
                "}"
        End Function

        Private Function Carregar(conteudo As String) As ActivationLoadResult
            Return ActivationLoader.Carregar(Arquivo(conteudo), Agora)
        End Function

        ''' <summary>Troca um pedaço do arquivo bom por outro.</summary>
        Private Shared Function Trocando(velho As String, novo As String) As String
            Dim s = Boa()
            Assert.IsTrue(s.Contains(velho), "o teste nao achou o que ia estragar: " & velho)
            Return s.Replace(velho, novo)
        End Function

        ' ==================================================================
        ' O controle: o arquivo bom carrega

        ''' <summary>
        ''' <b>A ativação boa carrega, e com os campos que estão no arquivo.</b>
        '''
        ''' O controle de todo o resto. Um carregador que recusasse tudo passaria
        ''' em cada um dos testes adversariais abaixo, e falharia só aqui.
        ''' </summary>
        <TestMethod>
        Public Sub A_ativacao_boa_CARREGA()
            Dim r = Carregar(Boa())

            Assert.IsTrue(r.Carregou, $"nao carregou: {r.Falha} {r.Campo}")
            Assert.AreEqual("ativacao-2026-08", r.Record.Id)
            Assert.AreEqual("openrouter", r.Record.Provedor)
            Assert.AreEqual("anthropic/claude-sonnet-4", r.Record.Modelo)
            Assert.IsFalse(r.Record.PoliticaCorporativaVerificada,
                           "o arquivo diz false, e false tem de chegar como false")
            Assert.IsTrue(r.Record.ExigirRetencaoZero)
            CollectionAssert.AreEqual({"anthropic"}, r.Record.ProvedoresPermitidos.ToArray())
            Assert.AreEqual(2, r.Record.Operacoes.Count)
            Assert.AreEqual(1, r.Record.Pastas.Count)
            Assert.AreEqual("F-teste", r.Record.Pastas(0).EntryId)
            Assert.IsTrue(r.Record.Completo() AndAlso r.Record.Coerente())
            Assert.IsTrue(r.Record.Vigente(Agora))
        End Sub

        ' ==================================================================
        ' O arquivo

        ''' <summary>
        ''' <b>Sem arquivo, sem ativação — e o motivo é "ausente".</b>
        '''
        ''' O caso normal, e o mais importante de distinguir: a IA nasce
        ''' desligada porque ninguém escreveu ativação, e não porque algo
        ''' quebrou.
        ''' </summary>
        <TestMethod>
        Public Sub Sem_arquivo_a_falha_e_AUSENTE()
            Dim pasta = Path.Combine(Path.GetTempPath(), "iris-ativ-" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(pasta)
            _pastas.Add(pasta)

            Dim r = ActivationLoader.Carregar(Path.Combine(pasta, "nao-existe.json"), Agora)

            Assert.IsFalse(r.Carregou)
            Assert.AreEqual(ActivationLoadFailure.Ausente, r.Falha)
            Assert.IsNull(r.Record)
        End Sub

        ''' <summary>
        ''' <b>Arquivo grande demais não é lido.</b>
        '''
        ''' Ativação é dezenas de linhas. Um arquivo de megabytes no lugar dela é
        ''' outra coisa — e ler para descobrir o que é já seria ter lido.
        ''' </summary>
        <TestMethod>
        Public Sub Arquivo_GRANDE_DEMAIS_nao_e_lido()
            Dim r = Carregar("{""x"": """ & New String("a"c, ActivationLoader.TetoDeBytes) & """}")

            Assert.AreEqual(ActivationLoadFailure.GrandeDemais, r.Falha)
        End Sub

        ''' <summary>
        ''' <b>Diretório no lugar do arquivo não vira ativação.</b>
        ''' </summary>
        <TestMethod>
        Public Sub Diretorio_no_lugar_do_arquivo_NAO_carrega()
            Dim pasta = Path.Combine(Path.GetTempPath(), "iris-ativ-" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(pasta)
            _pastas.Add(pasta)
            Dim alvo = Path.Combine(pasta, "ativacao.json")
            Directory.CreateDirectory(alvo)

            Dim r = ActivationLoader.Carregar(alvo, Agora)

            Assert.IsFalse(r.Carregou)
        End Sub

        ' ==================================================================
        ' O parse é estrito

        ''' <summary>
        ''' <b>Campo desconhecido recusa o arquivo inteiro.</b>
        '''
        ''' É a proteção contra o erro de digitação silencioso: quem escreve
        ''' <c>pasta</c> em vez de <c>pastas</c> teria, sem isto, uma ativação
        ''' com lista de pastas vazia — que nega tudo, mas por motivo que não
        ''' aparece em lugar nenhum.
        ''' </summary>
        <TestMethod>
        Public Sub Campo_DESCONHECIDO_recusa()
            Dim r = Carregar(Trocando("""versao"": 1,", """versao"": 1, ""pastaz"": [],"))

            Assert.AreEqual(ActivationLoadFailure.CampoDesconhecido, r.Falha)
            Assert.AreEqual("pastaz", r.Campo)
        End Sub

        ''' <summary>
        ''' <b>Campo repetido recusa.</b>
        '''
        ''' O <c>JsonDocument</c> aceita duplicata em silêncio e entrega a
        ''' última. Quem escreveu duas linhas de <c>modelo</c> não sabe qual
        ''' delas vale, e adivinhar por ele é decidir o modelo no lugar dele.
        ''' </summary>
        <TestMethod>
        Public Sub Campo_DUPLICADO_recusa()
            Dim r = Carregar(Trocando("""versao"": 1,", """versao"": 1, ""versao"": 2,"))

            Assert.AreEqual(ActivationLoadFailure.CampoDuplicado, r.Falha)
            Assert.AreEqual("versao", r.Campo)
        End Sub

        ''' <summary><b>Campo obrigatório ausente recusa, e diz qual.</b></summary>
        <TestMethod>
        Public Sub Campo_FALTANDO_recusa()
            Dim r = Carregar(Trocando("""modelo"": ""anthropic/claude-sonnet-4"",", ""))

            Assert.AreEqual(ActivationLoadFailure.CampoFaltando, r.Falha)
            Assert.AreEqual("modelo", r.Campo)
        End Sub

        ''' <summary>
        ''' <b>Booleano escrito como número ou texto recusa.</b>
        '''
        ''' <c>"politicaCorporativaVerificada": "false"</c> é string, e string
        ''' não-vazia é o tipo de coisa que um parser distraído lê como
        ''' verdadeira. O campo diz sim ou não, e sinônimo nenhum é aceito.
        ''' </summary>
        <TestMethod>
        Public Sub Booleano_como_TEXTO_ou_NUMERO_recusa()
            For Each falso In {"""false""", "0"}
                Dim r = Carregar(Trocando("""politicaCorporativaVerificada"": false",
                                          """politicaCorporativaVerificada"": " & falso))
                Assert.AreEqual(ActivationLoadFailure.TipoErrado, r.Falha, falso)
                Assert.AreEqual("politicaCorporativaVerificada", r.Campo)
            Next
        End Sub

        ''' <summary>
        ''' <b>Enum por número recusa.</b>
        '''
        ''' O <c>TryParse</c> do .NET aceita <c>"1"</c> e devolve <c>True</c>.
        ''' Número num arquivo escrito à mão não diz nada a quem relê, e um
        ''' membro inserido no meio do enum mudaria o que o arquivo antigo
        ''' autoriza sem ninguém tocar nele.
        ''' </summary>
        <TestMethod>
        Public Sub Enum_por_NUMERO_recusa()
            Dim r = Carregar(Trocando("""operacoes"": [""Resumir"", ""Redigir""]",
                                      """operacoes"": [""1""]"))

            Assert.AreEqual(ActivationLoadFailure.ValorInvalido, r.Falha)
            Assert.AreEqual("operacoes", r.Campo)
        End Sub

        ''' <summary><b>Enum que não existe recusa.</b></summary>
        <TestMethod>
        Public Sub Enum_DESCONHECIDO_recusa()
            Dim r = Carregar(Trocando("""operacoes"": [""Resumir"", ""Redigir""]",
                                      """operacoes"": [""Exfiltrar""]"))

            Assert.AreEqual(ActivationLoadFailure.ValorInvalido, r.Falha)
        End Sub

        ''' <summary>
        ''' <b>Data sem deslocamento recusa.</b>
        '''
        ''' <c>2026-09-24T00:00:00</c> vira hora local de quem lê. Uma
        ''' autorização que vence em horários diferentes conforme a máquina não
        ''' é uma autorização.
        ''' </summary>
        <TestMethod>
        Public Sub Data_SEM_deslocamento_recusa()
            Dim r = Carregar(Trocando("""ate"": ""2026-09-24T00:00:00Z""",
                                      """ate"": ""2026-09-24T00:00:00"""))

            Assert.AreEqual(ActivationLoadFailure.ValorInvalido, r.Falha)
            Assert.AreEqual("ate", r.Campo)
        End Sub

        ''' <summary><b>JSON com comentário ou vírgula sobrando recusa.</b></summary>
        <TestMethod>
        Public Sub JSON_com_comentario_ou_virgula_sobrando_recusa()
            Assert.AreEqual(ActivationLoadFailure.JsonInvalido,
                            Carregar(Trocando("""versao"": 1,", "/* nota */ ""versao"": 1,")).Falha)
            Assert.AreEqual(ActivationLoadFailure.JsonInvalido,
                            Carregar(Trocando("""contentBits"": [0]", """contentBits"": [0,]")).Falha)
        End Sub

        ''' <summary>
        ''' <b>Pasta com campo a mais, ou faltando, recusa.</b>
        '''
        ''' Pasta é um par — <c>storeId</c> e <c>entryId</c> —, e um objeto com
        ''' três chaves quer dizer que quem escreveu acha que declarou algo que
        ''' o Iris não vai ler.
        ''' </summary>
        <TestMethod>
        Public Sub Pasta_com_forma_ERRADA_recusa()
            Dim bom = "{""storeId"": ""S-1"", ""entryId"": ""F-teste""}"
            For Each ruim In {"{""storeId"": ""S-1""}",
                              "{""storeId"": ""S-1"", ""entryId"": ""F"", ""nome"": ""x""}",
                              """S-1/F-teste"""}
                Dim r = Carregar(Trocando(bom, ruim))
                Assert.AreEqual(ActivationLoadFailure.TipoErrado, r.Falha, ruim)
                Assert.AreEqual("pastas", r.Campo)
            Next
        End Sub

        ' ==================================================================
        ' As regras da própria autorização

        ''' <summary>
        ''' <b>Sem prazo, não há ativação.</b>
        '''
        ''' O campo é obrigatório e o registro é incompleto sem ele. Antes,
        ''' ausência de prazo queria dizer vigência eterna — a falha aberta de
        ''' sempre, com a ausência de dado virando permissão.
        ''' </summary>
        <TestMethod>
        Public Sub Sem_PRAZO_nao_ha_ativacao()
            Dim r = Carregar(Trocando("""ate"": ""2026-09-24T00:00:00Z"",", ""))

            Assert.AreEqual(ActivationLoadFailure.CampoFaltando, r.Falha)
            Assert.AreEqual("ate", r.Campo)
        End Sub

        ''' <summary><b>Prazo que termina antes de começar é incoerente.</b></summary>
        <TestMethod>
        Public Sub Prazo_INVERTIDO_e_incoerente()
            Dim r = Carregar(Trocando("""ate"": ""2026-09-24T00:00:00Z""",
                                      """ate"": ""2026-08-01T00:00:00Z"""))

            Assert.AreEqual(ActivationLoadFailure.Incoerente, r.Falha)
        End Sub

        ''' <summary>
        ''' <b>Prazo longo demais recusa.</b>
        '''
        ''' Uma autorização de dez anos é uma autorização que ninguém vai
        ''' reencontrar. O teto força a decisão a ser revisitada enquanto ainda
        ''' é lembrada.
        ''' </summary>
        <TestMethod>
        Public Sub Prazo_LONGO_DEMAIS_recusa()
            Dim r = Carregar(Trocando("""ate"": ""2026-09-24T00:00:00Z""",
                                      """ate"": ""2030-01-01T00:00:00Z"""))

            Assert.AreEqual(ActivationLoadFailure.PrazoLongoDemais, r.Falha)
            Assert.AreEqual("ate", r.Campo)
        End Sub

        ''' <summary>
        ''' <b>Endpoint que não é HTTPS recusa aqui também.</b>
        '''
        ''' O portão já recusa, e esta é a segunda barreira: uma ativação com
        ''' endereço em claro não devia nem chegar a existir como registro
        ''' carregado.
        ''' </summary>
        <TestMethod>
        Public Sub Endpoint_sem_HTTPS_recusa()
            Dim r = Carregar(Trocando("https://openrouter.ai", "http://openrouter.ai"))

            Assert.AreEqual(ActivationLoadFailure.EndpointInseguro, r.Falha)
        End Sub

        ''' <summary>
        ''' <b>Desfecho de leitura não elegível recusa.</b>
        '''
        ''' Autorizar sobre <c>Unreadable</c> é autorizar sobre "não consegui
        ''' ler" — decidir sem informação, com cara de decisão.
        ''' </summary>
        <TestMethod>
        Public Sub Leitura_NAO_ELEGIVEL_recusa()
            Dim r = Carregar(Trocando("""leituras"": [""Absent""]",
                                      """leituras"": [""Absent"", ""Unreadable""]"))

            Assert.AreEqual(ActivationLoadFailure.Incoerente, r.Falha)
        End Sub

        ''' <summary>
        ''' <b>Provedor permitido em branco recusa.</b>
        '''
        ''' Uma entrada vazia na lista casaria com nada e pareceria uma
        ''' restrição. Restrição que não restringe é pior que nenhuma.
        ''' </summary>
        <TestMethod>
        Public Sub Provedor_permitido_em_BRANCO_recusa()
            Dim r = Carregar(Trocando("""provedoresPermitidos"": [""anthropic""]",
                                      """provedoresPermitidos"": [""anthropic"", ""  ""]"))

            Assert.AreEqual(ActivationLoadFailure.Incoerente, r.Falha)
        End Sub

        ''' <summary>
        ''' <b>Sem pasta nenhuma, a ativação é incompleta.</b>
        '''
        ''' Lista vazia é o que sai de um erro de digitação no nome do campo, e
        ''' é a diferença entre "nega tudo por política" e "nega tudo porque
        ''' ninguém escreveu nada".
        ''' </summary>
        <TestMethod>
        Public Sub Sem_PASTA_a_ativacao_e_incompleta()
            Dim r = Carregar(Trocando("[{""storeId"": ""S-1"", ""entryId"": ""F-teste""}]", "[]"))

            Assert.AreEqual(ActivationLoadFailure.Incompleta, r.Falha)
        End Sub

    End Class

End Namespace

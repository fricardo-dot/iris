Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.Json
Imports Iris.Assist
Imports Iris.Integration.Assist.Http
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace Global.Iris.Tests

    ''' <summary>
    ''' <b>O corpo que sai para o OpenRouter — e o que não sai.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ESTE ARQUIVO EXISTE</b>
    '''
    ''' A capability cobre o corpo preparado, e não só o envelope. Isso fecha o
    ''' furo de "autorizar um artefato e transmitir outro" — mas só se a
    ''' preparação for exatamente o que ela diz ser. Um campo a mais aqui é um
    ''' campo que a autorização passou a cobrir sem ninguém ter decidido pô-lo.
    '''
    ''' O que se cobra: que o envelope atravesse <b>verbatim</b>, que não haja
    ''' mensagem <c>system</c>, que o bloco <c>provider</c> imponha o que a
    ''' ativação declara, e que a leitura da resposta seja estrita.
    ''' </summary>
    <TestClass>
    Public Class OpenRouterCodecTests

        Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 25, 12, 0, 0, TimeSpan.Zero)
        Private Shared ReadOnly Pasta As New FolderKey("F-teste", "S-1")

        Private Shared Function Ativacao(Optional zero As Boolean = True,
                                         Optional slugs As String() = Nothing,
                                         Optional modelo As String = "google/gemini-3.7-flash") _
                                         As ActivationRecord
            Return New ActivationRecord(
                "ativacao-1", 1, "Ricardo, dono da caixa", Agora.AddDays(-1),
                "openrouter", "https://openrouter.ai/api/v1/chat/completions", modelo,
                "não imposta pelo pedido", "retenção zero exigida",
                {AssistOperation.Resumir}, {Pasta}, Array.Empty(Of String)(),
                {LabelReadingKind.Absent}, {0},
                ate:=Agora.AddDays(30),
                politicaCorporativaVerificada:=False,
                exigirRetencaoZero:=zero,
                provedoresPermitidos:=If(slugs, {"google"}))
        End Function

        Private Shared Function Provedor(Optional zero As Boolean = True,
                                         Optional slugs As String() = Nothing,
                                         Optional modelo As String = "google/gemini-3.7-flash") _
                                         As OpenRouterAssistantProvider
            Return New OpenRouterAssistantProvider(Ativacao(zero, slugs, modelo),
                                                   Function() "chave-de-teste")
        End Function

        ''' <summary>Um envelope de verdade, montado pelo builder de verdade.</summary>
        Private Shared Function Envelope() As Byte()
            Dim s As New MessageSnapshot(New ItemKey("E-1", "S-1"), "CK-1",
                                         "assunto", "de@x.invalido", {"para@x.invalido"},
                                         "olá, tudo bem?", False, True, temAnexo:=False)
            Dim parte = ContentPipeline.Preparar(s)
            Assert.IsTrue(parte.Ok)
            Dim r = New EnvelopeBuilder().Montar(AssistOperation.Resumir, "resuma", {parte.Parte})
            Assert.IsTrue(r.Ok)
            Return r.Envelope.Bytes()
        End Function

        Private Shared Function Corpo(Optional p As OpenRouterAssistantProvider = Nothing,
                                      Optional env As Byte() = Nothing) As JsonDocument
            Dim bytes = If(p, Provedor()).Preparar(If(env, Envelope()))
            Assert.IsNotNull(bytes, "o codec devia ter preparado")
            Return JsonDocument.Parse(bytes)
        End Function

        ' ==================================================================
        ' O envelope atravessa inteiro

        ''' <summary>
        ''' <b>O envelope vai verbatim, como conteúdo da única mensagem.</b>
        '''
        ''' Byte a byte: o que a capability cobre é o envelope, e o que o modelo
        ''' recebe tem de ser ele, e não uma versão reescrita a caminho.
        ''' </summary>
        <TestMethod>
        Public Sub O_envelope_atravessa_VERBATIM()
            Dim env = Envelope()
            Using doc = Corpo(env:=env)
                Dim msgs = doc.RootElement.GetProperty("messages")
                Assert.AreEqual(1, msgs.GetArrayLength(),
                                "uma mensagem so: a segunda seria conteudo que ninguem autorizou")
                Assert.AreEqual("user", msgs(0).GetProperty("role").GetString())

                Dim conteudo = msgs(0).GetProperty("content").GetString()
                CollectionAssert.AreEqual(env, Encoding.UTF8.GetBytes(conteudo),
                                          "o envelope tem de sair byte a byte como entrou")
            End Using
        End Sub

        ''' <summary>
        ''' <b>Não há mensagem <c>system</c>.</b>
        '''
        ''' O envelope já carrega <c>instrucaoDoSistema</c> dentro dele, e foi
        ''' para isso que ela foi posta lá. Um papel <c>system</c> separado
        ''' seria texto que a capability não cobre entrando no pedido — e o
        ''' lugar exato onde alguém acrescentaria uma instrução com autoridade.
        ''' </summary>
        <TestMethod>
        Public Sub NAO_ha_mensagem_SYSTEM()
            Using doc = Corpo()
                For Each m In doc.RootElement.GetProperty("messages").EnumerateArray()
                    Assert.AreNotEqual("system", m.GetProperty("role").GetString())
                Next
                ' E a instrucao do Iris continua la — dentro do envelope.
                Dim conteudo = doc.RootElement.GetProperty("messages")(0).
                               GetProperty("content").GetString()
                ' Trecho ASCII de proposito: o Utf8JsonWriter escapa nao-ASCII
                ' (ç e companhia), entao a frase acentuada nao aparece
                ' literal nos bytes.
                StringAssert.Contains(conteudo, "DADO a ser processado")
            End Using
        End Sub

        ''' <summary>
        ''' <b>O corpo tem exatamente estas chaves, e nenhuma a mais.</b>
        '''
        ''' É o teste que acusa quando alguém acrescentar um campo — um
        ''' identificador de usuário, um nome de sessão, um metadado de
        ''' telemetria. Todos parecem inofensivos, e todos são conteúdo que sai
        ''' desta máquina sem ter passado pela autorização.
        ''' </summary>
        <TestMethod>
        Public Sub O_corpo_tem_EXATAMENTE_as_chaves_esperadas()
            Using doc = Corpo()
                Dim chaves = doc.RootElement.EnumerateObject().
                             Select(Function(p) p.Name).OrderBy(Function(n) n).ToArray()
                ' "usage" ENTROU DE PROPOSITO: sem pedir, o OpenRouter nao
                ' devolve o custo real da chamada, e a faixa nao teria como
                ' dizer quanto a acao custou.
                '
                ' Ele mora no Preparar, e nao em outro lugar, porque o corpo e
                ' montado ANTES de a capability ser emitida -- entao o que a
                ' autorizacao cobre ja inclui este campo. Um campo enfiado
                ' depois seria exatamente o defeito que a separacao
                ' Preparar/Enviar existe para impedir.
                CollectionAssert.AreEqual(
                    {"messages", "model", "provider", "temperature", "usage"}, chaves)

                Dim m = doc.RootElement.GetProperty("messages")(0).EnumerateObject().
                        Select(Function(p) p.Name).OrderBy(Function(n) n).ToArray()
                CollectionAssert.AreEqual({"content", "role"}, m)
            End Using
        End Sub

        ''' <summary><b>O modelo vem da ativação.</b></summary>
        <TestMethod>
        Public Sub O_modelo_vem_da_ATIVACAO()
            Using doc = Corpo(Provedor(modelo:="anthropic/claude-sonnet-5"))
                Assert.AreEqual("anthropic/claude-sonnet-5",
                                doc.RootElement.GetProperty("model").GetString())
            End Using
        End Sub

        ' ==================================================================
        ' O bloco provider é o que torna a ativação verdadeira

        ''' <summary>
        ''' <b><c>allow_fallbacks</c> é sempre falso.</b>
        '''
        ''' Sem ele, o OpenRouter cai para um provedor fora da lista quando o
        ''' fixado não atende — e a ativação estaria declarando um conjunto que
        ''' o pedido não impõe. Falhar é o desfecho certo; <b>degradar em
        ''' silêncio</b> é o que não pode acontecer.
        ''' </summary>
        <TestMethod>
        Public Sub Allow_fallbacks_e_SEMPRE_falso()
            ' Lista vazia saiu do laco: ela deixou de ser construivel, e o
            ' teste ao lado prova que o adaptador a recusa.
            For Each zero In {True, False}
                For Each slugs In {New String() {"google"},
                                   New String() {"google", "google-vertex"}}
                    Using doc = Corpo(Provedor(zero, slugs))
                        Assert.IsFalse(doc.RootElement.GetProperty("provider").
                                       GetProperty("allow_fallbacks").GetBoolean(),
                                       $"zero={zero} slugs={slugs.Length}")
                    End Using
                Next
            Next
        End Sub

        ''' <summary>
        ''' <b>Retenção zero declarada vira retenção zero exigida no pedido.</b>
        '''
        ''' Este é o teste que impede a ativação de virar frase decorativa.
        ''' </summary>
        <TestMethod>
        Public Sub Retencao_zero_DECLARADA_vira_EXIGIDA_no_pedido()
            Using doc = Corpo(Provedor(zero:=True))
                Dim prov = doc.RootElement.GetProperty("provider")
                Assert.IsTrue(prov.GetProperty("zdr").GetBoolean())
                Assert.AreEqual("deny", prov.GetProperty("data_collection").GetString())
            End Using
        End Sub

        ''' <summary>
        ''' Controle: ativação que <b>não</b> exige retenção zero não manda os
        ''' campos.
        '''
        ''' Sem ele, um codec que mandasse <c>zdr</c> sempre passaria no teste de
        ''' cima — e a prova de que a ativação controla o pedido seria vazia.
        ''' </summary>
        <TestMethod>
        Public Sub Controle_sem_exigir_retencao_zero_os_campos_NAO_vao()
            Using doc = Corpo(Provedor(zero:=False))
                Dim prov = doc.RootElement.GetProperty("provider")
                Dim ignorado As JsonElement = Nothing
                Assert.IsFalse(prov.TryGetProperty("zdr", ignorado))
                Assert.IsFalse(prov.TryGetProperty("data_collection", ignorado))
            End Using
        End Sub

        ''' <summary>
        ''' <b>Os provedores permitidos viram <c>only</c>.</b>
        '''
        ''' Gateway roteia: o endpoint autorizado é o dele, e quem processa o
        ''' conteúdo pode ser outro. Sem <c>only</c>, a autorização fala do
        ''' portão e não de quem vê a mensagem.
        ''' </summary>
        <TestMethod>
        Public Sub Os_provedores_permitidos_viram_ONLY()
            Using doc = Corpo(Provedor(slugs:={"google", "google-vertex"}))
                Dim only = doc.RootElement.GetProperty("provider").GetProperty("only")
                CollectionAssert.AreEqual({"google", "google-vertex"},
                                          only.EnumerateArray().
                                          Select(Function(e) e.GetString()).ToArray())
            End Using
        End Sub

        ''' <summary>
        ''' <b>Ativação sem provedor subjacente não vira adaptador.</b>
        '''
        ''' Este teste dizia o contrário: que lista vazia apenas omitia
        ''' <c>only</c>, e que omitir era o jeito de dizer "qualquer um". Isso
        ''' era a falha aberta — a ausência do campo ampliava a autorização.
        '''
        ''' O carregador passou a exigir o campo e a recusar lista vazia, e o
        ''' adaptador recusa também: barreira que só existe a montante some
        ''' quando alguém constrói o registro por outro caminho, e dentro deste
        ''' assembly isso é uma linha de código.
        ''' </summary>
        <TestMethod>
        Public Sub Ativacao_SEM_provedor_subjacente_nao_vira_adaptador()
            Assert.ThrowsException(Of ArgumentException)(
                Function() Provedor(slugs:=Array.Empty(Of String)()))
        End Sub

        ''' <summary>
        ''' <b>Ativação de outro provedor não vira adaptador do OpenRouter.</b>
        '''
        ''' Sem isto, um arquivo declarando outro provedor — com outro endereço —
        ''' receberia o protocolo do OpenRouter <b>e a credencial guardada sob o
        ''' nome dele</b>. Conteúdo da caixa indo para um endereço que alguém
        ''' escreveu num arquivo.
        ''' </summary>
        <TestMethod>
        Public Sub Ativacao_de_OUTRO_provedor_nao_vira_adaptador()
            Dim outra = New ActivationRecord(
                "ativacao-1", 1, "quem", Agora.AddDays(-1),
                "provedor-qualquer", "https://outro.invalido/v1", "modelo",
                "local", "sem retenção",
                {AssistOperation.Resumir}, {Pasta}, Array.Empty(Of String)(),
                {LabelReadingKind.Absent}, {0},
                ate:=Agora.AddDays(30), provedoresPermitidos:={"google"})

            Assert.IsFalse(OpenRouterAssistantProvider.Atende(outra))
            Assert.ThrowsException(Of ArgumentException)(
                Function() New OpenRouterAssistantProvider(outra, Function() "chave"))
        End Sub

        ''' <summary>
        ''' Controle: a ativação do OpenRouter <b>é</b> aceita, e a comparação
        ''' não é sensível a maiúsculas.
        ''' </summary>
        <TestMethod>
        Public Sub Controle_a_ativacao_do_OPENROUTER_e_aceita()
            Assert.IsTrue(OpenRouterAssistantProvider.Atende(Ativacao()))
            Assert.IsNotNull(Provedor())
        End Sub

        ''' <summary><b>Temperatura zero: o mesmo e-mail dá o mesmo resumo.</b></summary>
        <TestMethod>
        Public Sub A_temperatura_e_ZERO()
            Using doc = Corpo()
                Assert.AreEqual(0, doc.RootElement.GetProperty("temperature").GetInt32())
            End Using
        End Sub

        ' ==================================================================
        ' Preparar recusa

        ''' <summary>
        ''' <b>Envelope vazio ou ausente não vira corpo.</b>
        '''
        ''' <c>Nothing</c> aqui faz o transmissor parar antes de emitir
        ''' capability — nada sai, e isso se sabe.
        ''' </summary>
        <TestMethod>
        Public Sub Envelope_VAZIO_nao_vira_corpo()
            Assert.IsNull(Provedor().Preparar(Nothing))
            Assert.IsNull(Provedor().Preparar(Array.Empty(Of Byte)()))
        End Sub

        ''' <summary>
        ''' <b>Bytes que não são UTF-8 válido não viram corpo.</b>
        '''
        ''' O envelope é UTF-8 por construção. Se não for, alguém o produziu por
        ''' outro caminho — e substituir os bytes inválidos pelo caractere de
        ''' substituição mandaria um conteúdo que ninguém escreveu.
        ''' </summary>
        <TestMethod>
        Public Sub Bytes_que_nao_sao_UTF8_nao_viram_corpo()
            Assert.IsNull(Provedor().Preparar({&HFFUS, &HFEUS, &H80US, &H80US}.
                                              Select(Function(b) CByte(b)).ToArray()))
        End Sub

        ' ==================================================================
        ' A leitura da resposta

        ''' <summary>
        ''' Controle do grupo: a resposta bem formada dá o texto.
        '''
        ''' Sem ele, um extrator que devolvesse <c>Nothing</c> sempre passaria em
        ''' todos os testes de recusa abaixo.
        ''' </summary>
        <TestMethod>
        Public Sub Controle_resposta_BEM_FORMADA_da_o_texto()
            Assert.AreEqual("o resumo", Extrair(
                "{""choices"":[{""message"":{""role"":""assistant"",""content"":""o resumo""}}]}"))
        End Sub

        ''' <summary>
        ''' <b>Resposta que não dá para ler não vira texto.</b>
        '''
        ''' Cada um destes chega com HTTP 200: o conteúdo <b>saiu</b>. O desfecho
        ''' é <c>RespostaIlegivel</c>, e não "não começou" — dizer que não
        ''' começou seria mentir sobre a única coisa que importa.
        ''' </summary>
        <TestMethod>
        Public Sub Resposta_ILEGIVEL_nao_vira_texto()
            For Each ruim In {"",
                              "nao e json",
                              "{}",
                              "{""choices"":[]}",
                              "{""choices"":[{}]}",
                              "{""choices"":[{""message"":{}}]}",
                              "{""choices"":[{""message"":{""content"":null}}]}",
                              "{""choices"":[{""message"":{""content"":42}}]}",
                              "{""choices"":[{""message"":{""content"":[""a""]}}]}",
                              "{""choices"":""texto""}"}
                Assert.IsNull(Extrair(ruim), $"devia ser ilegivel: {ruim}")
            Next
        End Sub

        ''' <summary>
        ''' <b>Resposta vazia é resposta, e não ilegibilidade.</b>
        '''
        ''' <c>content: ""</c> quer dizer que o provedor não teve o que dizer —
        ''' e isso tem tratamento próprio na faixa. Confundir com ilegível
        ''' levaria a investigação para o lado errado.
        ''' </summary>
        <TestMethod>
        Public Sub Resposta_VAZIA_e_resposta()
            Assert.AreEqual("", Extrair("{""choices"":[{""message"":{""content"":""""}}]}"))
        End Sub

        ''' <summary>
        ''' Chama o extrator privado. Ele é privado de propósito — ninguém fora
        ''' do adaptador deve interpretar resposta de provedor.
        ''' </summary>
        Private Shared Function Extrair(bruto As String) As String
            Dim m = GetType(OpenRouterAssistantProvider).GetMethod(
                "Extrair", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static)
            Assert.IsNotNull(m, "o extrator mudou de nome")
            Return CStr(m.Invoke(Nothing, {bruto}))
        End Function

        ' ==============================================================
        ' A CONTABILIDADE
        '
        ' O extrator diz por escrito que nada alem de choices[0].message.content
        ' e lido, porque o resto e dado de fora. A regra continua valendo PARA
        ' TEXTO. O que entrou foram os NUMEROS de usage -- pelo mesmo motivo que
        ' o codigo HTTP entrou no diario: numero nao ecoa o e-mail de volta.

        ''' <summary>
        ''' <b>Custo e tokens atravessam.</b>
        ''' </summary>
        <TestMethod>
        Public Sub A_conta_do_provedor_ATRAVESSA()
            Dim r = Contabilidade(
                "{""choices"":[{""message"":{""content"":""oi""}}]," &
                """usage"":{""total_tokens"":1234,""cost"":0.00042}}")

            Assert.AreEqual(1234, r.Tokens)
            Assert.AreEqual(0.00042D, r.Custo)
        End Sub

        ''' <summary>
        ''' <b>Resposta sem <c>usage</c> não vira zero.</b>
        '''
        ''' Nem todo provedor conta, e nenhum é obrigado. Zero é uma afirmação:
        ''' dizer que custou nada quando ninguém contou é inventar a conta.
        ''' </summary>
        <TestMethod>
        Public Sub Sem_usage_a_conta_e_AUSENTE_e_nao_zero()
            Dim r = Contabilidade("{""choices"":[{""message"":{""content"":""oi""}}]}")

            Assert.IsFalse(r.Tokens.HasValue)
            Assert.IsFalse(r.Custo.HasValue)
        End Sub

        ''' <summary>
        ''' <b>Número disfarçado de texto não conta.</b>
        '''
        ''' <c>"cost": "1000"</c> é o provedor mandando uma string onde deveria
        ''' haver número. Aceitar seria deixar o outro lado escolher o que
        ''' aparece na ficha por um caminho que não é o dos números.
        ''' </summary>
        <TestMethod>
        Public Sub Custo_em_TEXTO_nao_conta()
            Dim r = Contabilidade(
                "{""usage"":{""total_tokens"":""muitos"",""cost"":""1000""}}")

            Assert.IsFalse(r.Tokens.HasValue)
            Assert.IsFalse(r.Custo.HasValue)
        End Sub

        ''' <summary>
        ''' <b>Custo negativo não entra.</b>
        '''
        ''' Não descreve custo de nada, e viraria "US$ -5,0000" na tela. Mesma
        ''' regra do código HTTP fora da faixa: número estranho num campo de
        ''' diagnóstico não pode virar afirmação.
        ''' </summary>
        <TestMethod>
        Public Sub Custo_NEGATIVO_nao_entra()
            Dim r = New ProviderOutcome(ProviderStatus.Respondeu, "oi", 200, -5D, -3)

            Assert.IsFalse(r.Custo.HasValue)
            Assert.IsFalse(r.Tokens.HasValue)
        End Sub

        ''' <summary>
        ''' <b>O <c>provider</c> da resposta continua NÃO sendo lido.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' É o controle que guarda a fronteira. <c>provider</c> é <b>texto
        ''' escolhido pelo outro lado</b>, e mostrá-lo na tela seria dado de fora
        ''' atravessando até o usuário — a §29.5 diz onde isso para.
        '''
        ''' Quem diz o agente e o modelo na faixa é a <b>ativação</b>, que o
        ''' usuário assinou. Este teste falha no dia em que alguém achar mais
        ''' fácil ler da resposta.
        ''' </summary>
        <TestMethod>
        Public Sub O_provider_da_RESPOSTA_nao_e_lido()
            Dim bruto = "{""choices"":[{""message"":{""content"":""oi""}}]," &
                        """provider"":""AGENTE-QUE-NAO-DEVIA-APARECER""," &
                        """model"":""MODELO-QUE-NAO-DEVIA-APARECER""," &
                        """usage"":{""total_tokens"":7,""cost"":0.1}}"

            Assert.AreEqual("oi", Extrair(bruto),
                "so o content, e nada mais, vira texto")

            Dim r = Contabilidade(bruto)
            Assert.AreEqual(7, r.Tokens, "os numeros entram")
            Assert.AreEqual(0.1D, r.Custo)
            ' E nao ha por onde o texto entrar: Contabilidade devolve dois
            ' numeros, e o tipo de retorno e a barreira.
        End Sub

        ''' <summary>Chama a contabilidade privada, pelo mesmo motivo do extrator.</summary>
        Private Shared Function Contabilidade(bruto As String) _
                                As (Custo As Decimal?, Tokens As Integer?)
            Dim m = GetType(OpenRouterAssistantProvider).GetMethod(
                "Contabilidade", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static)
            Assert.IsNotNull(m, "a contabilidade mudou de nome")
            Return CType(m.Invoke(Nothing, {bruto}), (Custo As Decimal?, Tokens As Integer?))
        End Function

    End Class

End Namespace

Imports System.Linq
Imports System.Text
Imports System.Threading
Imports Iris.Assist
Imports Iris.Integration.Assist.Http
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O transporte HTTP, contra um servidor de verdade — o 3.4.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE SERVIDOR E NÃO HANDLER FALSO</b>
'''
''' Um <c>HttpMessageHandler</c> falso provaria o código <i>em volta</i> do
''' <c>HttpClient</c>, e as propriedades que este marco precisa provar são
''' justamente as <b>dele</b>: redirect não seguido, timeout, cancelamento,
''' teto de resposta.
'''
''' Aqui há socket, requisição e resposta — tudo em <c>127.0.0.1</c>, e nenhum
''' byte sai da máquina.
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTE ARQUIVO NÃO PROVA</b>
'''
''' Nada sobre fornecedor real: nem que o contrato dele aceita estes bytes, nem
''' a semântica de streaming, nem os códigos de erro, nem os limites efetivos.
''' Isso é <b>pendência declarada</b>, e construir agora seria inventar
''' requisito.
''' </summary>
<TestClass>
Public Class TransporteTests

    Private Shared Function Destino(endereco As String) As AssistDestination
        Return New AssistDestination("provedor-de-teste", endereco, "modelo-de-teste")
    End Function

    ''' <summary>
    ''' O provedor apontado para o servidor falso. A exceção de loopback é
    ''' ligada <b>aqui</b>, e o teste do padrão prova que ela não está ligada
    ''' por conta própria.
    ''' </summary>
    Private Shared Function Provedor(s As ServidorFalso,
                                     Optional chave As String = "segredo-abc",
                                     Optional limite As TimeSpan = Nothing) _
                                     As HttpAssistantProvider
        Return New HttpAssistantProvider(Destino(s.Endereco), Function() chave,
                                         "Authorization",
                                         If(limite = Nothing, TimeSpan.FromSeconds(30), limite),
                                         permitirLoopbackSemTls:=True)
    End Function

    Private Shared Function Bytes(t As String) As Byte()
        Return Encoding.UTF8.GetBytes(t)
    End Function

    ''' <summary>
    ''' Espera uma condição do servidor, com teto.
    '''
    ''' O servidor falso atende numa thread própria. Sem esperar, um teste sobre
    ''' timeout mediria a corrida entre o cliente desistir e o servidor
    ''' registrar — e não a propriedade que interessa.
    ''' </summary>
    Private Shared Function Esperar(cond As Func(Of Boolean)) As Boolean
        Dim ate = DateTime.UtcNow.AddSeconds(5)
        While DateTime.UtcNow < ate
            If cond() Then Return True
            Thread.Sleep(20)
        End While
        Return cond()
    End Function

    ' ==================================================================
    ' Controle positivo

    ''' <summary>
    ''' <b>Os bytes que saem são exatamente os que entraram.</b>
    '''
    ''' É a garantia do 3.2 chegando até o fio: o transporte não reserializa,
    ''' não embrulha, não acrescenta. Se acrescentasse, o que foi autorizado
    ''' deixaria de ser o que sai.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Os_bytes_chegam_EXATAMENTE_como_sairam()
        Using s As New ServidorFalso()
            Using p = Provedor(s)
                Dim carga = Bytes("{""esquema"":""iris.assist.v1"",""x"":""ação 🙂""}")

                Dim r = p.Enviar(carga, CancellationToken.None)

                Assert.AreEqual(ProviderStatus.Respondeu, r.Status)
                Assert.AreEqual("resposta do modelo", r.Texto)
                CollectionAssert.AreEqual(carga, s.Ultimo.Corpo)
                Assert.AreEqual("POST", s.Ultimo.Metodo)
            End Using
        End Using
    End Sub

    ' ==================================================================
    ' As regras da §30

    ''' <summary>
    ''' <b>Redirect NÃO é seguido.</b>
    '''
    ''' Um 302 mandaria o corpo para um endereço que ninguém autorizou — e a
    ''' capability se prendeu ao endpoint, não a "onde ele apontar".
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Redirect_NAO_e_seguido()
        Using destino As New ServidorFalso()
            Using origem As New ServidorFalso()
                origem.RedirecionarPara = destino.Endereco

                Using p = Provedor(origem)
                    Dim r = p.Enviar(Bytes("carga"), CancellationToken.None)

                    Assert.AreEqual(ProviderStatus.Recusou, r.Status,
                        "302 nao e sucesso, e nao pode ser seguido")
                    Assert.AreEqual(302, r.Codigo.Value)
                    Assert.AreEqual(0, destino.Recebidos.Count,
                        "o conteudo NAO pode ter ido para o endereco do redirect")
                End Using
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' <b>Sem HTTPS, nem começa.</b>
    '''
    ''' E o ponto é o <b>padrão</b>: a exceção de loopback existe para o
    ''' servidor falso e tem de ser ligada explicitamente. Um provedor
    ''' construído do jeito normal recusa <c>http://</c>.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Sem_HTTPS_nem_comeca()
        Using s As New ServidorFalso()
            ' Sem permitirLoopbackSemTls — como a producao constroi.
            Using p As New HttpAssistantProvider(Destino(s.Endereco), Function() "chave")
                Dim r = p.Enviar(Bytes("carga"), CancellationToken.None)

                Assert.AreEqual(ProviderStatus.NaoComecou, r.Status)
                Assert.IsFalse(r.PodeTerChegado, "nada saiu, e isso se sabe")
                Assert.AreEqual(0, s.Recebidos.Count)
            End Using
        End Using
    End Sub

    ''' <summary>Sem credencial, também não começa.</summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Sem_credencial_nem_comeca()
        Using s As New ServidorFalso()
            Using p = Provedor(s, chave:="")
                Dim r = p.Enviar(Bytes("carga"), CancellationToken.None)

                Assert.AreEqual(ProviderStatus.NaoComecou, r.Status)
                Assert.AreEqual(0, s.Recebidos.Count)
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' <b>Timeout não é "não chegou".</b>
    '''
    ''' O servidor recebeu o corpo e demorou a responder. Do lado de cá não há
    ''' como distinguir isso de "não recebeu" — e é por isso que o desfecho
    ''' admite que pode ter chegado.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Timeout_ADMITE_que_pode_ter_chegado()
        Using s As New ServidorFalso()
            Using p = Provedor(s, limite:=TimeSpan.FromSeconds(2))
                ' AQUECIMENTO. A primeira chamada paga o custo de subir a
                ' conexao, e sob a suite inteira em paralelo esse custo passou
                ' dos 400 ms que o teste dava — entao ele media a corrida entre
                ' o cliente desistir e o TCP se estabelecer, e nao a
                ' propriedade. Com a conexao ja de pe, o que sobra e a demora
                ' do servidor.
                Assert.AreEqual(ProviderStatus.Respondeu,
                                p.Enviar(Bytes("aquecendo"), CancellationToken.None).Status)
                Assert.AreEqual(1, s.Recebidos.Count)

                s.Demora = TimeSpan.FromSeconds(10)
                Dim r = p.Enviar(Bytes("carga"), CancellationToken.None)

                Assert.AreEqual(ProviderStatus.Timeout, r.Status)
                Assert.IsTrue(r.PodeTerChegado,
                    "o servidor JA tinha o corpo — dizer que nao chegou seria mentira")

                ' E ele recebeu mesmo. A espera existe porque o servidor
                ' registra numa thread propria.
                Assert.IsTrue(Esperar(Function() s.Recebidos.Count = 2),
                              "o corpo TINHA de ter chegado ao servidor")
            End Using
        End Using
    End Sub

    ''' <summary>Cancelamento depois de começar também admite.</summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Cancelamento_ADMITE_que_pode_ter_chegado()
        Using s As New ServidorFalso()
            Using p = Provedor(s)
                ' Aquecimento, pelo mesmo motivo do teste de timeout.
                p.Enviar(Bytes("aquecendo"), CancellationToken.None)

                s.Demora = TimeSpan.FromSeconds(10)
                Using cts As New CancellationTokenSource(TimeSpan.FromSeconds(2))
                    Dim r = p.Enviar(Bytes("carga"), cts.Token)

                    Assert.AreEqual(ProviderStatus.Cancelado, r.Status)
                    Assert.IsTrue(r.PodeTerChegado)
                End Using
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Erro do provedor devolve o <b>código</b>, e o corpo do erro fica lá.
    '''
    ''' Corpo de erro pode <b>ecoar o que foi enviado</b>, e ele atravessaria a
    ''' fronteira até o diário e a tela. Só o número passa.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Erro_devolve_o_CODIGO_e_nao_o_corpo()
        Using s As New ServidorFalso()
            s.Codigo = 400
            s.Corpo = "ECO-DO-CONTEUDO-QUE-FOI-ENVIADO"

            Using p = Provedor(s)
                Dim r = p.Enviar(Bytes("carga"), CancellationToken.None)

                Assert.AreEqual(ProviderStatus.Recusou, r.Status)
                Assert.AreEqual(400, r.Codigo.Value)
                Assert.AreEqual("", r.Texto, "o corpo do erro NAO atravessa")
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' <b>Resposta maior que o teto NÃO vira sucesso truncado.</b>
    '''
    ''' Ela era cortada e devolvida como <c>Respondeu</c> — o que apresenta uma
    ''' resposta <b>parcial</b> como se fosse completa, e um resumo cortado no
    ''' meio parece um resumo.
    '''
    ''' Agora tem estado próprio, e sem texto.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Resposta_maior_que_o_teto_NAO_e_sucesso_truncado()
        Using s As New ServidorFalso()
            s.TamanhoDaResposta = HttpAssistantProvider.MaxResposta + 1

            Using p = Provedor(s)
                Dim r = p.Enviar(Bytes("carga"), CancellationToken.None)

                Assert.AreEqual(ProviderStatus.RespostaGrandeDemais, r.Status)
                Assert.AreEqual("", r.Texto, "meia resposta nao pode ser apresentada")
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' E o contraponto: resposta <b>exatamente</b> no teto passa.
    '''
    ''' Sem ele, um leitor que recusasse por engano ao encher o buffer passaria
    ''' no teste de cima e recusaria toda resposta grande e legítima.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Resposta_EXATAMENTE_no_teto_passa()
        Using s As New ServidorFalso()
            s.TamanhoDaResposta = HttpAssistantProvider.MaxResposta

            Using p = Provedor(s)
                Dim r = p.Enviar(Bytes("carga"), CancellationToken.None)

                Assert.AreEqual(ProviderStatus.Respondeu, r.Status)
                Assert.AreEqual(HttpAssistantProvider.MaxResposta, r.Texto.Length)
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' <b>Uma chamada, e só uma.</b>
    '''
    ''' Egress é mutação do mundo: repetir depois de começar manda o mesmo
    ''' conteúdo duas vezes. É a regra "leitura tem retry, mutação não" do
    ''' CLAUDE.md, e aqui vale mesmo quando o servidor responde erro — que é
    ''' justamente quando um cliente comum tentaria de novo.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Nenhum_RETRY_nem_com_erro_do_servidor()
        Using s As New ServidorFalso()
            s.Codigo = 503

            Using p = Provedor(s)
                p.Enviar(Bytes("carga"), CancellationToken.None)

                Assert.AreEqual(1, s.Recebidos.Count,
                    "503 e o codigo que mais convida a repetir, e nao se repete")
            End Using
        End Using
    End Sub

    ' ==================================================================
    ' A credencial

    ''' <summary>
    ''' A credencial vai no <b>cabeçalho</b>, e não na URL.
    '''
    ''' Query string aparece em log de servidor, em proxy corporativo e em
    ''' histórico — lugares que ninguém controla.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub A_credencial_vai_no_CABECALHO_e_nao_na_URL()
        Using s As New ServidorFalso()
            Using p = Provedor(s, chave:="segredo-abc")
                p.Enviar(Bytes("carga"), CancellationToken.None)

                ' COM O ESQUEMA, e exatamente uma vez. A chave crua no
                ' cabecalho obrigaria quem configura a colar "Bearer " dentro
                ' do segredo — e o dia em que alguem esquecesse, o Iris mandaria
                ' um Authorization malformado sem ninguem notar.
                Assert.AreEqual("Bearer segredo-abc", s.Ultimo.Cabecalhos("Authorization"))
                Assert.IsFalse(s.Ultimo.Caminho.Contains("segredo"),
                               "credencial em query string vaza em log e proxy")
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' <b>A credencial é lida na hora, e não guardada.</b>
    '''
    ''' Um campo com a credencial dentro é um campo que vaza em dump, em log de
    ''' exceção e em serialização acidental. Aqui o teste troca o valor que a
    ''' função devolve entre uma chamada e outra, e a segunda usa o novo.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub A_credencial_e_lida_na_HORA()
        Using s As New ServidorFalso()
            Dim atual = "primeira"
            Using p As New HttpAssistantProvider(Destino(s.Endereco), Function() atual,
                                                 "Authorization", TimeSpan.FromSeconds(30),
                                                 permitirLoopbackSemTls:=True)
                p.Enviar(Bytes("a"), CancellationToken.None)
                atual = "segunda"
                p.Enviar(Bytes("b"), CancellationToken.None)

                Assert.AreEqual("Bearer primeira", s.Recebidos(0).Cabecalhos("Authorization"))
                Assert.AreEqual("Bearer segunda", s.Recebidos(1).Cabecalhos("Authorization"))
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' <b>E ela não aparece no desfecho.</b> Nem no texto, nem no código.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub A_credencial_NAO_aparece_no_desfecho()
        Const chave = "SEGREDO-QUE-NAO-PODE-VAZAR-99"
        Using s As New ServidorFalso()
            s.Codigo = 401

            Using p = Provedor(s, chave:=chave)
                Dim r = p.Enviar(Bytes("carga"), CancellationToken.None)

                Assert.IsFalse(r.Texto.Contains(chave))
                Assert.IsFalse($"{r.Status}{r.Codigo}".Contains(chave))
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Servidor morto: conexão caiu, e a mensagem da exceção <b>não</b>
    ''' atravessa — ela carrega host, porta e às vezes mais.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Sub Servidor_morto_da_ConexaoCaiu_sem_detalhe()
        Dim endereco As String
        Using s As New ServidorFalso()
            endereco = s.Endereco
        End Using

        Using p As New HttpAssistantProvider(Destino(endereco), Function() "chave",
                                             "Authorization", TimeSpan.FromSeconds(30),
                                             permitirLoopbackSemTls:=True)
            Dim r = p.Enviar(Bytes("carga"), CancellationToken.None)

            Assert.AreEqual(ProviderStatus.ConexaoCaiu, r.Status)
            Assert.AreEqual("", r.Texto)
            Assert.IsTrue(r.PodeTerChegado,
                "conexao caindo NAO prova que nada saiu — o socket pode ter escrito")
        End Using
    End Sub

    ' ==================================================================

    ''' <summary>
    ''' O provedor da produção: recusa, e diz que <b>não começou</b>.
    '''
    ''' Existe em vez de <c>Nothing</c> porque <c>Nothing</c> vira
    ''' <c>NullReferenceException</c> em algum caminho esquecido, e "explodiu" e
    ''' "recusou por decisão" não são a mesma coisa para quem lê depois.
    ''' </summary>
    <TestMethod>
    Public Sub O_provedor_da_producao_recusa_por_DECISAO()
        Dim r = New AssistenteIndisponivel().Enviar(Bytes("carga"), CancellationToken.None)

        Assert.AreEqual(ProviderStatus.NaoComecou, r.Status)
        Assert.IsFalse(r.PodeTerChegado)
    End Sub

End Class

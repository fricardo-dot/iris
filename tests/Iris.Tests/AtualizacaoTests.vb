Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports Iris.App.ViewModels
Imports Iris.Update
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O caminho de atualização, do manifesto assinado até o arquivo no disco.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTES TESTES PRECISAM PROVAR</b>
'''
''' Um atualizador é a única porta pela qual código novo entra num programa que
''' lê o e-mail do dono. As propriedades que importam não são "funciona quando
''' está tudo certo" — são as recusas:
'''
''' <list type="bullet">
''' <item>manifesto assinado por outra chave é recusado;</item>
''' <item>manifesto alterado depois de assinado é recusado;</item>
''' <item>pacote cujo hash não bate com o manifesto é recusado <b>e não fica no
'''       disco</b>;</item>
''' <item>versão que não sobe não é oferecida;</item>
''' <item>"a assinatura não confere" e "não deu para saber" são desfechos
'''       <b>diferentes</b>, porque um é ataque e o outro é rede.</item>
''' </list>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE UM HANDLER FALSO, E NÃO O ServidorFalso</b>
'''
''' O <c>ServidorFalso</c> é HTTP em 127.0.0.1, e <c>ProcuraDeVersao</c> recusa
''' qualquer coisa que não seja <c>https</c> — de propósito. Um servidor TLS de
''' mentira exigiria certificado, e o que estes testes precisam exercitar é a
''' <i>decisão</i>, não o transporte: o transporte já tem os testes dele em
''' <c>TransporteTests</c>.
''' </summary>
<TestClass>
Public Class AtualizacaoTests

    ' ==================================================================
    ' Aparato
    ' ==================================================================

    Private Const Base As String = "https://exemplo.invalido/iris"

    ''' <summary>
    ''' Um par novo a cada chamada. Chave fixa em constante economizaria
    ''' milissegundos e faria dois testes compartilharem a mesma chave — e o
    ''' teste de "assinada por outra chave" precisa de duas de verdade.
    ''' </summary>
    Private Shared Function ParNovo() As ECDsa
        Return ECDsa.Create(ECCurve.NamedCurves.nistP256)
    End Function

    Private Shared Function Publica(deQuem As ECDsa) As Byte()
        Return deQuem.ExportSubjectPublicKeyInfo()
    End Function

    Private Shared Function Assinar(dados As Byte(), comQual As ECDsa) As Byte()
        Return comQual.SignData(dados, HashAlgorithmName.SHA256)
    End Function

    ''' <summary>
    ''' O JSON de um manifesto, montado à mão. <b>À mão de propósito</b>: montá-lo
    ''' com o mesmo código que o lê provaria que os dois concordam, e não que o
    ''' leitor entende o que o script de publicação escreve.
    ''' </summary>
    Private Shared Function MontarJson(numero As String,
                                       Optional aonde As String = Base & "/Iris.exe",
                                       Optional hash As String = Nothing,
                                       Optional quantos As Long = 40_000_000,
                                       Optional oQueMudou As String = "coisas") As Byte()
        Dim oHash = If(hash, New String("a"c, 64))
        Dim texto =
            "{" &
            """versao"": """ & numero & """," &
            """publicada"": ""2026-09-02T10:00:00.0000000Z""," &
            """notas"": """ & oQueMudou & """," &
            """endereco"": """ & aonde & """," &
            """sha256"": """ & oHash & """," &
            """bytes"": " & quantos.ToString(Globalization.CultureInfo.InvariantCulture) &
            "}"
        Return New UTF8Encoding(False).GetBytes(texto)
    End Function

    ''' <summary>
    ''' Responde o que o teste mandar, por caminho. Sem socket: o que está sob
    ''' teste é a decisão sobre o que chegou.
    ''' </summary>
    Private NotInheritable Class RespostasDeVersao
        Inherits HttpMessageHandler

        Friend ReadOnly Corpos As New Dictionary(Of String, Byte())(StringComparer.Ordinal)

        ''' <summary>
        ''' Respostas cujo corpo é um fluxo, e não um vetor. Um vetor tem fim; o
        ''' que os testes de teto precisam é de um que não tenha.
        ''' </summary>
        Friend ReadOnly Fluxos As New Dictionary(Of String, Stream)(StringComparer.Ordinal)

        Friend ReadOnly Pedidos As New List(Of String)()

        ''' <summary>Quando preenchido, toda requisição estoura com isto.</summary>
        Friend Property Explodir As Exception

        ''' <summary>
        ''' Quando preenchido, a resposta diz ter terminado NESTE endereço — que
        ''' é o que um redirecionamento faz. Serve para provar
        ''' <c>ExigirHttpsAteOFim</c>.
        ''' </summary>
        Friend Property EnderecoFinal As Uri

        ''' <summary>Devolve resposta sem <c>RequestMessage</c>.</summary>
        Friend Property EsquecerODePedido As Boolean

        Protected Overrides Function SendAsync(pedido As HttpRequestMessage,
                                               ct As CancellationToken) _
                                               As Task(Of HttpResponseMessage)
            Dim onde = pedido.RequestUri.ToString()
            SyncLock Pedidos
                Pedidos.Add(onde)
            End SyncLock

            If Explodir IsNot Nothing Then Throw Explodir

            Dim fluxo As Stream = Nothing
            If Fluxos.TryGetValue(onde, fluxo) Then
                Return Task.FromResult(Responder(pedido, HttpStatusCode.OK,
                                                 New StreamContent(fluxo)))
            End If

            Dim corpo As Byte() = Nothing
            If Not Corpos.TryGetValue(onde, corpo) Then
                Return Task.FromResult(Responder(pedido, HttpStatusCode.NotFound, Nothing))
            End If

            Return Task.FromResult(Responder(pedido, HttpStatusCode.OK,
                                             New ByteArrayContent(corpo)))
        End Function

        ''' <summary>
        ''' <b>Toda resposta carrega o pedido que a originou.</b>
        '''
        ''' Um handler de verdade preenche <c>RequestMessage</c> — é de lá que
        ''' sai o endereço FINAL, depois de redirecionamentos, e é o que
        ''' <c>ExigirHttpsAteOFim</c> confere. O dublê não preenchia, e a
        ''' produção tratava isso como "não sei o endereço final" e deixava
        ''' passar.
        '''
        ''' Quando a produção passou a recusar o que não sabe, sete testes caíram
        ''' de uma vez — e a causa era o dublê, não a produção. Um dublê infiel
        ''' num ponto vira, mais cedo ou mais tarde, uma conferência que ninguém
        ''' pode apertar.
        ''' </summary>
        Private Function Responder(pedido As HttpRequestMessage,
                                          codigo As HttpStatusCode,
                                          corpo As HttpContent) As HttpResponseMessage
            Dim r As New HttpResponseMessage(codigo)
            If Not EsquecerODePedido Then
                r.RequestMessage = pedido
                If EnderecoFinal IsNot Nothing Then
                    r.RequestMessage = New HttpRequestMessage(pedido.Method, EnderecoFinal)
                End If
            End If
            If corpo IsNot Nothing Then r.Content = corpo
            Return r
        End Function
    End Class

    ''' <summary>
    ''' <b>Um fluxo que nunca acaba, e que conta quanto lhe foi pedido.</b>
    '''
    ''' Sem ele, o teto do download era provado apenas pelo desfecho — e um
    ''' código que lesse o corpo inteiro e recusasse no fim daria o mesmo
    ''' desfecho. O que precisa ser provado é <i>quanto</i> chegou a ser lido.
    ''' </summary>
    Private NotInheritable Class FluxoSemFim
        Inherits Stream

        Friend Property Entregues As Long

        Public Overrides ReadOnly Property CanRead As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property CanSeek As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property CanWrite As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property Length As Long
            Get
                Throw New NotSupportedException()
            End Get
        End Property
        Public Overrides Property Position As Long
            Get
                Throw New NotSupportedException()
            End Get
            Set
                Throw New NotSupportedException()
            End Set
        End Property
        Public Overrides Sub Flush()
        End Sub
        Public Overrides Function Seek(o As Long, r As SeekOrigin) As Long
            Throw New NotSupportedException()
        End Function
        Public Overrides Sub SetLength(v As Long)
            Throw New NotSupportedException()
        End Sub
        Public Overrides Sub Write(b As Byte(), o As Integer, c As Integer)
            Throw New NotSupportedException()
        End Sub

        Public Overrides Function Read(buffer As Byte(), deslocamento As Integer,
                                       quantos As Integer) As Integer
            ' SEMPRE ENCHE O QUE LHE PEDIREM. E o servidor hostil do enunciado:
            ' ele nao para, entao quem tem de parar e quem le.
            Entregues += quantos
            Return quantos
        End Function
    End Class

    ''' <summary>
    ''' Monta a procura já com manifesto e assinatura no lugar. Devolve também o
    ''' handler, para o teste mexer no que foi publicado.
    ''' </summary>
    Private Shared Function Montar(corpoDoManifesto As Byte(),
                                   aAssinatura As Byte(),
                                   chavePublica As Byte(),
                                   jaInstalada As Version,
                                   <Runtime.InteropServices.Out> ByRef servidor _
                                       As RespostasDeVersao) As ProcuraDeVersao
        servidor = New RespostasDeVersao()
        servidor.Corpos(Base & "/iris.json") = corpoDoManifesto
        servidor.Corpos(Base & "/iris.json.sig") = aAssinatura
        Return New ProcuraDeVersao(New HttpClient(servidor), Base & "/iris.json",
                                   chavePublica, jaInstalada)
    End Function

    ' ==================================================================
    ' O manifesto e a assinatura
    ' ==================================================================

    <TestMethod>
    Public Sub Manifesto_assinado_pela_chave_certa_e_lido()
        Using dono = ParNovo()
            Dim corpo = MontarJson("1.2.3")
            Dim motivo As String = Nothing

            Dim lido = ManifestoDeVersao.Ler(corpo, Assinar(corpo, dono),
                                             Publica(dono), motivo)

            Assert.IsNotNull(lido, "recusou um manifesto legitimo: " & motivo)
            Assert.AreEqual(New Version(1, 2, 3), lido.Versao)
            Assert.AreEqual(Base & "/Iris.exe", lido.Endereco)
            Assert.AreEqual(40_000_000L, lido.Bytes)

            ' E OS CAMPOS QUE FALTAVAM. Sem eles, zerar Notas, Sha256 e Publicada
            ' no construtor deixava este teste verde -- e sao justamente os
            ' campos que a tela mostra e que o download compara.
            Assert.AreEqual("coisas", lido.Notas)
            Assert.AreEqual(New String("a"c, 64), lido.Sha256)
            Assert.AreEqual(New DateTimeOffset(2026, 9, 2, 10, 0, 0, TimeSpan.Zero),
                            lido.Publicada)
        End Using
    End Sub

    ''' <summary>
    ''' <b>A propriedade central.</b> Um manifesto perfeito, assinado por alguém
    ''' que não é o dono, é recusado.
    ''' </summary>
    <TestMethod>
    Public Sub Assinatura_de_OUTRA_chave_e_recusada()
        Using dono = ParNovo(), impostor = ParNovo()
            Dim corpo = MontarJson("9.9.9")
            Dim motivo As String = Nothing

            Dim lido = ManifestoDeVersao.Ler(corpo, Assinar(corpo, impostor),
                                             Publica(dono), motivo)

            Assert.IsNull(lido, "aceitou manifesto assinado por outra chave")
            StringAssert.Contains(motivo, "assinatura")
        End Using
    End Sub

    ''' <summary>
    ''' Assinado pelo dono e <b>alterado depois</b>. É o caso que a assinatura
    ''' destacada existe para pegar: trocar o endereço de download de um
    ''' manifesto verdadeiro.
    ''' </summary>
    <TestMethod>
    Public Sub Um_byte_trocado_depois_de_assinar_invalida_tudo()
        Using dono = ParNovo()
            Dim corpo = MontarJson("1.2.3")
            Dim aAssinatura = Assinar(corpo, dono)

            ' O manifesto que sai do servidor nao e o que foi assinado.
            Dim trocado = CType(corpo.Clone(), Byte())
            trocado(trocado.Length - 2) = Asc("9"c)

            Dim motivo As String = Nothing
            Assert.IsNull(ManifestoDeVersao.Ler(trocado, aAssinatura,
                                                Publica(dono), motivo),
                          "aceitou manifesto alterado depois da assinatura")
        End Using
    End Sub

    <TestMethod>
    Public Sub Sem_assinatura_nao_ha_leitura()
        Using dono = ParNovo()
            Dim motivo As String = Nothing
            Assert.IsNull(ManifestoDeVersao.Ler(MontarJson("1.2.3"), Array.Empty(Of Byte)(),
                                                Publica(dono), motivo))
            StringAssert.Contains(motivo, "assinatura")
        End Using
    End Sub

    <TestMethod>
    Public Sub Sem_chave_publica_nada_confere()
        Using dono = ParNovo()
            Dim corpo = MontarJson("1.2.3")
            Dim motivo As String = Nothing
            Assert.IsNull(ManifestoDeVersao.Ler(corpo, Assinar(corpo, dono),
                                                Array.Empty(Of Byte)(), motivo))
        End Using
    End Sub

    ''' <summary>
    ''' <b>Assinado pelo dono e recusado assim mesmo.</b>
    '''
    ''' A assinatura prova quem escreveu, e não que o que ele escreveu está certo.
    ''' Um <c>http://</c> assinado é um engano do dono — e um engano que entrega o
    ''' pacote a quem estiver no caminho.
    ''' </summary>
    <TestMethod>
    Public Sub Endereco_http_e_recusado_mesmo_ASSINADO()
        Using dono = ParNovo()
            Dim corpo = MontarJson("1.2.3", aonde:="http://exemplo.invalido/Iris.exe")
            Dim motivo As String = Nothing

            Assert.IsNull(ManifestoDeVersao.Ler(corpo, Assinar(corpo, dono),
                                                Publica(dono), motivo),
                          "aceitou download por http porque estava assinado")
            StringAssert.Contains(motivo, "https")
        End Using
    End Sub

    <TestMethod>
    Public Sub Sha_que_nao_e_sha_e_recusado()
        Using dono = ParNovo()
            For Each ruim In {"", "abc", New String("z"c, 64), New String("a"c, 63)}
                Dim corpo = MontarJson("1.2.3", hash:=ruim)
                Dim motivo As String = Nothing
                Assert.IsNull(ManifestoDeVersao.Ler(corpo, Assinar(corpo, dono),
                                                    Publica(dono), motivo),
                              "aceitou sha256 invalido: '" & ruim & "'")
            Next
        End Using
    End Sub

    <TestMethod>
    Public Sub Tamanho_implausivel_e_recusado()
        Using dono = ParNovo()
            For Each ruim In {0L, -1L, ManifestoDeVersao.TamanhoMaximo + 1L}
                Dim corpo = MontarJson("1.2.3", quantos:=ruim)
                Dim motivo As String = Nothing
                Assert.IsNull(ManifestoDeVersao.Ler(corpo, Assinar(corpo, dono),
                                                    Publica(dono), motivo),
                              "aceitou tamanho " & ruim)
            Next
        End Using
    End Sub

    ''' <summary>
    ''' Um manifesto grande demais é recusado <b>antes</b> de a assinatura ser
    ''' conferida — conferir custaria hash sobre o arquivo inteiro, que é
    ''' exatamente o que um manifesto de 40 MB estaria pedindo.
    ''' </summary>
    <TestMethod>
    Public Sub Manifesto_grande_demais_nem_chega_na_assinatura()
        Using dono = ParNovo(), impostor = ParNovo()
            ' 64 KiB + 1 byte, que e o teto do manifesto -- e nao os 40 MB do
            ' campo "bytes", que descreve o PACOTE.
            Dim enorme(ManifestoDeVersao.ManifestoMaximo) As Byte

            ' A ASSINATURA E DE OUTRA CHAVE, e e isso que prova a ORDEM.
            '
            ' Com uma assinatura valida, o motivo seria "grande demais" tanto
            ' antes quanto depois da conferencia criptografica, e o teste nao
            ' distinguiria as duas ordens. Com uma invalida, so a ordem certa
            ' produz "grande demais"; a trocada produziria "a assinatura nao
            ' confere".
            Dim motivo As String = Nothing
            Assert.IsNull(ManifestoDeVersao.Ler(enorme, Assinar(enorme, impostor),
                                                Publica(dono), motivo))
            StringAssert.Contains(motivo, "grande demais",
                                  "conferiu a assinatura antes do teto de tamanho")
        End Using
    End Sub

    ' ==================================================================
    ' A procura
    ' ==================================================================

    <TestMethod>
    Public Async Function Versao_maior_e_oferecida() As Task
        Using dono = ParNovo()
            Dim corpo = MontarJson("2.0.0")
            Dim servidor As RespostasDeVersao = Nothing
            Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                 New Version(1, 0, 0), servidor)

            Dim r = Await procura.Procurar(CancellationToken.None)

            Assert.AreEqual(DesfechoDaProcura.HaVersaoNova, r.Desfecho, r.Frase)
            Assert.AreEqual(New Version(2, 0, 0), r.Manifesto.Versao)
        End Using
    End Function

    <TestMethod>
    Public Async Function Versao_igual_ja_esta_em_dia() As Task
        Using dono = ParNovo()
            Dim corpo = MontarJson("1.0.0")
            Dim servidor As RespostasDeVersao = Nothing
            Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                 New Version(1, 0, 0), servidor)

            Assert.AreEqual(DesfechoDaProcura.JaEstaEmDia,
                            (Await procura.Procurar(CancellationToken.None)).Desfecho)
        End Using
    End Function

    ''' <summary>
    ''' <b>Versão menor não é oferecida.</b>
    '''
    ''' Um manifesto assinado que aponta para uma versão antiga é o ataque de
    ''' rebaixamento: nada nele está falsificado — é uma release verdadeira, só
    ''' que a de antes da correção que interessa a quem serviu o arquivo.
    ''' </summary>
    <TestMethod>
    Public Async Function Versao_MENOR_nao_e_oferecida() As Task
        Using dono = ParNovo()
            Dim corpo = MontarJson("0.9.0")
            Dim servidor As RespostasDeVersao = Nothing
            Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                 New Version(1, 0, 0), servidor)

            Dim r = Await procura.Procurar(CancellationToken.None)

            Assert.AreEqual(DesfechoDaProcura.JaEstaEmDia, r.Desfecho,
                            "ofereceu uma versao mais velha que a instalada")
        End Using
    End Function

    ''' <summary>
    ''' <b>"Não confere" e "não deu para saber" são desfechos diferentes.</b>
    '''
    ''' Um é a rede falhando, e a resposta é tentar de novo mais tarde. O outro é
    ''' um arquivo que existe, chegou inteiro, e não é do dono — e a resposta a
    ''' esse é outra. Colapsar os dois num só apagaria a única evidência que o
    ''' usuário teria.
    ''' </summary>
    <TestMethod>
    Public Async Function Assinatura_ruim_NAO_vira_nao_deu_para_saber() As Task
        Using dono = ParNovo(), impostor = ParNovo()
            Dim corpo = MontarJson("2.0.0")
            Dim servidor As RespostasDeVersao = Nothing
            Dim procura = Montar(corpo, Assinar(corpo, impostor), Publica(dono),
                                 New Version(1, 0, 0), servidor)

            Dim r = Await procura.Procurar(CancellationToken.None)

            Assert.AreEqual(DesfechoDaProcura.AssinaturaNaoConfere, r.Desfecho, r.Frase)
            Assert.IsNull(r.Manifesto, "devolveu o manifesto de uma assinatura recusada")
        End Using
    End Function

    <TestMethod>
    Public Async Function Servidor_fora_do_ar_e_nao_deu_para_saber() As Task
        Using dono = ParNovo()
            Dim corpo = MontarJson("2.0.0")
            Dim servidor As RespostasDeVersao = Nothing
            Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                 New Version(1, 0, 0), servidor)
            ' A MENSAGEM DA EXCECAO CARREGA UM SEGREDO, de proposito.
            '
            ' Antes, a excecao dizia so "sem rede", e a assercao procurava o
            ' dominio -- entao devolver ex.Message inteiro deixava o teste verde.
            ' Agora o segredo esta na mensagem, e so uma frase que NAO a repete
            ' passa.
            Const segredo = "Authorization: Bearer abc123"
            servidor.Explodir = New HttpRequestException(
                "falha ao chamar https://exemplo.invalido/iris.json (" & segredo & ")")

            Dim r = Await procura.Procurar(CancellationToken.None)

            Assert.AreEqual(DesfechoDaProcura.NaoDeuParaSaber, r.Desfecho)
            ' A FRASE VAI PARA A TELA. A mensagem de uma excecao de rede carrega
            ' endereco e as vezes cabecalhos; so o tipo pode sair.
            Assert.IsFalse(r.Frase.Contains(segredo),
                           "a frase da tela repetiu a mensagem da excecao: " & r.Frase)
            Assert.IsFalse(r.Frase.Contains("exemplo.invalido"),
                           "a frase da tela vazou o endereco: " & r.Frase)
            StringAssert.Contains(r.Frase, "HttpRequestException",
                                  "sem o tipo, a frase nao ajuda ninguem a diagnosticar")
        End Using
    End Function

    ''' <summary>
    ''' <b>E nenhum pedido chega a sair.</b>
    '''
    ''' Só o desfecho não provava nada: <c>NaoDeuParaSaber</c> é também o
    ''' resultado de 404, de rede caída e de JSON ilegível. Sem a contagem de
    ''' pedidos, remover a conferência de <c>https</c> deixaria o teste verde —
    ''' o fake responderia 404 e o <c>Catch</c> devolveria o mesmo desfecho.
    ''' </summary>
    <TestMethod>
    Public Async Function Endereco_de_versoes_precisa_ser_https() As Task
        Using dono = ParNovo()
            Dim servidor As New RespostasDeVersao()
            Dim procura = New ProcuraDeVersao(New HttpClient(servidor),
                                              "http://exemplo.invalido/iris.json",
                                              Publica(dono), New Version(1, 0, 0))

            Assert.AreEqual(DesfechoDaProcura.NaoDeuParaSaber,
                            (Await procura.Procurar(CancellationToken.None)).Desfecho)
            Assert.AreEqual(0, servidor.Pedidos.Count,
                            "falou com o servidor por http antes de recusar: " &
                            String.Join(", ", servidor.Pedidos))
        End Using
    End Function

    ' ==================================================================
    ' O download
    ' ==================================================================

    Private Shared Function PastaNova() As String
        Dim onde = Path.Combine(Path.GetTempPath(), "iris-atualizacao-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(onde)
        Return onde
    End Function

    <TestMethod>
    Public Async Function Pacote_conferido_e_gravado_com_o_nome_da_versao() As Task
        Dim onde = PastaNova()
        Try
            Using dono = ParNovo()
                Dim conteudo = Encoding.UTF8.GetBytes("um executavel de mentira")
                Dim hash = Convert.ToHexString(SHA256.HashData(conteudo)).ToLowerInvariant()
                Dim corpo = MontarJson("2.0.0", hash:=hash, quantos:=conteudo.LongLength)

                Dim servidor As RespostasDeVersao = Nothing
                Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                     New Version(1, 0, 0), servidor)
                servidor.Corpos(Base & "/Iris.exe") = conteudo

                Dim r = Await procura.Procurar(CancellationToken.None)
                Dim pacote = Await procura.Baixar(r.Manifesto, onde, CancellationToken.None)

                Assert.IsTrue(pacote.Veio, pacote.Motivo)
                Assert.AreEqual(Path.Combine(onde, "Iris-2.0.0.exe"), pacote.Caminho)
                CollectionAssert.AreEqual(conteudo, File.ReadAllBytes(pacote.Caminho))
            End Using
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Function

    ''' <summary>
    ''' <b>Hash diferente recusa — e não deixa nada no disco.</b>
    '''
    ''' A segunda metade é a que importa. Um parcial de executável sobrevivendo
    ''' com qualquer nome numa pasta de Downloads é um arquivo que alguém pode
    ''' executar sem saber que ele foi recusado.
    ''' </summary>
    <TestMethod>
    Public Async Function Pacote_com_hash_diferente_e_recusado_e_NAO_fica_no_disco() As Task
        Dim onde = PastaNova()
        Try
            Using dono = ParNovo()
                Dim conteudo = Encoding.UTF8.GetBytes("um executavel trocado")
                ' O manifesto promete OUTRO hash, do tamanho certo.
                Dim corpo = MontarJson("2.0.0", hash:=New String("b"c, 64),
                                       quantos:=conteudo.LongLength)

                Dim servidor As RespostasDeVersao = Nothing
                Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                     New Version(1, 0, 0), servidor)
                servidor.Corpos(Base & "/Iris.exe") = conteudo

                Dim r = Await procura.Procurar(CancellationToken.None)
                Dim pacote = Await procura.Baixar(r.Manifesto, onde, CancellationToken.None)

                Assert.IsFalse(pacote.Veio, "aceitou um pacote que nao bate com o manifesto")
                CollectionAssert.AreEqual(Array.Empty(Of String)(),
                                          Directory.GetFiles(onde),
                                          "sobrou arquivo de um download recusado: " &
                                          String.Join(", ", Directory.GetFiles(onde)))
            End Using
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Function

    ''' <summary>
    ''' O servidor manda mais do que o manifesto declara. O teto é conferido
    ''' <b>durante</b> a leitura: esperar o fim seria aceitar que um servidor
    ''' hostil escolha quanto disco usar.
    ''' </summary>
    <TestMethod>
    Public Async Function Pacote_maior_que_o_declarado_e_recusado() As Task
        Dim onde = PastaNova()
        Try
            Using dono = ParNovo()
                Const teto = 1_000L
                Dim corpo = MontarJson("2.0.0", quantos:=teto)

                Dim servidor As RespostasDeVersao = Nothing
                Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                     New Version(1, 0, 0), servidor)

                ' UM SERVIDOR QUE NAO PARA. Um vetor grande provaria o desfecho e
                ' nao a propriedade: um codigo que lesse o corpo inteiro e
                ' recusasse no fim daria o mesmo desfecho. Aqui nao ha fim.
                Dim semFim As New FluxoSemFim()
                servidor.Fluxos(Base & "/Iris.exe") = semFim

                Dim r = Await procura.Procurar(CancellationToken.None)
                Dim pacote = Await procura.Baixar(r.Manifesto, onde, CancellationToken.None)

                Assert.IsFalse(pacote.Veio, "gravou mais do que o manifesto declarava")
                CollectionAssert.AreEqual(Array.Empty(Of String)(), Directory.GetFiles(onde))

                ' E A MEDIDA. Com o clamp, o maximo que se pede e teto+1 -- o
                ' byte a mais existe so para o excesso ser detectado. Sem ele,
                ' a primeira leitura ja pediria os 80 KiB do buffer.
                Assert.IsTrue(semFim.Entregues <= teto + 1,
                              "leu " & semFim.Entregues & " bytes para um teto de " &
                              teto & ": o teto nao esta sendo aplicado NA leitura")
            End Using
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Function

    ' ==================================================================
    ' A tela
    ' ==================================================================

    ' ==================================================================
    ' A versão que este executável diz ser
    ' ==================================================================

    ''' <summary>
    ''' <b>O sufixo de commit é cortado — e a string do teste é a de verdade.</b>
    '''
    ''' <c>0.1.0+f29b9b8a…</c> foi lido do <c>Iris.exe</c> autocontido que o
    ''' <c>dotnet publish</c> produziu em 02/09/2026, e não inventado: o SDK cola
    ''' o commit no <c>InformationalVersion</c> por padrão.
    '''
    ''' Sem o corte, <c>Version.TryParse</c> falha calado e devolve 0.0.0 — o
    ''' Iris se acharia mais velho que qualquer coisa e ofereceria atualização
    ''' para sempre, inclusive para a versão que ele já é.
    ''' </summary>
    <TestMethod>
    Public Sub A_versao_instalada_ignora_o_sufixo_de_commit()
        Assert.AreEqual(New Version(0, 1, 0),
                        ProcuraDeVersao.Interpretar("0.1.0+f29b9b8a21079379dc057573f0984fae14c806e2"),
                        "nao cortou o commit que o SDK cola no InformationalVersion")
        Assert.AreEqual(New Version(1, 2, 3), ProcuraDeVersao.Interpretar("1.2.3"))
        Assert.AreEqual(New Version(1, 2, 3), ProcuraDeVersao.Interpretar("1.2.3-beta.1"))

        ' QUATRO COMPONENTES VIRAM TRES. O AssemblyVersion do .NET tem quatro, e
        ' a quarta nunca aparece na tela nem no nome do arquivo -- mantê-la faria
        ' 1.2.3.4 comparar como maior que 1.2.3 e exibir-se como "1.2.3".
        Assert.AreEqual(New Version(1, 2, 3), ProcuraDeVersao.Interpretar("1.2.3.4"))
    End Sub

    ''' <summary>
    ''' Ilegível vira 0.0.0 e não exceção — esta função roda na montagem da tela,
    ''' e uma exceção aqui derrubaria o Iris no arranque por causa de um atributo
    ''' de compilação.
    ''' </summary>
    <TestMethod>
    Public Sub Versao_ilegivel_vira_NOTHING_e_nao_excecao()
        For Each ruim In {Nothing, "", "   ", "abc", "1", "1.", "versao 2"}
            Assert.IsNull(ProcuraDeVersao.Interpretar(ruim),
                          "aceitou '" & If(ruim, "<Nothing>") & "' como versao")
        Next

        ' NOTHING, E NAO 0.0.0, e a mudanca nao e cosmetica: com zero, a regra
        ' antirrebaixamento passava a aceitar qualquer manifesto assinado, por
        ' mais antigo que fosse, porque tudo e maior que zero. Ela falhava aberta
        ' exatamente quando perdia a informacao de que precisa.
        '
        ' VersaoInstalada() continua colapsando em 0.0.0 para a TELA, que precisa
        ' mostrar alguma coisa -- mas NAO da para afirmar isso aqui: sob o MSTest,
        ' GetEntryAssembly() e o host, cujo InformationalVersion e 17.13.0. Uma
        ' assercao sobre ela mediria a versao do VSTest, e quebraria no dia em
        ' que ele fosse atualizado.
    End Sub

    <TestMethod>
    Public Sub Sem_chave_configurada_o_botao_de_verificar_nao_liga()
        Dim tela As New AtualizacaoViewModel(Nothing, "")

        Assert.IsFalse(tela.VerificarCommand.CanExecute(Nothing))
        StringAssert.Contains(tela.Frase, "não foi configurada")
        Assert.IsFalse(tela.HaVersaoNova)
    End Sub

    ''' <summary>
    ''' <b>Uma segunda procura apaga a oferta da primeira.</b>
    '''
    ''' Sem isto, procurar, achar a 2.0.0, e procurar de novo com o servidor fora
    ''' do ar deixaria na tela "não consegui falar com o servidor" ao lado de um
    ''' botão "Baixar" habilitado — e o botão baixaria a 2.0.0, sem que nada na
    ''' tela dissesse de onde ela veio.
    ''' </summary>
    <TestMethod>
    Public Async Function Segunda_procura_apaga_a_oferta_da_primeira() As Task
        Dim onde = PastaNova()
        Try
            Using dono = ParNovo()
                ' O CICLO COMPLETO ANTES DA SEGUNDA PROCURA: procurar, BAIXAR, e
                ' so entao procurar de novo. Sem o download, o teste nao provava
                ' que Esquecer() limpa "Baixado" -- e uma tela que mantivesse o
                ' arquivo da oferta anterior mostraria um caminho ao lado de uma
                ' frase que fala de outra coisa.
                Dim conteudo = Encoding.UTF8.GetBytes("um executavel de mentira")
                Dim hash = Convert.ToHexString(SHA256.HashData(conteudo)).ToLowerInvariant()
                Dim corpo = MontarJson("2.0.0", hash:=hash, quantos:=conteudo.LongLength)

                Dim servidor As RespostasDeVersao = Nothing
                Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                     New Version(1, 0, 0), servidor)
                servidor.Corpos(Base & "/Iris.exe") = conteudo

                Using tela As New AtualizacaoViewModel(procura, onde)
                    Await tela.VerificarCommand.ExecuteAsync(Nothing)
                    Assert.IsTrue(tela.HaVersaoNova, "nao achou a versao nova: " & tela.Frase)
                    Assert.IsTrue(tela.BaixarCommand.CanExecute(Nothing))

                    Await tela.BaixarCommand.ExecuteAsync(Nothing)
                    Assert.IsTrue(tela.TemBaixado, "nao baixou: " & tela.Frase)
                    Assert.IsTrue(tela.MostrarNaPastaCommand.CanExecute(Nothing))

                    servidor.Explodir = New HttpRequestException("caiu")
                    Await tela.VerificarCommand.ExecuteAsync(Nothing)

                    Assert.IsFalse(tela.HaVersaoNova, "manteve a oferta de uma procura que falhou")
                    Assert.IsFalse(tela.BaixarCommand.CanExecute(Nothing),
                                   "o botao Baixar sobreviveu a uma procura sem resposta")
                    Assert.AreEqual("", tela.Notas)
                    Assert.AreEqual("", tela.Baixado,
                                    "a tela ainda aponta para o arquivo da oferta anterior")
                    Assert.IsFalse(tela.MostrarNaPastaCommand.CanExecute(Nothing))
                End Using
            End Using
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Function

    ''' <summary>
    ''' <b>O caminho feliz da tela, do clique ao arquivo.</b>
    '''
    ''' Ele não existia: os testes da tela cobriam só as recusas, e o
    ''' <c>BaixarAsync</c> inteiro — a frase final, o caminho, o
    ''' <c>MostrarNaPastaCommand</c> — não era tocado por ninguém.
    ''' </summary>
    <TestMethod>
    Public Async Function A_tela_baixa_e_diz_onde_o_arquivo_ficou() As Task
        Dim onde = PastaNova()
        Try
            Using dono = ParNovo()
                Dim conteudo = Encoding.UTF8.GetBytes("outro executavel de mentira")
                Dim hash = Convert.ToHexString(SHA256.HashData(conteudo)).ToLowerInvariant()
                Dim corpo = MontarJson("3.1.4", hash:=hash, quantos:=conteudo.LongLength,
                                       oQueMudou:="consertos")

                Dim servidor As RespostasDeVersao = Nothing
                Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                     New Version(1, 0, 0), servidor)
                servidor.Corpos(Base & "/Iris.exe") = conteudo

                Using tela As New AtualizacaoViewModel(procura, onde)
                    Assert.IsFalse(tela.BaixarCommand.CanExecute(Nothing),
                                   "da para baixar antes de procurar")

                    Await tela.VerificarCommand.ExecuteAsync(Nothing)
                    Assert.AreEqual("consertos", tela.Notas)
                    Assert.IsTrue(tela.TemNotas)

                    Await tela.BaixarCommand.ExecuteAsync(Nothing)

                    Assert.AreEqual(Path.Combine(onde, "Iris-3.1.4.exe"), tela.Baixado)
                    Assert.IsTrue(File.Exists(tela.Baixado))
                    ' A FRASE DIZ O QUE FALTA FAZER. "Baixado com sucesso" e
                    ' verdade e nao ajuda: o programa novo so passa a existir
                    ' quando alguem executa o arquivo.
                    StringAssert.Contains(tela.Frase, "execute o arquivo")
                    Assert.IsFalse(tela.Ocupado)
                End Using
            End Using
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Function

    ' ==================================================================
    ' As duas pontas se encontram
    ' ==================================================================

    ''' <summary>
    ''' <b>Assinado pela ferramenta de publicação, lido pelo cliente.</b>
    '''
    ''' Este é o teste que faltava, e a falta não era pequena: até aqui, "os dois
    ''' lados usam o mesmo formato de assinatura" era uma afirmação sobre a
    ''' documentação do .NET — <c>SignData</c> e <c>VerifyData</c> usam
    ''' IeeeP1363 por padrão — e não sobre este programa. Os testes antigos
    ''' assinavam com o mesmo <c>ECDsa</c> que verificavam, então concordariam
    ''' entre si em qualquer formato.
    '''
    ''' Aqui quem assina é <c>Iris.Assinatura.Assinador</c>, que é o código que
    ''' <c>tools/publicar-versao.ps1</c> chama, e quem lê é
    ''' <c>ManifestoDeVersao.Ler</c>, que é o código que roda na máquina do
    ''' usuário. A chave pública atravessa como Base64, que é a forma em que ela
    ''' é colada em <c>ChaveDeAtualizacao.vb</c>.
    ''' </summary>
    <TestMethod>
    Public Sub A_ferramenta_assina_e_o_cliente_le()
        Dim publicaBase64 As String = Nothing
        Dim privadaEmPem = Assinatura.Assinador.GerarPar(publicaBase64)

        Dim corpo = MontarJson("2.0.0", oQueMudou:="notas com acento: ção, ümlaut, 日本")
        Dim aAssinatura = Assinatura.Assinador.Assinar(privadaEmPem, corpo)

        ' Convert.FromBase64String e exatamente o que ChaveDeAtualizacao.Bytes
        ' faz com a string que o dono cola no codigo.
        Dim motivo As String = Nothing
        Dim lido = ManifestoDeVersao.Ler(corpo, aAssinatura,
                                         Convert.FromBase64String(publicaBase64), motivo)

        Assert.IsNotNull(lido, "o cliente recusou o que a ferramenta assinou: " & motivo)
        Assert.AreEqual(New Version(2, 0, 0), lido.Versao)
        Assert.AreEqual("notas com acento: ção, ümlaut, 日本", lido.Notas,
                        "os acentos nao sobreviveram a viagem")
        Assert.AreEqual(64, aAssinatura.Length,
                        "uma assinatura P-256 em IeeeP1363 tem 64 bytes; " &
                        "128 e mais seria DER, e o cliente nao le DER")
    End Sub

    ''' <summary>
    ''' E a chave errada continua sendo recusada quando quem assina é a
    ''' ferramenta de verdade — sem isto, o teste de cima provaria só que
    ''' <c>Ler</c> aceita coisas.
    ''' </summary>
    <TestMethod>
    Public Sub A_ferramenta_assinando_com_OUTRA_chave_e_recusada()
        Dim doDono As String = Nothing
        Assinatura.Assinador.GerarPar(doDono)

        Dim doImpostor As String = Nothing
        Dim privadaDoImpostor = Assinatura.Assinador.GerarPar(doImpostor)

        Dim corpo = MontarJson("2.0.0")
        Dim motivo As String = Nothing

        Assert.IsNull(ManifestoDeVersao.Ler(
            corpo, Assinatura.Assinador.Assinar(privadaDoImpostor, corpo),
            Convert.FromBase64String(doDono), motivo))
        StringAssert.Contains(motivo, "assinatura")
    End Sub

    ''' <summary>
    ''' A ferramenta recusa curva que não seja P-256, dos dois lados. Uma chave
    ''' de outra curva produziria assinatura que nenhuma cópia do Iris confere, e
    ''' o erro apareceria na máquina do usuário como "não foi publicado por você"
    ''' — dito sobre um arquivo que foi.
    '''
    ''' <b>O que este teste prova é a metade do tamanho da chave.</b> A P-384 é
    ''' recusada pelo <c>KeySize</c>; a conferência do tamanho da SPKI, que existe
    ''' para outras curvas <i>de 256 bits</i>, sobrevive à sabotagem deste teste.
    ''' Está declarado como tal em <c>Assinador.ExigirP256</c>.
    ''' </summary>
    <TestMethod>
    Public Sub Curva_diferente_e_recusada_nas_duas_pontas()
        Using outraCurva = ECDsa.Create(ECCurve.NamedCurves.nistP384)
            Dim pem = outraCurva.ExportPkcs8PrivateKeyPem()

            Assert.ThrowsException(Of CryptographicException)(
                Sub() Assinatura.Assinador.Assinar(pem, {1, 2, 3}),
                "a ferramenta assinou com uma curva que o cliente nao confere")

            ' E o cliente tambem recusa, mesmo que a assinatura fosse coerente.
            Dim corpo = MontarJson("2.0.0")
            Dim motivo As String = Nothing
            Assert.IsNull(ManifestoDeVersao.Ler(
                corpo, outraCurva.SignData(corpo, HashAlgorithmName.SHA256),
                outraCurva.ExportSubjectPublicKeyInfo(), motivo))
        End Using
    End Sub

    ''' <summary>
    ''' <b>Lixo colado atrás de uma chave válida não é chave válida.</b>
    '''
    ''' <c>ImportSubjectPublicKeyInfo</c> diz quantos bytes leu e ignora o resto.
    ''' Sem conferir, aquilo seria conteúdo que ninguém olhou viajando dentro do
    ''' executável.
    ''' </summary>
    <TestMethod>
    Public Sub Chave_publica_com_lixo_atras_e_recusada()
        Using dono = ParNovo()
            Dim corpo = MontarJson("2.0.0")
            Dim inchada = Publica(dono).Concat(New Byte() {9, 9, 9}).ToArray()

            Dim motivo As String = Nothing
            Assert.IsNull(ManifestoDeVersao.Ler(corpo, Assinar(corpo, dono), inchada, motivo),
                          "aceitou uma chave com bytes sobrando")
        End Using
    End Sub

    ' ==================================================================
    ' A versão do manifesto
    ' ==================================================================

    ''' <summary>
    ''' <b>Três números, nem dois nem quatro.</b>
    '''
    ''' <c>Version.TryParse</c> aceita os três. "1.2" faz <c>ToString(3)</c>
    ''' <b>lançar</b>, numa frase de tela e longe da causa. "1.2.3.4" é pior por
    ''' ser silencioso: ela é maior que 1.2.3 na comparação, mas as duas
    ''' aparecem como "1.2.3" e produzem o mesmo nome de arquivo — o dono veria
    ''' o Iris oferecendo a versão que ele já tem.
    ''' </summary>
    <TestMethod>
    Public Sub Versao_do_manifesto_tem_de_ter_TRES_numeros()
        Using dono = ParNovo()
            For Each ruim In {"1.2", "1.2.3.4", "1", "", "dois"}
                Dim corpo = MontarJson(ruim)
                Dim motivo As String = Nothing
                Assert.IsNull(ManifestoDeVersao.Ler(corpo, Assinar(corpo, dono),
                                                    Publica(dono), motivo),
                              "aceitou a versao '" & ruim & "'")
            Next
        End Using
    End Sub

    ''' <summary>
    ''' <b>Não saber a própria versão não pode virar "sou a 0.0.0".</b>
    '''
    ''' Seguir com zero faria toda versão publicada parecer mais nova, e um
    ''' manifesto antigo legitimamente assinado seria oferecido como atualização.
    ''' A proteção contra rebaixamento falharia aberta exatamente quando perde a
    ''' informação de que precisa.
    ''' </summary>
    <TestMethod>
    Public Async Function Sem_saber_a_versao_instalada_nao_ha_oferta() As Task
        Using dono = ParNovo()
            Dim corpo = MontarJson("2.0.0")
            Dim servidor As New RespostasDeVersao()
            servidor.Corpos(Base & "/iris.json") = corpo
            servidor.Corpos(Base & "/iris.json.sig") = Assinar(corpo, dono)

            ' SEM instalada: e o construtor Friend le isso como "nao sei qual
            ' versao sou". Nao e "descubra" -- quem descobre e o construtor
            ' publico. A distincao existe porque, sem ela, este teste passava
            ' pelo motivo errado: o host do MSTest tem InformationalVersion
            ' 17.13.0, maior que qualquer manifesto de mentira, entao o desfecho
            ' era JaEstaEmDia e a assercao ficava verde sem tocar na guarda.
            Dim procura = New ProcuraDeVersao(New HttpClient(servidor),
                                              Base & "/iris.json", Publica(dono))

            Dim r = Await procura.Procurar(CancellationToken.None)

            Assert.AreEqual(DesfechoDaProcura.NaoDeuParaSaber, r.Desfecho,
                            "ofereceu atualizacao sem saber contra o que comparar")
            Assert.AreEqual(0, servidor.Pedidos.Count,
                            "foi a rede antes de descobrir que nao sabe se comparar")
        End Using
    End Function

    <TestMethod>
    Public Sub A_versao_ilegivel_e_DESCONHECIDA_e_nao_zero()
        ' Interpretar devolve Nothing, e nao 0.0.0: quem decide precisa da
        ' diferenca entre "sou a 0.0.0" e "nao sei qual sou".
        Assert.IsNull(ProcuraDeVersao.Interpretar("abc"))
        Assert.IsNull(ProcuraDeVersao.Interpretar(""))
        Assert.IsNull(ProcuraDeVersao.Interpretar(Nothing))

        ' E o que ela devolve tem sempre tres componentes, para ToString(3) nao
        ' lancar numa frase de tela.
        Assert.AreEqual(New Version(1, 2, 0), ProcuraDeVersao.Interpretar("1.2"))
        Assert.AreEqual("1.2.0", ProcuraDeVersao.Interpretar("1.2").ToString(3))
    End Sub

    ' ==================================================================
    ' O que o script realmente escreve
    ' ==================================================================

    ''' <summary>
    ''' <b>Roda o <c>montar-manifesto.ps1</c> e lê o que ele gravou.</b>
    '''
    ''' A primeira versão deste teste era uma cópia literal, feita à mão, de um
    ''' <c>iris.json</c> de 02/09/2026 — e o comentário dela dizia "se o script
    ''' mudar a forma, este teste cai. É para cair." <b>Era falso</b>: o teste
    ''' construía a própria string e nunca tocava no script. Trocar
    ''' <c>ConvertTo-Json</c> por outra coisa, gravar com BOM, ou serializar
    ''' <c>bytes</c> como texto não derrubava nada.
    '''
    ''' Agora ele executa o script de verdade, no <c>powershell.exe</c> desta
    ''' máquina, e entrega os bytes gravados ao mesmo
    ''' <c>ManifestoDeVersao.Ler</c> que roda na máquina do usuário.
    '''
    ''' <b>É lento e depende do PowerShell</b>, e vale: é o único ponto em que a
    ''' forma produzida e a forma esperada se encontram. Se o PowerShell não
    ''' estiver disponível, o teste <b>falha</b> em vez de ser pulado — pular
    ''' seria voltar a não cobrir nada, só que em silêncio.
    ''' </summary>
    <TestMethod>
    Public Sub O_QUE_O_SCRIPT_ESCREVE_e_lido_pelo_cliente()
        Dim onde = PastaNova()
        Try
            Dim arquivo = Path.Combine(onde, "iris.json")
            Dim oQueMudou = "Notas com aspas "" e acento: coração, ümlaut, 日本."

            Rodar(Path.Combine(RaizDoRepositorio(), "tools", "montar-manifesto.ps1"),
                  "-Versao", "0.2.0",
                  "-Notas", oQueMudou,
                  "-Endereco", "https://exemplo.invalido/Iris-0.2.0.exe",
                  "-Sha256", New String("c"c, 64),
                  "-Bytes", "66422058",
                  "-Publicada", "2026-09-02T13:55:33.8968678Z",
                  "-Destino", arquivo)

            Dim corpo = File.ReadAllBytes(arquivo)

            ' SEM BOM, e conferido nos bytes crus. Com BOM a assinatura ainda
            ' bateria -- os tres bytes seriam assinados junto -- e o
            ' JsonDocument.Parse tropecaria neles logo depois, produzindo "o
            ' manifesto nao e JSON legivel" DEPOIS de a assinatura conferir.
            Assert.IsFalse(corpo.Length >= 3 AndAlso
                           corpo(0) = &HEF AndAlso corpo(1) = &HBB AndAlso corpo(2) = &HBF,
                           "o script gravou BOM")

            Dim publicaBase64 As String = Nothing
            Dim privadaEmPem = Assinatura.Assinador.GerarPar(publicaBase64)

            Dim motivo As String = Nothing
            Dim lido = ManifestoDeVersao.Ler(
                corpo, Assinatura.Assinador.Assinar(privadaEmPem, corpo),
                Convert.FromBase64String(publicaBase64), motivo)

            Assert.IsNotNull(lido, "o cliente nao le o que o script escreve: " & motivo &
                                   vbLf & Encoding.UTF8.GetString(corpo))
            Assert.AreEqual(New Version(0, 2, 0), lido.Versao)
            ' bytes COMO NUMERO: se o script passar a escrever "66422058" entre
            ' aspas, TryGetInt64 estoura e o manifesto inteiro e recusado.
            Assert.AreEqual(66_422_058L, lido.Bytes)
            Assert.AreEqual(New String("c"c, 64), lido.Sha256)
            Assert.AreEqual(oQueMudou, lido.Notas, "acentos ou aspas nao sobreviveram")
            Assert.AreEqual(2026, lido.Publicada.Year)
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Sub

    ''' <summary>A raiz do repositório, subindo a partir do diretório do teste.</summary>
    Private Shared Function RaizDoRepositorio() As String
        Dim onde As String = AppContext.BaseDirectory
        While onde IsNot Nothing
            If Directory.Exists(Path.Combine(onde, "tools")) AndAlso
               File.Exists(Path.Combine(onde, "Iris.slnx")) Then
                Return onde
            End If
            onde = Path.GetDirectoryName(onde)
        End While
        Throw New InvalidOperationException("nao achei a raiz do repositorio")
    End Function

    ''' <summary>
    ''' Executa um script do <c>tools/</c> e estoura com a saída se ele falhar.
    ''' <c>-NoProfile</c> porque o perfil do usuário não pode influir no que o
    ''' teste mede.
    ''' </summary>
    Private Shared Sub Rodar(script As String, ParamArray argumentos As String())
        Dim linha As New List(Of String) From {
            "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
            "-File", script}
        linha.AddRange(argumentos)

        Dim como As New ProcessStartInfo("powershell.exe") With {
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .UseShellExecute = False,
            .CreateNoWindow = True
        }
        For Each a In linha
            como.ArgumentList.Add(a)
        Next

        Using quem = Process.Start(como)
            Dim saiu = quem.StandardOutput.ReadToEnd()
            Dim erro = quem.StandardError.ReadToEnd()
            Assert.IsTrue(quem.WaitForExit(120_000), "o script nao terminou em 2 minutos")
            Assert.AreEqual(0, quem.ExitCode,
                            Path.GetFileName(script) & " falhou:" & vbLf & erro & vbLf & saiu)
        End Using
    End Sub

    ' ==================================================================
    ' O que a reconferência pós-Move faz
    ' ==================================================================

    ''' <summary>
    ''' <b>Controle negativo da segunda conferência.</b>
    '''
    ''' Sabotar o <c>If Not Confere(...)</c> para <c>If True</c> deixava um teste
    ''' vermelho, e isso provava só que a linha é <i>alcançada</i>. Apagá-la
    ''' inteira deixava a suíte verde — ninguém provava que ela <i>pega</i>
    ''' alguma coisa.
    ''' </summary>
    <TestMethod>
    Public Sub A_reconferencia_pega_arquivo_trocado()
        Dim onde = PastaNova()
        Try
            Dim arquivo = Path.Combine(onde, "pacote.bin")
            Dim conteudo = Encoding.UTF8.GetBytes("o que foi assinado")
            File.WriteAllBytes(arquivo, conteudo)
            Dim certo = Convert.ToHexString(SHA256.HashData(conteudo)).ToLowerInvariant()

            Assert.IsTrue(ProcuraDeVersao.Confere(arquivo, certo),
                          "recusou o arquivo que confere")

            ' O ARQUIVO TROCADO NO CAMINHO -- que e o que a segunda leitura
            ' existe para pegar.
            File.WriteAllBytes(arquivo, Encoding.UTF8.GetBytes("o que alguem pos no lugar"))
            Assert.IsFalse(ProcuraDeVersao.Confere(arquivo, certo),
                           "aceitou um arquivo trocado depois da primeira conferencia")

            File.Delete(arquivo)
            Assert.IsFalse(ProcuraDeVersao.Confere(arquivo, certo),
                           "um arquivo que nao da para ler nao e um arquivo conferido")
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Sub

    ' ==================================================================
    ' O que o descarte da tela realmente faz
    ' ==================================================================

    ''' <summary>
    ''' Um fluxo que só termina quando o <c>CancellationToken</c> é cancelado.
    ''' Serve para haver algo <i>em voo</i> para observar.
    ''' </summary>
    Private NotInheritable Class FluxoQueEspera
        Inherits Stream

        Friend ReadOnly Comecou As New TaskCompletionSource(
            TaskCreationOptions.RunContinuationsAsynchronously)

        ''' <summary>
        ''' <b>True quando a leitura de verdade foi cancelada.</b>
        '''
        ''' Sem isto, o teste do descarte provava que a <i>tarefa do comando</i>
        ''' terminou — e um <c>BaixarAsync</c> que devolvesse uma tarefa pronta
        ''' deixando a rede correndo solta o deixava verde.
        ''' </summary>
        Friend Property Cancelaram As Boolean

        Public Overrides ReadOnly Property CanRead As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property CanSeek As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property CanWrite As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property Length As Long
            Get
                Throw New NotSupportedException()
            End Get
        End Property
        Public Overrides Property Position As Long
            Get
                Throw New NotSupportedException()
            End Get
            Set
                Throw New NotSupportedException()
            End Set
        End Property
        Public Overrides Sub Flush()
        End Sub
        Public Overrides Function Seek(o As Long, r As SeekOrigin) As Long
            Throw New NotSupportedException()
        End Function
        Public Overrides Sub SetLength(v As Long)
            Throw New NotSupportedException()
        End Sub
        Public Overrides Sub Write(b As Byte(), o As Integer, c As Integer)
            Throw New NotSupportedException()
        End Sub
        Public Overrides Function Read(b As Byte(), o As Integer, c As Integer) As Integer
            Throw New NotSupportedException()
        End Function

        Public Overrides Function ReadAsync(
                destino As Memory(Of Byte),
                Optional ct As CancellationToken = Nothing) As ValueTask(Of Integer)
            Comecou.TrySetResult()

            ' CANCELAMENTO CONTINUA CANCELAMENTO.
            '
            ' A primeira versao devolvia Task.Delay(Infinite, ct).ContinueWith(
            ' Function(t) 0) -- e aquele ContinueWith ENGOLIA o estado cancelado
            ' e devolvia zero, que para quem le e FIM DE ARQUIVO. A producao
            ' terminava por tamanho errado, nao por cancelamento, e o teste dizia
            ' provar cancelamento. Apagar o Catch OperationCanceledException da
            ' producao o deixava verde.
            '
            ' Agora a tarefa termina CANCELADA, como um fluxo de rede termina --
            ' e registra que terminou assim.
            ct.Register(Sub() Cancelaram = True)
            Return New ValueTask(Of Integer)(
                Task.Run(Async Function() As Task(Of Integer)
                             Await Task.Delay(Timeout.Infinite, ct)
                             Return 0
                         End Function, ct))
        End Function
    End Class

    ''' <summary>
    ''' <b>Descartar a tela cancela o download que está em voo.</b>
    '''
    ''' A primeira versão deste teste chamava <c>Dispose</c> duas vezes numa tela
    ''' ociosa e verificava que não estourava. Esvaziar o <c>Dispose</c> inteiro
    ''' o deixava verde — ele não observava nada em voo, que é literalmente o que
    ''' o nome dele promete.
    ''' </summary>
    <TestMethod>
    Public Async Function Descartar_a_tela_cancela_o_que_esta_em_voo() As Task
        Dim onde = PastaNova()
        Try
            Using dono = ParNovo()
                Dim corpo = MontarJson("2.0.0", quantos:=1_000_000)
                Dim servidor As RespostasDeVersao = Nothing
                Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                     New Version(1, 0, 0), servidor)

                Dim travado As New FluxoQueEspera()
                servidor.Fluxos(Base & "/Iris.exe") = travado

                Dim tela As New AtualizacaoViewModel(procura, onde)
                Await tela.VerificarCommand.ExecuteAsync(Nothing)
                Assert.IsTrue(tela.HaVersaoNova, tela.Frase)

                Dim baixando = tela.BaixarCommand.ExecuteAsync(Nothing)

                ' COM PRAZO TAMBEM AQUI. Se a producao deixar de ler o corpo,
                ' este Await nao volta -- e um teste que pendura nao e vermelho,
                ' e suite parada.
                Assert.AreSame(travado.Comecou.Task,
                    Await Task.WhenAny(travado.Comecou.Task,
                                       Task.Delay(TimeSpan.FromSeconds(10))),
                    "o download nem chegou a comecar a ler")

                ' EM VOO: e agora da para observar o que so existe durante.
                Assert.IsTrue(tela.Ocupado, "nao ficou ocupado durante o download")
                Assert.IsFalse(tela.VerificarCommand.CanExecute(Nothing),
                               "da para verificar no meio de um download")

                Dim aFraseDoMomento = tela.Frase

                ' CONTA O QUE A TELA AVISAR A PARTIR DO DESCARTE -- inclusive
                ' DURANTE ele.
                '
                ' O handler entrava DEPOIS do Dispose, e assim qualquer
                ' PropertyChanged disparado dentro do proprio Dispose escapava da
                ' conta. O contrato e que nada seja publicado do descarte em
                ' diante, e a contagem tem de comecar antes dele.
                Dim avisosDepois = 0
                Dim contar As ComponentModel.PropertyChangedEventHandler =
                    Sub(quem, oQue) Threading.Interlocked.Increment(avisosDepois)

                AddHandler tela.PropertyChanged, contar
                tela.Dispose()
                ' COM PRAZO, e o prazo e CURTO.
                '
                ' Sem cancelamento, "baixando" nao termina nunca -- o fluxo so sai
                ' por cancelamento. O WhenAny volta pelo prazo e a assercao fica
                ' vermelha; sem ele, a suite ficaria PARADA, que foi o que
                ' aconteceu na primeira vez que sabotei o Dispose.
                '
                ' Cinco segundos, e nao vinte: um prazo largo aceita como sucesso
                ' um cancelamento patologicamente lento -- um Dispose que agendasse
                ' o Cancel para dezenove segundos depois passaria.
                Dim naPaciencia = Await Task.WhenAny(baixando, Task.Delay(TimeSpan.FromSeconds(5)))
                RemoveHandler tela.PropertyChanged, contar
                Assert.AreSame(baixando, naPaciencia,
                               "o descarte nao cancelou o download em 5s")

                Assert.AreEqual(0, avisosDepois,
                                "a tela publicou mudanca de propriedade depois de descartada")

                Assert.AreEqual(aFraseDoMomento, tela.Frase,
                                "a continuacao escreveu na tela depois do descarte")
                ' A LEITURA DE REDE FOI CANCELADA, e nao so a tarefa do comando.
                Assert.IsTrue(travado.Cancelaram,
                    "o comando terminou, mas a leitura de rede seguiu viva")
                Assert.IsFalse(tela.TemBaixado, "deu por baixado um download cancelado")
                CollectionAssert.AreEqual(Array.Empty(Of String)(), Directory.GetFiles(onde),
                                          "sobrou arquivo de um download cancelado")
                Assert.IsFalse(tela.VerificarCommand.CanExecute(Nothing),
                               "os comandos continuam ligados depois do descarte")

                ' E DESCARTAR DE NOVO NAO ESTOURA: o Application_Exit e o Dispose
                ' do MainViewModel podem alcancar o mesmo objeto, e uma excecao
                ' aqui interromperia o descarte de tudo que vem depois.
                tela.Dispose()
            End Using
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Function

    ''' <summary>
    ''' <b>Uma curva de 256 bits que NÃO é a P-256.</b>
    '''
    ''' Este é o controle negativo que faltava. Enquanto o único teste de curva
    ''' errada usava P-384, ele era recusado pelo <c>KeySize</c> — e apagar a
    ''' conferência da curva deixava a suíte verde. A <c>brainpoolP256r1</c> tem
    ''' exatamente 256 bits, passa pelo <c>KeySize</c>, e só o OID a distingue.
    '''
    ''' Depende de o CNG desta máquina conhecer a curva; o Windows 10 em diante
    ''' conhece. Se um dia não conhecer, este teste falha em vez de mentir.
    ''' </summary>
    <TestMethod>
    Public Sub Curva_de_256_bits_que_nao_e_P256_e_recusada()
        Using outra = ECDsa.Create(ECCurve.CreateFromFriendlyName("brainpoolP256r1"))
            Assert.AreEqual(256, outra.KeySize,
                            "esta curva precisa ter 256 bits, senao o teste nao prova nada")

            Assert.ThrowsException(Of CryptographicException)(
                Sub() Assinatura.Assinador.Assinar(outra.ExportPkcs8PrivateKeyPem(), {1, 2, 3}),
                "a ferramenta assinou com uma curva que nao e a P-256")

            ' E o cliente tambem, mesmo com a assinatura coerente com a chave.
            Dim corpo = MontarJson("2.0.0")
            Dim motivo As String = Nothing
            Assert.IsNull(ManifestoDeVersao.Ler(
                corpo, outra.SignData(corpo, HashAlgorithmName.SHA256),
                outra.ExportSubjectPublicKeyInfo(), motivo),
                "o cliente aceitou uma curva que nao e a P-256")
        End Using
    End Sub

    ''' <summary>
    ''' Descartada e <b>ociosa</b>, ela continua sem comandos.
    '''
    ''' O teste do descarte com download em voo não cobria isto: lá o
    ''' <c>Ocupado</c> fica <c>True</c> para sempre, então o <c>CanExecute</c>
    ''' seria <c>False</c> mesmo sem a guarda de descarte. Aqui não há nada
    ''' ocupado, e só a guarda segura.
    ''' </summary>
    <TestMethod>
    Public Async Function Tela_descartada_e_ociosa_nao_tem_comando_ligado() As Task
        Dim onde = PastaNova()
        Try
          Using dono = ParNovo()
            Dim corpo = MontarJson("2.0.0")
            Dim servidor As RespostasDeVersao = Nothing
            Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                 New Version(1, 0, 0), servidor)

            Dim tela As New AtualizacaoViewModel(procura, onde)
            Await tela.VerificarCommand.ExecuteAsync(Nothing)
            Assert.IsTrue(tela.BaixarCommand.CanExecute(Nothing), "nao chegou a ter oferta")
            Assert.IsFalse(tela.Ocupado, "deveria estar ociosa aqui")
            Dim pedidosAteAqui = servidor.Pedidos.Count

            tela.Dispose()

            Assert.IsFalse(tela.BaixarCommand.CanExecute(Nothing),
                           "da para baixar depois de a tela ser descartada")
            Assert.IsFalse(tela.VerificarCommand.CanExecute(Nothing),
                           "da para verificar depois de a tela ser descartada")
            Assert.IsFalse(tela.MostrarNaPastaCommand.CanExecute(Nothing))

            ' E EXECUTAR MESMO ASSIM NAO PODE ESTOURAR NEM MEXER EM NADA.
            '
            ' CanExecute e conselho, e nao tranca: um binding do WPF guarda a
            ' referencia do comando, e ExecuteAsync pode ser chamado. Sem a
            ' guarda no corpo, isto tocaria _ateFechar.Token depois do Dispose --
            ' que LANCA ObjectDisposedException -- e antes disso ja teria escrito
            ' Ocupado e Frase.
            Dim antes = tela.Frase
            Await tela.VerificarCommand.ExecuteAsync(Nothing)
            Await tela.BaixarCommand.ExecuteAsync(Nothing)
            Assert.AreEqual(antes, tela.Frase,
                            "executar um comando depois do descarte mexeu na tela")
            Assert.IsFalse(tela.Ocupado)
            Assert.AreEqual(0, servidor.Pedidos.Count - pedidosAteAqui,
                            "foi a rede depois de a tela ser descartada")
          End Using
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Function

    ' ==================================================================
    ' A casca de linha de comando do assinador
    ' ==================================================================

    ''' <summary>
    ''' Roda o utilitário como os scripts o rodam. <b>"RodarAssinador", e não
    ''' "Assinador"</b>: o segundo eclipsaria a classe <c>Assinatura.Assinador</c>,
    ''' porque VB é insensível a maiúsculas. CLAUDE.md, primeira seção. Devolve saída padrão, saída de
    ''' erro e código de saída.
    ''' </summary>
    Private Shared Function RodarAssinador(<Runtime.InteropServices.Out> ByRef saiu As String,
                                      <Runtime.InteropServices.Out> ByRef errou As String,
                                      ParamArray argumentos As String()) As Integer
        Dim como As New ProcessStartInfo("dotnet") With {
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .UseShellExecute = False,
            .CreateNoWindow = True
        }
        como.ArgumentList.Add("exec")
        como.ArgumentList.Add(Path.Combine(AppContext.BaseDirectory, "Iris.Assinatura.dll"))
        For Each a In argumentos
            como.ArgumentList.Add(a)
        Next

        Using quem = Process.Start(como)
            ' OS DOIS CANAIS AO MESMO TEMPO: drenar um e depois o outro trava se
            ' o segundo encher o buffer antes de o primeiro fechar.
            Dim lendoASaida = quem.StandardOutput.ReadToEndAsync()
            Dim lendoOErro = quem.StandardError.ReadToEndAsync()
            Task.WaitAll(lendoASaida, lendoOErro)
            saiu = lendoASaida.Result.Trim()
            errou = lendoOErro.Result.Trim()
            Assert.IsTrue(quem.WaitForExit(60_000), "o assinador nao terminou")
            Return quem.ExitCode
        End Using
    End Function

    ''' <summary>
    ''' <b>O caminho que o dono percorre, pela linha de comando.</b>
    '''
    ''' Os testes de ida e volta chamam a biblioteca; os scripts chamam o
    ''' executável. Entre os dois há a leitura de argumentos, os códigos de saída
    ''' e a gravação dos arquivos — que não tinham teste nenhum.
    ''' </summary>
    <TestMethod>
    Public Sub A_linha_de_comando_do_assinador_faz_o_ciclo_inteiro()
        Dim onde = PastaNova()
        Try
            Dim chave = Path.Combine(onde, "chave.pem")
            Dim saiu As String = Nothing, errou As String = Nothing

            Assert.AreEqual(0, RodarAssinador(saiu, errou, "gerar", "--destino", chave), errou)
            Dim publicaBase64 = saiu
            Assert.IsTrue(publicaBase64.Length > 40, "nao imprimiu a chave publica")

            ' A PRIVADA FICA NO ARQUIVO, e so la -- nas DUAS saidas.
            '
            ' A primeira versao so olhava a saida padrao, e imprimir o PEM em
            ' stderr a deixava verde. Um console mostra as duas do mesmo jeito, e
            ' as duas vao para o buffer de rolagem.
            Dim oPem = File.ReadAllText(chave)
            StringAssert.Contains(oPem, "BEGIN PRIVATE KEY")
            Dim miolo = oPem.Replace("-----BEGIN PRIVATE KEY-----", "").
                             Replace("-----END PRIVATE KEY-----", "").
                             Replace(vbCr, "").Replace(vbLf, "").Trim()
            For Each saida In {publicaBase64, errou}
                ' Sem diferenciar maiusculas: "private key" tambem e vazamento.
                Assert.IsFalse(saida.Contains("PRIVATE", StringComparison.OrdinalIgnoreCase),
                               "saiu 'PRIVATE' por um dos canais: " & saida)
                Assert.IsFalse(saida.Contains(miolo),
                               "a chave privada vazou por um dos canais")
            Next
            Assert.AreEqual("", errou, "o caminho feliz escreveu em stderr")

            ' RECUSA GERAR POR CIMA: uma chave nova invalida tudo o que ja foi
            ' publicado, e isso nao pode acontecer por um comando repetido.
            Assert.AreNotEqual(0, RodarAssinador(saiu, errou, "gerar", "--destino", chave))
            StringAssert.Contains(errou, "já existe")
            StringAssert.Contains(File.ReadAllText(chave), "BEGIN PRIVATE KEY")

            ' A publica e recuperavel sem gerar par novo, e e a MESMA.
            Assert.AreEqual(0, RodarAssinador(saiu, errou, "publica", "--chave", chave), errou)
            Assert.AreEqual(publicaBase64, saiu, "a chave publica mudou entre duas leituras")

            ' Assinar produz o .sig ao lado, com 64 bytes.
            Dim manifesto = Path.Combine(onde, "iris.json")
            Dim corpo = MontarJson("2.0.0")
            File.WriteAllBytes(manifesto, corpo)
            Assert.AreEqual(0, RodarAssinador(saiu, errou, "assinar",
                                         "--chave", chave, "--arquivo", manifesto), errou)

            Dim aAssinatura = File.ReadAllBytes(manifesto & ".sig")
            Assert.AreEqual(64, aAssinatura.Length,
                            "uma assinatura P-256 em IeeeP1363 tem 64 bytes")
            ' E NAO SOBROU PARCIAL: o .sig e gravado com outro nome e promovido.
            Assert.AreEqual(0, Directory.GetFiles(onde, "*.parcial").Length,
                            "sobrou um arquivo parcial da assinatura")

            ' E o cliente le o que a linha de comando produziu.
            Dim motivo As String = Nothing
            Assert.IsNotNull(ManifestoDeVersao.Ler(corpo, aAssinatura,
                                                   Convert.FromBase64String(publicaBase64),
                                                   motivo),
                             "o cliente recusou o que o executavel assinou: " & motivo)

            ' Verbo desconhecido e falta de opcao saem com codigo proprio.
            Assert.AreEqual(2, RodarAssinador(saiu, errou, "voar"))
            StringAssert.Contains(errou, "uso:")
            Assert.AreEqual(1, RodarAssinador(saiu, errou, "publica"))
            StringAssert.Contains(errou, "--chave")
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Sub

    ''' <summary>
    ''' <b>E um redirecionamento que continua em https é ACEITO.</b>
    '''
    ''' É o controle positivo dos dois testes acima: sem ele, uma produção que
    ''' recusasse <i>todo</i> redirecionamento — ou toda procura — passaria nos
    ''' dois. E recusar redirecionamento quebraria o
    ''' <c>releases/latest/download/</c>, que existe para redirecionar.
    ''' </summary>
    <TestMethod>
    Public Async Function Redirecionamento_que_segue_em_https_e_aceito() As Task
        Using dono = ParNovo()
            Dim corpo = MontarJson("2.0.0")
            Dim servidor As RespostasDeVersao = Nothing
            Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                 New Version(1, 0, 0), servidor)

            ' Outro host, outro caminho -- e https.
            servidor.EnderecoFinal = New Uri("https://objects.exemplo.invalido/a/b")

            Dim r = Await procura.Procurar(CancellationToken.None)
            Assert.AreEqual(DesfechoDaProcura.HaVersaoNova, r.Desfecho, r.Frase)
        End Using
    End Function

    ' ==================================================================
    ' Os elos: cada peça é testada, e alguém tem de chamar a peça certa
    ' ==================================================================

    ''' <summary>
    ''' <b>O publicador chama o script do manifesto, e passa tudo.</b>
    '''
    ''' O teste que executa <c>montar-manifesto.ps1</c> prova que <i>aquele
    ''' script</i> produz algo que o cliente lê. Não prova que a publicação o
    ''' usa: trocar a chamada por um <c>ConvertTo-Json</c> escrito à mão em
    ''' <c>publicar-versao.ps1</c> deixava tudo verde.
    '''
    ''' Isto é uma verificação de <b>texto do script</b>, e não de execução —
    ''' rodar a publicação inteira custa três minutos e um <c>.exe</c> de 63 MB.
    ''' É o mesmo <i>recurso</i> de <c>MarcaDaMutacaoTests</c> — olhar onde a
    ''' chamada está —, e é <b>menos rigoroso que ele</b>: aquele confere
    ''' adjacência, contagens e ausência de ramo constante; este lê linhas sem
    ''' entender bloco nem fluxo. Comentários são descartados e os campos são
    ''' cobrados dentro do comando, o que fecha os enganos que aparecem de
    ''' verdade; um <c>if ($false)</c> em volta da chamada, não.
    ''' </summary>
    <TestMethod>
    Public Sub O_publicador_chama_o_script_do_manifesto()
        Dim script = SemComentarios(Path.Combine(
            RaizDoRepositorio(), "tools", "publicar-versao.ps1"), "#")

        ' A CHAMADA, e nao o nome em qualquer lugar. A primeira versao procurava
        ' as substrings no arquivo inteiro: deixar "montar-manifesto.ps1" num
        ' comentario e montar o JSON a mao a deixava verde.
        Dim aChamada = Array.FindIndex(
            script, Function(l) l.Contains("montar-manifesto.ps1"))
        Assert.IsTrue(aChamada >= 0,
            "a publicacao nao chama o script do manifesto: o teste que prova a " &
            "forma do iris.json passou a cobrir codigo que ninguem usa")

        ' E OS CAMPOS SAO DAQUELA CHAMADA. No PowerShell a crase no fim da linha
        ' continua o comando; entao o comando e a linha da chamada mais as
        ' seguintes, ate uma que nao termine em crase.
        Dim comando = script(aChamada)
        Dim i = aChamada
        While script(i).TrimEnd().EndsWith("`") AndAlso i + 1 < script.Length
            i += 1
            comando &= " " & script(i)
        End While

        ' Faltar um campo faria o manifesto sair sem ele -- e o -Bytes ausente,
        ' por exemplo, sai como zero, que o cliente recusa por "tamanho nao
        ' plausivel" DEPOIS de a assinatura conferir.
        For Each campo In {"-Versao", "-Notas", "-Endereco", "-Sha256", "-Bytes", "-Destino"}
            StringAssert.Contains(comando, campo & " ",
                "a chamada de montar-manifesto.ps1 nao passa " & campo &
                vbLf & comando)
        Next

        ' E ASSINA O ARQUIVO QUE ELE ACABOU DE MONTAR.
        Assert.IsTrue(script.Any(Function(l) l.Contains(
            "assinar --chave $Chave --arquivo $arquivoDoManifesto")),
            "a publicacao assina outra coisa que nao o manifesto que montou")
    End Sub

    ''' <summary>
    ''' As linhas de um arquivo <b>sem os comentários</b>, e sem as vazias.
    '''
    ''' Existe porque os meta-testes de posição são buscas de texto, e busca de
    ''' texto aceita comentário: escrever a chamada num comentário satisfazia
    ''' todos eles. Não é análise sintática — uma <c>#</c> dentro de uma string
    ''' ainda confunde — mas fecha o caso que aparece de verdade.
    ''' </summary>
    Private Shared Function SemComentarios(arquivo As String,
                                           marca As String) As String()
        Return File.ReadAllLines(arquivo).
            Select(Function(l)
                       Dim onde = l.IndexOf(marca, StringComparison.Ordinal)
                       Return If(onde >= 0, l.Substring(0, onde), l)
                   End Function).
            Where(Function(l) l.Trim().Length > 0).
            ToArray()
    End Function

    ''' <summary>
    ''' <b>Os dois scripts chamam o assinador pelo mesmo caminho.</b>
    '''
    ''' <c>montar-assinador.ps1</c> existe para haver um lugar só; um deles
    ''' voltar a compilar por conta própria seria a cópia divergente que ele
    ''' evita.
    ''' </summary>
    <TestMethod>
    Public Sub Os_scripts_chamam_o_assinador_pelo_mesmo_caminho()
        For Each qual In {"gerar-chave-de-assinatura.ps1", "publicar-versao.ps1"}
            Dim script = SemComentarios(
                Path.Combine(RaizDoRepositorio(), "tools", qual), "#")

            ' A CHAMADA E ATRIBUIDA A ALGUMA COISA. "$ferramenta = & (...)" e a
            ' forma; so o nome aparecer nao prova que a saida e usada.
            Assert.IsTrue(script.Any(Function(l) l.Contains("montar-assinador.ps1") AndAlso
                                                 l.Contains("=") AndAlso l.Contains("&")),
                qual & " nao usa a saida do montar-assinador.ps1")

            ' E NENHUM DELES COMPILA O ASSINADOR POR CONTA PROPRIA.
            '
            ' A proibicao e sobre O ASSINADOR, e nao sobre compilar: o
            ' publicar-versao.ps1 chama "dotnet publish" para o proprio Iris, que
            ' e o trabalho dele. Banir o verbo inteiro reprovava o script certo --
            ' foi o que aconteceu na primeira versao desta assercao.
            For Each porFora In {"dotnet build", "dotnet publish", "dotnet msbuild", "msbuild "}
                Assert.IsFalse(script.Any(Function(l) l.Contains(porFora) AndAlso
                                                      l.Contains("Iris.Assinatura")),
                    qual & " compila o assinador com '" & porFora & "' em vez de " &
                    "usar o montar-assinador.ps1 -- e uma copia que vai divergir")
            Next
        Next
    End Sub

    ''' <summary>
    ''' <b>A reconferência está DEPOIS do <c>Move</c>, e antes do sucesso.</b>
    '''
    ''' O teste de <c>Confere</c> prova que a função pega arquivo trocado.
    ''' <b>Não prova que ela é chamada no ponto certo</b>: movê-la para antes do
    ''' <c>Move</c>, ou apagá-la, deixava a suíte verde — provocar a troca no
    ''' meio do fluxo real exigiria um gancho no sistema de arquivos.
    '''
    ''' Então a posição é verificada onde ela mora: no texto.
    ''' </summary>
    <TestMethod>
    Public Sub A_reconferencia_vem_depois_do_Move_e_antes_do_sucesso()
        ' SEM COMENTARIOS: este arquivo fala bastante sobre Move e sobre
        ' Confere, e a versao anterior aceitava qualquer uma dessas mencoes.
        Dim fonte = SemComentarios(Path.Combine(
            RaizDoRepositorio(), "src", "Iris.Update", "ProcuraDeVersao.vb"), "'")

        Dim linhaDoMove = Array.FindIndex(
            fonte, Function(l) l.Contains("File.Move(temporario, destino"))
        Assert.IsTrue(linhaDoMove >= 0, "nao achei a promocao do arquivo temporario")

        ' "If Not Confere(destino," e nao "Confere(destino,": o resultado tem de
        ' ser USADO. Chamar e descartar satisfazia a busca antiga.
        Dim linhaDaConferencia = Array.FindIndex(
            fonte, Function(l) l.Contains("If Not Confere(destino,"))
        Assert.IsTrue(linhaDaConferencia >= 0,
                      "nao ha reconferencia do arquivo promovido, ou o resultado " &
                      "dela nao decide nada")

        Assert.IsTrue(linhaDaConferencia > linhaDoMove,
            "a reconferencia esta ANTES do Move: ela conferiria o temporario, e " &
            "a janela que ela existe para fechar e justamente a da promocao")

        Dim linhaDoSucesso = Array.FindIndex(
            fonte, Function(l) l.Contains("Return PacoteBaixado.Sim(destino)"))
        Assert.IsTrue(linhaDoSucesso > linhaDaConferencia,
            "o sucesso e devolvido antes de o arquivo promovido ser reconferido")
    End Sub

    ' ==================================================================
    ' O endereço final
    ' ==================================================================

    ''' <summary>
    ''' <b>Um redirecionamento que termine fora do https é recusado.</b>
    '''
    ''' O dublê passou a preencher <c>RequestMessage</c>, o que o tornou fiel —
    ''' e ainda assim ninguém testava a conferência: ele devolvia sempre o mesmo
    ''' pedido <c>https</c> que recebeu, então apagar <c>ExigirHttpsAteOFim</c>
    ''' não mudava nada.
    ''' </summary>
    <TestMethod>
    Public Async Function Redirecionamento_que_sai_do_https_e_recusado() As Task
        Using dono = ParNovo()
            Dim corpo = MontarJson("2.0.0")
            Dim servidor As RespostasDeVersao = Nothing
            Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                 New Version(1, 0, 0), servidor)

            ' O pedido saiu em https e TERMINOU em http, que e o que um
            ' redirecionamento mal-comportado faz.
            servidor.EnderecoFinal = New Uri("http://exemplo.invalido/iris.json")

            Dim r = Await procura.Procurar(CancellationToken.None)
            Assert.AreEqual(DesfechoDaProcura.NaoDeuParaSaber, r.Desfecho, r.Frase)
            Assert.IsNull(r.Manifesto)
        End Using
    End Function

    ''' <summary>
    ''' E não saber onde o pedido terminou também é recusa: a pós-condição é "o
    ''' endereço final é https", e não conseguir dizer qual foi não a satisfaz.
    ''' </summary>
    <TestMethod>
    Public Async Function Sem_saber_o_endereco_final_tambem_e_recusa() As Task
        Using dono = ParNovo()
            Dim corpo = MontarJson("2.0.0")
            Dim servidor As RespostasDeVersao = Nothing
            Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                 New Version(1, 0, 0), servidor)

            servidor.EsquecerODePedido = True

            Dim r = Await procura.Procurar(CancellationToken.None)
            Assert.AreEqual(DesfechoDaProcura.NaoDeuParaSaber, r.Desfecho)
            Assert.IsNull(r.Manifesto, "devolveu manifesto de um pedido sem endereco final")
        End Using
    End Function

    ''' <summary>
    ''' <b>Cancelar é cancelar, e não "falhou".</b>
    '''
    ''' O teste do descarte observa a tela, e a tela não escreve nada depois de
    ''' descartada — então "cancelado" e "falhou" ficavam indistinguíveis por
    ''' lá: transformar o <c>Catch OperationCanceledException</c> em erro comum
    ''' o deixava verde. Aqui o desfecho é olhado direto.
    ''' </summary>
    <TestMethod>
    Public Async Function Baixar_cancelado_diz_que_foi_interrompido() As Task
        Dim onde = PastaNova()
        Try
            Using dono = ParNovo()
                Dim corpo = MontarJson("2.0.0", quantos:=1_000_000)
                Dim servidor As RespostasDeVersao = Nothing
                Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                     New Version(1, 0, 0), servidor)

                Dim travado As New FluxoQueEspera()
                servidor.Fluxos(Base & "/Iris.exe") = travado

                Dim r = Await procura.Procurar(CancellationToken.None)
                Using parar As New CancellationTokenSource()
                    Dim baixando = procura.Baixar(r.Manifesto, onde, parar.Token)
                    Await travado.Comecou.Task
                    parar.Cancel()

                    Dim pacote = Await baixando
                    Assert.IsFalse(pacote.Veio)
                    StringAssert.Contains(pacote.Motivo, "interrompido",
                        "um cancelamento virou erro comum: " & pacote.Motivo)
                End Using
                CollectionAssert.AreEqual(Array.Empty(Of String)(), Directory.GetFiles(onde),
                                          "sobrou arquivo de um download cancelado")
            End Using
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Function

    ''' <summary>
    ''' E o <b>pacote</b> também: <c>ExigirHttpsAteOFim</c> está nos dois
    ''' caminhos de leitura, e o teste do manifesto só cobria um deles.
    ''' </summary>
    <TestMethod>
    Public Async Function Pacote_que_termina_fora_do_https_e_recusado() As Task
        Dim onde = PastaNova()
        Try
            Using dono = ParNovo()
                Dim conteudo = Encoding.UTF8.GetBytes("um executavel de mentira")
                Dim hash = Convert.ToHexString(SHA256.HashData(conteudo)).ToLowerInvariant()
                Dim corpo = MontarJson("2.0.0", hash:=hash, quantos:=conteudo.LongLength)

                Dim servidor As RespostasDeVersao = Nothing
                Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                     New Version(1, 0, 0), servidor)
                servidor.Corpos(Base & "/Iris.exe") = conteudo

                ' A procura corre normal; o redirecionamento so estraga na hora
                ' de buscar o PACOTE.
                Dim r = Await procura.Procurar(CancellationToken.None)
                Assert.AreEqual(DesfechoDaProcura.HaVersaoNova, r.Desfecho, r.Frase)

                servidor.EnderecoFinal = New Uri("http://exemplo.invalido/Iris.exe")
                Dim pacote = Await procura.Baixar(r.Manifesto, onde, CancellationToken.None)

                Assert.IsFalse(pacote.Veio, "baixou o pacote por um endereco http")
                CollectionAssert.AreEqual(Array.Empty(Of String)(), Directory.GetFiles(onde))
            End Using
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Function

End Class

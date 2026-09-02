Imports System.Collections.Generic
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
        Friend ReadOnly Pedidos As New List(Of String)()

        ''' <summary>Quando preenchido, toda requisição estoura com isto.</summary>
        Friend Property Explodir As Exception

        Protected Overrides Function SendAsync(pedido As HttpRequestMessage,
                                               ct As CancellationToken) _
                                               As Task(Of HttpResponseMessage)
            Dim onde = pedido.RequestUri.ToString()
            SyncLock Pedidos
                Pedidos.Add(onde)
            End SyncLock

            If Explodir IsNot Nothing Then Throw Explodir

            Dim corpo As Byte() = Nothing
            If Not Corpos.TryGetValue(onde, corpo) Then
                Return Task.FromResult(New HttpResponseMessage(HttpStatusCode.NotFound))
            End If

            Dim r As New HttpResponseMessage(HttpStatusCode.OK) With {
                .Content = New ByteArrayContent(corpo)
            }
            Return Task.FromResult(r)
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
        Using dono = ParNovo()
            Dim enorme(ManifestoDeVersao.ManifestoMaximo) As Byte
            Dim motivo As String = Nothing
            Assert.IsNull(ManifestoDeVersao.Ler(enorme, Assinar(enorme, dono),
                                                Publica(dono), motivo))
            StringAssert.Contains(motivo, "grande demais")
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
            servidor.Explodir = New HttpRequestException("sem rede")

            Dim r = Await procura.Procurar(CancellationToken.None)

            Assert.AreEqual(DesfechoDaProcura.NaoDeuParaSaber, r.Desfecho)
            ' A FRASE VAI PARA A TELA. A mensagem de uma excecao de rede carrega
            ' endereco e as vezes cabecalhos; so o tipo pode sair.
            Assert.IsFalse(r.Frase.Contains("exemplo.invalido"),
                           "a frase da tela vazou o endereco: " & r.Frase)
        End Using
    End Function

    <TestMethod>
    Public Async Function Endereco_de_versoes_precisa_ser_https() As Task
        Using dono = ParNovo()
            Dim procura = New ProcuraDeVersao(New HttpClient(New RespostasDeVersao()),
                                              "http://exemplo.invalido/iris.json",
                                              Publica(dono), New Version(1, 0, 0))

            Assert.AreEqual(DesfechoDaProcura.NaoDeuParaSaber,
                            (Await procura.Procurar(CancellationToken.None)).Desfecho)
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
                Dim conteudo = New Byte(200_000) {}
                Dim corpo = MontarJson("2.0.0", quantos:=1_000)

                Dim servidor As RespostasDeVersao = Nothing
                Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                     New Version(1, 0, 0), servidor)
                servidor.Corpos(Base & "/Iris.exe") = conteudo

                Dim r = Await procura.Procurar(CancellationToken.None)
                Dim pacote = Await procura.Baixar(r.Manifesto, onde, CancellationToken.None)

                Assert.IsFalse(pacote.Veio, "gravou mais do que o manifesto declarava")
                CollectionAssert.AreEqual(Array.Empty(Of String)(), Directory.GetFiles(onde))
            End Using
        Finally
            Directory.Delete(onde, recursive:=True)
        End Try
    End Function

    ' ==================================================================
    ' A tela
    ' ==================================================================

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
        Using dono = ParNovo()
            Dim corpo = MontarJson("2.0.0")
            Dim servidor As RespostasDeVersao = Nothing
            Dim procura = Montar(corpo, Assinar(corpo, dono), Publica(dono),
                                 New Version(1, 0, 0), servidor)
            Dim tela As New AtualizacaoViewModel(procura, PastaNova())

            Await tela.VerificarCommand.ExecuteAsync(Nothing)
            Assert.IsTrue(tela.HaVersaoNova, "nao achou a versao nova: " & tela.Frase)
            Assert.IsTrue(tela.BaixarCommand.CanExecute(Nothing))

            servidor.Explodir = New HttpRequestException("caiu")
            Await tela.VerificarCommand.ExecuteAsync(Nothing)

            Assert.IsFalse(tela.HaVersaoNova, "manteve a oferta de uma procura que falhou")
            Assert.IsFalse(tela.BaixarCommand.CanExecute(Nothing),
                           "o botao Baixar sobreviveu a uma procura sem resposta")
            Assert.AreEqual("", tela.Notas)
        End Using
    End Function

End Class

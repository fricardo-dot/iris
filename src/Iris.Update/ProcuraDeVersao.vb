Imports System.Collections.Generic
Imports System.IO
Imports System.Net.Http
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Threading

Namespace Global.Iris.Update

    ''' <summary>O que a procura por versões achou.</summary>
    Public Enum DesfechoDaProcura
        ''' <summary>Zero: nunca é sucesso.</summary>
        NaoDecidido = 0
        ''' <summary>Esta já é a versão publicada. Nada a fazer.</summary>
        JaEstaEmDia
        ''' <summary>Há versão nova, e ela está descrita.</summary>
        HaVersaoNova
        ''' <summary>Não deu para saber — rede, endereço, resposta ilegível.</summary>
        NaoDeuParaSaber
        ''' <summary>
        ''' O arquivo veio, e <b>não foi publicado por quem devia</b>. Isto não é
        ''' "não deu para saber": é uma resposta que existe e não confere.
        ''' </summary>
        AssinaturaNaoConfere
    End Enum

    ''' <summary>O resultado, com a frase que a tela mostra.</summary>
    Public NotInheritable Class ResultadoDaProcura
        Public ReadOnly Property Desfecho As DesfechoDaProcura
        Public ReadOnly Property Manifesto As ManifestoDeVersao
        ''' <summary>Em português, e sempre preenchida.</summary>
        Public ReadOnly Property Frase As String

        Friend Sub New(desfecho As DesfechoDaProcura, manifesto As ManifestoDeVersao,
                       frase As String)
            Me.Desfecho = desfecho
            Me.Manifesto = manifesto
            Me.Frase = If(frase, "")
        End Sub
    End Class

    ''' <summary>
    ''' O pacote, ou o motivo de não ter vindo.
    '''
    ''' Um tipo em vez de <c>ByRef</c>: método <c>Async</c> não aceita
    ''' <c>ByRef</c> em VB, e devolver <c>Nothing</c> com o motivo noutro canal
    ''' seria o caminho para alguém ler o caminho sem ler o motivo.
    ''' </summary>
    Public NotInheritable Class PacoteBaixado
        ''' <summary>Onde o arquivo ficou. Vazio quando não veio.</summary>
        Public ReadOnly Property Caminho As String
        ''' <summary>Em português. Vazio quando veio.</summary>
        Public ReadOnly Property Motivo As String

        Public ReadOnly Property Veio As Boolean
            Get
                Return Caminho.Length > 0
            End Get
        End Property

        Private Sub New(caminho As String, motivo As String)
            Me.Caminho = If(caminho, "")
            Me.Motivo = If(motivo, "")
        End Sub

        Friend Shared Function Sim(caminho As String) As PacoteBaixado
            Return New PacoteBaixado(caminho, "")
        End Function

        Friend Shared Function Nao(motivo As String) As PacoteBaixado
            Return New PacoteBaixado("", motivo)
        End Function
    End Class

    ''' <summary>
    ''' <b>Procura uma versão nova, e não instala nada.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ELE DELIBERADAMENTE NÃO FAZ</b>
    '''
    ''' Não troca o executável, não roda instalador, não reinicia o programa. Ele
    ''' baixa, confere e <b>diz onde está</b>.
    '''
    ''' Substituir sozinho é mais cômodo e exige um processo auxiliar que roda
    ''' fora do Iris, com permissão de escrever sobre o próprio executável — mais
    ''' uma peça a proteger, apontada para dentro de um programa que lê e-mail. A
    ''' comodidade não paga isso na primeira versão.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A ORDEM DAS CONFERÊNCIAS</b>
    '''
    ''' <list type="number">
    ''' <item>manifesto e assinatura, os dois com teto de tamanho;</item>
    ''' <item>a assinatura confere — antes de o JSON ser interpretado;</item>
    ''' <item>a versão é <b>maior</b> que a instalada — nunca igual, nunca
    ''' menor;</item>
    ''' <item>só então o pacote é baixado, e o SHA-256 dele tem de bater com o
    ''' que o manifesto <i>assinado</i> declara.</item>
    ''' </list>
    '''
    ''' <b>"Maior" e não "diferente"</b>: um manifesto antigo, legitimamente
    ''' assinado, pode ser reapresentado por quem controle o caminho da rede. Se
    ''' "diferente" bastasse, isso rebaixaria a instalação para uma versão com um
    ''' defeito já corrigido — assinada, e portanto aceita.
    ''' </summary>
    Public NotInheritable Class ProcuraDeVersao
        Implements IDisposable

        Private ReadOnly _cliente As HttpClient
        Private ReadOnly _enderecoDoManifesto As String
        Private ReadOnly _chavePublica As Byte()
        Private ReadOnly _instalada As Version

        ''' <summary>True quando o HttpClient nasceu aqui dentro.</summary>
        Private ReadOnly _meu As Boolean

        ''' <summary>
        ''' <b>A versão que este executável é.</b>
        '''
        ''' Lida do <c>InformationalVersion</c>, que é o que o
        ''' <c>Directory.Build.props</c> escreve. Sufixos como <c>+abc123</c> que
        ''' o SDK acrescenta são cortados antes de interpretar.
        ''' </summary>
        Public Shared Function VersaoInstalada() As Version
            Try
                Dim bruta = Assembly.GetEntryAssembly()?.
                    GetCustomAttribute(Of AssemblyInformationalVersionAttribute)()?.
                    InformationalVersion

                If String.IsNullOrWhiteSpace(bruta) Then Return New Version(0, 0, 0)

                Dim corte = bruta.IndexOfAny({"+"c, "-"c})
                If corte > 0 Then bruta = bruta.Substring(0, corte)

                Dim v As Version = Nothing
                Return If(Version.TryParse(bruta, v), v, New Version(0, 0, 0))
            Catch
                Return New Version(0, 0, 0)
            End Try
        End Function

        ''' <summary>
        ''' <b>O construtor da produção — e ele monta o próprio HttpClient.</b>
        '''
        ''' Não é conveniência: montá-lo no composition root faria o
        ''' <c>Iris.App</c> referenciar <c>System.Net.Http</c> diretamente, e a
        ''' lista de quem pode abrir socket passaria de dois para três. Foi
        ''' exatamente o que aconteceu na primeira tentativa, e quem contou foi o
        ''' <c>EgressArquiteturaTests</c>.
        '''
        ''' A capacidade de rede fica onde ela é auditada. O que o composition
        ''' root escolhe é <i>se</i> monta um destes, e com que chave.
        '''
        ''' <b>Tempo curto de propósito</b>: perguntar a versão é um clique numa
        ''' tela aberta. Se o servidor não responde em meio minuto, a resposta
        ''' útil é "não deu para saber", e não um painel travado.
        ''' </summary>
        Public Sub New(enderecoDoManifesto As String, chavePublica As Byte(),
                       Optional instalada As Version = Nothing)
            Me.New(New HttpClient() With {.Timeout = TimeSpan.FromSeconds(30)},
                   enderecoDoManifesto, chavePublica, instalada)
            _meu = True
        End Sub

        ''' <summary>
        ''' <b>Friend</b>, e é o caminho dos testes: eles trocam o handler para
        ''' responder sem socket. Público, ele seria a porta por onde um
        ''' <c>HttpClient</c> de fora entraria — e a auditoria de quem tem rede
        ''' passaria a depender de quem chama, e não de quem referencia.
        ''' </summary>
        Friend Sub New(cliente As HttpClient, enderecoDoManifesto As String,
                       chavePublica As Byte(), Optional instalada As Version = Nothing)
            If cliente Is Nothing Then Throw New ArgumentNullException(NameOf(cliente))
            _cliente = cliente
            _enderecoDoManifesto = If(enderecoDoManifesto, "")
            _chavePublica = chavePublica
            _instalada = If(instalada, VersaoInstalada())
        End Sub

        ''' <summary>
        ''' Só descarta o cliente que ele mesmo montou. O que veio de fora é de
        ''' quem o passou — descartá-lo aqui derrubaria um cliente compartilhado
        ''' que o dono ainda usa.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            If _meu Then _cliente.Dispose()
        End Sub

        ''' <summary>
        ''' Pergunta se há versão nova. Nunca lança: devolve desfecho.
        ''' </summary>
        Public Async Function Procurar(ct As CancellationToken) As Task(Of ResultadoDaProcura)
            If Not _enderecoDoManifesto.StartsWith("https://",
                                                   StringComparison.OrdinalIgnoreCase) Then
                Return Parar(DesfechoDaProcura.NaoDeuParaSaber,
                             "o endereço de versões não é https")
            End If

            Dim manifesto As Byte()
            Dim assinatura As Byte()
            Try
                manifesto = Await Baixar(_enderecoDoManifesto,
                                         ManifestoDeVersao.ManifestoMaximo, ct)
                assinatura = Await Baixar(_enderecoDoManifesto & ".sig",
                                          ManifestoDeVersao.ManifestoMaximo, ct)
            Catch ex As OperationCanceledException
                Return Parar(DesfechoDaProcura.NaoDeuParaSaber, "a procura foi interrompida")
            Catch ex As Exception
                ' SO O TIPO. A mensagem de uma excecao de rede carrega o endereco
                ' e as vezes cabecalhos, e esta frase vai para a tela.
                Return Parar(DesfechoDaProcura.NaoDeuParaSaber,
                             "não consegui falar com o servidor de versões (" &
                             ex.GetType().Name & ")")
            End Try

            Dim motivo As String = Nothing
            Dim lido = ManifestoDeVersao.Ler(manifesto, assinatura, _chavePublica, motivo)
            If lido Is Nothing Then
                ' ASSINATURA ERRADA NAO E "NAO DEU PARA SABER". Uma e a rede
                ' falhando; a outra e um arquivo que existe e nao e seu, e o dono
                ' tem de ver a diferenca.
                Dim desfecho = If(motivo.Contains("assinatura"),
                                  DesfechoDaProcura.AssinaturaNaoConfere,
                                  DesfechoDaProcura.NaoDeuParaSaber)
                Return Parar(desfecho, motivo)
            End If

            If lido.Versao <= _instalada Then
                Return New ResultadoDaProcura(
                    DesfechoDaProcura.JaEstaEmDia, lido,
                    $"Você já está na versão {_instalada.ToString(3)}.")
            End If

            Return New ResultadoDaProcura(
                DesfechoDaProcura.HaVersaoNova, lido,
                $"Há a versão {lido.Versao.ToString(3)} — você tem a " &
                $"{_instalada.ToString(3)}.")
        End Function

        ''' <summary>
        ''' <b>Baixa o pacote e confere o SHA-256 contra o manifesto assinado.</b>
        '''
        ''' Grava num arquivo temporário e só o renomeia ao final: um download
        ''' interrompido não pode deixar meio pacote com nome de pacote inteiro —
        ''' é o mesmo desenho de <c>MessageReading.GravarComTemporario</c>, e pelo
        ''' mesmo motivo.
        '''
        ''' Devolve o caminho, ou <c>Nothing</c> com o motivo.
        ''' </summary>
        Public Async Function Baixar(manifesto As ManifestoDeVersao, pasta As String,
                                     ct As CancellationToken) As Task(Of PacoteBaixado)
            If manifesto Is Nothing Then
                Return PacoteBaixado.Nao("não há versão para baixar")
            End If

            Dim destino = Path.Combine(pasta, $"Iris-{manifesto.Versao.ToString(3)}.exe")
            Dim temporario = destino & ".parcial"

            Try
                Directory.CreateDirectory(pasta)

                ' PARA O ARQUIVO, E NAO PARA A MEMORIA.
                '
                ' Materializar o pacote inteiro custaria o dobro do tamanho dele
                ' em RAM -- o buffer e a copia que ToArray devolve -- e o teto
                ' aqui e de 300 MB. O hash sai da mesma passada, incremental, de
                ' forma que os bytes conferidos sao exatamente os gravados.
                Dim veio = Await BaixarPara(temporario, manifesto.Endereco,
                                            manifesto.Bytes, ct)

                Dim gravados = New FileInfo(temporario).Length
                If gravados <> manifesto.Bytes Then
                    Limpar(temporario)
                    Return PacoteBaixado.Nao(
                        "o pacote não tem o tamanho que o manifesto declara")
                End If

                If Not String.Equals(veio, manifesto.Sha256, StringComparison.Ordinal) Then
                    ' O PACOTE NAO E O QUE FOI ASSINADO. O manifesto conferiu; o
                    ' arquivo que ele descreve, nao. Alguem trocou o pacote, ou o
                    ' download veio corrompido -- e os dois levam a mesma acao.
                    '
                    ' E O PARCIAL MORRE AQUI: um executavel nao conferido nao pode
                    ' sobreviver ao fim desta funcao, porque quem olhar a pasta
                    ' depois nao tem como saber que ele foi recusado.
                    Limpar(temporario)
                    Return PacoteBaixado.Nao(
                        "o pacote baixado não é o que o manifesto assinado descreve")
                End If

                If File.Exists(destino) Then File.Delete(destino)
                File.Move(temporario, destino)
                Return PacoteBaixado.Sim(destino)
            Catch ex As OperationCanceledException
                Limpar(temporario)
                Return PacoteBaixado.Nao("o download foi interrompido")
            Catch ex As Exception
                Limpar(temporario)
                Return PacoteBaixado.Nao("não consegui baixar (" & ex.GetType().Name & ")")
            End Try
        End Function

        ''' <summary>
        ''' Lê no máximo <paramref name="teto"/> bytes. O teto é conferido
        ''' <b>durante</b> a leitura, e não depois: um servidor que ignore o
        ''' <c>Content-Length</c> pode mandar para sempre.
        ''' </summary>
        Private Async Function Baixar(endereco As String, teto As Long,
                                      ct As CancellationToken) As Task(Of Byte())
            Using resposta = Await _cliente.GetAsync(
                    endereco, HttpCompletionOption.ResponseHeadersRead, ct)
                resposta.EnsureSuccessStatusCode()

                Using entrada = Await resposta.Content.ReadAsStreamAsync(ct)
                    Using saida As New MemoryStream()
                        Dim buffer(81_919) As Byte
                        While True
                            Dim lidos = Await entrada.ReadAsync(buffer, ct)
                            If lidos = 0 Then Exit While
                            If saida.Length + lidos > teto Then
                                Throw New InvalidDataException("resposta maior que o teto")
                            End If
                            saida.Write(buffer, 0, lidos)
                        End While
                        Return saida.ToArray()
                    End Using
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Grava a resposta em <paramref name="destino"/> e devolve o SHA-256 do
        ''' que gravou, em hexadecimal minúsculo.
        '''
        ''' Mesmo teto durante a leitura da irmã acima, e pelo mesmo motivo: o
        ''' <c>Content-Length</c> é uma promessa do servidor, não um limite.
        ''' </summary>
        Private Async Function BaixarPara(destino As String, endereco As String,
                                          teto As Long,
                                          ct As CancellationToken) As Task(Of String)
            Using resposta = Await _cliente.GetAsync(
                    endereco, HttpCompletionOption.ResponseHeadersRead, ct)
                resposta.EnsureSuccessStatusCode()

                Using hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256)
                    Using entrada = Await resposta.Content.ReadAsStreamAsync(ct)
                        Using saida As New FileStream(destino, FileMode.Create,
                                                      FileAccess.Write, FileShare.None)
                            Dim buffer(81_919) As Byte
                            Dim total As Long = 0
                            While True
                                Dim lidos = Await entrada.ReadAsync(buffer, ct)
                                If lidos = 0 Then Exit While
                                total += lidos
                                If total > teto Then
                                    Throw New InvalidDataException("resposta maior que o teto")
                                End If
                                hash.AppendData(buffer, 0, lidos)
                                Await saida.WriteAsync(buffer.AsMemory(0, lidos), ct)
                            End While
                        End Using
                    End Using

                    Return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant()
                End Using
            End Using
        End Function

        Private Shared Sub Limpar(caminho As String)
            Try
                If File.Exists(caminho) Then File.Delete(caminho)
            Catch
            End Try
        End Sub

        Private Shared Function Parar(desfecho As DesfechoDaProcura,
                                      frase As String) As ResultadoDaProcura
            Return New ResultadoDaProcura(desfecho, Nothing, frase)
        End Function

    End Class

End Namespace

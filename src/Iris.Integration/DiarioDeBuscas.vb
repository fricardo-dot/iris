Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Threading

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>O que foi procurado, para o oráculo se juntar sozinho.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO EXISTE</b>
    '''
    ''' A metade aberta da Fase 4 — busca por sentido — precisa de um oráculo
    ''' que só o dono da caixa tem: <i>qual mensagem eu queria quando digitei
    ''' isto</i>. Em 30/08/2026 eu pedi a ele de memória, e não veio nenhum caso
    ''' — o que é a resposta normal de quem é perguntado sobre uma busca que
    ''' falhou semanas atrás.
    '''
    ''' Pedir de memória é o método fraco. Este é o forte: o oráculo se junta
    ''' sozinho, com a distribuição real das buscas em vez da amostra do que
    ''' alguém lembrou.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O SINAL É A REFORMULAÇÃO, E NÃO O CLIQUE</b>
    '''
    ''' A tela de busca não abre resultado — ela só mostra. Isso parecia uma
    ''' limitação e é uma vantagem: o sinal que interessa não é <i>qual
    ''' mensagem foi aberta</i>, é <b>o usuário ter reformulado</b>.
    '''
    ''' Digitar "cobrança", não achar, digitar "fatura" e parar — esse par
    ''' <b>é</b> a falha semântica, inteira, sem precisar saber que mensagem
    ''' era. E de quebra o diário não registra assunto de mensagem nenhuma.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ELE GUARDA — E POR QUE "SÓ O TERMO" NÃO É CONFORTO</b>
    '''
    ''' Guarda o instante, o termo <b>como foi digitado</b>, e quantos achados
    ''' saíram. Não guarda assunto, remetente, <c>EntryID</c> nem pasta, e o
    ''' <see cref="Registrar"/> não tem parâmetro onde isso caberia.
    '''
    ''' <b>Mas o termo é dado sensível, e a primeira versão desta tela dizia
    ''' "só o texto digitado" como se aquilo tranquilizasse.</b> Numa caixa
    ''' corporativa, o que se digita numa busca é nome de pessoa, número de
    ''' contrato, valor, cliente. O arquivo é <b>texto claro</b>, fica
    ''' <b>indefinidamente</b>, e qualquer processo rodando como o usuário — ou
    ''' qualquer ferramenta corporativa de backup ou administração — pode lê-lo.
    '''
    ''' A revisão externa chamou isso de consentimento mal informado, e estava
    ''' certa: a frase era literalmente verdadeira e escolhida para acalmar.
    ''' Agora a tela diz o risco.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E DÁ PARA DESLIGAR — apagar não é a mesma coisa</b>
    '''
    ''' Apagar tira o que já foi coletado; a busca seguinte recria o arquivo.
    ''' Coleta contínua de comportamento precisa de <b>retirada de
    ''' consentimento</b>, e não só de faxina — e a primeira versão desta classe
    ''' só tinha faxina.
    '''
    ''' <see cref="Desligar"/> grava um marcador ao lado do diário, e ele
    ''' sobrevive ao fechamento do programa. Enquanto existir, nada é anotado.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E ELE NUNCA DERRUBA A BUSCA — DOS DOIS LADOS</b>
    '''
    ''' Aqui as falhas são engolidas, e o motivo fica em
    ''' <see cref="UltimaFalha"/> para a tela mostrar.
    '''
    ''' <b>Isso não basta, e a revisão externa mostrou por quê:</b> um contrato
    ''' de "nunca lança" que só esta classe cumpre deixa de valer no dia em que
    ''' alguém trocar a implementação. A barreira de verdade está no
    ''' <c>BuscaViewModel</c>, que protege <b>toda</b> chamada ao diário.
    ''' </summary>
    Public Interface IDiarioDeBuscas
        ''' <summary>Anota uma busca. Não faz nada se estiver desligado.</summary>
        Sub Registrar(termo As String, exatos As Integer, aproximados As Integer)

        ''' <summary>Onde o arquivo está, para a tela poder dizer.</summary>
        ReadOnly Property Caminho As String

        ''' <summary>Quantas buscas o arquivo tem, ou <c>Nothing</c> se não deu para contar.</summary>
        Function Quantas() As Integer?

        ''' <summary>Apaga o que já foi coletado. Devolve o motivo da falha, ou <c>Nothing</c>.</summary>
        Function Apagar() As String

        ''' <summary>O motivo da última falha, ou vazio.</summary>
        ReadOnly Property UltimaFalha As String

        ''' <summary>
        ''' A coleta está ligada?
        '''
        ''' <b>Separada de <see cref="Apagar"/> de propósito.</b> Apagar é
        ''' faxina; desligar é retirar consentimento. Confundir as duas é o que
        ''' faz um botão "apagar" parecer que resolve — e ele não resolve: a
        ''' busca seguinte recria o arquivo.
        ''' </summary>
        ReadOnly Property Ligado As Boolean

        ''' <summary>Para de anotar. Persiste entre execuções.</summary>
        Function Desligar() As String

        ''' <summary>Volta a anotar.</summary>
        Function Ligar() As String
    End Interface

    ''' <summary>O diário de verdade, em arquivo.</summary>
    Public NotInheritable Class DiarioDeBuscasEmArquivo
        Implements IDiarioDeBuscas

        Private ReadOnly _caminho As String
        Private ReadOnly _marcador As String
        Private ReadOnly _agora As Func(Of DateTimeOffset)
        Private ReadOnly _trava As New Object()

        ''' <summary>
        ''' <b>Trava entre PROCESSOS, e não só entre threads.</b>
        '''
        ''' O <c>SyncLock</c> protege uma instância. Duas janelas do Iris abertas
        ''' são dois processos, e no Windows dois <c>AppendAllText</c>
        ''' concorrentes não intercalam bytes — um deles <b>falha</b> por
        ''' compartilhamento. Com a exceção engolida, o resultado seria uma
        ''' amostra com buraco que ninguém veria.
        '''
        ''' O nome inclui o caminho, com os separadores trocados: nome de mutex
        ''' não aceita <c>\</c>, e dois diários em arquivos diferentes não
        ''' precisam esperar um pelo outro.
        ''' </summary>
        Private ReadOnly _entreProcessos As Mutex

        Private _ultimaFalha As String = ""

        ''' <summary>
        ''' Contagem em memória, para a tela não varrer o arquivo a cada busca.
        '''
        ''' <c>Nothing</c> é "ainda não contei" ou "não consegui contar" — e
        ''' nunca zero, que é "contei e não há".
        ''' </summary>
        Private _contagem As Integer?

        Public Sub New(Optional caminho As String = Nothing,
                       Optional agora As Func(Of DateTimeOffset) = Nothing)
            _caminho = If(caminho, CaminhoPadrao())
            _marcador = _caminho & ".desligado"
            _agora = If(agora, Function() DateTimeOffset.Now)
            _entreProcessos = New Mutex(initiallyOwned:=False,
                                        name:="Iris.buscas." &
                                              _caminho.Replace("\"c, "_"c).
                                                       Replace("/"c, "_"c).
                                                       Replace(":"c, "_"c))
        End Sub

        ''' <summary>
        ''' Ao lado do cache, e não ao lado do executável — o executável pode
        ''' estar em Program Files, onde escrever exige elevação.
        ''' </summary>
        Public Shared Function CaminhoPadrao() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "buscas.jsonl")
        End Function

        Public ReadOnly Property Caminho As String Implements IDiarioDeBuscas.Caminho
            Get
                Return _caminho
            End Get
        End Property

        Public ReadOnly Property UltimaFalha As String Implements IDiarioDeBuscas.UltimaFalha
            Get
                SyncLock _trava
                    Return _ultimaFalha
                End SyncLock
            End Get
        End Property

        ''' <summary>
        ''' Ligado enquanto não houver marcador. <b>Falha ao conferir vale como
        ''' DESLIGADO</b> — não conseguir ler o consentimento não é autorização
        ''' para coletar.
        ''' </summary>
        Public ReadOnly Property Ligado As Boolean Implements IDiarioDeBuscas.Ligado
            Get
                Try
                    Return Not File.Exists(_marcador)
                Catch
                    Return False
                End Try
            End Get
        End Property

        Public Function Desligar() As String Implements IDiarioDeBuscas.Desligar
            Try
                Directory.CreateDirectory(Path.GetDirectoryName(_marcador))
                File.WriteAllText(_marcador,
                    "Enquanto este arquivo existir, o Iris nao anota buscas." &
                    Environment.NewLine &
                    "Apague-o, ou use o botao na tela, para voltar a anotar." &
                    Environment.NewLine, Encoding.UTF8)
                Return Nothing
            Catch ex As Exception
                Return "não consegui desligar o registro (" & ex.GetType().Name & ")"
            End Try
        End Function

        Public Function Ligar() As String Implements IDiarioDeBuscas.Ligar
            Try
                If File.Exists(_marcador) Then File.Delete(_marcador)
                Return Nothing
            Catch ex As Exception
                Return "não consegui ligar o registro (" & ex.GetType().Name & ")"
            End Try
        End Function

        ''' <summary>
        ''' Uma linha, anexada. <b>Append e não reescrita</b>: uma queda no meio
        ''' perde no máximo a última linha, e o leitor pula linha quebrada.
        ''' Reescrever o arquivo inteiro a cada busca poria em risco tudo o que
        ''' já foi coletado, para gravar uma linha.
        ''' </summary>
        Public Sub Registrar(termo As String, exatos As Integer, aproximados As Integer) _
            Implements IDiarioDeBuscas.Registrar

            ' DESLIGADO NAO ANOTA. Conferido a cada busca, e nao guardado em
            ' campo: o marcador pode ser criado ou apagado por fora, e um
            ' consentimento em cache seria o consentimento de ontem.
            If Not Ligado Then Return

            ' TERMO VAZIO NAO E BUSCA. Registrar o Enter dado num campo em
            ' branco encheria o arquivo de linha que nao ensina nada.
            If String.IsNullOrWhiteSpace(termo) Then Return

            Dim segurou = False
            Try
                Dim linha = JsonSerializer.Serialize(New With {
                    .quando = _agora().ToString("o", CultureInfo.InvariantCulture),
                    .termo = termo,
                    .exatos = exatos,
                    .aproximados = aproximados})

                SyncLock _trava
                    Try
                        segurou = _entreProcessos.WaitOne(TimeSpan.FromSeconds(2))
                    Catch ex As AbandonedMutexException
                        ' Outro processo morreu segurando a trava. O arquivo
                        ' pode ter meia linha; o leitor pula linha quebrada, e
                        ' o append continua valendo. Seguir e melhor que parar.
                        segurou = True
                    End Try

                    Directory.CreateDirectory(Path.GetDirectoryName(_caminho))
                    File.AppendAllText(_caminho, linha & Environment.NewLine, Encoding.UTF8)

                    If _contagem.HasValue Then _contagem = _contagem.Value + 1
                    _ultimaFalha = ""
                End SyncLock
            Catch ex As Exception
                ' Falha de escrita nao derruba a busca -- mas o motivo fica, e
                ' a contagem perde a confianca: uma linha que nao entrou faz o
                ' numero em memoria divergir do arquivo.
                SyncLock _trava
                    _ultimaFalha = "não consegui anotar a busca (" & ex.GetType().Name & ")"
                    _contagem = Nothing
                End SyncLock
            Finally
                If segurou Then
                    Try
                        _entreProcessos.ReleaseMutex()
                    Catch
                    End Try
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Quantas linhas o arquivo tem. <c>Nothing</c> é "não consegui
        ''' contar", e não zero — arquivo travado não é arquivo vazio.
        '''
        ''' <b>Varre o arquivo uma vez e guarda.</b> A tela pergunta isto a cada
        ''' busca, e varrer um arquivo que só cresce, na thread da interface,
        ''' ficaria mais caro a cada dia — a revisão externa apontou isso antes
        ''' de acontecer.
        ''' </summary>
        Public Function Quantas() As Integer? Implements IDiarioDeBuscas.Quantas
            SyncLock _trava
                If _contagem.HasValue Then Return _contagem
            End SyncLock

            Dim lido As Integer?
            Try
                ' AUSENTE E ILEGIVEL NAO SAO A MESMA COISA, e o File.Exists
                ' sozinho colapsava as duas: ele devolve False tanto para "nao
                ' existe" -- que e zero de verdade -- quanto para "existe e nao
                ' e arquivo", que e "nao consegui ler". Foi o teste que pegou.
                If Directory.Exists(_caminho) Then
                    lido = Nothing
                ElseIf Not File.Exists(_caminho) Then
                    lido = 0
                Else
                    lido = File.ReadLines(_caminho).
                           Count(Function(l) Not String.IsNullOrWhiteSpace(l))
                End If
            Catch
                lido = Nothing
            End Try

            SyncLock _trava
                _contagem = lido
            End SyncLock
            Return lido
        End Function

        Public Function Apagar() As String Implements IDiarioDeBuscas.Apagar
            Dim segurou = False
            Try
                SyncLock _trava
                    Try
                        segurou = _entreProcessos.WaitOne(TimeSpan.FromSeconds(2))
                    Catch ex As AbandonedMutexException
                        segurou = True
                    End Try

                    If File.Exists(_caminho) Then File.Delete(_caminho)
                    _contagem = 0
                    _ultimaFalha = ""
                End SyncLock
                Return Nothing
            Catch ex As Exception
                SyncLock _trava
                    _contagem = Nothing
                End SyncLock
                Return "não consegui apagar (" & ex.GetType().Name & ")"
            Finally
                If segurou Then
                    Try
                        _entreProcessos.ReleaseMutex()
                    Catch
                    End Try
                End If
            End Try
        End Function
    End Class

End Namespace

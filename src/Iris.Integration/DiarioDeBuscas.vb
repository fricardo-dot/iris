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
        Implements IDiarioDeBuscas, IDisposable

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

        ''' <summary>
        ''' O carimbo do arquivo quando a contagem foi feita: tamanho e
        ''' instante da última escrita.
        '''
        ''' <b>O cache sem carimbo era cache cego.</b> Duas janelas do Iris, ou
        ''' uma edição por fora, deixavam a tela mostrando um número de antes —
        ''' e a revisão externa apontou os três casos: outra instância anexa,
        ''' outra instância apaga, alguém edita o arquivo. Um <c>stat</c> é
        ''' barato; varrer o arquivo é que não era.
        '''
        ''' <b>E o que ele AINDA não pega, dito em vez de escondido:</b> uma
        ''' substituição por conteúdo de <i>mesmo tamanho</i> dentro da
        ''' resolução do relógio do sistema de arquivos. A data de criação foi
        ''' acrescentada e pega o apagar-e-recriar; edição no lugar, com o mesmo
        ''' tamanho e no mesmo instante, escaparia. Pegar isso exigiria somar o
        ''' conteúdo — que é justamente a varredura que este carimbo existe para
        ''' evitar. O uso normal é só append, e append muda o tamanho.
        ''' </summary>
        Private _carimbo As (Tamanho As Long, Quando As DateTime, Nasceu As DateTime)?

        Public Sub New(Optional caminho As String = Nothing,
                       Optional agora As Func(Of DateTimeOffset) = Nothing)
            _caminho = If(caminho, CaminhoPadrao())
            _marcador = _caminho & ".desligado"
            _agora = If(agora, Function() DateTimeOffset.Now)
            ' NOME CURTO E DETERMINISTICO. O caminho inteiro sanitizado
            ' podia estourar o limite de nome de objeto do kernel e fazer o
            ' CONSTRUTOR lancar -- e ai o diario derrubaria a aplicacao no
            ' arranque, que e o oposto do que ele promete. Os testes usam
            ' caminho temporario curto e nunca teriam pego.
            ' O MUTEX PODE NAO NASCER, E ISSO NAO DERRUBA O PROGRAMA.
            '
            ' O nome curto resolveu o limite do kernel; nao resolve ACL. Num
            ' perfil corporativo, o namespace de objetos do Windows pode recusar
            ' a criacao com UnauthorizedAccessException -- e o construtor
            ' lancando aqui derruba a composicao da janela inteira, que e o
            ' oposto do que este diario promete. Achado por revisao externa em
            ' 02/09/2026.
            '
            ' Sem ele, a trava entre PROCESSOS some e sobra a de dentro do
            ' processo. Duas instancias do Iris podem entao intercalar escritas
            ' no arquivo de buscas -- que e degradacao, e nao perda: o pior caso
            ' e uma linha malformada, que a leitura ja descarta. Nao abrir seria
            ' pior que isso.
            Try
                Using sha = Security.Cryptography.SHA256.Create()
                    Dim bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(_caminho.ToLowerInvariant()))
                    _entreProcessos = New Mutex(initiallyOwned:=False,
                                                name:="Iris.buscas." &
                                                      Convert.ToHexString(bytes, 0, 12))
                End Using
            Catch
                _entreProcessos = Nothing
            End Try
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

        ''' <summary>
        ''' <b>Desligar entra na MESMA trava que gravar.</b>
        '''
        ''' Sem isto havia uma corrida real, e a revisão externa a classificou
        ''' como ALTO: o <c>Registrar</c> conferia <c>Ligado</c>, esperava o
        ''' mutex, e nesse meio-tempo outro processo desligava — a busca era
        ''' anotada <b>depois</b> de o dono retirar o consentimento.
        '''
        ''' Uma retirada de consentimento que admite "mais uma" não é uma
        ''' retirada.
        ''' </summary>
        Public Function Desligar() As String Implements IDiarioDeBuscas.Desligar
            Dim segurou = False
            Dim porque As String = Nothing
            Try
                segurou = Segurar(porque)

                ' SEM A TRAVA, NAO DESLIGA -- e esta linha faltava.
                '
                ' O primeiro corte chamava Segurar e IGNORAVA o retorno, e a
                ' revisao externa achou a corrida pelo outro lado: com a espera
                ' estourada, o Desligar criava o marcador sem exclusao e
                ' devolvia SUCESSO, enquanto o Registrar do outro processo --
                ' que ja tinha conferido Ligado -- seguia e gravava.
                '
                ' Dizer "desliguei" sem ter desligado e pior que recusar: quem
                ' recusa tenta de novo, quem foi enganado nao tenta.
                If Not segurou Then
                    Return "não consegui desligar agora (" & porque & "). Tente de novo."
                End If

                Directory.CreateDirectory(Path.GetDirectoryName(_marcador))
                File.WriteAllText(_marcador,
                    "Enquanto este arquivo existir, o Iris nao anota buscas." &
                    Environment.NewLine &
                    "Apague-o, ou use o botao na tela, para voltar a anotar." &
                    Environment.NewLine, Encoding.UTF8)
                Return Nothing
            Catch ex As Exception
                Return "não consegui desligar o registro (" & ex.GetType().Name & ")"
            Finally
                Soltar(segurou)
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
        ''' Pega a trava entre processos, ou devolve <c>False</c>.
        '''
        ''' <b>Espera curta, e recusa em vez de forçar.</b> Ela roda na thread
        ''' da interface: dois segundos de espera — o valor da primeira versão —
        ''' seriam dois segundos de janela congelada por causa de uma anotação.
        ''' Anexar uma linha leva microssegundos; 250 ms já é folga enorme.
        '''
        ''' E quem chama <b>tem</b> de olhar o retorno. A primeira versão
        ''' ignorava o <c>False</c> e gravava assim mesmo — ou seja, a exclusão
        ''' entre processos sumia exatamente quando havia disputa, que é a única
        ''' hora em que ela serve.
        ''' </summary>
        Private Function Segurar(ByRef motivo As String) As Boolean
            motivo = Nothing

            ' SEM MUTEX, SEGUE SEM ELE. Ver o construtor: num perfil que recusa
            ' criar o objeto, a trava entre processos nao existe e resta a de
            ' dentro do processo. Degradacao declarada, e nao silencio.
            If _entreProcessos Is Nothing Then Return True

            Try
                If _entreProcessos.WaitOne(TimeSpan.FromMilliseconds(250)) Then Return True
                motivo = "outra janela do Iris estava escrevendo"
                Return False
            Catch ex As AbandonedMutexException
                ' Outro processo morreu segurando a trava. No .NET isto quer
                ' dizer que ESTA thread adquiriu -- entao o Release depois esta
                ' certo. O arquivo pode ter meia linha; o leitor pula linha
                ' quebrada, e o append continua valendo.
                Return True
            Catch ex As Exception
                ' TIMEOUT E FALHA DE INFRAESTRUTURA NAO SAO A MESMA COISA.
                ' O primeiro corte devolvia False para os dois, e o chamador
                ' dizia "outra janela estava escrevendo" sobre um mutex
                ' descartado ou um handle quebrado -- diagnostico errado no
                ' lugar onde a pessoa vai procurar o motivo.
                motivo = "a trava do registro falhou (" & ex.GetType().Name & ")"
                Return False
            End Try
        End Function

        Private Sub Soltar(segurou As Boolean)
            If Not segurou Then Return
            If _entreProcessos Is Nothing Then Return
            Try
                _entreProcessos.ReleaseMutex()
            Catch
            End Try
        End Sub

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
                    Dim porque As String = Nothing
                    segurou = Segurar(porque)

                    ' NAO CONSEGUIU A TRAVA: RECUSA, e nao grava solto.
                    ' Uma busca nao anotada e um buraco na amostra, e o leitor
                    ' o mostra como linha ilegivel ou ausencia; uma escrita
                    ' concorrente sem trava e corrupcao silenciosa.
                    If Not segurou Then
                        _ultimaFalha = "não consegui anotar a busca (" & porque & ")"
                        _contagem = Nothing
                        Return
                    End If

                    ' O CONSENTIMENTO E CONFERIDO DENTRO DA TRAVA, e nao so na
                    ' entrada. Entre a conferencia de fora e este ponto, outro
                    ' processo pode ter desligado -- e anotar depois disso e
                    ' anotar sem consentimento.
                    '
                    ' A janela e estreita demais para um teste de um processo so
                    ' abrir sozinho -- o primeiro controle negativo desta guarda
                    ' PASSOU por isso. O ponto de injecao a abre.
                    Iris.Cache.CrashInjection.Talvez(
                        Iris.Cache.CrashInjection.EntreConferirOConsentimentoEGravar)
                    If Not Ligado Then Return

                    Directory.CreateDirectory(Path.GetDirectoryName(_caminho))
                    File.AppendAllText(_caminho, linha & Environment.NewLine, Encoding.UTF8)

                    _contagem = Nothing   ' recontar; outra janela pode ter escrito
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
                Soltar(segurou)
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
            Dim atual = CarimboAtual()

            SyncLock _trava
                If _contagem.HasValue AndAlso _carimbo.HasValue AndAlso
                   atual.HasValue AndAlso _carimbo.Value.Equals(atual.Value) Then
                    Return _contagem
                End If
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
                _carimbo = CarimboAtual()
            End SyncLock
            Return lido
        End Function

        ''' <summary>
        ''' Tamanho e instante da última escrita, ou <c>Nothing</c> se o arquivo
        ''' não existe ou não se deixa olhar. Arquivo ausente tem carimbo
        ''' próprio — senão "sumiu" e "não consegui ver" ficariam iguais.
        ''' </summary>
        Private Function CarimboAtual() As (Tamanho As Long, Quando As DateTime, Nasceu As DateTime)?
            Try
                Dim fi As New FileInfo(_caminho)
                If Not fi.Exists Then Return (-1L, DateTime.MinValue, DateTime.MinValue)
                Return (fi.Length, fi.LastWriteTimeUtc, fi.CreationTimeUtc)
            Catch
                Return Nothing
            End Try
        End Function

        Public Function Apagar() As String Implements IDiarioDeBuscas.Apagar
            Dim segurou = False
            Try
                SyncLock _trava
                    Dim porque As String = Nothing
                    segurou = Segurar(porque)
                    If Not segurou Then
                        Return "não consegui apagar agora (" & porque & "). Tente de novo."
                    End If

                    If File.Exists(_caminho) Then File.Delete(_caminho)
                    _contagem = Nothing
                    _ultimaFalha = ""
                End SyncLock
                Return Nothing
            Catch ex As Exception
                SyncLock _trava
                    _contagem = Nothing
                End SyncLock
                Return "não consegui apagar (" & ex.GetType().Name & ")"
            Finally
                Soltar(segurou)
            End Try
        End Function

        ''' <summary>
        ''' O mutex é um handle do sistema. Uma instância de vida longa não
        ''' pesa; a suíte cria dezenas, e handle acumulado até a finalização é
        ''' sujeira que ninguém vê até ver.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            _entreProcessos?.Dispose()
        End Sub
    End Class

End Namespace

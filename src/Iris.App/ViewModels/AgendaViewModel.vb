Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Globalization
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>A agenda dos próximos dias.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>OS PRÓXIMOS SETE DIAS, E NÃO "O CALENDÁRIO"</b>
    '''
    ''' A medição de 28/08/2026 contou <b>434 compromissos</b> nesta caixa e
    ''' <b>30,9 ms por item</b> — quase o dobro dos ~16 ms que a Fase 0 mediu
    ''' por mensagem. Ler o calendário inteiro custaria ~13 s numa chamada só,
    ''' na fila única da STA, com a janela parada.
    '''
    ''' Uma janela curta não é uma limitação da tela: é a única forma de a tela
    ''' existir sem o cache que a Fase 6 ainda não tem.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ISTO NÃO É UM CACHE — E "AO VIVO" NÃO QUER DIZER "COMPLETO"</b>
    '''
    ''' A agenda lê ao vivo do Outlook, como a lista de mensagens e ao contrário
    ''' do acervo. Ela não carrega a ressalva do <b>acervo histórico</b>, porque
    ''' não há retrato guardado para ressalvar.
    '''
    ''' <b>Mas eu escrevi aqui, até a revisão de 28/08/2026, que ela por isso
    ''' "não carrega ressalva de cobertura". Era isenção que eu me dei.</b>
    ''' Ler ao vivo prova <i>frescor</i> em relação ao OOM local; não prova
    ''' nada sobre o servidor.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E AGORA HÁ MEDIÇÃO, QUE MUDA METADE DISSO</b>
    '''
    ''' <c>tools/medir-cobertura-calendario.ps1</c>, na tarde de 28/08/2026:
    '''
    ''' <code>
    '''   correio ...... corta em 2026-07-28, ~31 dias
    '''   calendário ... compromissos de 2024-06-07 a 2026-12-15
    ''' </code>
    '''
    ''' <b>411 dos 434</b> compromissos são anteriores ao corte do correio, e a
    ''' distribuição por ano é contínua.
    '''
    ''' <b>O que isso sustenta, e nada além:</b> o corte de ~31 dias <i>não
    ''' aparece</i> neste calendário. Este comentário chegou a dizer "921 dias,
    ''' sem corte" e "não alcança o calendário" — as duas mais fortes que a
    ''' medição, que não procura cortes mais antigos que o item mais velho que
    ''' ela achou e não correlaciona <c>StoreID</c> entre os dois roteiros.
    '''
    ''' Então repetir aqui a ressalva de um mês, que vale para o correio, seria
    ''' ressalva emprestada. <b>Mas o alcance disso é o calendário padrão local,
    ''' e a agenda abre qualquer pasta classificada como calendário</b> — numa
    ''' caixa compartilhada, ou noutro store, ninguém mediu nada. O que continua
    ''' verdade dos dois lados: a contagem do servidor segue inalcançável pelo
    ''' OOM, então <b>ausência continua proibida</b> — por falta de prova, e não
    ''' por janela.
    '''
    ''' Então a agenda diz o que ela <b>sabe</b>: quantos compromissos leu na
    ''' janela, quantos vieram de séries, quantos itens recusou, e se a
    ''' enumeração foi <b>interrompida</b> antes do fim. E diz que não sabe
    ''' dizer o que existe além do que o Outlook expõe.
    ''' </summary>
    Public NotInheritable Class AgendaViewModel
        Inherits ObservableObject
        Implements IDisposable

        Private ReadOnly _fonte As IAgendaSource
        Private ReadOnly _agora As Func(Of DateTimeOffset)

        Private ReadOnly _escritor As IAgendaWriter
        Private _selecionado As LinhaDaAgenda
        Private _confirmando As LinhaDaAgenda
        Private _gravando As Boolean
        Private _novoAssunto As String = ""
        Private _novoInicio As DateTimeOffset
        Private _novaDuracao As Integer = 30

        Private _pasta As FolderKey
        Private _carregando As Boolean
        Private _erro As String = ""
        Private _resumo As String = ""
        Private _dias As Integer = 7
        Private _disposed As Boolean

        ''' <summary>
        ''' A geração da leitura corrente.
        '''
        ''' Mesma família das outras deste projeto: o usuário troca a janela de
        ''' 7 para 30 dias enquanto a primeira leitura está em voo, e a antiga
        ''' volta depois. Publicar a lista velha sobre a nova mostraria uma
        ''' agenda de outro período com cara de atual.
        ''' </summary>
        Private _geracao As Integer

        ''' <summary>
        ''' O relógio é injetável para o teste não depender de que dia é hoje.
        ''' Um teste de agenda que use <c>Now</c> passa em agosto e falha em
        ''' dezembro por um motivo que não é o código.
        ''' </summary>
        Public Sub New(fonte As IAgendaSource,
                       Optional agora As Func(Of DateTimeOffset) = Nothing,
                       Optional escritor As IAgendaWriter = Nothing)
            _fonte = fonte
            _agora = If(agora, Function() DateTimeOffset.Now)

            ' SEM ESCRITOR A AGENDA CONTINUA INTEIRA, so de leitura. E o que
            ' os testes de leitura usam, e e o que a Fase 6 entregou primeiro.
            _escritor = escritor

            AtualizarCommand = New AsyncRelayCommand(AddressOf CarregarAsync,
                                                     Function() _pasta IsNot Nothing AndAlso Not _carregando)
            CriarCommand = New AsyncRelayCommand(AddressOf CriarAsync, Function() PodeCriar)
            PedirExclusaoCommand = New RelayCommand(AddressOf PedirExclusao, Function() PodeApagar)
            CancelarExclusaoCommand = New RelayCommand(Sub() Confirmando = Nothing)
            ApagarCommand = New AsyncRelayCommand(AddressOf ApagarAsync,
                                                 Function() Confirmando IsNot Nothing)
        End Sub

        ''' <summary>
        ''' <b>Criar, editar e apagar compromisso — e por que a agenda pergunta
        ''' antes de apagar.</b>
        '''
        ''' A leitura da Fase 6 entrou em 28/08; a escrita esperou o desenho que
        ''' não pudesse mandar e-mail por descuido. Ver
        ''' <c>Iris.Outlook.CalendarWriting</c> para a invariante.
        '''
        ''' Aqui a única decisão de tela que importa é a confirmação: apagar não
        ''' tem desfazer visível — o compromisso vai para os Itens Excluídos e
        ''' ninguém olha lá. Um clique não pode bastar.
        '''
        ''' <b>A validação NÃO é repetida aqui.</b> Quem recusa assunto vazio e
        ''' fim antes do início é o <c>CalendarWriting</c>, e esta tela mostra o
        ''' motivo que ele devolve. Guarda duplicada é guarda que ninguém prova
        ''' — foi a lição da vigésima passada de revisão, e ela vale igual aqui.
        ''' </summary>
        Public ReadOnly Property CriarCommand As IAsyncRelayCommand
        Public ReadOnly Property PedirExclusaoCommand As RelayCommand
        Public ReadOnly Property CancelarExclusaoCommand As RelayCommand
        Public ReadOnly Property ApagarCommand As IAsyncRelayCommand

        ''' <summary>
        ''' A agenda sabe escrever? Sem escritor ela é só leitura, e a faixa não
        ''' mostra nada de escrita — em vez de mostrar botão que não funciona.
        ''' </summary>
        Public ReadOnly Property PodeEscrever As Boolean
            Get
                Return _escritor IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property PodeCriar As Boolean
            Get
                Return PodeEscrever AndAlso _pasta IsNot Nothing AndAlso
                       Not _carregando AndAlso Not _gravando
            End Get
        End Property

        Public ReadOnly Property PodeApagar As Boolean
            Get
                Return PodeEscrever AndAlso Selecionado IsNot Nothing AndAlso Not _gravando
            End Get
        End Property

        Public Property Selecionado As LinhaDaAgenda
            Get
                Return _selecionado
            End Get
            Set(value As LinhaDaAgenda)
                If SetProperty(_selecionado, value) Then
                    ' Trocar de compromisso CANCELA a confirmacao pendente.
                    ' Sem isto, confirmar apagaria o item que se acabou de
                    ' escolher, e nao o que a pergunta citou.
                    Confirmando = Nothing
                    OnPropertyChanged(NameOf(PodeApagar))
                    PedirExclusaoCommand.NotifyCanExecuteChanged()
                End If
            End Set
        End Property

        ''' <summary>
        ''' O compromisso cuja exclusão está sendo confirmada, ou
        ''' <c>Nothing</c>. É <b>a linha</b>, e não um booleano, justamente para
        ''' a pergunta poder citar o assunto: "apagar <i>Reunião de equipe</i>?"
        ''' </summary>
        Public Property Confirmando As LinhaDaAgenda
            Get
                Return _confirmando
            End Get
            Private Set(value As LinhaDaAgenda)
                If SetProperty(_confirmando, value) Then
                    OnPropertyChanged(NameOf(PerguntaDaExclusao))
                    OnPropertyChanged(NameOf(EstaConfirmando))
                    ApagarCommand.NotifyCanExecuteChanged()
                End If
            End Set
        End Property

        Public ReadOnly Property EstaConfirmando As Boolean
            Get
                Return _confirmando IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property PerguntaDaExclusao As String
            Get
                If _confirmando Is Nothing Then Return ""
                Return $"Apagar ""{_confirmando.Assunto}""? Ele vai para os Itens Excluídos."
            End Get
        End Property

        ''' <summary>O assunto do compromisso a criar. Vazio é recusado — pelo
        ''' <c>CalendarWriting</c>, e não por aqui.</summary>
        Public Property NovoAssunto As String
            Get
                Return _novoAssunto
            End Get
            Set(value As String)
                SetProperty(_novoAssunto, If(value, ""))
            End Set
        End Property

        Public Property NovoInicio As DateTimeOffset
            Get
                Return _novoInicio
            End Get
            Set(value As DateTimeOffset)
                SetProperty(_novoInicio, value)
            End Set
        End Property

        Public Property NovaDuracaoMinutos As Integer
            Get
                Return _novaDuracao
            End Get
            Set(value As Integer)
                SetProperty(_novaDuracao, value)
            End Set
        End Property

        Public ReadOnly Property Compromissos As New ObservableCollection(Of LinhaDaAgenda)()
        Public ReadOnly Property AtualizarCommand As IAsyncRelayCommand

        Public ReadOnly Property TemPasta As Boolean
            Get
                Return _pasta IsNot Nothing
            End Get
        End Property

        Public Property Carregando As Boolean
            Get
                Return _carregando
            End Get
            Private Set(value As Boolean)
                If SetProperty(_carregando, value) Then
                    AtualizarCommand.NotifyCanExecuteChanged()
                End If
            End Set
        End Property

        Public Property Erro As String
            Get
                Return _erro
            End Get
            Private Set(value As String)
                If SetProperty(_erro, If(value, "")) Then
                    OnPropertyChanged(NameOf(TemErro))
                End If
            End Set
        End Property

        Public ReadOnly Property TemErro As Boolean
            Get
                Return Not String.IsNullOrEmpty(_erro)
            End Get
        End Property

        ''' <summary>
        ''' O que a leitura viu, incluindo o que ela recusou.
        '''
        ''' "12 compromissos lidos, 5 de séries" é diferente de "12
        ''' compromissos", e a diferença muda o que o usuário conclui de uma
        ''' agenda cheia. <b>O exemplo aqui dizia "12 compromissos"</b>, sem o
        ''' "lidos" — a quarta versão da mesma história, achada pela revisão
        ''' depois de a tela e o XAML já concordarem.
        ''' </summary>
        Public Property Resumo As String
            Get
                Return _resumo
            End Get
            Private Set(value As String)
                SetProperty(_resumo, If(value, ""))
            End Set
        End Property

        ''' <summary>
        ''' <b>Aponta para a pasta de calendário e lê.</b>
        '''
        ''' <c>Nothing</c> esvazia — pelo mesmo motivo que o acervo esvazia
        ''' quando a seleção some: números sem dono, descrevendo um calendário
        ''' que ninguém está olhando.
        ''' </summary>
        Public Sub Apontar(pasta As FolderKey)
            _pasta = pasta
            _geracao += 1
            OnPropertyChanged(NameOf(TemPasta))
            AtualizarCommand.NotifyCanExecuteChanged()

            Compromissos.Clear()
            Erro = ""
            Resumo = ""

            ' O ESTADO DE ESCRITA E DA PASTA, e nao da tela. Trocar de
            ' calendario com uma confirmacao de exclusao pendente apagaria um
            ' compromisso de outra pasta -- o mesmo defeito que a lista de
            ' mensagens teve com o total da pasta anterior.
            Selecionado = Nothing
            Confirmando = Nothing
            _novoInicio = ProximaMeiaHora()
            OnPropertyChanged(NameOf(NovoInicio))
            OnPropertyChanged(NameOf(PodeCriar))
            OnPropertyChanged(NameOf(PodeApagar))
            CriarCommand.NotifyCanExecuteChanged()
            PedirExclusaoCommand.NotifyCanExecuteChanged()
        End Sub

        ''' <summary>
        ''' O próximo instante "redondo", para o formulário não abrir num horário
        ''' com segundos. Vem do relógio injetável, e não de <c>Now</c>.
        ''' </summary>
        Private Function ProximaMeiaHora() As DateTimeOffset
            Dim n = _agora()
            Dim base = New DateTimeOffset(n.Year, n.Month, n.Day, n.Hour, 0, 0, n.Offset)
            Return If(n.Minute < 30, base.AddMinutes(30), base.AddHours(1))
        End Function

        Private Sub PedirExclusao()
            Confirmando = Selecionado
        End Sub

        ''' <summary>
        ''' Cria o compromisso na pasta que está aberta.
        '''
        ''' Depois de criar, <b>recarrega</b>: a agenda tem de mostrar o item
        ''' novo, e reler é mais barato e mais honesto que enfiar na lista o que
        ''' eu acho que foi gravado. O <c>AppointmentInfo</c> devolvido já vem
        ''' com a identidade relida, mas quem manda na tela é o calendário.
        ''' </summary>
        Private Async Function CriarAsync() As Task
            If Not PodeCriar OrElse _disposed Then Return

            Dim rascunho As New Model.AppointmentDraft With {
                .Subject = _novoAssunto,
                .De = _novoInicio,
                .Ate = _novoInicio.AddMinutes(Math.Max(0, _novaDuracao))
            }

            Await GravarAsync(Function(c) EmResultado(_escritor.CreateAppointmentAsync(_pasta, rascunho, c)),
                              Sub()
                                  NovoAssunto = ""
                              End Sub)
        End Function

        Private Async Function ApagarAsync() As Task
            Dim alvo = Confirmando
            If alvo Is Nothing OrElse _disposed Then Return

            Await GravarAsync(Function(c) _escritor.DeleteAppointmentAsync(
                                  New AppointmentKey(alvo.Chave), c),
                              Sub()
                                  Confirmando = Nothing
                                  Selecionado = Nothing
                              End Sub)
        End Function

        Private Shared Async Function EmResultado(tarefa As Task(Of OperationResult(Of Model.AppointmentInfo))) _
            As Task(Of OperationResult(Of Boolean))
            Dim r = Await tarefa
            If Not r.Succeeded Then Return OperationResult(Of Boolean).Fail(r.Kind, r.Detail)
            Return OperationResult(Of Boolean).Ok(True)
        End Function

        ''' <summary>
        ''' O corpo comum das duas escritas: trava, executa, mostra o motivo se
        ''' recusar, e <b>recarrega</b> se der certo.
        '''
        ''' <c>_gravando</c> existe para o segundo clique não virar a segunda
        ''' criação. Mutação não tem retry neste projeto justamente porque criar
        ''' não é idempotente, e uma tela que deixa clicar duas vezes desfaz essa
        ''' garantia por fora.
        ''' </summary>
        Private Async Function GravarAsync(op As Func(Of Threading.CancellationToken, Task(Of OperationResult(Of Boolean))),
                                           aoDarCerto As Action) As Task
            _gravando = True
            AvisarEstadoDeEscrita()
            Erro = ""

            Try
                Dim r = Await op(Threading.CancellationToken.None)
                If Not r.Succeeded Then
                    Erro = If(String.IsNullOrWhiteSpace(r.Detail),
                              Traduzir(r.Kind), r.Detail)
                    Return
                End If
                aoDarCerto()
            Catch ex As Exception
                Erro = "não consegui gravar no calendário (" & ex.GetType().Name & ")."
                Return
            Finally
                _gravando = False
                AvisarEstadoDeEscrita()
            End Try

            Await CarregarAsync()
        End Function

        Private Sub AvisarEstadoDeEscrita()
            OnPropertyChanged(NameOf(PodeCriar))
            OnPropertyChanged(NameOf(PodeApagar))
            CriarCommand.NotifyCanExecuteChanged()
            PedirExclusaoCommand.NotifyCanExecuteChanged()
            ApagarCommand.NotifyCanExecuteChanged()
        End Sub

        Public Async Function CarregarAsync() As Task
            If _pasta Is Nothing OrElse _disposed Then Return

            Dim minha = Threading.Interlocked.Increment(_geracao)
            Carregando = True
            Erro = ""
            Try
                Dim de = _agora()
                Dim ate = de.AddDays(_dias)

                Dim r = Await _fonte.GetAppointmentsAsync(_pasta, de, ate, CancellationToken.None)

                ' DEPOIS DO AWAIT. A janela pode ter fechado, e a pasta pode ter
                ' mudado — publicar aqui mostraria a agenda de outro período com
                ' cara de atual.
                If _disposed OrElse minha <> Volatile.Read(_geracao) Then Return

                If Not r.Succeeded Then
                    Erro = Traduzir(r.Kind)
                    Compromissos.Clear()
                    Resumo = ""
                    Return
                End If

                Compromissos.Clear()
                For Each a In r.Value.Items.OrderBy(Function(x) x.Start)
                    Compromissos.Add(New LinhaDaAgenda(a))
                Next
                Resumo = Descrever(r.Value)
            Catch ex As Exception
                ' A GERACAO TAMBEM AQUI, e nao so o descarte.
                '
                ' Uma leitura VELHA que falhe depois de uma nova ter terminado
                ' limparia a lista boa e publicaria erro sobre ela. O caminho
                ' feliz ja conferia; o de falha nao, e a revisao externa pegou.
                If _disposed OrElse minha <> Volatile.Read(_geracao) Then Return
                Erro = "Não consegui ler a agenda agora."
                Compromissos.Clear()
                Resumo = ""
            Finally
                ' E O `Carregando` E DO DONO DO VOO.
                '
                ' Apagar o indicador em qualquer geracao faria uma leitura
                ' velha desligar o "carregando" da nova -- a tela pararia de
                ' mostrar progresso com trabalho em voo, e o comando voltaria a
                ' habilitar, permitindo uma terceira leitura concorrente.
                '
                ' E a mesma licao do assistente: quem limpa o Ocupado e o dono
                ' do voo, nao a geracao corrente.
                If Not _disposed AndAlso minha = Volatile.Read(_geracao) Then
                    Carregando = False
                End If
            End Try
        End Function

        ''' <summary>
        ''' O resumo, e ele conta o que foi recusado.
        '''
        ''' <c>Nothing</c> em <c>Skipped</c> não é zero: zero afirma que nada
        ''' foi recusado, nulo diz que aquela leitura não contou. É a mesma
        ''' distinção que o manifesto do acervo faz com as descartadas.
        ''' </summary>
        Private Shared Function Descrever(j As AppointmentWindow) As String
            Dim partes As New List(Of String)()

            ' ZERO NA TELA ERA UMA AFIRMACAO DE AUSENCIA, e o projeto inteiro
            ' proibe essa. A revisao externa achou o caminho completo: o
            ' comentario da classe reconheceu que a medicao so alcanca o
            ' calendario padrao local, o XAML continuava dizendo "por isso ela
            ' nao tem ressalva de cobertura", e a tela mostrava "0
            ' compromisso(s)" numa pasta compartilhada sem ressalva nenhuma.
            '
            ' "LIDO" e a palavra que faz a diferenca, e ela custa nada.
            '
            ' E ELA VALE PARA OS DOIS RAMOS. A primeira correcao qualificou so
            ' o zero, e a revisao seguinte apontou que ficaram TRES versoes da
            ' mesma historia: o comentario da classe prometendo "quantos
            ' compromissos LEU", o XAML qualificando so o caso zero, e a tela
            ' apresentando o numero positivo como TOTAL. Um "12 compromissos"
            ' numa pasta cuja cobertura ninguem mediu e a mesma afirmacao que
            ' o "0 compromissos" era, so que mais dificil de notar.
            If j.Items.Count = 0 Then
                partes.Add($"nenhum compromisso LIDO até {j.Ate.LocalDateTime:dd/MM} — " &
                           "o que não é o mesmo que não haver")
            Else
                partes.Add($"{j.Items.Count} compromisso(s) lido(s) até {j.Ate.LocalDateTime:dd/MM}")
            End If

            If j.FromRecurrence > 0 Then
                partes.Add($"{j.FromRecurrence} de séries")
            End If

            ' NULO E ZERO SAO COISAS DIFERENTES, e este If colapsava as duas
            ' -- no arquivo cujo comentário de cima declara justamente que elas
            ' não são a mesma coisa. Hoje o CalendarReading sempre atribui o
            ' contador, então é dívida latente; mas latente é o estado em que
            ' todos os outros defeitos desta família estavam quando chegaram à
            ' tela.
            If Not j.Skipped.HasValue Then
                partes.Add("não sei quantos itens foram recusados nesta leitura")
            ElseIf j.Skipped.Value > 0 Then
                partes.Add($"{j.Skipped.Value} item(ns) que não consegui ler ou que " &
                           "não eram compromisso")
            End If

            ' AUSENCIA QUE VIROU VALOR, e ela e diferente da recusa.
            '
            ' Recusado e "o item inteiro nao entrou"; fabricado e "o item
            ' entrou com celula inventada" -- AllDayEvent = False, sem
            ' participante, assunto vazio. A listagem ja mostrava este numero
            ' desde 28/08 e o calendario nao: eu instrumentei uma superficie e
            ' nao procurei a irma dela.
            '
            ' Campo que ninguem mostra na tela e o mesmo defeito num lugar
            ' diferente, entao ele sobe aqui.
            If j.FabricatedCells > 0 Then
                partes.Add($"{j.FabricatedCells} campo(s) que o Outlook não entregou")
            End If

            ' TRUNCAMENTO VEM PRIMEIRO NA LEITURA, mesmo vindo por último na
            ' frase: é a informação que muda o que o usuário conclui da lista.
            If j.Truncada Then
                partes.Add("LISTA INCOMPLETA: " &
                           If(String.IsNullOrEmpty(j.MotivoDoCorte),
                              "a leitura parou antes do fim", j.MotivoDoCorte))
            End If

            Return String.Join(" · ", partes)
        End Function

        Private Shared Function Traduzir(kind As ErrorKind) As String
            Select Case kind
                Case ErrorKind.NotFound
                    Return "A pasta de calendário não foi encontrada nesta sessão."
                Case ErrorKind.NotConnected
                    Return "Sem conexão com o Outlook."
                Case ErrorKind.Busy
                    Return "O Outlook está ocupado. Tente de novo em instantes."
                Case ErrorKind.Denied
                    Return "Sem permissão para ler este calendário."
                Case Else
                    Return "Não consegui ler a agenda."
            End Select
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            ' Sobe a geração: uma leitura em voo que volte depois já é de outra
            ' geração e não escreve numa tela que saiu.
            Threading.Interlocked.Increment(_geracao)
        End Sub

    End Class

    ''' <summary>Um compromisso já no formato da tela.</summary>
    Public NotInheritable Class LinhaDaAgenda

        Private Shared ReadOnly Daqui As CultureInfo = CultureInfo.GetCultureInfo("pt-BR")

        Public ReadOnly Property Assunto As String
        Public ReadOnly Property Quando As String
        Public ReadOnly Property Local As String
        Public ReadOnly Property Organizador As String

        ''' <summary>
        ''' Marca de ocorrência de série. Fica visível porque uma ocorrência
        ''' não é um compromisso avulso: mexer nela mexe na série, e o usuário
        ''' precisa saber disso antes de clicar em qualquer coisa.
        ''' </summary>
        Public ReadOnly Property EhSerie As Boolean

        ''' <summary>
        ''' Convite ainda não respondido. É a única informação da agenda que
        ''' pede ação, então ela é a que a tela destaca.
        ''' </summary>
        Public ReadOnly Property Pendente As Boolean

        ''' <summary>
        ''' A chave do compromisso, para a tela poder apagá-lo.
        '''
        ''' Ela é <c>ItemKey</c> e não <c>AppointmentKey</c> porque a linha é um
        ''' dado de tela; quem constrói a chave tipada é quem vai chamar a
        ''' operação — e aí o compilador impede passar uma mensagem no lugar.
        ''' </summary>
        Public ReadOnly Property Chave As ItemKey

        Friend Sub New(a As AppointmentInfo)
            Chave = a.Key
            Assunto = If(String.IsNullOrWhiteSpace(a.Subject), "(sem assunto)", a.Subject)
            Local = If(a.Location, "")
            Organizador = If(a.Organizer, "")
            EhSerie = a.IsRecurring
            Pendente = a.ResponseStatus = AppointmentResponse.NaoRespondeu

            ' CULTURA FIXA. Formatar com a cultura ambiente faria a tela mudar
            ' com a configuração da máquina, e este projeto já teve um teste
            ' que media o host em vez do código.
            Dim inicio = a.Start.LocalDateTime
            If a.AllDayEvent Then
                Quando = inicio.ToString("ddd dd/MM", Daqui) & " · dia inteiro"
            Else
                Quando = inicio.ToString("ddd dd/MM HH:mm", Daqui) &
                         "–" & a.End.LocalDateTime.ToString("HH:mm", Daqui)
            End If
        End Sub

    End Class

End Namespace

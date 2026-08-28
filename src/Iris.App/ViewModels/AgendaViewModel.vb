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
    ''' Então a agenda <b>não</b> mostra uma fatia de um mês, e dizer que ela
    ''' mostra seria repetir para o calendário a ressalva que só vale para o
    ''' correio. O que continua verdade é o outro lado: a contagem do servidor
    ''' segue inalcançável pelo OOM, então <b>ausência continua proibida</b> —
    ''' por falta de prova, e não por janela.
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
        Public Sub New(fonte As IAgendaSource, Optional agora As Func(Of DateTimeOffset) = Nothing)
            _fonte = fonte
            _agora = If(agora, Function() DateTimeOffset.Now)
            AtualizarCommand = New AsyncRelayCommand(AddressOf CarregarAsync,
                                                     Function() _pasta IsNot Nothing AndAlso Not _carregando)
        End Sub

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
        ''' "12 compromissos, 5 de séries" é diferente de "12 compromissos", e
        ''' a diferença muda o que o usuário conclui de uma agenda cheia.
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
            partes.Add($"{j.Items.Count} compromisso(s) até {j.Ate.LocalDateTime:dd/MM}")

            If j.FromRecurrence > 0 Then
                partes.Add($"{j.FromRecurrence} de séries")
            End If

            If j.Skipped.HasValue AndAlso j.Skipped.Value > 0 Then
                partes.Add($"{j.Skipped.Value} item(ns) que não consegui ler")
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

        Friend Sub New(a As AppointmentInfo)
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

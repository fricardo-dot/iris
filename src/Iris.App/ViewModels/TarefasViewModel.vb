Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Globalization
Imports System.Threading
Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>Tarefas — e as duas etapas que a Fase 5 exige.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>"A IA SUGERE, VOCÊ CONFIRMA, O IRIS CRIA"</b>
    '''
    ''' O ESCOPO escreveu a fase assim, e a parte que importa é o meio: <b>nunca
    ''' criação silenciosa em massa</b>. Uma sugestão que virasse tarefa sozinha
    ''' encheria a lista de coisas que ninguém pediu, e a lista de tarefas é
    ''' justamente onde lixo custa caro — ela existe para dizer o que falta
    ''' fazer.
    '''
    ''' Aqui a sugestão <b>preenche o formulário</b> e para. Criar é outro
    ''' comando, outro clique, e o texto continua editável até lá.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E A SUGESTÃO DE HOJE NÃO É DA IA</b>
    '''
    ''' <see cref="ProporDaMensagem"/> monta a proposta a partir do
    ''' <i>assunto</i> da mensagem selecionada — determinístico, local, sem
    ''' provedor nenhum.
    '''
    ''' Isso é decisão, e não meio-caminho: fazer a IA extrair tarefas de uma
    ''' mensagem é uma <b>operação nova</b>, e a cerimônia de ativação autoriza
    ''' operações nomeadas sobre pastas escolhidas. Reusar a autorização do
    ''' resumo para extrair tarefas alargaria uma permissão que ninguém
    ''' concordou em alargar — que é exatamente a razão pela qual a Fase 4
    ''' continua parada. O dia em que existir uma entrada de ativação para
    ''' isso, o <c>ProporDaMensagem</c> ganha um irmão que chama o assistente.
    ''' </summary>
    Public NotInheritable Class TarefasViewModel
        Inherits ObservableObject
        Implements IDisposable

        Private Const Teto As Integer = 200

        Private ReadOnly _broker As ITarefasBroker
        Private ReadOnly _agora As Func(Of DateTimeOffset)

        Private _pasta As FolderKey
        Private _carregando As Boolean
        Private _gravando As Boolean
        Private _erro As String = ""
        Private _resumo As String = ""
        Private _novoAssunto As String = ""
        Private _venceEm As DateTimeOffset?
        Private _selecionada As LinhaDeTarefa
        Private _disposed As Boolean
        Private _geracao As Integer

        Public Sub New(broker As ITarefasBroker,
                       Optional agora As Func(Of DateTimeOffset) = Nothing)
            _broker = broker
            _agora = If(agora, Function() DateTimeOffset.Now)

            AbrirCommand = New AsyncRelayCommand(AddressOf AbrirAsync,
                                                 Function() Not _carregando)
            AtualizarCommand = New AsyncRelayCommand(AddressOf CarregarAsync,
                                                     Function() _pasta IsNot Nothing AndAlso Not _carregando)
            CriarCommand = New AsyncRelayCommand(AddressOf CriarAsync, Function() PodeCriar)
            ConcluirCommand = New AsyncRelayCommand(AddressOf ConcluirAsync, Function() PodeConcluir)
        End Sub

        Public ReadOnly Property Tarefas As New ObservableCollection(Of LinhaDeTarefa)()
        Public ReadOnly Property AbrirCommand As IAsyncRelayCommand
        Public ReadOnly Property AtualizarCommand As IAsyncRelayCommand
        Public ReadOnly Property CriarCommand As IAsyncRelayCommand
        Public ReadOnly Property ConcluirCommand As IAsyncRelayCommand

        Public ReadOnly Property TemPasta As Boolean
            Get
                Return _pasta IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property Carregando As Boolean
            Get
                Return _carregando
            End Get
        End Property

        Public ReadOnly Property PodeCriar As Boolean
            Get
                Return _pasta IsNot Nothing AndAlso Not _gravando AndAlso Not _carregando
            End Get
        End Property

        ''' <summary>
        ''' <b>Concluir uma tarefa ATRIBUÍDA não é oferecido.</b>
        '''
        ''' O <c>TaskWriting</c> recusa de qualquer jeito — a guarda mora lá, e
        ''' não é repetida aqui. O que esta propriedade faz é diferente: evita
        ''' oferecer um botão que vai ser recusado. Botão que não funciona é
        ''' pior que botão ausente, porque promete.
        ''' </summary>
        Public ReadOnly Property PodeConcluir As Boolean
            Get
                Return _selecionada IsNot Nothing AndAlso
                       Not _selecionada.Atribuida AndAlso
                       Not _selecionada.Concluida AndAlso
                       Not _gravando
            End Get
        End Property

        Public Property Selecionada As LinhaDeTarefa
            Get
                Return _selecionada
            End Get
            Set(value As LinhaDeTarefa)
                If SetProperty(_selecionada, value) Then
                    OnPropertyChanged(NameOf(PodeConcluir))
                    OnPropertyChanged(NameOf(AvisoDaSelecionada))
                    ConcluirCommand.NotifyCanExecuteChanged()
                End If
            End Set
        End Property

        ''' <summary>
        ''' O que a tela diz sobre a tarefa escolhida. Vazio quando não há o que
        ''' dizer — e não vazio justamente quando o botão está desabilitado, que
        ''' é quando o usuário precisa saber por quê.
        ''' </summary>
        Public ReadOnly Property AvisoDaSelecionada As String
            Get
                If _selecionada Is Nothing Then Return ""
                If _selecionada.Atribuida Then Return TarefaAtribuida
                If _selecionada.Concluida Then Return "Esta tarefa já está concluída."
                Return ""
            End Get
        End Property

        Friend Const TarefaAtribuida As String =
            "Esta tarefa está atribuída a alguém: mexer nela manda atualização " &
            "por e-mail, e o Iris não envia. Use o Outlook."

        Public Property Erro As String
            Get
                Return _erro
            End Get
            Private Set(value As String)
                If SetProperty(_erro, If(value, "")) Then OnPropertyChanged(NameOf(TemErro))
            End Set
        End Property

        Public ReadOnly Property TemErro As Boolean
            Get
                Return _erro.Length > 0
            End Get
        End Property

        Public Property Resumo As String
            Get
                Return _resumo
            End Get
            Private Set(value As String)
                SetProperty(_resumo, If(value, ""))
            End Set
        End Property

        Public Property NovoAssunto As String
            Get
                Return _novoAssunto
            End Get
            Set(value As String)
                If SetProperty(_novoAssunto, If(value, "")) Then
                    OnPropertyChanged(NameOf(TemProposta))
                End If
            End Set
        End Property

        Public Property VenceEm As DateTimeOffset?
            Get
                Return _venceEm
            End Get
            Set(value As DateTimeOffset?)
                SetProperty(_venceEm, value)
            End Set
        End Property

        Public ReadOnly Property TemProposta As Boolean
            Get
                Return _novoAssunto.Trim().Length > 0
            End Get
        End Property

        ''' <summary>
        ''' <b>ETAPA UM: propor. Isto NÃO cria nada.</b>
        '''
        ''' Preenche o formulário a partir do assunto da mensagem e para. O
        ''' usuário lê, muda se quiser, e só então clica em criar.
        '''
        ''' É determinístico e local: nenhum byte sai da máquina para montar
        ''' esta proposta. Ver o comentário da classe para por que a versão com
        ''' IA depende de uma entrada nova na cerimônia de ativação.
        ''' </summary>
        Public Sub ProporDaMensagem(assunto As String)
            Dim limpo = If(assunto, "").Trim()
            If limpo.Length = 0 Then limpo = "(mensagem sem assunto)"

            NovoAssunto = limpo
            VenceEm = Nothing
            Erro = ""
        End Sub

        ''' <summary>
        ''' Descobre a pasta de Tarefas e lê. Botão próprio porque a política de
        ''' visibilidade mantém essa pasta <b>fora</b> da árvore: não há o que
        ''' selecionar.
        '''
        ''' <b>O <c>_carregando</c> sobe ANTES da descoberta</b>, e não só no
        ''' <c>CarregarAsync</c>. Antes disso o <c>CanExecute</c> do botão ficava
        ''' verdadeiro durante toda a primeira espera, e dois cliques
        ''' disparavam duas descobertas que atribuíam <c>_pasta</c> em corrida.
        ''' A geração protege os carregamentos; ela não protegia esta atribuição.
        ''' </summary>
        Public Async Function AbrirAsync() As Task
            If _disposed OrElse _carregando Then Return

            Erro = ""
            _carregando = True
            OnPropertyChanged(NameOf(Carregando))
            AvisarComandos()

            Dim r As OperationResult(Of FolderKey)
            Try
                r = Await _broker.GetDefaultTasksFolderAsync(CancellationToken.None)
            Finally
                _carregando = False
                OnPropertyChanged(NameOf(Carregando))
                AvisarComandos()
            End Try

            ' DEPOIS DA ESPERA, DE NOVO. Descarte durante a descoberta
            ' ressuscitaria a tela: instalaria _pasta, notificaria propriedades
            ' e entraria num carregamento que ninguém pediu.
            If _disposed Then Return

            If Not r.Succeeded Then
                Erro = "não consegui achar a pasta de Tarefas (" & r.Kind.ToString() & ")."
                Return
            End If

            _pasta = r.Value
            OnPropertyChanged(NameOf(TemPasta))
            OnPropertyChanged(NameOf(PodeCriar))
            AtualizarCommand.NotifyCanExecuteChanged()
            CriarCommand.NotifyCanExecuteChanged()

            Await CarregarAsync()
        End Function

        Public Async Function CarregarAsync() As Task
            If _pasta Is Nothing OrElse _disposed Then Return

            Dim minha = Interlocked.Increment(_geracao)
            _carregando = True
            OnPropertyChanged(NameOf(Carregando))
            AvisarComandos()

            Try
                Dim r = Await _broker.GetTasksAsync(_pasta, Teto, CancellationToken.None)
                If minha <> Volatile.Read(_geracao) OrElse _disposed Then Return

                If Not r.Succeeded Then
                    Erro = "não consegui ler as tarefas (" & r.Kind.ToString() & ")."
                    Tarefas.Clear()
                    Resumo = ""
                    Return
                End If

                Tarefas.Clear()
                For Each t In r.Value.Items
                    Tarefas.Add(New LinhaDeTarefa(t))
                Next
                Selecionada = Nothing
                Erro = ""
                Resumo = Descrever(r.Value)
            Finally
                If Not _disposed AndAlso minha = Volatile.Read(_geracao) Then
                    _carregando = False
                    OnPropertyChanged(NameOf(Carregando))
                    AvisarComandos()
                End If
            End Try
        End Function

        ''' <summary>
        ''' O resumo, e ele conta o que ficou de fora — mesma disciplina da
        ''' agenda: recusado e truncado são ditos, e "não contei" não vira zero.
        ''' </summary>
        Private Shared Function Descrever(lista As TaskList) As String
            Dim partes As New List(Of String)()

            If lista.Items.Count = 0 Then
                partes.Add("nenhuma tarefa LIDA — o que não é o mesmo que não haver")
            Else
                partes.Add($"{lista.Items.Count} tarefa(s) lida(s)")
            End If

            If Not lista.Skipped.HasValue Then
                partes.Add("não sei quantos itens foram recusados nesta leitura")
            ElseIf lista.Skipped.Value > 0 Then
                partes.Add($"{lista.Skipped.Value} item(ns) que não consegui ler")
            End If

            If lista.Truncada Then
                partes.Add("LISTA INCOMPLETA: " &
                           If(String.IsNullOrEmpty(lista.MotivoDoCorte),
                              "a leitura parou antes do fim", lista.MotivoDoCorte))
            End If

            Return String.Join(" · ", partes)
        End Function

        ''' <summary>
        ''' <b>ETAPA DOIS: criar.</b> Uma tarefa, deste formulário, com este
        ''' clique. Nunca em lote.
        ''' </summary>
        ''' <b>Pública, e não privada.</b> O <c>AsyncRelayCommand</c> serializa
        ''' as execuções que passam por ele — o que significa que a guarda
        ''' interna só é alcançável por quem <i>não</i> passa pelo comando.
        ''' Deixá-la privada tornaria a guarda intestável, e um teste que só
        ''' consegue clicar no botão prova o comportamento do toolkit, e não o
        ''' desta classe. Foi exatamente isso que o primeiro corte destes testes
        ''' fez: passou com a guarda removida.
        Public Async Function CriarAsync() As Task
            If Not PodeCriar OrElse _disposed Then Return

            Dim rascunho As New TaskDraft With {
                .Subject = _novoAssunto,
                .Vence = _venceEm
            }

            _gravando = True
            AvisarComandos()
            Erro = ""

            Try
                Dim r = Await _broker.CreateTaskAsync(_pasta, rascunho, CancellationToken.None)

                ' DESCARTE DURANTE A GRAVAÇÃO. A tarefa foi criada -- isso não
                ' se desfaz, e nem deve. O que não pode é a continuação mexer
                ' numa tela que já foi embora.
                If _disposed Then Return

                If Not r.Succeeded Then
                    Erro = If(String.IsNullOrWhiteSpace(r.Detail),
                              "não consegui criar a tarefa (" & r.Kind.ToString() & ").",
                              r.Detail)
                    Return
                End If
                NovoAssunto = ""
                VenceEm = Nothing
            Catch ex As Exception
                Erro = "não consegui criar a tarefa (" & ex.GetType().Name & ")."
                Return
            Finally
                _gravando = False
                AvisarComandos()
            End Try

            Await CarregarAsync()
        End Function

        ''' <summary>
        ''' Conclui a tarefa selecionada.
        '''
        ''' <b>A guarda é a mesma do <c>CanExecute</c>, repetida aqui de
        ''' propósito.</b> O <c>CanExecute</c> governa o botão; ele não governa
        ''' quem chama <c>ExecuteAsync</c> direto — automação, teste, ou a
        ''' janela entre dois eventos de comando. Sem esta linha, uma segunda
        ''' execução chegaria ao broker com <c>_gravando</c> já verdadeiro.
        '''
        ''' O <c>TaskWriting</c> ainda recusaria tarefa atribuída, e é ele a
        ''' barreira que impede o e-mail. Mas uma tela que promete uma condição
        ''' e não a sustenta deixa a barreira sozinha, e barreira sozinha é
        ''' exatamente o que esta base evita.
        ''' </summary>
        Public Async Function ConcluirAsync() As Task
            If Not PodeConcluir OrElse _disposed Then Return

            Dim alvo = Selecionada
            If alvo Is Nothing Then Return

            _gravando = True
            AvisarComandos()
            Erro = ""

            Try
                Dim r = Await _broker.CompleteTaskAsync(New TaskKey(alvo.Chave),
                                                        CancellationToken.None)
                If _disposed Then Return

                If Not r.Succeeded Then
                    Erro = If(String.IsNullOrWhiteSpace(r.Detail),
                              "não consegui concluir a tarefa (" & r.Kind.ToString() & ").",
                              r.Detail)
                    Return
                End If
            Catch ex As Exception
                Erro = "não consegui concluir a tarefa (" & ex.GetType().Name & ")."
                Return
            Finally
                _gravando = False
                AvisarComandos()
            End Try

            Await CarregarAsync()
        End Function

        Private Sub AvisarComandos()
            OnPropertyChanged(NameOf(PodeCriar))
            OnPropertyChanged(NameOf(PodeConcluir))
            AbrirCommand.NotifyCanExecuteChanged()
            AtualizarCommand.NotifyCanExecuteChanged()
            CriarCommand.NotifyCanExecuteChanged()
            ConcluirCommand.NotifyCanExecuteChanged()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            _disposed = True
            Interlocked.Increment(_geracao)
        End Sub

    End Class

    ''' <summary>Uma tarefa já no formato da tela.</summary>
    Public NotInheritable Class LinhaDeTarefa

        Private Shared ReadOnly Daqui As CultureInfo = CultureInfo.GetCultureInfo("pt-BR")

        Public ReadOnly Property Chave As ItemKey
        Public ReadOnly Property Assunto As String
        Public ReadOnly Property Prazo As String
        Public ReadOnly Property Concluida As Boolean

        ''' <summary>
        ''' Atribuída. Fica <b>visível</b> porque é o que explica o botão
        ''' desabilitado — e porque avisa que aquela linha pertence a uma
        ''' conversa por e-mail da qual o Iris não participa.
        ''' </summary>
        Public ReadOnly Property Atribuida As Boolean

        Friend Sub New(t As TaskInfo)
            Chave = t.Key
            Assunto = If(String.IsNullOrWhiteSpace(t.Subject), "(sem assunto)", t.Subject)
            Concluida = t.Concluida
            Atribuida = t.Atribuida

            ' SEM PRAZO E DITO, e nao virado numa data qualquer. O Outlook
            ' guarda "sem prazo" como 4501-01-01, e o TaskWriting ja traduziu
            ' isso para Nothing -- aqui so falta a tela nao inventar.
            Prazo = If(t.Vence.HasValue,
                       t.Vence.Value.LocalDateTime.ToString("dd/MM/yyyy", Daqui),
                       "sem prazo")
        End Sub

    End Class

End Namespace

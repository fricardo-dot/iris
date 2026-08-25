Imports System.Collections.Generic
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Assist
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>A IA na janela — e, hoje, o motivo de ela não estar disponível.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE O USUÁRIO VÊ HOJE</b>
    '''
    ''' Uma frase dizendo que a IA externa não está habilitada, e por quê. Não é
    ''' "recurso em construção": o mecanismo está inteiro e testado, e o que
    ''' falta é a decisão dele — a política da empresa e um provedor à escolha
    ''' dele. É a §28.2.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A RECONCILIAÇÃO É PRÉ-CONDIÇÃO, NÃO TRABALHO DA TELA</b>
    '''
    ''' O que ficou <i>em voo</i> numa execução que morreu vira ambíguo na
    ''' abertura seguinte — e isso é recuperação de segurança, não um número
    ''' bonito. Se ela falhar ou não terminar, <b>o egress fica fechado</b>: a
    ''' tela não oferece a operação.
    '''
    ''' A tela só <b>mostra</b> o resultado, porque "pode ter saído conteúdo
    ''' desta caixa e ninguém sabe" não é detalhe de log.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>RESPOSTA VELHA NÃO APARECE EM CONTEXTO NOVO</b>
    '''
    ''' Enquanto a IA responde, o usuário troca de mensagem, pede outra geração
    ''' ou cancela. Cada pedido carrega uma <b>geração</b>, e um resultado só é
    ''' publicado se a geração dele ainda for a corrente.
    '''
    ''' Sem isso, o resumo de uma mensagem apareceria embaixo de outra — e um
    ''' resumo errado com cara de certo é pior que resumo nenhum.
    ''' </summary>
    Public NotInheritable Class AssistenteViewModel
        Inherits ObservableObject

        Private ReadOnly _ui As Dispatcher
        Private ReadOnly _transmissor As AssistTransmitter
        Private ReadOnly _politica As DisclosurePolicy
        Private ReadOnly _relogio As Func(Of DateTimeOffset)

        ''' <summary>
        ''' A geração corrente. Cada troca de contexto a incrementa, e um
        ''' resultado de geração vencida é <b>descartado</b>.
        ''' </summary>
        Private _geracao As Integer

        Private _cancelamento As CancellationTokenSource

        Private ReadOnly _contexto As IAssistContext
        Private ReadOnly _rascunho As IRascunho

        Public Sub New(ui As Dispatcher, transmissor As AssistTransmitter,
                       politica As DisclosurePolicy, relogio As Func(Of DateTimeOffset),
                       reconciliacao As ReconciliationResult,
                       contexto As IAssistContext, rascunho As IRascunho)
            _ui = ui
            _transmissor = transmissor
            _politica = politica
            _relogio = relogio
            _contexto = If(contexto, CType(New ContextoIndisponivel(), IAssistContext))
            _rascunho = rascunho

            ' Comandos declarados a mao, e nao pelo gerador do
            ' CommunityToolkit: ele so roda em C#. Em VB o atributo compila e
            ' NAO gera nada — e o sintoma seria um botao que nunca faz nada.
            ResumirCommand = New AsyncRelayCommand(AddressOf Resumir, Function() PodePedir)
            RedigirCommand = New AsyncRelayCommand(AddressOf Redigir, Function() PodeRedigir)
            CancelarCommand = New RelayCommand(AddressOf Cancelar, Function() PodeCancelar)
            DesfazerCommand = New RelayCommand(AddressOf Desfazer, Function() PodeDesfazer)
            ' Me. OBRIGATORIO: sem ele, 'reconciliacao' eclipsa 'Reconciliacao'
            ' — VB e case-insensitive, e a atribuicao vira o parametro para ele
            ' mesmo. O compilador nao avisa; o sintoma aparece longe, como uma
            ' propriedade Nothing na primeira leitura.
            Me.Reconciliacao = reconciliacao
        End Sub

        ' ==============================================================
        ' O estado

        ''' <summary>O que aconteceu na reconciliação da abertura.</summary>
        Public ReadOnly Property Reconciliacao As ReconciliationResult

        ''' <summary>
        ''' <b>Visível sempre, habilitado só quando dá.</b>
        '''
        ''' Um botão que some quando a IA está desligada esconde a funcionalidade
        ''' <i>e</i> o motivo — e o motivo é exatamente o que o usuário precisa
        ''' ler, no lugar onde ele procuraria a ação.
        ''' </summary>
        Public ReadOnly Property ResumirCommand As AsyncRelayCommand
        Public ReadOnly Property RedigirCommand As AsyncRelayCommand
        Public ReadOnly Property CancelarCommand As RelayCommand
        Public ReadOnly Property DesfazerCommand As RelayCommand

        Private _ocupado As Boolean
        ''' <summary>Há um pedido em andamento.</summary>
        Public Property Ocupado As Boolean
            Get
                Return _ocupado
            End Get
            Private Set(value As Boolean)
                SetProperty(_ocupado, value)
                OnPropertyChanged(NameOf(PodePedir))
                OnPropertyChanged(NameOf(PodeRedigir))
                OnPropertyChanged(NameOf(PodeCancelar))
                Avisar()
            End Set
        End Property

        Private _resultado As String = ""
        ''' <summary>
        ''' O texto do modelo. <b>Texto</b>, e nada além disso.
        '''
        ''' Ele vem de um lugar que leu o e-mail — que por sua vez veio de fora.
        ''' A tela o mostra como texto simples: não vira Markdown ativo, não vira
        ''' link clicável, não vira comando. É a barreira da §29.5, e ela mora
        ''' aqui e no XAML, não numa instrução ao modelo.
        ''' </summary>
        Public Property Resultado As String
            Get
                Return _resultado
            End Get
            Private Set(value As String)
                SetProperty(_resultado, value)
                OnPropertyChanged(NameOf(TemResultado))
            End Set
        End Property

        Public ReadOnly Property TemResultado As Boolean
            Get
                Return Resultado.Length > 0
            End Get
        End Property

        Private _aviso As String = ""
        ''' <summary>
        ''' O que a tela precisa dizer — em português, e sem código de enum.
        ''' </summary>
        Public Property Aviso As String
            Get
                Return _aviso
            End Get
            Private Set(value As String)
                SetProperty(_aviso, value)
                OnPropertyChanged(NameOf(TemAviso))
                OnPropertyChanged(NameOf(TemAlgoADizer))
            End Set
        End Property

        Public ReadOnly Property TemAviso As Boolean
            Get
                Return Aviso.Length > 0
            End Get
        End Property

        ''' <summary>
        ''' <b>A faixa tem algo a dizer?</b> Aviso <b>ou</b> reconciliação.
        '''
        ''' A visibilidade olhava só o <see cref="Aviso"/>, e com ativação válida
        ''' ele fica vazio — então uma reconciliação que achou envios ambíguos
        ''' ficaria <b>invisível</b> justamente no caso em que ela tem algo grave
        ''' a contar.
        ''' </summary>
        Public ReadOnly Property TemAlgoADizer As Boolean
            Get
                Return TemAviso OrElse Reconciliacao.Aviso.Length > 0
            End Get
        End Property

        ''' <summary>
        ''' <b>A IA está disponível?</b>
        '''
        ''' Duas condições, e as duas são fechadas por padrão: a reconciliação
        ''' terminou, e o portão aceita o voo. Em produção a segunda é
        ''' <c>False</c> — <c>ActivationRecord.DaProducao</c> é <c>Nothing</c>.
        ''' </summary>
        Public ReadOnly Property Disponivel As Boolean
            Get
                Return Reconciliacao.Terminou AndAlso _portaoAceita
            End Get
        End Property

        Private _portaoAceita As Boolean

        Public ReadOnly Property PodePedir As Boolean
            Get
                Return Disponivel AndAlso Not Ocupado
            End Get
        End Property

        Public ReadOnly Property PodeCancelar As Boolean
            Get
                Return Ocupado
            End Get
        End Property

        ' ==============================================================

        ''' <summary>
        ''' Pergunta ao portão se o voo passa, <b>sem classificar item nenhum</b>,
        ''' e guarda o motivo para a tela.
        '''
        ''' É o preflight: sem autorização não se gasta uma ida ao COM lendo
        ''' rótulo de coisa nenhuma.
        ''' </summary>
        ''' <summary>Reavalia com o contexto corrente.</summary>
        Public Sub Avaliar()
            Avaliar(_contexto.Pedido(AssistOperation.Resumir))
        End Sub

        Public Sub Avaliar(pedido As PreflightRequest)
            Dim d = _politica.Preflight(pedido, _relogio())
            _portaoAceita = d.Permitido
            Aviso = If(d.Permitido, "", d.Explicacao)
            OnPropertyChanged(NameOf(Disponivel))
            OnPropertyChanged(NameOf(PodePedir))
            OnPropertyChanged(NameOf(PodeRedigir))
            Avisar()
        End Sub

        ''' <summary>Os comandos reavaliam o que podem fazer.</summary>
        Private Sub Avisar()
            ResumirCommand.NotifyCanExecuteChanged()
            RedigirCommand.NotifyCanExecuteChanged()
            CancelarCommand.NotifyCanExecuteChanged()
        End Sub

        ''' <summary>
        ''' <b>Troca de contexto.</b> Incrementa a geração, limpa o resultado e
        ''' cancela o que estiver em voo.
        '''
        ''' Chamada quando o usuário muda de mensagem. Sem ela, o resumo da
        ''' anterior apareceria embaixo da nova.
        ''' </summary>
        Public Sub Trocou()
            _geracao += 1
            Resultado = ""
            Cancelar()
        End Sub

        Public Sub Cancelar()
            _cancelamento?.Cancel()
        End Sub

        ''' <summary>
        ''' <b>Resumir.</b> Visível sempre, habilitado só quando dá.
        '''
        ''' Um botão que some quando a IA está desligada esconderia a
        ''' funcionalidade e o motivo dela estar desligada — e o motivo é
        ''' exatamente o que o usuário precisa ler, no lugar onde ele procuraria a
        ''' ação. Um botão que executa e sempre recusa seria o outro extremo.
        ''' Visível e desabilitado, com o motivo ao lado, é o meio.
        ''' </summary>
        Public Async Function Resumir() As Task
            Await Executar(AssistOperation.Resumir, "Resuma estas mensagens.")
        End Function

        ''' <summary>
        ''' <b>Redigir.</b> Escreve no rascunho — e guarda o que estava lá.
        '''
        ''' Escrever por cima do que o usuário digitou é mutação local, e mutação
        ''' local sem volta é a que ele descobre tarde demais. O texto anterior
        ''' fica guardado, e <see cref="Desfazer"/> o devolve.
        '''
        ''' <b>Nada é enviado por e-mail</b>: a redação para no compositor.
        ''' </summary>
        Public Async Function Redigir() As Task
            Await Executar(AssistOperation.Redigir, "Redija uma resposta.")
        End Function

        ''' <summary>Devolve o rascunho como estava antes da redação.</summary>
        Public Sub Desfazer()
            If _rascunho Is Nothing OrElse _anterior Is Nothing Then Return
            _rascunho.Texto = _anterior
            _anterior = Nothing
            OnPropertyChanged(NameOf(PodeDesfazer))
            DesfazerCommand.NotifyCanExecuteChanged()
        End Sub

        ''' <summary>O que estava no rascunho antes da última redação.</summary>
        Private _anterior As String

        Public ReadOnly Property PodeRedigir As Boolean
            Get
                Return PodePedir AndAlso _rascunho IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property PodeDesfazer As Boolean
            Get
                Return _anterior IsNot Nothing
            End Get
        End Property

        ''' <summary>
        ''' O caminho comum dos dois comandos: monta o contexto, executa, e para
        ''' a redação, escreve no rascunho guardando o que estava lá.
        ''' </summary>
        Private Async Function Executar(operacao As AssistOperation,
                                        instrucao As String) As Task
            Dim antesDoPedido = If(_rascunho Is Nothing, Nothing, _rascunho.Texto)

            Await Pedir(_contexto.Pedido(operacao),
                        AddressOf _contexto.Classificar,
                        Function() _contexto.Montar(operacao, instrucao))

            If operacao <> AssistOperation.Redigir Then Return
            If _rascunho Is Nothing OrElse Not TemResultado Then Return

            _anterior = antesDoPedido
            _rascunho.Texto = Resultado
            OnPropertyChanged(NameOf(PodeDesfazer))
            DesfazerCommand.NotifyCanExecuteChanged()
        End Function

        ''' <summary>
        ''' Pede a operação. O resultado só é publicado se a geração ainda for a
        ''' corrente quando ele voltar.
        ''' </summary>
        Public Async Function Pedir(pedido As PreflightRequest,
                                    classificar As Func(Of IReadOnlyList(Of MessageClassification)),
                                    montar As Func(Of EnvelopeResult)) As Task

            If Not PodePedir Then Return

            Dim minha = _geracao
            Ocupado = True
            Aviso = ""

            Dim cts As New CancellationTokenSource()
            _cancelamento = cts

            Try
                Dim r = Await Task.Run(Function() _transmissor.Executar(
                    pedido, classificar, montar, cts.Token))

                ' A GERACAO. Se o usuario trocou de mensagem enquanto isto
                ' rodava, o resultado e de outro contexto — e mostrar seria pior
                ' que nao mostrar: um resumo errado com cara de certo.
                If minha <> _geracao Then Return

                Publicar(r)
            Finally
                cts.Dispose()
                If ReferenceEquals(_cancelamento, cts) Then _cancelamento = Nothing
                If minha = _geracao Then Ocupado = False
            End Try
        End Function

        Private Sub Publicar(r As AssistOutcome)
            Select Case r.Kind
                Case AssistOutcomeKind.Respondeu
                    Resultado = r.Texto
                    Aviso = ""
                Case AssistOutcomeKind.Negado
                    Resultado = ""
                    Aviso = "A IA não foi usada: " & EmPortugues(r.MotivoDoPortao)
                Case AssistOutcomeKind.Ambiguo,
                     AssistOutcomeKind.AmbiguoSemFechamentoDoDiario
                    Resultado = ""
                    ' Sem asterisco: num TextBlock ele aparece literalmente, e
                    ' "**nao da para saber**" na tela e desleixo visivel.
                    Aviso = "A operação não terminou, e não dá para saber se o " &
                            "conteúdo chegou ao provedor. Isso ficou registrado."
                Case Else
                    Resultado = ""
                    Aviso = "A operação não foi feita, e nada saiu deste computador."
            End Select
        End Sub

        ''' <summary>
        ''' O motivo do portão em português.
        '''
        ''' A tradução mora <b>aqui</b>, e não no diário: lá o motivo é código,
        ''' justamente para não haver campo por onde texto arbitrário entre.
        ''' </summary>
        Public Shared Function EmPortugues(m As DisclosureReason) As String
            Select Case m
                Case DisclosureReason.SemAtivacao
                    Return "a IA externa não está habilitada."
                Case DisclosureReason.AtivacaoIncompleta, DisclosureReason.AtivacaoInvalida
                    Return "a autorização registrada não está em ordem."
                Case DisclosureReason.AtivacaoForaDeVigencia
                    Return "a autorização não vale nesta data."
                Case DisclosureReason.EndpointInseguro, DisclosureReason.EndpointNaoAutorizado
                    Return "o endereço de destino não é o autorizado."
                Case DisclosureReason.OperacaoNaoAutorizada
                    Return "esta operação não está autorizada."
                Case DisclosureReason.ProvedorNaoAutorizado
                    Return "o provedor ou o modelo não é o autorizado."
                Case DisclosureReason.PastaNaoAutorizada
                    Return "esta pasta não está autorizada."
                Case DisclosureReason.AnexoForaDeEscopo
                    Return "há mensagem com anexo, e anexo não é tratado."
                Case DisclosureReason.RotuloNaoPermitido, DisclosureReason.HistoricoNaoDeclarado
                    Return "há mensagem com classificação de sensibilidade não autorizada."
                Case DisclosureReason.ContentBitsDesconhecido, DisclosureReason.ContentBitsNaoAceito
                    Return "não dá para saber se alguma mensagem está protegida."
                Case DisclosureReason.LeituraNaoAceita,
                     DisclosureReason.LeituraEstruturalmenteInsegura,
                     DisclosureReason.ClassificacaoIncoerente
                    Return "não foi possível classificar alguma mensagem com segurança."
                Case DisclosureReason.SemEvidenciaDeVersao, DisclosureReason.IdentidadeNaoBate
                    Return "não dá para saber qual versão de alguma mensagem foi classificada."
                Case DisclosureReason.MensagemDeOutraPasta
                    Return "há mensagem de outra pasta no pedido."
                Case DisclosureReason.PedidoVazio
                    Return "não há nada a enviar."
                Case Else
                    Return "não foi autorizado."
            End Select
        End Function

    End Class

    ''' <summary>
    ''' O que a reconciliação da abertura achou — e se ela chegou a terminar.
    '''
    ''' <see cref="Terminou"/> é <b>pré-condição de egress</b>: reconciliação que
    ''' falhou deixa o diário sem saber o que ficou em voo, e transmitir por cima
    ''' disso é acrescentar incerteza a incerteza.
    ''' </summary>
    Public NotInheritable Class ReconciliationResult

        Public ReadOnly Property Terminou As Boolean
        ''' <summary>Quantas divulgações viraram ambíguas.</summary>
        Public ReadOnly Property Ambiguas As Integer

        Private Sub New(terminou As Boolean, ambiguas As Integer)
            Me.Terminou = terminou
            Me.Ambiguas = ambiguas
        End Sub

        ''' <summary>
        ''' Roda a reconciliação. <b>Não lança</b>: falha aqui fecha o egress em
        ''' vez de derrubar a abertura do programa.
        ''' </summary>
        Public Shared Function Rodar(diario As IDisclosureJournal,
                                     agora As DateTimeOffset) As ReconciliationResult
            Try
                Return New ReconciliationResult(True, diario.Reconciliar(agora))
            Catch
                Return New ReconciliationResult(False, 0)
            End Try
        End Function

        ''' <summary>Quando não há diário nenhum — o cache não abriu.</summary>
        Public Shared Function NaoRodou() As ReconciliationResult
            Return New ReconciliationResult(False, 0)
        End Function

        ''' <summary>O que a tela diz sobre isso. Vazio quando não há o que dizer.</summary>
        Public ReadOnly Property Aviso As String
            Get
                If Not Terminou Then
                    Return "Não foi possível conferir o registro de envios à IA. " &
                           "Enquanto isso, a IA externa fica desligada."
                End If
                If Ambiguas = 0 Then Return ""
                If Ambiguas = 1 Then
                    Return "Um envio à IA ficou sem desfecho conhecido numa execução " &
                           "anterior. Pode ter saído conteúdo, e não dá para saber."
                End If
                Return $"{Ambiguas} envios à IA ficaram sem desfecho conhecido numa " &
                       "execução anterior. Pode ter saído conteúdo, e não dá para saber."
            End Get
        End Property

    End Class

End Namespace

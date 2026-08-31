Imports System.Threading.Tasks
Imports System.Collections.Generic
Imports Iris.Assist
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>De onde a operação tira o que precisa</b> — o pedido, as mensagens
    ''' classificadas e os bytes.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE É UMA PORTA E NÃO CÓDIGO DENTRO DO VIEWMODEL</b>
    '''
    ''' Classificar exige ir ao COM ler o rótulo; montar exige o corpo, que só a
    ''' borda do provider sabe capturar numa leitura só. Nada disso pode viver
    ''' no ViewModel sem arrastar o Outlook para dentro da tela.
    '''
    ''' E a ordem importa: quem chama <see cref="Classificar"/> é o
    ''' <c>DisclosureGate</c>, <b>depois</b> de o preflight passar. Por isso são
    ''' funções, e não valores prontos.
    ''' </summary>
    Public Interface IAssistContext

        ''' <summary>O voo: operação, pasta e destino.</summary>
        Function Pedido(operacao As AssistOperation) As PreflightRequest

        ''' <summary>
        ''' As mensagens do contexto, já classificadas.
        '''
        ''' <b>Só é invocada se o portão deixar</b> — é o que impede ir ao COM
        ''' sem autorização para tocar em item nenhum.
        ''' </summary>
        Function Classificar() As IReadOnlyList(Of MessageClassification)

        ''' <summary>Os bytes. Um envelope, materializado uma vez.</summary>
        Function Montar(operacao As AssistOperation, instrucao As String) As EnvelopeResult

    End Interface

    ''' <summary>
    ''' <b>O contexto que a produção tem: nenhum.</b>
    '''
    ''' Ligar o contexto de verdade — broker lê corpo e <c>PR_CHANGE_KEY</c> numa
    ''' operação, pipeline prepara, envelope monta — é trabalho que só faz
    ''' sentido depois de haver para onde mandar. Fazê-lo agora seria escrever
    ''' um caminho que nada exercita, e o §28.2 diz por que não há para onde
    ''' mandar.
    '''
    ''' Enquanto isso ele recusa, e recusa <b>de um jeito que não engana</b>: o
    ''' destino vazio nunca casa com autorização nenhuma, e o envelope não
    ''' monta. É pendência declarada, não silêncio.
    ''' </summary>
    Friend NotInheritable Class ContextoIndisponivel
        Implements IAssistContext

        Public Function Pedido(operacao As AssistOperation) As PreflightRequest _
                               Implements IAssistContext.Pedido
            Return New PreflightRequest(operacao, Nothing, New AssistDestination("", "", ""))
        End Function

        Public Function Classificar() As IReadOnlyList(Of MessageClassification) _
                                        Implements IAssistContext.Classificar
            Return Array.Empty(Of MessageClassification)()
        End Function

        Public Function Montar(operacao As AssistOperation, instrucao As String) _
                               As EnvelopeResult Implements IAssistContext.Montar
            Return New EnvelopeBuilder(teto:=1).Montar(operacao, instrucao,
                                                       Array.Empty(Of MessagePart)())
        End Function

    End Class

    ''' <summary>
    ''' O rascunho onde a redação é escrita.
    '''
    ''' Porta mínima de propósito: o assistente não precisa saber o que é um
    ''' compositor, e o compositor não precisa saber que existe IA.
    ''' </summary>
    Public Interface IRascunho
        Property Texto As String

        ''' <summary>
        ''' <b>Identidade da sessão de edição.</b> Muda a cada rascunho novo.
        '''
        ''' Sem ela, "o rascunho não mudou" era provado só pelo texto — e dois
        ''' rascunhos diferentes com o mesmo texto (o caso comum: os dois
        ''' vazios) passavam pela prova.
        ''' </summary>
        ReadOnly Property Sessao As Long

        ''' <summary>
        ''' Dá para escrever nele <b>agora</b>. Compositor fechado, ou travado
        ''' durante a confirmação de envio, responde <c>False</c>.
        ''' </summary>
        ReadOnly Property PodeEditar As Boolean

        ''' <summary>
        ''' <b>Da para ABRIR uma resposta agora?</b>
        '''
        ''' Existe porque "mandar a resposta para um rascunho" com o
        ''' compositor fechado nao e uma recusa: e um passo que o programa
        ''' sabe dar. Exigir que o usuario abrisse a resposta antes era uma
        ''' pre-condicao inventada pela forma do codigo -- a porta so sabia
        ''' escrever em compositor aberto -- e nao pelo que ele pediu.
        ''' </summary>
        ReadOnly Property PodeAbrir As Boolean

        ''' <summary>Abre uma resposta para a mensagem aberta.</summary>
        Function AbrirAsync() As Task

        ''' <summary>
        ''' <b>O rascunho mudou</b> — texto, sessão ou editabilidade.
        '''
        ''' Sem este evento, <c>PodeDesfazer</c> ficava correto e <b>invisível</b>:
        ''' o <c>RelayCommand</c> não se reconsulta sozinho, então digitar por
        ''' cima da redação deixava o botão "Desfazer" habilitado até alguma
        ''' outra mudança de estado passar por perto. Clicar recusaria com
        ''' segurança — e a promessa da §38.6, de que a ação fica desabilitada
        ''' quando indisponível, estaria quebrada.
        ''' </summary>
        Event Mudou As EventHandler
    End Interface

    ''' <summary>
    ''' O rascunho do compositor, visto pela porta mínima.
    '''
    ''' O assistente não precisa saber o que é um compositor, e o compositor não
    ''' precisa saber que existe IA. Um adaptador de três linhas custa menos que
    ''' as duas classes se conhecerem.
    ''' </summary>
    Friend NotInheritable Class RascunhoDoCompositor
        Implements IRascunho

        Private ReadOnly _compositor As ComposerViewModel
        Private ReadOnly _abrir As Func(Of Task)
        Private ReadOnly _podeAbrir As Func(Of Boolean)

        ''' <summary>
        ''' <paramref name="abrir"/> e <paramref name="podeAbrir"/> sao o
        ''' "Responder" da barra de cima, vistos por esta porta. Entram como
        ''' funcao, e nao como referencia ao MainViewModel: o assistente nao
        ''' precisa saber que existe uma janela, e o adaptador ja existia
        ''' justamente para as duas classes nao se conhecerem.
        ''' </summary>
        Friend Sub New(compositor As ComposerViewModel,
                       abrir As Func(Of Task),
                       podeAbrir As Func(Of Boolean))
            _compositor = compositor
            _abrir = abrir
            _podeAbrir = podeAbrir
            AddHandler _compositor.PropertyChanged, AddressOf AoMudar
        End Sub

        Public Event Mudou As EventHandler Implements IRascunho.Mudou

        ''' <summary>
        ''' O que do compositor muda o rascunho: o texto, o estado de edição, e
        ''' o fim da sessão.
        '''
        ''' <c>UserText</c> é o que faltava. Ele notifica <c>PropertyChanged</c>
        ''' desde sempre, e ninguém escutava — quem escutava o compositor era o
        ''' <c>MainViewModel</c>, e só para <c>PodeEditar</c>, <c>IsOpen</c> e
        ''' <c>State</c>.
        ''' </summary>
        Private Sub AoMudar(remetente As Object,
                            arg As ComponentModel.PropertyChangedEventArgs)
            Select Case arg.PropertyName
                Case NameOf(ComposerViewModel.UserText),
                     NameOf(ComposerViewModel.PodeEditar),
                     NameOf(ComposerViewModel.IsOpen),
                     NameOf(ComposerViewModel.State)
                    RaiseEvent Mudou(Me, EventArgs.Empty)
            End Select
        End Sub

        Public Property Texto As String Implements IRascunho.Texto
            Get
                Return _compositor.UserText
            End Get
            Set(value As String)
                _compositor.UserText = value
            End Set
        End Property

        ''' <summary>
        ''' A geração do compositor — o contador que ele já mantinha para
        ''' largar continuações em voo quando o rascunho acaba. É exatamente a
        ''' identidade de sessão que o assistente precisa, e reaproveitá-la
        ''' evita duas noções de "rascunho novo" que um dia discordariam.
        ''' </summary>
        Public ReadOnly Property Sessao As Long Implements IRascunho.Sessao
            Get
                Return _compositor.Geracao
            End Get
        End Property

        Public ReadOnly Property PodeEditar As Boolean Implements IRascunho.PodeEditar
            Get
                Return _compositor.PodeEditar
            End Get
        End Property

        Public ReadOnly Property PodeAbrir As Boolean Implements IRascunho.PodeAbrir
            Get
                Return Not _compositor.PodeEditar AndAlso _podeAbrir()
            End Get
        End Property

        Public Function AbrirAsync() As Task Implements IRascunho.AbrirAsync
            Return _abrir()
        End Function
    End Class

End Namespace

Imports System.Collections.Generic
Imports System.Globalization
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
        Implements IDisposable

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

        ''' <summary>
        ''' Quem põe texto na área de transferência.
        '''
        ''' Injetado para o ViewModel não depender de <c>Clipboard</c>: a área
        ''' de transferência exige STA e um desktop de verdade, e um teste que
        ''' precise disso deixa de ser teste de ViewModel.
        ''' </summary>
        Private ReadOnly _copiador As Action(Of String)

        ''' <summary>
        ''' Quando o voo corrente começou. <c>Nothing</c> fora de voo.
        '''
        ''' Vem do <c>_relogio</c> injetado, e não de <c>DateTimeOffset.Now</c>:
        ''' o tempo decorrido é conferível em teste sem esperar de verdade.
        ''' </summary>
        Private _inicioDoVoo As DateTimeOffset?

        ''' <summary>
        ''' Quanto o último voo demorou. <c>Nothing</c> enquanto ele corre.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O CRONÔMETRO TEM DE PARAR DE ANDAR</b>
        '''
        ''' Sem isto, <see cref="Decorrido"/> continuava calculando
        ''' <c>relógio − início</c> <b>depois</b> do voo: parar o
        ''' <c>DispatcherTimer</c> só faz a tela deixar de reperguntar, e não
        ''' congela o valor. Uma chamada de 2,5 s passava a dizer 30 s, 5 min,
        ''' uma hora — e a propriedade que se documenta como "há quanto tempo o
        ''' pedido corrente está rodando" virava um relógio de parede.
        '''
        ''' A ficha escapava por acidente, porque materializa a string antes do
        ''' <c>Finally</c>. Escapar por acidente não é escapar.
        ''' </summary>
        Private _duracaoDoVoo As TimeSpan?

        ''' <summary>
        ''' Só reavisa que <see cref="Decorrido"/> mudou. Não mede nada — quem
        ''' mede é o relógio, e o cálculo é feito na leitura.
        ''' </summary>
        Private ReadOnly _pulso As DispatcherTimer

        Public Sub New(ui As Dispatcher, transmissor As AssistTransmitter,
                       politica As DisclosurePolicy, relogio As Func(Of DateTimeOffset),
                       reconciliacao As ReconciliationResult,
                       contexto As IAssistContext, rascunho As IRascunho,
                       Optional avisoDaAtivacao As String = "",
                       Optional copiador As Action(Of String) = Nothing)
            _ui = ui
            _transmissor = transmissor
            _politica = politica
            _relogio = relogio
            _contexto = If(contexto, CType(New ContextoIndisponivel(), IAssistContext))
            _rascunho = rascunho
            _copiador = copiador
            Me.AvisoDaAtivacao = If(avisoDaAtivacao, "")

            ' SEM _ui NAO HA PULSO, E ISSO NAO E DEFEITO.
            '
            ' Os testes montam o ViewModel com ui = Nothing. Um DispatcherTimer
            ' ali estouraria, e o cronometro nao e a coisa sendo testada --
            ' Decorrido continua conferivel, porque quem mede e o relogio
            ' injetado e a conta acontece na LEITURA. O pulso so avisa a tela.
            If _ui IsNot Nothing Then
                _pulso = New DispatcherTimer(TimeSpan.FromMilliseconds(100),
                                             DispatcherPriority.Normal,
                                             Sub(remetente As Object, arg As EventArgs)
                                                 OnPropertyChanged(NameOf(Decorrido))
                                             End Sub,
                                             _ui)
                ' O construtor de quatro argumentos JA COMECA o timer.
                _pulso.Stop()
            End If

            ' O RASCUNHO MUDOU: reconsulta os comandos.
            '
            ' PodeDesfazer depende do texto, da sessao e da editabilidade do
            ' rascunho, e nenhum deles muda por acao do assistente. Sem escutar,
            ' o estado ficaria certo e invisivel — o RelayCommand nao se
            ' reconsulta sozinho.
            If _rascunho IsNot Nothing Then
                AddHandler _rascunho.Mudou, Sub(remetente As Object, arg As EventArgs)
                                                Avisar()
                                            End Sub
            End If

            ' Comandos declarados a mao, e nao pelo gerador do
            ' CommunityToolkit: ele so roda em C#. Em VB o atributo compila e
            ' NAO gera nada — e o sintoma seria um botao que nunca faz nada.
            ResumirCommand = New AsyncRelayCommand(AddressOf Resumir, Function() PodePedir)
            RedigirCommand = New AsyncRelayCommand(AddressOf Redigir, Function() PodeRedigir)
            CancelarCommand = New RelayCommand(AddressOf Cancelar, Function() PodeCancelar)
            DesfazerCommand = New RelayCommand(AddressOf Desfazer, Function() PodeDesfazer)
            CopiarCommand = New RelayCommand(AddressOf Copiar, Function() PodeCopiar)
            EnviarParaRascunhoCommand = New AsyncRelayCommand(AddressOf EnviarParaRascunho,
                                                             Function() PodeEnviarParaRascunho)
            CopiarRespostaCommand = New RelayCommand(AddressOf CopiarResposta,
                                                    Function() TemResposta)
            ' Me. OBRIGATORIO: sem ele, 'reconciliacao' eclipsa 'Reconciliacao'
            ' — VB e case-insensitive, e a atribuicao vira o parametro para ele
            ' mesmo. O compilador nao avisa; o sintoma aparece longe, como uma
            ' propriedade Nothing na primeira leitura.
            ' RECONHECIMENTO APLICADO NA CONSTRUCAO, e nao lido pela tela a
            ' cada binding: ele mora em disco, e a tela pergunta muitas vezes.
            Me.Reconciliacao = reconciliacao.ComReconhecimento(LerReconhecidas())

            ReconhecerAmbiguasCommand = New RelayCommand(AddressOf ReconhecerAmbiguas,
                                                        Function() Reconciliacao.TemNovidade)
        End Sub

        ''' <summary>
        ''' <b>"Eu vi este aviso."</b>
        '''
        ''' Não desfaz a ambiguidade, e não some com ela: grava quantas já foram
        ''' vistas, para o parágrafo virar ícone. Envio ambíguo novo faz o
        ''' parágrafo voltar inteiro.
        ''' </summary>
        Public ReadOnly Property ReconhecerAmbiguasCommand As IRelayCommand

        ''' <summary>
        ''' Onde o reconhecimento fica. Ao lado do cache, e em texto: o dono
        ''' precisa poder olhar e apagar, como no diário de buscas.
        ''' </summary>
        Friend Shared Function CaminhoDoReconhecimento() As String
            Return IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "ambiguas-reconhecidas.txt")
        End Function

        ''' <summary>
        ''' Quantas o dono já reconheceu. <b>Falha ao ler vale ZERO</b> — não
        ''' conseguir ler o reconhecimento não é ter reconhecido, e o custo de
        ''' errar para este lado é mostrar um aviso a mais.
        ''' </summary>
        Friend Shared Function LerReconhecidas() As Integer
            Try
                Dim caminho = CaminhoDoReconhecimento()
                If Not IO.File.Exists(caminho) Then Return 0
                Dim n As Integer
                If Integer.TryParse(IO.File.ReadAllText(caminho).Trim(), n) Then
                    Return Math.Max(0, n)
                End If
                Return 0
            Catch
                Return 0
            End Try
        End Function

        Private Sub ReconhecerAmbiguas()
            Try
                Dim caminho = CaminhoDoReconhecimento()
                IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(caminho))
                IO.File.WriteAllText(caminho, Reconciliacao.Ambiguas.ToString())
                Reconciliacao = Reconciliacao.ComReconhecimento(Reconciliacao.Ambiguas)
            Catch
                ' NAO CONSEGUIU GRAVAR: o aviso FICA. Reconhecimento que nao
                ' persiste seria pior que nenhum -- o parágrafo sumiria nesta
                ' sessao e voltaria na proxima, ensinando que o botao nao
                ' funciona.
            End Try

            OnPropertyChanged(NameOf(Reconciliacao))
            OnPropertyChanged(NameOf(TemAvisoDaReconciliacao))
            OnPropertyChanged(NameOf(TemMarcaDaReconciliacao))
            OnPropertyChanged(NameOf(TemAlgoADizer))
            ReconhecerAmbiguasCommand.NotifyCanExecuteChanged()
        End Sub

        ' ==============================================================
        ' O estado

        ''' <summary>O que aconteceu na reconciliação da abertura.</summary>
        Public Property Reconciliacao As ReconciliationResult

        ''' <summary>
        ''' <b>O que a cerimônia de ativação tem a dizer — sempre visível.</b>
        '''
        ''' Duas coisas cabem aqui, e as duas precisam ficar na tela o tempo
        ''' todo em vez de aparecer só quando algo falha:
        '''
        ''' <list type="bullet">
        ''' <item>o <b>motivo</b> de o arquivo de ativação não ter sido aceito —
        ''' campo com erro de digitação, prazo vencido, JSON malformado. Sem
        ''' isso, quem escreveu o arquivo errado vê a mesma frase de quem nunca
        ''' escreveu arquivo nenhum;</item>
        ''' <item>que a <b>política corporativa não foi verificada</b>. Isso não
        ''' impede a ativação — é decisão do dono da caixa —, mas some da vista
        ''' se ficar só dentro do arquivo, e some justamente enquanto a IA
        ''' funciona bem e ninguém tem motivo para reler nada.</item>
        ''' </list>
        '''
        ''' Separado do <see cref="Aviso"/> de propósito: aquele é substituído a
        ''' cada operação, e este não pode ser.
        ''' </summary>
        Public ReadOnly Property AvisoDaAtivacao As String

        Public ReadOnly Property TemAvisoDaAtivacao As Boolean
            Get
                Return AvisoDaAtivacao.Length > 0
            End Get
        End Property

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
        Public ReadOnly Property CopiarCommand As RelayCommand

        ''' <summary>Manda a resposta redigida para o rascunho aberto.</summary>
        Public ReadOnly Property EnviarParaRascunhoCommand As AsyncRelayCommand

        ''' <summary>
        ''' Copia a RESPOSTA. Separado do <see cref="CopiarCommand"/>, que copia
        ''' o resumo: são dois quadros, e um botão que copiasse "o que estiver
        ''' ali" copiaria o errado metade das vezes.
        ''' </summary>
        Public ReadOnly Property CopiarRespostaCommand As RelayCommand

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
                OnPropertyChanged(NameOf(PorQueNaoRedige))
                OnPropertyChanged(NameOf(TemMotivoParaNaoRedigir))
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
        ''' <summary>
        ''' <b>O que NÃO foi lido para produzir este resultado.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE ELA PRECISOU EXISTIR</b>
        '''
        ''' Até 30/08/2026 o portão negava qualquer mensagem com anexo, e imagem
        ''' embutida contava como anexo. Medido numa pasta real: <b>13 de 13</b>
        ''' mensagens recusadas, dez delas só por causa do logo da assinatura.
        '''
        ''' A regra passou a distinguir — e é aí que esta propriedade entra. Uma
        ''' <b>captura de tela colada no corpo</b> é embutida do mesmo jeito que
        ''' um logo, e pode carregar o teor inteiro da mensagem. Deixar de negar
        ''' sem declarar teria trocado uma recusa honesta por um resumo
        ''' silenciosamente parcial.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>CURTA DE PROPÓSITO</b>
        '''
        ''' Numa caixa corporativa quase toda mensagem tem logo de assinatura,
        ''' então esta linha aparece quase sempre. Ressalva que aparece sempre e
        ''' é longa vira ruído — e ruído ensina a não ler a ressalva de verdade.
        ''' Uma linha, um número, sem adjetivo.
        ''' </summary>
        Public Property RessalvaDoConteudo As String
            Get
                Return _ressalvaDoConteudo
            End Get
            Private Set(value As String)
                SetProperty(_ressalvaDoConteudo, If(value, ""))
                OnPropertyChanged(NameOf(TemRessalvaDoConteudo))
            End Set
        End Property
        Private _ressalvaDoConteudo As String = ""

        Public ReadOnly Property TemRessalvaDoConteudo As Boolean
            Get
                Return _ressalvaDoConteudo.Length > 0
            End Get
        End Property

        ''' <summary>
        ''' A frase, a partir do que as classificações disseram.
        '''
        ''' <c>Nothing</c> em alguma delas vira "não sei quantas" — e não some,
        ''' nem vira zero. É a mesma regra das outras contagens desta base.
        ''' </summary>
        Friend Shared Function DizerOQueFicouDeFora(
                msgs As IReadOnlyList(Of MessageClassification)) As String
            If msgs Is Nothing OrElse msgs.Count = 0 Then Return ""

            If msgs.Any(Function(m) Not m.Embutidas.HasValue) Then
                Return "Não sei quantas imagens embutidas esta mensagem tem; " &
                       "nenhuma delas foi lida."
            End If

            Dim total = msgs.Sum(Function(m) m.Embutidas.Value)
            If total = 0 Then Return ""
            Return If(total = 1,
                      "1 imagem embutida não foi lida.",
                      $"{total} imagens embutidas não foram lidas.")
        End Function

        ''' <summary>
        ''' <b>A resposta redigida — separada do resumo, e não no lugar dele.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O QUE MUDOU, E POR QUÊ</b>
        '''
        ''' Antes, redigir <b>substituía</b> o resumo na tela e <b>escrevia
        ''' direto no rascunho</b>. Duas coisas erradas de uma vez: o resumo que
        ''' o usuário acabou de pagar sumia, e a aplicação no rascunho acontecia
        ''' sem ele pedir — o botão dizia "redigir", e ele também aplicava.
        '''
        ''' Agora são dois quadros e dois atos. Redigir mostra; <b>Enviar para
        ''' rascunho</b> aplica. Um botão, uma coisa.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>E A REDAÇÃO USA O RESUMO</b>
        '''
        ''' Por isso ela só é oferecida depois dele. O modelo recebe o e-mail
        ''' <i>e</i> o resumo — o resumo não amplia o que sai da caixa (o e-mail
        ''' já saiu na primeira chamada), mas dá ao segundo pedido o que o
        ''' primeiro concluiu.
        ''' </summary>
        Public Property Resposta As String
            Get
                Return _resposta
            End Get
            Private Set(value As String)
                SetProperty(_resposta, If(value, ""))
                OnPropertyChanged(NameOf(TemResposta))
                OnPropertyChanged(NameOf(PodeEnviarParaRascunho))
                OnPropertyChanged(NameOf(PorQueNaoEnvia))
                OnPropertyChanged(NameOf(TemMotivoParaNaoEnviar))
                EnviarParaRascunhoCommand?.NotifyCanExecuteChanged()
                CopiarRespostaCommand?.NotifyCanExecuteChanged()
            End Set
        End Property
        Private _resposta As String = ""

        Public ReadOnly Property TemResposta As Boolean
            Get
                Return _resposta.Length > 0
            End Get
        End Property

        ''' <summary>
        ''' <b>A primeira linha do resumo, sozinha.</b>
        '''
        ''' A instrução pede uma frase, linha em branco, e o resto. Aqui a
        ''' frase é recortada para a tela poder mostrá-la em destaque sem
        ''' pedir <b>outro</b> resumo ao modelo — seria a mesma leitura
        ''' cobrada duas vezes.
        '''
        ''' <b>Recorte, e não promessa.</b> Se o modelo ignorar o formato e
        ''' devolver um parágrafo, isto devolve a primeira linha dele, que é
        ''' o começo do parágrafo. Fica pior que o pedido e melhor que vazio,
        ''' e nunca inventa uma frase que o modelo não escreveu.
        ''' </summary>
        Public ReadOnly Property ResumoDeUmaLinha As String
            Get
                Return PrimeiraLinha(_resultado)
            End Get
        End Property

        ''' <summary>
        ''' <b>O corte, num lugar só.</b>
        '''
        ''' Compartilhado porque o duplo dos testes de renderização também
        ''' precisa dele: um duplo que cortasse por conta própria poderia
        ''' cortar <i>diferente</i>, e aí a tela estaria certa no teste e
        ''' errada na janela — que é o defeito que duplo nenhum deve poder
        ''' esconder.
        ''' </summary>
        Friend Shared Function PrimeiraLinha(texto As String) As String
            If String.IsNullOrEmpty(texto) Then Return ""
            Dim quebra = texto.IndexOf(vbLf, StringComparison.Ordinal)
            If quebra < 0 Then Return texto.Trim()
            Return texto.Substring(0, quebra).Trim()
        End Function

        ''' <summary>O que vem depois da primeira linha.</summary>
        Friend Shared Function DepoisDaPrimeiraLinha(texto As String) As String
            If String.IsNullOrEmpty(texto) Then Return ""
            Dim quebra = texto.IndexOf(vbLf, StringComparison.Ordinal)
            If quebra < 0 Then Return ""
            Return texto.Substring(quebra + 1).Trim()
        End Function

        ''' <summary>O que vem depois da primeira linha. Vazio se não vier nada.</summary>
        Public ReadOnly Property ResumoDetalhado As String
            Get
                Return DepoisDaPrimeiraLinha(_resultado)
            End Get
        End Property

        Public ReadOnly Property TemResumoDetalhado As Boolean
            Get
                Return ResumoDetalhado.Length > 0
            End Get
        End Property

        Public Property Resultado As String
            Get
                Return _resultado
            End Get
            Private Set(value As String)
                SetProperty(_resultado, value)
                OnPropertyChanged(NameOf(TemResultado))
                OnPropertyChanged(NameOf(ResumoDeUmaLinha))
                OnPropertyChanged(NameOf(ResumoDetalhado))
                OnPropertyChanged(NameOf(TemResumoDetalhado))
                ' O Copiar depende DISTO, e so disto. Sem avisar aqui, o botao
                ' so acordaria quando outra coisa qualquer chamasse Avisar().
                Avisar()
            End Set
        End Property

        ''' <summary>
        ''' <b>Há resultado?</b> — e espaço em branco <b>não</b> é resultado.
        '''
        ''' Era <c>Length > 0</c>, e por isso uma resposta de três espaços ou de
        ''' uma quebra de linha escapava do aviso de "respondeu sem texto",
        ''' deixava a faixa visualmente vazia, e — pior — era <b>aplicada por
        ''' cima do rascunho do usuário</b> na redação. Trocar o texto dele por
        ''' espaços é perda de trabalho com cara de sucesso.
        ''' </summary>
        Public ReadOnly Property TemResultado As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(Resultado)
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
                OnPropertyChanged(NameOf(TemAvisoDeOperacao))
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
                Return TemAviso OrElse Reconciliacao.Aviso.Length > 0 OrElse
                       TemAvisoDaAtivacao
            End Get
        End Property

        ''' <summary>
        ''' <b>Há aviso de operação?</b> — sem contar o da cerimônia.
        '''
        ''' ------------------------------------------------------------------
        ''' Existe porque o aviso da cerimônia ganhou <b>linha própria</b>, no
        ''' rodapé. Enquanto a linha do aviso comum era governada por
        ''' <see cref="TemAlgoADizer"/>, ela ficava visível por causa do rodapé
        ''' — com os dois <c>TextBlock</c> dela <b>vazios</b>, ocupando altura.
        '''
        ''' O sintoma era um vão de dedo entre os botões e a resposta, que
        ''' ninguém conseguia explicar olhando o XAML: o espaço era de um
        ''' elemento presente, e não de margem.
        '''
        ''' <see cref="TemAlgoADizer"/> continua significando "a faixa tem algo
        ''' a dizer", que é outra pergunta e tem outros usos.
        ''' </summary>
        Public ReadOnly Property TemAvisoDeOperacao As Boolean
            Get
                Return TemAviso OrElse Reconciliacao.Aviso.Length > 0
            End Get
        End Property

        Public ReadOnly Property TemAvisoDaReconciliacao As Boolean
            Get
                Return Reconciliacao.Aviso.Length > 0
            End Get
        End Property

        ''' <summary>
        ''' A versão curta, para quando o dono já reconheceu. Ela <b>não</b>
        ''' substitui a longa: as duas nunca aparecem juntas, e qual das duas
        ''' aparece é o reconhecimento que decide.
        ''' </summary>
        Public ReadOnly Property TemMarcaDaReconciliacao As Boolean
            Get
                Return Reconciliacao.TemMarca
            End Get
        End Property

        ''' <summary>
        ''' <b>A IA está disponível?</b> — para <b>alguma</b> das operações.
        '''
        ''' Duas condições, e as duas são fechadas por padrão: a reconciliação
        ''' terminou, e o portão aceita pelo menos um dos dois voos. Em produção
        ''' a segunda é <c>False</c> — <c>ActivationRecord.DaProducao</c> é
        ''' <c>Nothing</c>.
        ''' </summary>
        Public ReadOnly Property Disponivel As Boolean
            Get
                Return Reconciliacao.Terminou AndAlso (_portaoResumir OrElse _portaoRedigir)
            End Get
        End Property

        ''' <summary>
        ''' O portão aceita <b>resumir</b>, e o portão aceita <b>redigir</b> —
        ''' separados.
        '''
        ''' Havia um só, calculado com <c>AssistOperation.Resumir</c> e usado
        ''' para habilitar os dois botões. Como a ativação lista as operações
        ''' autorizadas uma a uma (<c>Operacoes.Contains</c>), uma ativação só
        ''' para resumo habilitava visualmente a redação — que seria negada
        ''' depois, com o motivo aparecendo tarde —, e uma ativação só para
        ''' redação a deixava inalcançável.
        ''' </summary>
        Private _portaoResumir As Boolean
        Private _portaoRedigir As Boolean

        Public ReadOnly Property PodePedir As Boolean
            Get
                Return Reconciliacao.Terminou AndAlso _portaoResumir AndAlso Not Ocupado
            End Get
        End Property

        Public ReadOnly Property PodeCancelar As Boolean
            Get
                Return Ocupado
            End Get
        End Property

        ''' <summary>
        ''' Copiar exige <b>ter o que copiar</b>, e nada além disso.
        '''
        ''' Não depende do portão: o texto já está na tela, e já foi pago. Amarrar
        ''' a cópia à autorização faria a IA vencer entre o resumo aparecer e o
        ''' usuário conseguir levá-lo embora.
        ''' </summary>
        Public ReadOnly Property PodeCopiar As Boolean
            Get
                Return TemResultado AndAlso _copiador IsNot Nothing
            End Get
        End Property

        ''' <summary>
        ''' <b>Há quanto tempo o pedido corrente está rodando.</b>
        '''
        ''' Existe porque uma chamada a modelo demora segundos, e sem número na
        ''' tela "está pensando" e "travou" são a mesma coisa para quem olha. O
        ''' botão Cancelar ao lado só é uma escolha de verdade se o usuário
        ''' souber quanto já esperou.
        '''
        ''' A conta é feita na leitura, sobre o relógio injetado: o cronômetro
        ''' não guarda estado que possa divergir do voo.
        ''' </summary>
        Public ReadOnly Property Decorrido As String
            Get
                Dim quanto As TimeSpan
                If _duracaoDoVoo.HasValue Then
                    quanto = _duracaoDoVoo.Value
                ElseIf _inicioDoVoo.HasValue Then
                    quanto = _relogio() - _inicioDoVoo.Value
                Else
                    Return ""
                End If

                Dim s = quanto.TotalSeconds
                If s < 0 Then s = 0
                Return s.ToString("0.0", Daqui) & " s"
            End Get
        End Property

        Private _ficha As String = ""
        ''' <summary>
        ''' <b>Quem atendeu, e quanto custou.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O AGENTE E O MODELO VÊM DA ATIVAÇÃO, NÃO DA RESPOSTA</b>
        '''
        ''' O OpenRouter devolve um campo <c>provider</c> no corpo da resposta,
        ''' e seria mais fácil mostrar aquilo. Mas é <b>texto escolhido pelo
        ''' outro lado</b>, e a §29.5 diz onde o dado de fora para. O que
        ''' aparece aqui é o que <b>o usuário assinou</b> na cerimônia — e é
        ''' também a resposta mais útil, porque a pergunta "para onde meu e-mail
        ''' foi" se responde com a autorização, não com a nota fiscal.
        '''
        ''' Da resposta entram só os <b>números</b>: custo e tokens.
        '''
        ''' Vazio quando não houve voo. Nada de "0 tokens, US$ 0,00" antes do
        ''' primeiro pedido — zero é uma afirmação, e ali não há nada a afirmar.
        ''' </summary>
        Public Property Ficha As String
            Get
                Return _ficha
            End Get
            Private Set(valor As String)
                SetProperty(_ficha, If(valor, ""))
                OnPropertyChanged(NameOf(TemFicha))
            End Set
        End Property

        Public ReadOnly Property TemFicha As Boolean
            Get
                Return Ficha.Length > 0
            End Get
        End Property

        ''' <summary>
        ''' Monta a ficha. <c>Decimal</c> com quatro casas: uma chamada de
        ''' modelo barato custa frações de centavo, e arredondar para dois
        ''' transformaria todo custo real em "US$ 0,00" — que é pior que não
        ''' mostrar, porque afirma gratuidade.
        ''' </summary>
        ''' <summary>
        ''' <b>A cultura da ficha é fixa, e não a da máquina.</b>
        '''
        ''' Toda a interface do Iris é escrita em português; formatar número
        ''' pela cultura ambiente faria "US$ 0.0004" aparecer no meio de frases
        ''' em português numa máquina configurada em inglês, e o mesmo texto
        ''' mudar de forma conforme o Windows de quem abre.
        '''
        ''' E torna o número conferível: um teste que dependa da cultura do
        ''' processo está testando o host, não este código.
        ''' </summary>
        Private Shared ReadOnly Daqui As CultureInfo = CultureInfo.GetCultureInfo("pt-BR")

        Private Function Fichar(destino As AssistDestination,
                                r As AssistOutcome) As String
            Dim partes As New List(Of String)()
            If destino IsNot Nothing Then
                If destino.Provedor.Length > 0 Then partes.Add(destino.Provedor)
                If destino.Modelo.Length > 0 Then partes.Add(destino.Modelo)
            End If
            If r.Tokens.HasValue Then partes.Add(r.Tokens.Value.ToString("N0", Daqui) & " tokens")
            ' "informado", e nao o valor seco: o numero e a palavra do
            ' provedor, e nao um fato conferido. Ver ProviderOutcome.Custo.
            If r.Custo.HasValue Then
                partes.Add("US$ " & r.Custo.Value.ToString("N4", Daqui) & " informado")
            End If
            If _inicioDoVoo.HasValue Then partes.Add(Decorrido)
            Return String.Join("  ·  ", partes)
        End Function

        ''' <summary>
        ''' <b>Leva a resposta do modelo para a área de transferência.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' Texto, e só texto — o mesmo que está na tela. Não é egress: a área
        ''' de transferência é local, e quem pediu foi o usuário.
        '''
        ''' Falha da área de transferência não derruba nada: ela é disputada
        ''' entre processos e recusa por motivos que não têm nada a ver com o
        ''' Iris. Deixar a exceção subir mataria a janela por causa de um botão
        ''' de conveniência.
        ''' </summary>
        Private Sub CopiarResposta()
            If Not TemResposta Then Return
            Try
                _copiador(Resposta)
            Catch
                Aviso = "Não consegui usar a área de transferência. " &
                        "O texto continua aí para copiar à mão."
            End Try
        End Sub

        Private Sub Copiar()
            If Not PodeCopiar Then Return
            Try
                _copiador(Resultado)
            Catch
                Aviso = "Não consegui usar a área de transferência. " &
                        "O texto continua aí para copiar à mão."
            End Try
        End Sub

        ' ==============================================================

        ''' <summary>
        ''' Pergunta ao portão se o voo passa, <b>sem classificar item nenhum</b>,
        ''' e guarda o motivo para a tela.
        '''
        ''' É o preflight: sem autorização não se gasta uma ida ao COM lendo
        ''' rótulo de coisa nenhuma.
        ''' </summary>
        ''' <summary>
        ''' Reavalia com o contexto corrente — <b>as duas operações</b>.
        ''' </summary>
        Public Sub Avaliar()
            Dim resumir = _politica.Preflight(
                _contexto.Pedido(AssistOperation.Resumir), _relogio())
            Dim redigir = _politica.Preflight(
                _contexto.Pedido(AssistOperation.Redigir), _relogio())

            _portaoResumir = resumir.Permitido
            _portaoRedigir = redigir.Permitido
            Aviso = Explicar(resumir, redigir)

            OnPropertyChanged(NameOf(Disponivel))
            OnPropertyChanged(NameOf(PodePedir))
            OnPropertyChanged(NameOf(PodeRedigir))
            Avisar()
        End Sub

        ''' <summary>
        ''' O que a faixa diz quando as duas operações não concordam.
        '''
        ''' Se nenhuma passa, o motivo é um só e é ele que aparece. Se só uma
        ''' passa, o botão da outra fica desabilitado — e um botão desabilitado
        ''' sem motivo ao lado é a forma mais silenciosa de esconder uma recusa.
        ''' </summary>
        Private Shared Function Explicar(resumir As DisclosureDecision,
                                         redigir As DisclosureDecision) As String
            If resumir.Permitido AndAlso redigir.Permitido Then Return ""
            If Not resumir.Permitido AndAlso Not redigir.Permitido Then
                Return resumir.Explicacao
            End If
            If resumir.Permitido Then
                Return "Redigir resposta não está disponível: " & redigir.Explicacao
            End If
            Return "Resumir não está disponível: " & resumir.Explicacao
        End Function

        ''' <summary>Os comandos reavaliam o que podem fazer.</summary>
        Private Sub Avisar()
            ResumirCommand.NotifyCanExecuteChanged()
            RedigirCommand.NotifyCanExecuteChanged()
            CancelarCommand.NotifyCanExecuteChanged()
            ' O DESFAZER TAMBEM: ele depende da sessao e do estado do rascunho,
            ' e fechar o compositor tem de apaga-lo da tela.
            DesfazerCommand.NotifyCanExecuteChanged()
            OnPropertyChanged(NameOf(PodeDesfazer))
            ' E O COPIAR. Ele nasceu sem esta linha, e o botao ficava
            ' desabilitado para sempre: PodeCopiar virava True quando a
            ' resposta chegava, e ninguem avisava o RelayCommand.
            '
            ' O teste que eu tinha escrito perguntava vm.PodeCopiar -- a
            ' PROPRIEDADE -- e passava. E o proprio arquivo ja documenta essa
            ' armadilha, no Desfazer: "perguntar CanExecute nao pegaria isso:
            ' a resposta estaria certa e o botao errado". Escrevi o mesmo
            ' defeito duas telas abaixo do aviso sobre ele.
            CopiarCommand.NotifyCanExecuteChanged()
            OnPropertyChanged(NameOf(PodeCopiar))

            ' E O ENVIAR PARA RASCUNHO -- O MESMO DEFEITO, DUAS TELAS ABAIXO
            ' DO PARAGRAFO QUE O DESCREVE.
            '
            ' PodeEnviarParaRascunho depende do compositor estar EDITAVEL, e
            ' o compositor e a unica coisa da lista que muda por fora: abrir
            ' uma resposta e o que liga o botao. Sem estas linhas, quem
            ' clicava em Responder via o botao continuar apagado -- e a frase
            ' embaixo dele continuar mandando abrir uma resposta que ja
            ' estava aberta.
            '
            ' O evento Mudou do rascunho ja chegava aqui. So nao saia daqui.
            EnviarParaRascunhoCommand?.NotifyCanExecuteChanged()
            OnPropertyChanged(NameOf(PodeEnviarParaRascunho))
            OnPropertyChanged(NameOf(PorQueNaoEnvia))
            OnPropertyChanged(NameOf(TemMotivoParaNaoEnviar))
        End Sub

        ''' <summary>
        ''' <b>Troca de contexto.</b> Incrementa a geração, limpa o resultado e
        ''' cancela o que estiver em voo.
        '''
        ''' Chamada quando o usuário muda de mensagem. Sem ela, o resumo da
        ''' anterior apareceria embaixo da nova.
        ''' </summary>
        Public Sub Trocou()
            Trocou(Nothing)
        End Sub

        ''' <summary>
        ''' <b>Troca de contexto, guardando o que já foi pago.</b>
        '''
        ''' Resumo e resposta custam dinheiro e tempo. Clicar noutra
        ''' mensagem e voltar apagava os dois, e o usuário pagava de novo
        ''' pela mesma pergunta — ou, pior, não voltava a perguntar.
        '''
        ''' <b>A memória é por <c>ItemKey</c>, e essa é a regra inteira.</b>
        ''' O perigo aqui nunca foi perder o resumo: é mostrar o resumo de
        ''' uma mensagem embaixo de outra, que é pior que resumo nenhum
        ''' porque ninguém desconfia. Chave ausente ou vazia — troca de
        ''' pasta, desmarcação — <b>não restaura nada</b>, e é por isso que
        ''' a sobrecarga sem argumento continua existindo: quem não sabe
        ''' de que mensagem está falando não tem o que restaurar.
        '''
        ''' Ela mora na memória do processo e não em disco: é conteúdo de
        ''' mensagem, e o cache em disco guarda metadado (D1).
        ''' <see cref="EsquecerASessao"/> a apaga quando a sessão do Outlook
        ''' é outra e as identidades deixam de valer.
        ''' </summary>
        Public Sub Trocou(chave As ItemKey)
            ' FECHADO E FECHADO. Sem esta linha, um evento de selecao que
            ' chegasse depois do Dispose -- e eles chegam, porque a janela
            ' desmonta em ordem propria -- voltaria a disparar resumo
            ' automatico num assistente que ninguem esta mais olhando.
            If _descartado Then Return

            _geracao += 1
            Cancelar()

            Guardar()
            _chaveAtual = If(chave IsNot Nothing AndAlso Not chave.IsEmpty, chave, Nothing)
            Restaurar()

            ' E REAVALIA. Invalidar sem reavaliar deixava os comandos refletindo
            ' o contexto anterior: pasta nova pode ter outra autorizacao, e o
            ' botao continuaria habilitado — ou desabilitado — pelo motivo
            ' errado.
            Avaliar()

            ' E, se o interruptor estiver ligado, resume sozinho. Sem Await:
            ' a troca de mensagem nao pode esperar por rede. Quem precisa
            ' esperar e o teste, por EsperarOResumoAutomatico.
            '
            ' O CTS NOVO NASCE AQUI, e o anterior morre aqui: e a troca de
            ' mensagem que cancela a espera da mensagem anterior, e essa e a
            ' unica coisa que impede um pedido por linha atravessada.
            Dim anterior = _esperaDoResumo
            _esperaDoResumo = New CancellationTokenSource()
            anterior?.Cancel()
            anterior?.Dispose()

            _resumoAutomatico = ResumirSozinho(_geracao, _esperaDoResumo.Token)
        End Sub

        ''' <summary>
        ''' <b>Resumir ao abrir</b> — e ele nasce DESLIGADO.
        '''
        ''' Ligado, abrir uma mensagem passa a mandar conteúdo para fora sem
        ''' clique nenhum. Isso é uma mudança de categoria, e não de conforto:
        ''' por isso é decisão gravada em disco, e por isso
        ''' <b>falha ao conferir vale como desligado</b> — não conseguir ler o
        ''' consentimento nunca é autorização.
        '''
        ''' O marcador é o mesmo desenho do diário de buscas, com o sinal
        ''' trocado: lá a ausência liga, aqui a presença. O padrão de cada um
        ''' é o lado seguro do seu caso.
        ''' </summary>
        Public Property ResumirAoAbrir As Boolean
            Get
                Try
                    Return IO.File.Exists(CaminhoDoResumoAutomatico())
                Catch
                    Return False
                End Try
            End Get
            Set(value As Boolean)
                Try
                    Dim caminho = CaminhoDoResumoAutomatico()
                    If value Then
                        IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(caminho))
                        IO.File.WriteAllText(caminho,
                            "Enquanto este arquivo existir, o Iris resume ao abrir a mensagem." &
                            Environment.NewLine &
                            "Apague-o, ou use o botao na tela, para voltar a resumir so a pedido." &
                            Environment.NewLine, Text.Encoding.UTF8)
                    ElseIf IO.File.Exists(caminho) Then
                        IO.File.Delete(caminho)
                    End If
                Catch
                    ' Nao deu para gravar: o estado continua sendo o do disco, e
                    ' o disco nao mudou. A tela relê e mostra a verdade.
                End Try
                OnPropertyChanged(NameOf(ResumirAoAbrir))
            End Set
        End Property

        Friend Shared Function CaminhoDoResumoAutomatico() As String
            Return IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "resumir-ao-abrir.txt")
        End Function

        ''' <summary>
        ''' <b>A espera antes de resumir sozinho.</b>
        '''
        ''' Descer a lista com a seta dispararia um pedido por mensagem
        ''' atravessada. Cancelar não desfaz requisição que já saiu — o duplo
        ''' do provedor desta base existe justamente para lembrar disso — então
        ''' a defesa é <b>não sair</b>: a troca seguinte cancela a espera antes
        ''' de haver pedido.
        '''
        ''' Ajustável porque o teste não pode esperar de verdade.
        ''' </summary>
        Friend Property EsperaAntesDeResumir As TimeSpan = TimeSpan.FromMilliseconds(800)

        Private _resumoAutomatico As Task = Task.CompletedTask

        ''' <summary>
        ''' <b>O cancelamento DA ESPERA — e ele não é o do voo.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O PRIMEIRO CORTE USAVA <c>_cancelamento</c>, E ISSO NÃO FUNCIONA</b>
        '''
        ''' <c>_cancelamento</c> é criado dentro de <c>Pedir</c> e volta a
        ''' <c>Nothing</c> quando o voo acaba. No momento em que a espera
        ''' começa, ele é uma de duas coisas, e as duas quebram:
        '''
        ''' <list type="bullet">
        ''' <item><b><c>Nothing</c></b>, no caso comum — nenhum voo em
        ''' andamento. <c>_cancelamento.Token</c> lança
        ''' <c>NullReferenceException</c> dentro de uma task que ninguém
        ''' espera, e o resumo automático simplesmente não acontece.</item>
        ''' <item><b>o CTS do voo anterior</b>, que o <c>Cancelar()</c> do
        ''' próprio <c>Trocou</c> acabou de cancelar. O <c>Delay</c> nasce
        ''' cancelado, e o resumo também não acontece.</item>
        ''' </list>
        '''
        ''' O recurso ficou <b>inteiramente quebrado em produção</b> — a espera
        ''' padrão é de 800 ms, então o caminho do <c>Delay</c> é o único que
        ''' roda lá — e a suíte ficou verde porque <b>todos</b> os testes
        ''' zeravam a espera e pulavam o <c>Delay</c>. O único que a usava
        ''' cobrava <i>zero chamadas</i>, e zero era o que a exceção produzia.
        ''' Teste verde não é prova, e o controle positivo do caminho lento é
        ''' o que faltava.
        '''
        ''' Achado por revisão externa em 31/08/2026.
        ''' </summary>
        Private _esperaDoResumo As CancellationTokenSource

        ''' <summary>
        ''' O resumo automático da última troca. Existe para o teste poder
        ''' esperá-lo: na janela ninguém espera, e é isso que se quer.
        ''' </summary>
        Friend Function EsperarOResumoAutomatico() As Task
            Return _resumoAutomatico
        End Function

        ''' <summary>
        ''' <b>Quatro perguntas antes de gastar uma chamada</b>, e a primeira
        ''' é a memória: mensagem já resumida não se resume de novo, senão
        ''' ir e voltar na lista cobra duas vezes pelo mesmo texto.
        ''' </summary>
        Private Async Function ResumirSozinho(minha As Integer,
                                              paradaDaEspera As CancellationToken) As Task
            If Not ResumirAoAbrir Then Return
            If _chaveAtual Is Nothing Then Return
            If TemResultado Then Return
            If Not PodePedir Then Return

            Try
                If EsperaAntesDeResumir > TimeSpan.Zero Then
                    Await Task.Delay(EsperaAntesDeResumir, paradaDaEspera)
                End If
            Catch ex As OperationCanceledException
                Return
            End Try

            ' A GERACAO DE NOVO, depois da espera. Trocar de mensagem durante
            ' os 800 ms cancela o Delay, mas cancelamento e corrida: sem esta
            ' conferencia, uma troca que chegasse entre o fim da espera e a
            ' chamada poria o resumo desta mensagem embaixo da proxima.
            If minha <> _geracao Then Return
            If TemResultado OrElse Not PodePedir Then Return

            Await Resumir()
        End Function

        ''' <summary>
        ''' <b>A sessão é outra: a memória não vale mais.</b>
        '''
        ''' <c>EntryID</c> é identidade de uma ligação com o Outlook. Numa
        ''' sessão nova a mesma cadeia pode apontar para outra coisa, ou
        ''' para nada — guardar seria apostar que aponta para a mesma.
        ''' <b>Não se chama <c>Esquecer</c></b> porque esse nome já é do
        ''' desfazer, três telas abaixo, e falam de coisas diferentes.
        ''' </summary>
        Public Sub EsquecerASessao()
            ' A GERACAO SOBE, E O QUE ESTIVER EM VOO E CANCELADO.
            '
            ' Sessao nova do Outlook e a troca de contexto mais forte que
            ' existe -- mais forte que trocar de mensagem -- e este metodo
            ' limpava a memoria sem tocar na geracao. Um resumo em voo voltava
            ' e era publicado como se nada tivesse mudado, e um envio ao
            ' rascunho atravessava a guarda de geracao intacto.
            '
            ' Achado por revisao externa em 31/08/2026.
            _geracao += 1
            Cancelar()

            _memoria.Clear()
            _ordem.Clear()
            _chaveAtual = Nothing
            Resultado = ""
            Resposta = ""
            Ficha = ""
            RessalvaDoConteudo = ""
        End Sub

        Private NotInheritable Class TrabalhoDaIa
            Public Property Resumo As String = ""
            Public Property Redacao As String = ""
            Public Property Ficha As String = ""
            Public Property Ressalva As String = ""
        End Class

        ' TETO DE VINTE, e descarte pela ORDEM DE CHEGADA.
        '
        ' Sem teto, uma varredura de pasta grande acumularia o corpo de
        ' resumo de cada mensagem visitada pelo tempo que o programa ficasse
        ' aberto. Vinte cobre o vaivem de quem le uma conversa e volta; alem
        ' disso e memoria por acaso, e nao por decisao.
        Private Const TetoDaMemoria As Integer = 20
        Private ReadOnly _memoria As New Dictionary(Of ItemKey, TrabalhoDaIa)()
        Private ReadOnly _ordem As New List(Of ItemKey)()
        Private _chaveAtual As ItemKey

        ''' <summary>O que está na tela vai para a chave que o pôs lá.</summary>
        Private Sub Guardar()
            If _chaveAtual Is Nothing Then Return

            ' Nada na tela apaga o que havia: e a diferenca entre 'ainda nao
            ' pedi' e 'pedi e nao deu'. Guardar vazio por cima de um resumo
            ' bom seria a memoria trabalhando contra si mesma.
            If Not TemResultado AndAlso Not TemResposta Then
                If _memoria.Remove(_chaveAtual) Then _ordem.Remove(_chaveAtual)
                Return
            End If

            If Not _memoria.ContainsKey(_chaveAtual) Then
                _ordem.Add(_chaveAtual)
                Do While _ordem.Count > TetoDaMemoria
                    _memoria.Remove(_ordem(0))
                    _ordem.RemoveAt(0)
                Loop
            End If

            _memoria(_chaveAtual) = New TrabalhoDaIa() With {
                .Resumo = Resultado, .Redacao = Resposta,
                .Ficha = Ficha, .Ressalva = RessalvaDoConteudo}
        End Sub

        ''' <summary>
        ''' O que a chave nova tem guardado — <b>ou o vazio</b>. Nunca o que
        ''' estava na tela: sair sem limpar é justamente o defeito.
        ''' </summary>
        Private Sub Restaurar()
            Dim guardado As TrabalhoDaIa = Nothing
            If _chaveAtual IsNot Nothing Then _memoria.TryGetValue(_chaveAtual, guardado)

            Resultado = If(guardado?.Resumo, "")
            Resposta = If(guardado?.Redacao, "")
            Ficha = If(guardado?.Ficha, "")
            RessalvaDoConteudo = If(guardado?.Ressalva, "")
        End Sub

        Public Sub Cancelar()
            _cancelamento?.Cancel()

            ' E A ESPERA DO RESUMO AUTOMATICO. Sao dois cancelamentos porque
            ' sao duas coisas: um voo em andamento, e uma espera que ainda nao
            ' virou voo. Cancelar so o primeiro deixaria o botao Cancelar sem
            ' efeito visivel e um pedido saindo logo depois.
            _esperaDoResumo?.Cancel()
        End Sub

        Private _descartado As Boolean

        ''' <summary>
        ''' <b>A janela fechou.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' Este ViewModel não era descartado por ninguém — nem
        ''' <c>MainViewModel.Dispose</c> o cancelava. Uma transmissão em voo
        ''' podia terminar <b>depois do fechamento</b>, publicar
        ''' <c>Resultado</c> e <c>Ficha</c>, mexer em <c>Ocupado</c> e deixar o
        ''' <c>_pulso</c> batendo.
        '''
        ''' É a mesma família das continuações do acervo e do leitor, e é a mais
        ''' desconfortável delas: o que está em voo aqui é <b>egress</b>. Cancelar
        ''' não desfaz uma requisição já enviada — o diário é quem sabe disso, e
        ''' ele continua registrando o desfecho pela reconciliação da próxima
        ''' abertura. O que muda é que a tela para de ser escrita.
        '''
        ''' A geração sobe primeiro: um voo que volte depois já é de outra
        ''' geração e cai na guarda que sempre existiu.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            If _descartado Then Return
            _descartado = True

            _geracao += 1
            Try
                _cancelamento?.Cancel()
            Catch
            End Try

            ' A ESPERA TAMBEM. Fechar a janela com uma espera pendente
            ' deixaria um Delay vivo por ate 800 ms e um pedido saindo
            ' DEPOIS de a tela ter sumido -- conteudo indo para fora sem
            ' ninguem para ver o resultado.
            Try
                _esperaDoResumo?.Cancel()
                _esperaDoResumo?.Dispose()
            Catch
            End Try
            ' E O CAMPO VAI JUNTO. Deixa-lo apontando para um CTS descartado
            ' fazia o Cancelar() seguinte -- que nao engole excecao como este
            ' Dispose -- lancar ObjectDisposedException do nada, num caminho
            ' que o usuario alcanca clicando em Cancelar depois de fechar.
            _esperaDoResumo = Nothing
            If _pulso IsNot Nothing Then _pulso.Stop()
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
        ''' <summary>
        ''' <b>A instrução do resumo — e o histórico que ela precisou pedir.</b>
        '''
        ''' Era só "Resuma estas mensagens." O corpo que o Iris manda é o
        ''' <c>mail.Body</c> inteiro, sem truncar, e num "RE:" ele já traz a
        ''' conversa citada embaixo — o modelo <b>recebia</b> o histórico e
        ''' ninguém pedia que ele o usasse. O resumo saía sobre a mensagem nova
        ''' e nada mais.
        '''
        ''' <b>E a instrução pede que ele SEPARE.</b> Um resumo que misture o
        ''' que acabou de chegar com o que já estava na conversa faz o leitor
        ''' achar que tudo é novidade — e para quem lê uma caixa de entrada,
        ''' "o que mudou" é a única pergunta que importa.
        ''' </summary>
        Friend Const InstrucaoDeResumo As String =
            "Comece por UMA LINHA que resuma tudo -- uma frase, sem rótulo, " &
            "sem 'Resumo:' na frente -- e depois deixe uma linha em branco " &
            "antes do resto. " &
            "Resuma estas mensagens. O corpo pode trazer a conversa anterior " &
            "citada abaixo da mensagem mais recente: percorra a conversa INTEIRA, " &
            "de baixo para cima, e diga primeiro o que há de NOVO, depois o " &
            "histórico necessário para entender. " &
            "PRESERVE OS DADOS CONCRETOS: códigos, números, quantidades, datas, " &
            "nomes de itens e valores devem aparecer como no original, e não " &
            "resumidos em alguns códigos ou alguns itens — num e-mail de " &
            "trabalho são eles o conteúdo. Se um item foi corrigido depois, diga " &
            "o valor antigo e o novo."

        ''' <summary>
        ''' A instrução da resposta. Pelo mesmo motivo: responder sem olhar o
        ''' que já foi dito produz resposta que repete ou contradiz a conversa.
        ''' </summary>
        Friend Const InstrucaoDeResposta As String =
            "Redija uma resposta. Leve em conta a conversa citada no corpo, se " &
            "houver, para não repetir o que já foi dito nem contradizer o que " &
            "já foi combinado."

        Public Async Function Resumir() As Task
            Await Executar(AssistOperation.Resumir, InstrucaoDeResumo)
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
            ''' O RESUMO VAI JUNTO. Ele não amplia o que sai da caixa -- o
            ''' e-mail já saiu na primeira chamada --, mas dá ao segundo pedido
            ''' o que o primeiro concluiu, em vez de refazer o trabalho.
            Await Executar(AssistOperation.Redigir,
                           InstrucaoDeResposta & Environment.NewLine &
                           "Resumo já produzido desta conversa: " & Resultado)
        End Function

        ''' <summary>
        ''' Devolve o rascunho como estava antes da redação — <b>se ainda for o
        ''' mesmo rascunho, e se ele ainda estiver como a IA o deixou</b>.
        '''
        ''' O desfazer guardava só o texto anterior, e isso o tornava um
        ''' escrevedor cego com um alvo móvel: depois de redigir em A, fechar A e
        ''' abrir B deixava o botão habilitado, e clicar nele escrevia o texto
        ''' antigo de A dentro de B. A guarda da aplicação protegia a ida e
        ''' deixava a volta aberta.
        '''
        ''' Recusar em silêncio seria a outra forma de errar: o botão está ali,
        ''' o usuário clicou, e nada acontecer não se distingue de estar
        ''' quebrado. Por isso a recusa explica.
        ''' </summary>
        Public Sub Desfazer()
            If _rascunho Is Nothing OrElse _anterior Is Nothing Then Return

            If Not PodeDesfazer Then
                Aviso = "Não dá mais para desfazer esta redação: o rascunho não é " &
                        "mais o mesmo, ou você já mexeu no texto depois dela."
                Esquecer()
                Return
            End If

            _rascunho.Texto = _anterior
            Esquecer()
        End Sub

        ''' <summary>Não há mais o que desfazer.</summary>
        Private Sub Esquecer()
            _anterior = Nothing
            _aplicado = Nothing
            _sessaoDaRedacao = 0
            OnPropertyChanged(NameOf(PodeDesfazer))
            DesfazerCommand.NotifyCanExecuteChanged()
        End Sub

        ''' <summary>O que estava no rascunho antes da última redação.</summary>
        Private _anterior As String
        ''' <summary>O texto que a IA escreveu — o estado que o desfazer espera.</summary>
        Private _aplicado As String
        ''' <summary>A sessão do rascunho em que a redação foi aplicada.</summary>
        Private _sessaoDaRedacao As Long

        ''' <summary>
        ''' <b>Dá para redigir?</b> — e o rascunho precisa aceitar escrita.
        '''
        ''' Havia só <c>_rascunho IsNot Nothing</c>, e em produção o adaptador
        ''' existe sempre: o botão ficava habilitado com o compositor fechado, ou
        ''' durante a confirmação de envio, quando os campos estão travados
        ''' justamente para que ninguém mexa no que o usuário já aprovou.
        ''' </summary>
        ''' <summary>
        ''' <b>Dá para redigir?</b> — e agora o que ela exige é o <b>resumo</b>,
        ''' não o rascunho.
        '''
        ''' Redigir deixou de escrever no rascunho: ele mostra a resposta num
        ''' quadro, e aplicar é outro botão. Então a pergunta mudou de "há onde
        ''' escrever?" para "há o que usar como contexto?".
        '''
        ''' <b>Exigir o resumo é decisão de produto, e não uma limitação
        ''' técnica.</b> A redação sozinha já receberia o e-mail inteiro; o
        ''' resumo acrescenta o que a primeira chamada concluiu. Sem ele, as
        ''' duas operações fariam o mesmo trabalho duas vezes.
        '''
        ''' <b>E por isso ela cede quando resumir não é autorizado.</b> Uma
        ''' ativação que libere só <c>Redigir</c> é legítima, e a exigência
        ''' a tornaria inútil: pedir o resumo primeiro é impossível se
        ''' resumir não passa pelo portão. Requisito de qualidade não pode
        ''' virar bloqueio de configuração — foi um teste que existia antes
        ''' desta mudança que apontou isso.
        ''' </summary>
        Public ReadOnly Property PodeRedigir As Boolean
            Get
                Return Reconciliacao.Terminou AndAlso _portaoRedigir AndAlso
                       Not Ocupado AndAlso
                       (TemResultado OrElse Not _portaoResumir)
            End Get
        End Property

        ''' <summary>
        ''' <b>Dá para mandar a resposta ao rascunho?</b>
        '''
        ''' Aqui sim o rascunho é obrigatório — é o destino. As guardas que
        ''' moravam no fim da redação vieram para cá inteiras: mesmo rascunho,
        ''' aceitando escrita, e o texto dele ainda sendo o que era.
        ''' </summary>
        Public ReadOnly Property PodeEnviarParaRascunho As Boolean
            Get
                ' PodeAbrir entra no OU: com o compositor fechado o botao NAO
                ' recusa -- ele abre a resposta e escreve nela. Exigir que o
                ' usuario abrisse antes era pre-condicao da forma do codigo, e
                ' nao do que ele pediu.
                Return TemResposta AndAlso Not Ocupado AndAlso
                       _rascunho IsNot Nothing AndAlso
                       (_rascunho.PodeEditar OrElse _rascunho.PodeAbrir)
            End Get
        End Property

        ''' <summary>Por que o "Enviar para rascunho" está apagado.</summary>
        Public ReadOnly Property PorQueNaoEnvia As String
            Get
                If PodeEnviarParaRascunho OrElse Not TemResposta Then Return ""
                ' NAO DIZ "o rascunho aberto", porque PodeEditar responde False
                ' para DOIS estados: compositor FECHADO e compositor travado na
                ' confirmacao de envio. Em producao o adaptador existe sempre,
                ' entao o caso comum -- nao ha compositor nenhum -- caia na
                ' frase que afirmava haver um, e mandava procurar defeito num
                ' rascunho que nao existia.
                If _rascunho Is Nothing OrElse
                   (Not _rascunho.PodeEditar AndAlso Not _rascunho.PodeAbrir) Then
                    Return "Não há onde escrever esta resposta: nenhuma mensagem " &
                           "aberta para responder, ou o rascunho aberto está " &
                           "travado na confirmação de envio e não aceita edição " &
                           "até você sair de lá."
                End If
                Return ""
            End Get
        End Property

        Public ReadOnly Property TemMotivoParaNaoEnviar As Boolean
            Get
                Return PorQueNaoEnvia.Length > 0
            End Get
        End Property

        ''' <summary>
        ''' <b>POR QUE o "Redigir resposta" está apagado.</b>
        '''
        ''' Botão cinza sem explicação é pior que botão ausente: ele promete uma
        ''' coisa e não diz o que falta. Esta base já corrigiu isso nas tarefas
        ''' atribuídas, e o usuário teve que perguntar aqui — o que é a prova de
        ''' que o botão não estava se explicando.
        '''
        ''' A ordem das perguntas é a ordem em que elas importam: primeiro o que
        ''' o usuário resolve sozinho, depois o que depende de autorização, por
        ''' último o que é só esperar.
        ''' </summary>
        Public ReadOnly Property PorQueNaoRedige As String
            Get
                If PodeRedigir Then Return ""
                If Not TemResultado AndAlso _portaoResumir Then
                    Return "Resuma primeiro. A redação usa o resumo como " &
                           "contexto, além da própria mensagem."
                End If
                If Not _portaoRedigir Then
                    Return "Redigir não está entre as operações autorizadas nesta ativação."
                End If
                If Not Reconciliacao.Terminou Then
                    Return "O registro de envios à IA não pôde ser conferido, " &
                           "e a IA externa fica desligada enquanto isso."
                End If
                Return ""   ' Ocupado: a barra de progresso ja diz.
            End Get
        End Property

        Public ReadOnly Property TemMotivoParaNaoRedigir As Boolean
            Get
                Return PorQueNaoRedige.Length > 0
            End Get
        End Property

        ''' <summary>
        ''' <b>Dá para desfazer?</b> — quatro condições, e todas são a mesma
        ''' pergunta: <i>o que eu desfaria ainda é o que eu fiz?</i>
        '''
        ''' Há o que desfazer; o rascunho aceita escrita agora; é o <b>mesmo</b>
        ''' rascunho em que a redação foi aplicada; e o texto ainda é o que a IA
        ''' escreveu.
        '''
        ''' A última fecha o desfazer assim que o usuário digita por cima da
        ''' redação — de propósito. Restaurar ali apagaria a edição dele para
        ''' desfazer algo que ele já desfez à mão.
        ''' </summary>
        Public ReadOnly Property PodeDesfazer As Boolean
            Get
                Return _anterior IsNot Nothing AndAlso
                       _rascunho IsNot Nothing AndAlso _rascunho.PodeEditar AndAlso
                       _rascunho.Sessao = _sessaoDaRedacao AndAlso
                       String.Equals(_rascunho.Texto, _aplicado, StringComparison.Ordinal)
            End Get
        End Property

        ''' <summary>
        ''' O caminho comum dos dois comandos: monta o contexto, executa, e
        ''' guarda o resultado no quadro certo -- resumo num, resposta noutro.
        '''
        ''' <b>A redação NÃO escreve mais no rascunho.</b> Isso virou o
        ''' <see cref="EnviarParaRascunho"/>, que é um ato do usuário: o botão
        ''' dizia "redigir" e também aplicava, e ninguém tinha pedido a
        ''' segunda coisa.
        ''' </summary>
        Private Async Function Executar(operacao As AssistOperation,
                                        instrucao As String) As Task
            Dim antesDoPedido = If(_rascunho Is Nothing, Nothing, _rascunho.Texto)
            Dim sessaoDoPedido = If(_rascunho Is Nothing, 0L, _rascunho.Sessao)

            ' A CLASSIFICACAO E CAPTURADA DE PASSAGEM.
            '
            ' O Pedir ja a pede para o portao; envolve-la aqui evita ir ao COM
            ' uma segunda vez so para contar imagem embutida -- e evita que a
            ' contagem venha de uma leitura DIFERENTE da que o portao aprovou,
            ' que e como uma ressalva passa a falar de outra mensagem.
            Dim classificadas As IReadOnlyList(Of MessageClassification) = Nothing
            RessalvaDoConteudo = ""

            ''' O RESUMO NAO PODE SER APAGADO PELA REDACAO.
            '''
            ''' As duas operacoes publicam pelo mesmo caminho, e ele escreve em
            ''' Resultado. Guardar e devolver aqui e feio; a alternativa era
            ''' espalhar a nocao de operacao por dentro do Publicar, que nao a
            ''' tem e nao deveria ter.
            Dim resumoAntes = Resultado

            Await Pedir(_contexto.Pedido(operacao),
                        Function()
                            classificadas = _contexto.Classificar()
                            Return classificadas
                        End Function,
                        Function() _contexto.Montar(operacao, instrucao))

            If operacao = AssistOperation.Redigir Then
                If TemResultado Then
                    Resposta = Resultado
                    Resultado = resumoAntes
                Else
                    ' A redacao falhou: o resumo continua valendo.
                    Resultado = resumoAntes
                End If
                Return
            End If

            If TemResultado Then RessalvaDoConteudo = DizerOQueFicouDeFora(classificadas)
        End Function

        ''' <summary>
        ''' <b>Manda a resposta redigida para o rascunho.</b>
        '''
        ''' Isto acontecia <b>junto</b> com a redação, e era o defeito: o botão
        ''' dizia "redigir" e também aplicava. Agora redigir mostra, e aplicar
        ''' é este ato — um botão, uma coisa.
        '''
        ''' As guardas vieram inteiras, e todas perguntam a mesma coisa: <i>o
        ''' rascunho onde eu vou escrever ainda é o que o usuário está
        ''' olhando?</i> Escrever por cima do que ele digitou é mutação local,
        ''' e mutação local sem volta é a que se descobre tarde.
        '''
        ''' <b>Nada é enviado por e-mail</b>: a redação para no compositor.
        ''' </summary>
        Public Async Function EnviarParaRascunho() As Task
            If Not TemResposta OrElse _rascunho Is Nothing Then Return

            ' O QUE FOI CLICADO FICA GUARDADO ANTES DO PRIMEIRO Await.
            '
            ' Abrir uma resposta e assincrono, e durante ele o usuario pode
            ' trocar de mensagem. Sem estas duas linhas, a sequencia era:
            ' clica em enviar na mensagem A, o compositor de A comeca a abrir,
            ' o usuario clica na mensagem B, o Trocou restaura a resposta de B,
            ' e o codigo abaixo -- que le Resposta AGORA -- escrevia a resposta
            ' de B dentro do rascunho aberto para A.
            '
            ' Achado por revisao externa em 31/08/2026. E a mesma familia da
            ' guarda de geracao do resumo: tudo que atravessa um Await precisa
            ' perguntar, do outro lado, se ainda esta falando da mesma coisa.
            ' UM ENVIO POR VEZ. Com o compositor fechado, dois cliques abrem
            ' duas respostas, e as duas passam pela guarda de geracao -- que
            ' fala da MENSAGEM, e nao da instancia do compositor. O ultimo a
            ' continuar escreve no que estiver aberto, que pode ser o da outra
            ' abertura.
            If _enviando Then Return
            _enviando = True
            Try

            Dim aEnviar = Resposta
            Dim minha = _geracao
            ' NAO HA GUARDA DE SESSAO AQUI, e a ausencia e decidida.
            '
            ' Escrevi uma e ela NAO PODIA FALHAR: no ramo do compositor ja
            ' aberto nao existe Await nenhum entre ler a sessao e escrever,
            ' entao os dois lados da comparacao liam o mesmo valor. Guarda
            ' que nunca dispara e o bloqueio que nunca bloqueia -- pior que
            ' guarda nenhuma, porque parece cobertura.
            '
            ' O que a revisao externa apontou como risco de sessao esta
            ' coberto por outras duas: _enviando impede a segunda abertura,
            ' e EsquecerASessao agora sobe a geracao. E a sessao continua
            ' valendo onde ela e a pergunta certa: no PodeDesfazer.

            ' NAO HA RESPOSTA ABERTA? ABRE UMA.
            '
            ' O botao dizia "enviar para rascunho" e exigia que o usuario
            ' abrisse o rascunho ANTES. Tres rodadas de "ainda nao funciona"
            ' foram gastas consertando a plumbagem em volta de uma
            ' pre-condicao que so existia porque esta porta so sabia
            ' ESCREVER em compositor aberto. Quem pede "manda para um
            ' rascunho" esta pedindo o rascunho tambem.
            '
            ' Abre RESPONDER, e nao responder a todos: acrescentar
            ' destinatarios que ninguem pediu e do tipo de coisa que so se
            ' descobre depois de enviada. Quem quer todos abre "Responder a
            ' todos" antes -- compositor ja aberto e usado como esta.
            If Not _rascunho.PodeEditar Then
                If Not _rascunho.PodeAbrir Then
                    Aviso = "Não há onde escrever esta resposta: nenhuma " &
                            "mensagem aberta para responder, ou o rascunho " &
                            "está travado na confirmação de envio."
                    Return
                End If

                Await _rascunho.AbrirAsync()

                ' E CONFERE DE NOVO. Abrir e assincrono: pode falhar, pode
                ' demorar, e o usuario pode ter mexido no meio. Escrever
                ' assumindo que deu certo poria a resposta em lugar nenhum --
                ' ou, pior, no rascunho errado.
                If Not _rascunho.PodeEditar Then
                    Aviso = "A resposta não abriu. Tente Responder ali em cima."
                    Return
                End If

                ' E A MENSAGEM AINDA E A MESMA? A geracao muda a cada troca, e
                ' o rascunho que acabou de abrir e o da mensagem que estava na
                ' tela quando o botao foi clicado -- nao o da que esta agora.
                If minha <> _geracao Then
                    Aviso = "Você trocou de mensagem enquanto a resposta abria. " &
                            "Nada foi escrito; clique em Enviar para rascunho de novo."
                    Return
                End If
            End If

            ''' O QUE ESTAVA LA FICA GUARDADO, e o Desfazer o devolve. A
            ''' sessao entra junto: fechar este rascunho e abrir outro com o
            ''' mesmo texto -- os dois vazios e o caso comum -- faria o
            ''' desfazer escrever numa mensagem que nao e esta.
            ' O TEXTO CLICADO, e nao o que esta na tela agora. Sao a mesma
            ' coisa quando nada aconteceu no meio -- e quando algo aconteceu,
            ' a guarda de geracao acima ja saiu fora.
            _anterior = _rascunho.Texto
            _sessaoDaRedacao = _rascunho.Sessao
            _rascunho.Texto = aEnviar
            _aplicado = aEnviar

            Aviso = ""
            OnPropertyChanged(NameOf(PodeDesfazer))
            DesfazerCommand.NotifyCanExecuteChanged()

            Finally
                _enviando = False
            End Try
        End Function

        ''' <summary>Há um envio ao rascunho em andamento.</summary>
        Private _enviando As Boolean

        ''' <summary>
        ''' Pede a operação. O resultado só é publicado se a geração ainda for a
        ''' corrente quando ele voltar.
        ''' </summary>
        Public Async Function Pedir(pedido As PreflightRequest,
                                    classificar As Func(Of IReadOnlyList(Of MessageClassification)),
                                    montar As Func(Of EnvelopeResult)) As Task

            ' A EXECUCAO E POR OPERACAO, COMO A HABILITACAO.
            '
            ' Aqui estava `If Not PodePedir Then Return`, e `PodePedir` quer
            ' dizer "pode RESUMIR". Duas consequencias, as duas erradas em
            ' direcoes opostas: com ativacao so para redigir, o botao ficava
            ' habilitado e clicar nele nao fazia nada; e a exigencia de rascunho
            ' editavel vivia so no CanExecute, de modo que uma chamada direta a
            ' Redigir() atravessava e podia TRANSMITIR conteudo sem haver lugar
            ' valido para aplicar a resposta.
            '
            ' Botao desabilitado e conveniencia. A recusa tem de estar aqui.
            If Not PodeExecutar(pedido.Operacao) Then Return

            Dim minha = _geracao
            Ocupado = True
            Aviso = ""

            ' O CRONOMETRO COMECA AQUI, e a ficha velha morre junto.
            '
            ' Deixar a ficha do pedido anterior na tela durante o proximo
            ' mostraria custo e tempo de OUTRA chamada ao lado de um cronometro
            ' correndo -- dois numeros de coisas diferentes, com cara de serem
            ' do mesmo pedido.
            _inicioDoVoo = _relogio()
            _duracaoDoVoo = Nothing
            Ficha = ""
            OnPropertyChanged(NameOf(Decorrido))
            If _pulso IsNot Nothing Then _pulso.Start()

            Dim cts As New CancellationTokenSource()
            _cancelamento = cts

            Try
                Dim r = Await Task.Run(Function() _transmissor.Executar(
                    pedido, classificar, montar, cts.Token))

                ' A GERACAO. Se o usuario trocou de mensagem enquanto isto
                ' rodava, o resultado e de outro contexto — e mostrar seria pior
                ' que nao mostrar: um resumo errado com cara de certo.
                If _descartado OrElse minha <> _geracao Then Return

                Publicar(r)
                Ficha = Fichar(pedido.Destino, r)
            Finally
                ' QUEM LIMPA O `Ocupado` E O DONO DO VOO, NAO A GERACAO.
                '
                ' Aqui estava `If minha = _geracao Then Ocupado = False`: trocar
                ' de mensagem durante uma operacao incrementava a geracao, a
                ' condicao ficava falsa, e `Ocupado` NUNCA voltava para False —
                ' o assistente ficava travado para sempre, com todos os botoes
                ' desabilitados e sem nada na tela explicando por que.
                '
                ' A geracao decide se o RESULTADO vale. Quem decide se o estado
                ' de "ocupado" e meu para limpar e ser eu o voo corrente, e isso
                ' e a identidade do CancellationTokenSource.
                ' QUEM PARA O CRONOMETRO E O DONO DO VOO, pelo mesmo motivo
                ' que o Ocupado: parar por geracao deixaria o pulso batendo
                ' para sempre depois de uma troca de mensagem.
                Dim meu = ReferenceEquals(_cancelamento, cts)
                cts.Dispose()
                If meu Then
                    _cancelamento = Nothing
                    ' Ocupado dispara notificacao e reconsulta de comando. Num
                    ' ViewModel descartado isso e mexer em tela que ja saiu.
                    If Not _descartado Then Ocupado = False
                    If _pulso IsNot Nothing Then _pulso.Stop()
                    ' CONGELA. Parar o pulso so faz a tela deixar de
                    ' reperguntar; sem fixar a duracao, Decorrido continuaria
                    ' contando para sempre.
                    If _inicioDoVoo.HasValue Then
                        _duracaoDoVoo = _relogio() - _inicioDoVoo.Value
                    End If
                    ' A DURACAO fica gravada de qualquer jeito -- e estado
                    ' interno, e congela-la e o certo. O que nao pode e AVISAR:
                    ' notificacao de propriedade num ViewModel descartado
                    ' contradiz a intencao do Dispose logo acima.
                    If Not _descartado Then OnPropertyChanged(NameOf(Decorrido))
                End If
            End Try
        End Function

        ''' <summary>
        ''' A condição de execução de <b>cada</b> operação.
        '''
        ''' Operação que não está na lista — <c>Nenhuma</c>, ou uma que venha a
        ''' existir — recusa. Fechado por padrão, como o resto da §29.
        ''' </summary>
        Private Function PodeExecutar(operacao As AssistOperation) As Boolean
            Select Case operacao
                Case AssistOperation.Resumir
                    Return PodePedir
                Case AssistOperation.Redigir
                    Return PodeRedigir
                Case Else
                    Return False
            End Select
        End Function

        Private Sub Publicar(r As AssistOutcome)
            Select Case r.Kind
                Case AssistOutcomeKind.Respondeu
                    ' LIMPO, e nao renderizado. Ver TextoDoModelo: aqui so se
                    ' APAGA marcador. Interpretar Markdown transformaria texto
                    ' de terceiro em arvore visual, e o proximo passo obvio --
                    ' "ja que fazemos negrito, faz link tambem" -- daria ao
                    ' e-mail um jeito de fazer o Iris buscar coisa na rede.
                    Resultado = TextoDoModelo.Limpar(r.Texto)
                    ' RESPOSTA VAZIA TEM DE APARECER COMO ALGUMA COISA.
                    '
                    ' "Respondeu" com texto vazio fechava o diario como sucesso
                    ' e nao deixava nada na tela: nem resultado, nem aviso. A
                    ' operacao simplesmente sumia, e o usuario nao teria como
                    ' distinguir "o provedor nao tinha o que dizer" de "o botao
                    ' nao funcionou".
                    '
                    ' Nao e ambiguo: o conteudo SAIU e a resposta CHEGOU. O
                    ' diario fecha como concluida, e a frase diz exatamente
                    ' isso.
                    Aviso = If(TemResultado, "",
                               "O provedor respondeu sem texto. O conteúdo saiu " &
                               "desta máquina e a operação foi concluída — só não " &
                               "veio resposta.")
                Case AssistOutcomeKind.Negado
                    Resultado = ""
                    Aviso = "A IA não foi usada: " & EmPortugues(r.MotivoDoPortao)
                Case AssistOutcomeKind.Ambiguo,
                     AssistOutcomeKind.AmbiguoSemFechamentoDoDiario
                    Resultado = ""
                    ' Sem asterisco: num TextBlock ele aparece literalmente, e
                    ' "**nao da para saber**" na tela e desleixo visivel.
                    '
                    ' O CODIGO HTTP VAI JUNTO quando houve um. So o numero:
                    ' corpo de erro de provedor ECOA o que foi enviado, e a
                    ' faixa da IA nao e lugar para o e-mail do usuario voltar.
                    ' Mas sem numero nenhum a frase nao diz o que fazer a
                    ' seguir -- 401 manda recadastrar a chave, 404 manda rever
                    ' a restricao de provedor -- e da primeira vez descobrir
                    ' isso custou tres ferramentas de linha de comando.
                    Aviso = "A operação não terminou, e não dá para saber se o " &
                            "conteúdo chegou ao provedor. Isso ficou registrado." &
                            If(r.CodigoHttp.HasValue,
                               $" O provedor respondeu HTTP {r.CodigoHttp.Value}.", "")
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
                    ' TRES MOTIVOS NUMA FRASE SO, e um deles nao era esse.
                    '
                    ' LeituraNaoAceita quer dizer que a classificacao FOI FEITA e
                    ' o resultado nao esta entre os autorizados -- o oposto de
                    ' "nao foi possivel classificar". Medido em 30/08: 13 de 13
                    ' mensagens de uma pasta real classificavam como Blank, a
                    ' ativacao autorizava so Absent, e a tela dizia que nao tinha
                    ' conseguido classificar. A frase mandou procurar defeito no
                    ' lugar errado.
                Case DisclosureReason.LeituraNaoAceita
                    Return "a classificação de alguma mensagem não está entre as " &
                           "autorizadas nesta ativação. Ela foi lida; o que falta é " &
                           "a autorização cobrir esse resultado."
                Case DisclosureReason.LeituraEstruturalmenteInsegura,
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
        ''' <summary>
        ''' Quantas divulgações <b>estão</b> ambíguas — não quantas viraram
        ''' ambíguas nesta abertura.
        '''
        ''' A diferença apagava o aviso: contar a transição fazia a segunda
        ''' abertura depois de uma queda devolver zero, com as ambíguas
        ''' gravadas no banco e o egresso religado em silêncio.
        ''' </summary>
        Public ReadOnly Property Ambiguas As Integer

        Private Sub New(terminou As Boolean, ambiguas As Integer,
                        Optional reconhecidas As Integer = 0)
            Me.Terminou = terminou
            Me.Ambiguas = ambiguas
            Me.Reconhecidas = reconhecidas
        End Sub

        ''' <summary>
        ''' O mesmo resultado, com o reconhecimento do dono aplicado.
        '''
        ''' Devolve um objeto novo em vez de mudar este: o resultado da
        ''' reconciliação é o que o banco disse, e o reconhecimento é o que a
        ''' pessoa disse. Misturar os dois num só objeto mutável faria a segunda
        ''' leitura não saber mais qual era qual.
        ''' </summary>
        Public Function ComReconhecimento(reconhecidas As Integer) As ReconciliationResult
            Return New ReconciliationResult(Terminou, Ambiguas, reconhecidas)
        End Function

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

        ''' <summary>
        ''' <b>Quantas ambíguas o dono já reconheceu ter visto.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE ISTO EXISTE</b>
        '''
        ''' O aviso dizia, com razão, "este aviso não desaparece sozinho" — e não
        ''' desaparecia mesmo, porque <b>não havia como dizer que se viu</b>. Isso
        ''' estava anotado no ESCOPO como dívida desde antes: a divulgação
        ''' ambígua ficava na tela para sempre, e um parágrafo permanente é um
        ''' parágrafo que se aprende a não ler.
        '''
        ''' Esconder atrás de um ícone <b>sem</b> reconhecimento seria pior: uma
        ''' divulgação não reconhecida deixaria de ser dita. O reconhecimento é o
        ''' que torna o ícone honesto.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>E RECONHECER O QUE JÁ HÁ NÃO SILENCIA O QUE VIER</b>
        '''
        ''' Guarda-se o <b>número</b> reconhecido, e não um "já vi" booleano. Se
        ''' aparecerem envios ambíguos novos, <c>Ambiguas</c> passa o
        ''' reconhecido e o texto volta inteiro.
        '''
        ''' Um booleano teria calado a próxima ambiguidade com o clique da
        ''' anterior — que é a forma mais silenciosa possível de perder uma
        ''' divulgação.
        ''' </summary>
        Public ReadOnly Property Reconhecidas As Integer

        ''' <summary>
        ''' Há ambiguidade que o dono <b>ainda não</b> reconheceu?
        '''
        ''' É isto que decide entre o parágrafo e o ícone — e não o fato de
        ''' haver ambiguidade, que continua verdadeiro depois de reconhecida.
        ''' </summary>
        Public ReadOnly Property TemNovidade As Boolean
            Get
                Return Ambiguas > Reconhecidas
            End Get
        End Property

        ''' <summary>
        ''' O ícone, quando tudo o que há já foi reconhecido. <b>Não some</b>: a
        ''' ambiguidade continua sendo verdade, e o que mudou foi só o tamanho
        ''' com que ela é dita.
        ''' </summary>
        Public ReadOnly Property Marca As String
            Get
                ' O "Not Terminou" E REDUNDANTE HOJE, e esta escrito porque
                ' redundancia que ninguem sabe que e redundante vira armadilha.
                '
                ' Os dois caminhos que deixam Terminou falso -- NaoRodou e o
                ' Catch do Rodar -- carregam Ambiguas = 0, entao a segunda
                ' condicao ja os pega. Foi medido desfazendo-o: o teste
                ' Nao_ter_conferido_nao_vira_marca continuou passando.
                '
                ' Ele fica porque a pergunta que importa aqui e "eu conferi?",
                ' e amarra-la ao zero e depender de uma coincidencia dos
                ' construtores. Se um dia Terminou=False vier com contagem, a
                ' marca nao pode aparecer: "nao sei" nao se reconhece.
                If Not Terminou OrElse Ambiguas = 0 OrElse TemNovidade Then Return ""
                Return If(Ambiguas = 1,
                          "⚠ 1 envio à IA sem desfecho conhecido (você já viu este aviso)",
                          $"⚠ {Ambiguas} envios à IA sem desfecho conhecido (você já viu este aviso)")
            End Get
        End Property

        Public ReadOnly Property TemMarca As Boolean
            Get
                Return Marca.Length > 0
            End Get
        End Property

        ''' <summary>O que a tela diz sobre isso. Vazio quando não há o que dizer.</summary>
        Public ReadOnly Property Aviso As String
            Get
                If Not Terminou Then
                    Return "Não foi possível conferir o registro de envios à IA. " &
                           "Enquanto isso, a IA externa fica desligada."
                End If
                If Ambiguas = 0 Then Return ""

                ' JA RECONHECIDO VIRA MARCA, e a marca mora noutra propriedade.
                ' O texto inteiro volta sozinho se aparecer ambiguidade nova.
                If Not TemNovidade Then Return ""

                If Ambiguas = 1 Then
                    Return "Um envio à IA ficou sem desfecho conhecido numa execução " &
                           "anterior. Pode ter saído conteúdo, e não dá para saber. " &
                           "Este aviso não desaparece sozinho."
                End If
                Return $"{Ambiguas} envios à IA ficaram sem desfecho conhecido em " &
                       "execuções anteriores. Pode ter saído conteúdo, e não dá para " &
                       "saber. Este aviso não desaparece sozinho."
            End Get
        End Property

    End Class

End Namespace

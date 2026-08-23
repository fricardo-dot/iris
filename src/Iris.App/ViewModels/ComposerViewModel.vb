Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' Em que passo da composição o usuário está.
    '''
    ''' É um enum e não três Booleans porque os estados são mutuamente
    ''' exclusivos: com Booleans separados existiria a combinação
    ''' "confirmando e enviando ao mesmo tempo", que não quer dizer nada e
    ''' mais cedo ou mais tarde apareceria na tela.
    ''' </summary>
    Public Enum ComposerState
        Closed
        Editing
        ''' <summary>Revisando para quem vai, antes de enviar.</summary>
        ConfirmingSend
        Sending
        ''' <summary>Fechando com alteração não salva.</summary>
        ConfirmingClose
        ''' <summary>
        ''' O envio falhou SEM se saber se saiu. Estado terminal: este
        ''' rascunho não envia mais.
        ''' </summary>
        SendUnknown
    End Enum

    ''' <summary>
    ''' O compositor.
    '''
    ''' É a única parte do Iris que escreve no mundo real, e por isso é a
    ''' que mais precisa de disciplina:
    '''
    '''   • O rascunho é do OUTLOOK. Responder, responder a todos e
    '''     encaminhar usam Reply/ReplyAll/Forward do Object Model. Remontar
    '''     destinatários à mão significaria errar em quem recebe.
    '''   • O corpo citado e a assinatura NUNCA são reescritos. O usuário
    '''     edita só o texto dele, que entra acima da citação.
    '''   • A chave do rascunho é relida a cada Save, porque o EntryID muda.
    '''     Guardar a chave antiga daria NotFound no envio.
    '''   • Salvar tem debounce e no máximo uma gravação em voo. Uma fila de
    '''     Save ocuparia a fila única da STA sem produzir nada útil.
    '''   • Enviar exige confirmação, acontece uma vez e não tem retry.
    '''   • Envio ambíguo é terminal. Reenviar no escuro é o único erro
    '''     irreversível deste projeto.
    ''' </summary>
    Public NotInheritable Class ComposerViewModel
        Inherits ObservableObject
        Implements IDisposable

        ''' <summary>
        ''' 1,5 s. Curto o bastante para o rascunho sobreviver a um
        ''' fechamento acidental, longo o bastante para não gravar a cada
        ''' palavra digitada.
        ''' </summary>
        Private Const AutosaveMs As Integer = 1500

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _observe As Action(Of Task, String)
        Private ReadOnly _pickFile As IPickFileService
        Private ReadOnly _autosave As DispatcherTimer

        Private _key As DraftKey
        Private _saveTask As Task
        Private _savePendente As Boolean

        ''' <summary>
        ''' Conta as edições do usuário. Depois de gravar, comparar este
        ''' número com o que valia quando a gravação COMEÇOU diz, sem
        ''' adivinhação, se ainda há coisa por salvar. Um Boolean simples
        ''' erraria: uma tecla digitada durante a gravação seria apagada
        ''' pelo sucesso da gravação anterior.
        ''' </summary>
        Private _edicoes As Long

        Private _carregando As Boolean
        Private _disposed As Boolean

        ''' <param name="autosaveMs">
        ''' Só os testes passam outro valor. Esperar 1,5 s de relógio real
        ''' em cada teste de debounce tornaria a suíte lenta e, pior,
        ''' intermitente — e um teste intermitente acaba sendo ignorado, que
        ''' é o mesmo que não ter teste.
        ''' </param>
        Public Sub New(broker As IOutlookBroker, ui As Dispatcher,
                       observe As Action(Of Task, String), pickFile As IPickFileService,
                       Optional autosaveMs As Integer = AutosaveMs)
            _broker = broker
            _observe = observe
            _pickFile = pickFile

            _autosave = New DispatcherTimer(DispatcherPriority.Background, ui) With {
                .Interval = TimeSpan.FromMilliseconds(autosaveMs)
            }
            AddHandler _autosave.Tick, AddressOf OnAutosaveTick

            AttachCommand = New AsyncRelayCommand(AddressOf AnexarAsync, Function() PodeEditar)
            RequestSendCommand = New AsyncRelayCommand(AddressOf PedirEnvioAsync, Function() PodeEditar)
            ConfirmSendCommand = New AsyncRelayCommand(AddressOf ConfirmarEnvioAsync,
                                                       Function() State = ComposerState.ConfirmingSend)
            CancelSendCommand = New RelayCommand(Sub() State = ComposerState.Editing,
                                                 Function() State = ComposerState.ConfirmingSend)
            CloseCommand = New RelayCommand(AddressOf Fechar)
            SaveAndCloseCommand = New AsyncRelayCommand(AddressOf SalvarEFecharAsync)
            DiscardCommand = New AsyncRelayCommand(AddressOf DescartarAsync)
            KeepEditingCommand = New RelayCommand(Sub() State = ComposerState.Editing)
        End Sub

        Public ReadOnly Property AttachCommand As IAsyncRelayCommand
        Public ReadOnly Property RequestSendCommand As IAsyncRelayCommand
        Public ReadOnly Property ConfirmSendCommand As IAsyncRelayCommand
        Public ReadOnly Property CancelSendCommand As IRelayCommand
        Public ReadOnly Property CloseCommand As IRelayCommand
        Public ReadOnly Property SaveAndCloseCommand As IAsyncRelayCommand
        Public ReadOnly Property DiscardCommand As IAsyncRelayCommand
        Public ReadOnly Property KeepEditingCommand As IRelayCommand

        Public ReadOnly Property Attachments As New ObservableCollection(Of AttachmentInfo)()

        Private _state As ComposerState = ComposerState.Closed

        Public Property State As ComposerState
            Get
                Return _state
            End Get
            Private Set(value As ComposerState)
                If SetProperty(_state, value) Then
                    OnPropertyChanged(NameOf(IsOpen))
                    OnPropertyChanged(NameOf(PodeEditar))
                    OnPropertyChanged(NameOf(IsConfirmingSend))
                    OnPropertyChanged(NameOf(IsConfirmingClose))
                    OnPropertyChanged(NameOf(IsSending))
                    OnPropertyChanged(NameOf(IsSendUnknown))
                    OnPropertyChanged(NameOf(ShowOverlay))
                    NotificarComandos()
                End If
            End Set
        End Property

        Public ReadOnly Property IsOpen As Boolean
            Get
                Return _state <> ComposerState.Closed
            End Get
        End Property

        ''' <summary>
        ''' Editar só vale no estado de edição. Durante a confirmação e o
        ''' envio os campos ficam travados: mudar o destinatário depois de o
        ''' usuário ter aprovado a lista tornaria a confirmação uma mentira.
        ''' </summary>
        Public ReadOnly Property PodeEditar As Boolean
            Get
                Return _state = ComposerState.Editing
            End Get
        End Property

        Public ReadOnly Property IsConfirmingSend As Boolean
            Get
                Return _state = ComposerState.ConfirmingSend
            End Get
        End Property

        Public ReadOnly Property IsConfirmingClose As Boolean
            Get
                Return _state = ComposerState.ConfirmingClose
            End Get
        End Property

        Public ReadOnly Property IsSending As Boolean
            Get
                Return _state = ComposerState.Sending
            End Get
        End Property

        Public ReadOnly Property IsSendUnknown As Boolean
            Get
                Return _state = ComposerState.SendUnknown
            End Get
        End Property

        ''' <summary>Alguma pergunta cobrindo o compositor.</summary>
        Public ReadOnly Property ShowOverlay As Boolean
            Get
                Return _state = ComposerState.ConfirmingSend OrElse
                       _state = ComposerState.ConfirmingClose OrElse
                       _state = ComposerState.Sending
            End Get
        End Property

        Private _title As String = "Nova mensagem"

        Public Property Title As String
            Get
                Return _title
            End Get
            Private Set(value As String)
                SetProperty(_title, value)
            End Set
        End Property

        ' --- Campos editáveis -------------------------------------------
        '
        ' Estes quatro são os ÚNICOS que o usuário digita, e o broker nunca
        ' escreve de volta neles depois de gravar. Escrever de volta pularia
        ' o cursor de lugar e desfaria o que foi digitado durante a
        ' gravação.

        Private _subject As String = ""
        Private _toLine As String = ""
        Private _ccLine As String = ""
        Private _userText As String = ""

        Public Property Subject As String
            Get
                Return _subject
            End Get
            Set(value As String)
                If SetProperty(_subject, If(value, "")) Then Editou()
            End Set
        End Property

        Public Property ToLine As String
            Get
                Return _toLine
            End Get
            Set(value As String)
                If SetProperty(_toLine, If(value, "")) Then Editou()
            End Set
        End Property

        Public Property CcLine As String
            Get
                Return _ccLine
            End Get
            Set(value As String)
                If SetProperty(_ccLine, If(value, "")) Then Editou()
            End Set
        End Property

        Public Property UserText As String
            Get
                Return _userText
            End Get
            Set(value As String)
                If SetProperty(_userText, If(value, "")) Then Editou()
            End Set
        End Property

        ' --- Somente exibição -------------------------------------------

        Private _quotedPreview As String = ""

        ''' <summary>
        ''' A citação e a assinatura que o Outlook gerou, em texto, só para
        ''' o usuário ver o que vai junto. Não é editável de propósito: o
        ''' Iris escreve ACIMA dela e nunca a reescreve.
        ''' </summary>
        Public Property QuotedPreview As String
            Get
                Return _quotedPreview
            End Get
            Private Set(value As String)
                If SetProperty(_quotedPreview, If(value, "")) Then
                    OnPropertyChanged(NameOf(HasQuoted))
                End If
            End Set
        End Property

        Public ReadOnly Property HasQuoted As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(_quotedPreview)
            End Get
        End Property

        Private _status As String = ""

        Public Property Status As String
            Get
                Return _status
            End Get
            Private Set(value As String)
                If SetProperty(_status, If(value, "")) Then
                    OnPropertyChanged(NameOf(HasStatus))
                End If
            End Set
        End Property

        Public ReadOnly Property HasStatus As Boolean
            Get
                Return Not String.IsNullOrEmpty(_status)
            End Get
        End Property

        Private _preview As SendPreview

        Public Property Preview As SendPreview
            Get
                Return _preview
            End Get
            Private Set(value As SendPreview)
                If SetProperty(_preview, value) Then
                    OnPropertyChanged(NameOf(PreviewRecipients))
                    OnPropertyChanged(NameOf(PreviewAccount))
                End If
            End Set
        End Property

        Public ReadOnly Property PreviewRecipients As IEnumerable(Of RecipientInfo)
            Get
                If _preview Is Nothing Then Return Array.Empty(Of RecipientInfo)()
                Return _preview.Recipients
            End Get
        End Property

        Public ReadOnly Property PreviewAccount As String
            Get
                Return If(_preview Is Nothing, "", _preview.SendingAccount)
            End Get
        End Property

        Private _isDirty As Boolean

        ''' <summary>Há edição do usuário ainda não gravada no store.</summary>
        Public Property IsDirty As Boolean
            Get
                Return _isDirty
            End Get
            Private Set(value As Boolean)
                If SetProperty(_isDirty, value) Then
                    OnPropertyChanged(NameOf(SaveHint))
                End If
            End Set
        End Property

        Public ReadOnly Property SaveHint As String
            Get
                Return If(_isDirty, "Alterações não salvas", "Rascunho salvo")
            End Get
        End Property

        ' ================================================================
        ' Abertura
        ' ================================================================

        Public Function NewMessageAsync() As Task
            Return AbrirAsync("Nova mensagem",
                              Function(ct) _broker.CreateDraftAsync(New DraftContent(), ct))
        End Function

        Public Function ReplyAsync(item As ItemKey, replyAll As Boolean) As Task
            Return AbrirAsync(If(replyAll, "Responder a todos", "Responder"),
                              Function(ct) _broker.CreateReplyDraftAsync(item, replyAll, ct))
        End Function

        Public Function ForwardAsync(item As ItemKey) As Task
            Return AbrirAsync("Encaminhar",
                              Function(ct) _broker.CreateForwardDraftAsync(item, ct))
        End Function

        ''' <summary>
        ''' Cria o rascunho e só então abre a janela. O rascunho existe no
        ''' store ANTES de o usuário digitar qualquer coisa — é isso que faz
        ''' um fechamento acidental não custar trabalho, e é isso que dá uma
        ''' chave estável para o autosave usar.
        '''
        ''' Se a criação falhar, o compositor não abre. Abrir um compositor
        ''' sem rascunho por trás produziria uma tela onde tudo parece
        ''' funcionar e nada é gravado.
        ''' </summary>
        Private Async Function AbrirAsync(titulo As String,
                                          criar As Func(Of CancellationToken, Task(Of OperationResult(Of DraftInfo)))) _
            As Task

            If IsOpen Then
                Status = "Feche a mensagem aberta antes de começar outra."
                Return
            End If

            Title = titulo
            Status = "Criando rascunho…"

            Dim resultado = Await criar(CancellationToken.None)
            If Not resultado.Succeeded Then
                State = ComposerState.Closed
                Status = "Não foi possível criar o rascunho. " & Traduzir(resultado.Kind)
                Return
            End If

            Aplicar(resultado.Value, primeiraVez:=True)
            Status = ""
            IsDirty = False
            State = ComposerState.Editing
        End Function

        ''' <summary>
        ''' Copia do rascunho para a tela.
        '''
        ''' Na primeira vez copia tudo, inclusive os campos editáveis — é o
        ''' que traz destinatários e assunto que o Outlook preencheu ao
        ''' responder. Depois disso copia SÓ o que o usuário não digita:
        ''' chave, citação, anexos e conta. Sobrescrever o que ele está
        ''' digitando seria perder texto.
        ''' </summary>
        Private Sub Aplicar(info As DraftInfo, primeiraVez As Boolean)
            _key = info.Key

            If primeiraVez Then
                _carregando = True
                Try
                    Subject = info.Subject
                    ToLine = info.ToLine
                    CcLine = info.CcLine
                    UserText = info.UserText
                Finally
                    _carregando = False
                End Try
            End If

            QuotedPreview = info.QuotedBody
            SincronizarAnexos(info.Attachments)
        End Sub

        ''' <summary>
        ''' Só mexe na coleção se ela realmente mudou.
        '''
        ''' Um Clear seguido de Add repovoa a lista a cada autosave, ou seja,
        ''' a cada 1,5 s enquanto o usuário digita — o ItemsControl
        ''' reconstruiria os itens e a lista piscaria sem nada ter mudado.
        ''' </summary>
        Private Sub SincronizarAnexos(itens As List(Of AttachmentInfo))
            Dim novos As New List(Of AttachmentInfo)()
            If itens IsNot Nothing Then
                For Each a In itens
                    ' Anexo embutido é parte do corpo HTML, não arquivo que
                    ' o usuário adicionou. Listá-lo faria a citação de uma
                    ' resposta parecer que veio com anexos.
                    If Not a.IsInline Then novos.Add(a)
                Next
            End If

            If novos.Count = Attachments.Count Then
                Dim iguais = True
                For i = 0 To novos.Count - 1
                    If Not Equals(novos(i).Key, Attachments(i).Key) Then
                        iguais = False
                        Exit For
                    End If
                Next
                If iguais Then Return
            End If

            Attachments.Clear()
            For Each a In novos
                Attachments.Add(a)
            Next
        End Sub

        ' ================================================================
        ' Autosave
        ' ================================================================

        Private Sub Editou()
            ' Carregar o rascunho na tela dispara os mesmos setters que
            ' digitar. Sem esta trava, abrir uma resposta já nasceria
            ' "suja" e gravaria de volta o que acabou de ler.
            If _carregando Then Return
            If Not PodeEditar Then Return

            Interlocked.Increment(_edicoes)
            IsDirty = True

            ' Reiniciar: o debounce conta a partir da ÚLTIMA tecla, não da
            ' primeira. Sem o Stop, o timer do WPF continua o ciclo antigo e
            ' a gravação sai no meio da digitação.
            _autosave.Stop()
            _autosave.Start()
        End Sub

        Private Sub OnAutosaveTick(sender As Object, e As EventArgs)
            _autosave.Stop()
            _observe(SalvarAsync(), "composer.autosave")
        End Sub

        ''' <summary>
        ''' Uma gravação em voo por vez, com no máximo um estado pendente.
        ''' Enfileirar cada edição produziria versões intermediárias que
        ''' ninguém quer e ocuparia a fila única da STA, que é a mesma que
        ''' serve a lista e a leitura.
        ''' </summary>
        Private Function SalvarAsync() As Task
            If _saveTask IsNot Nothing AndAlso Not _saveTask.IsCompleted Then
                _savePendente = True
                Return _saveTask
            End If

            _saveTask = GravarAteEstabilizarAsync()
            Return _saveTask
        End Function

        Private Async Function GravarAteEstabilizarAsync() As Task
            Do
                _savePendente = False

                ' Fotografa o contador ANTES da chamada. O que for digitado
                ' durante a gravação tem número maior e continua sujo.
                Dim marca = Interlocked.Read(_edicoes)

                Dim conteudo As New DraftContent With {
                    .Subject = Subject,
                    .UserText = UserText,
                    .ToLine = ToLine,
                    .CcLine = CcLine
                }

                Dim resultado = Await _broker.UpdateDraftAsync(_key, conteudo, CancellationToken.None)

                If resultado.Succeeded Then
                    ' A chave muda a cada Save. Reler é obrigatório: guardar
                    ' a antiga daria NotFound na hora de enviar.
                    Aplicar(resultado.Value, primeiraVez:=False)
                    If Interlocked.Read(_edicoes) = marca Then IsDirty = False
                    If Status.StartsWith(FalhaAoSalvar, StringComparison.Ordinal) Then Status = ""
                Else
                    ' Não fecha e não descarta: o texto continua na tela. A
                    ' próxima tecla dispara outra tentativa.
                    Status = FalhaAoSalvar & Traduzir(resultado.Kind)
                End If
            Loop While _savePendente
        End Function

        Private Const FalhaAoSalvar As String = "Não foi possível salvar agora. "

        ''' <summary>
        ''' Quantas rodadas de gravação a descarga tenta antes de desistir.
        ''' Existe porque o usuário pode continuar digitando durante a
        ''' descarga, e sem teto isso seria um laço que só termina quando
        ''' ele parar de digitar.
        ''' </summary>
        Private Const MaxRodadasDeDescarga As Integer = 3

        ''' <summary>
        ''' Garante que o store tem exatamente o que está na tela.
        '''
        ''' Chamado antes de conferir o envio e antes de fechar salvando.
        ''' Sem isto, a confirmação mostraria destinatários de uma versão
        ''' antiga e o Outlook enviaria essa versão antiga — o usuário
        ''' aprovaria uma coisa e sairia outra.
        '''
        ''' Precisa CONVERGIR, e não gravar uma vez. Gravar leva tempo, e o
        ''' que for digitado nesse meio continua sujo — uma rodada só
        ''' devolveria "pronto" com a tela ainda na frente do store.
        ''' </summary>
        ''' <returns>True se o store bate com a tela.</returns>
        Private Async Function DescarregarAsync() As Task(Of Boolean)
            _autosave.Stop()

            If _saveTask IsNot Nothing AndAlso Not _saveTask.IsCompleted Then
                Await _saveTask
            End If

            Dim rodadas = 0
            Do While IsDirty
                rodadas += 1
                If rodadas > MaxRodadasDeDescarga Then Return False

                Dim antes = Interlocked.Read(_edicoes)
                _saveTask = GravarAteEstabilizarAsync()
                Await _saveTask

                ' Continuar sujo SEM edição nova quer dizer que a gravação
                ' falhou. Insistir só repetiria a falha; o Status já explica
                ' o que houve, e o autosave tenta de novo na próxima tecla.
                If IsDirty AndAlso Interlocked.Read(_edicoes) = antes Then Return False
            Loop

            Return True
        End Function

        ' ================================================================
        ' Anexar
        ' ================================================================

        Private Async Function AnexarAsync() As Task
            Dim escolhido = _pickFile.AskWhichFileToAttach()
            If String.IsNullOrEmpty(escolhido) Then Return

            Dim resultado = Await _broker.AddDraftAttachmentAsync(_key, escolhido, CancellationToken.None)
            If resultado.Succeeded Then
                Attachments.Add(resultado.Value)
                Status = ""
            Else
                ' O nome vai para a tela porque é um arquivo do próprio
                ' usuário, escolhido por ele no diálogo — não é conteúdo
                ' vindo de e-mail.
                Status = "Não foi possível anexar " & System.IO.Path.GetFileName(escolhido) &
                         ". " & Traduzir(resultado.Kind)
            End If
        End Function

        ' ================================================================
        ' Envio
        ' ================================================================

        Private Async Function PedirEnvioAsync() As Task
            If State = ComposerState.SendUnknown Then Return

            Status = "Conferindo destinatários…"

            ' A confirmação tem de descrever o que vai sair AGORA. Conferir
            ' com o store atrasado faria o usuário aprovar uma versão e o
            ' Outlook mandar outra.
            If Not Await DescarregarAsync() Then
                AvisarDescargaIncompleta()
                Return
            End If

            Dim resultado = Await _broker.PrepareSendAsync(_key, CancellationToken.None)
            If Not resultado.Succeeded Then
                Status = "Não foi possível conferir o envio. " & Traduzir(resultado.Kind)
                Return
            End If

            Dim p = resultado.Value
            Preview = p
            Status = ""

            If p.Recipients.Count = 0 Then
                Status = "Sem destinatários."
                Return
            End If

            ' Destinatário não resolvido BLOQUEIA. Um nome que o Outlook não
            ' reconheceu pode virar endereço errado no envio, e não existe
            ' desfazer.
            If Not p.AllResolved Then
                Dim pendentes = p.Recipients.Where(Function(r) Not r.Resolved).
                                             Select(Function(r) r.DisplayName)
                Status = "O Outlook não reconheceu: " & String.Join("; ", pendentes) &
                         ". Corrija antes de enviar."
                Return
            End If

            State = ComposerState.ConfirmingSend
        End Function

        ''' <summary>
        ''' Envia UMA vez. Sem retry, em nenhuma hipótese.
        '''
        ''' Falha ambígua leva a um estado terminal: o compositor continua
        ''' aberto, com o texto todo, mas o botão de enviar não volta. O
        ''' usuário confere os Itens Enviados e decide — reenviar no escuro
        ''' é o único erro deste projeto que não tem volta.
        ''' </summary>
        Private Async Function ConfirmarEnvioAsync() As Task
            State = ComposerState.Sending
            Status = "Enviando…"

            Dim resultado = Await _broker.SendDraftAsync(_key, CancellationToken.None)

            If resultado.Succeeded Then
                Status = ""
                Encerrar()
                Return
            End If

            If resultado.IsAmbiguous Then
                State = ComposerState.SendUnknown
                Status = "O envio falhou e NÃO dá para saber se a mensagem saiu. " &
                         "Confira Itens Enviados no Outlook antes de tentar de novo. " &
                         "O Iris não vai reenviar sozinho."
                Return
            End If

            State = ComposerState.Editing
            Status = "Não foi possível enviar. " & Traduzir(resultado.Kind) &
                     " A mensagem continua aqui como rascunho."
        End Function

        ' ================================================================
        ' Fechamento
        ' ================================================================

        Private Sub Fechar()
            If State = ComposerState.SendUnknown Then
                ' Já é terminal: o rascunho fica no Outlook, intocado, para
                ' o usuário reconciliar.
                Encerrar()
                Return
            End If

            If IsDirty OrElse _autosave.IsEnabled OrElse
               (_saveTask IsNot Nothing AndAlso Not _saveTask.IsCompleted) Then
                State = ComposerState.ConfirmingClose
                Return
            End If

            Encerrar()
        End Sub

        Private Async Function SalvarEFecharAsync() As Task
            State = ComposerState.Editing
            Status = "Salvando…"

            If Not Await DescarregarAsync() Then
                ' Não fecha em cima de uma gravação que falhou: fechar aqui
                ' seria perder o texto justamente no caso em que ele não
                ' está guardado.
                AvisarDescargaIncompleta()
                Return
            End If

            Status = ""
            Encerrar()
        End Function

        ''' <summary>
        ''' Descarta o rascunho de vez. Só chega aqui por escolha explícita
        ''' na pergunta de fechamento.
        ''' </summary>
        Private Async Function DescartarAsync() As Task
            _autosave.Stop()

            Dim resultado = Await _broker.DeleteDraftAsync(_key, CancellationToken.None)

            ' NotFound é sucesso disfarçado: o rascunho já não está lá, que
            ' é exatamente o que se queria.
            If Not resultado.Succeeded AndAlso resultado.Kind <> ErrorKind.NotFound Then
                State = ComposerState.Editing
                Status = "Não foi possível descartar o rascunho. " & Traduzir(resultado.Kind)
                Return
            End If

            Encerrar()
        End Function

        Private Sub Encerrar()
            _autosave.Stop()

            _carregando = True
            Try
                Subject = ""
                ToLine = ""
                CcLine = ""
                UserText = ""
            Finally
                _carregando = False
            End Try

            QuotedPreview = ""
            Attachments.Clear()
            Preview = Nothing
            _key = Nothing
            _saveTask = Nothing
            _savePendente = False
            IsDirty = False
            State = ComposerState.Closed
        End Sub

        ''' <summary>
        ''' A descarga desistiu. Quase sempre a gravação falhou e o Status
        ''' já diz por quê; só quando ele está vazio — o caso raro de o
        ''' usuário digitar sem parar durante a descarga — é que falta
        ''' explicação, e ficar mudo aqui daria um botão que não faz nada.
        ''' </summary>
        Private Sub AvisarDescargaIncompleta()
            If HasStatus AndAlso Not Status.EndsWith("…", StringComparison.Ordinal) Then Return
            Status = "Ainda há alterações por salvar. Aguarde um instante e tente de novo."
        End Sub

        Private Sub NotificarComandos()
            AttachCommand.NotifyCanExecuteChanged()
            RequestSendCommand.NotifyCanExecuteChanged()
            ConfirmSendCommand.NotifyCanExecuteChanged()
            CancelSendCommand.NotifyCanExecuteChanged()
        End Sub

        Private Shared Function Traduzir(kind As ErrorKind) As String
            Select Case kind
                Case ErrorKind.NotConnected : Return "Sem conexão com o Outlook."
                Case ErrorKind.Busy : Return "O Outlook está ocupado."
                Case ErrorKind.NotFound : Return "O rascunho não está mais lá."
                Case ErrorKind.Denied : Return "Bloqueado pela política."
                Case ErrorKind.Cancelled : Return "Cancelado."
                Case ErrorKind.NotImplemented : Return "Ainda não implementado."
                Case Else : Return "Erro inesperado."
            End Select
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            _autosave.Stop()
            RemoveHandler _autosave.Tick, AddressOf OnAutosaveTick
        End Sub

    End Class

End Namespace

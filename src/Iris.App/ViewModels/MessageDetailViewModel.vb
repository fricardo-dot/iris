Imports System.Collections.ObjectModel
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
    ''' O painel de leitura.
    '''
    ''' Três atrasos deliberados, cada um com um motivo:
    '''
    '''   • 120 ms antes de pedir o detalhe. Navegar com as setas do teclado
    '''     percorre dezenas de mensagens; pedir corpo a cada uma encheria a
    '''     fila única da STA e a lista travaria (F1-F).
    '''   • Geração própria da seleção. O corpo de A não pode aparecer
    '''     depois que o usuário já selecionou B — e a resposta pode chegar
    '''     nessa ordem.
    '''   • 1 segundo antes de marcar como lida, cancelado se a seleção
    '''     mudar. Passar por uma mensagem com a seta não é lê-la.
    '''
    ''' Nesta versão o corpo é TEXTO PURO. HTML de e-mail é conteúdo hostil
    ''' vindo de fora, e o WebView2 endurecido entra depois, com a lista de
    ''' travas toda verificável — começar por ele seria abrir a superfície
    ''' antes de a leitura sequer funcionar.
    ''' </summary>
    Public NotInheritable Class MessageDetailViewModel
        Inherits ObservableObject
        Implements IDisposable

        Private Const SelecaoDebounceMs As Integer = 120
        Private Const MarcarLidaAposMs As Integer = 1000

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _ui As Dispatcher
        Private ReadOnly _observe As Action(Of Task, String)
        Private ReadOnly _saveFile As ISaveFileService
        Private ReadOnly _debounce As DispatcherTimer
        Private ReadOnly _marcarLida As DispatcherTimer

        Private _generation As Long = 0
        Private _pendente As MessageRowViewModel
        Private _linhaAtual As MessageRowViewModel

        Private _detail As MessageDetail
        Private _isLoading As Boolean
        Private _errorMessage As String = ""
        Private _disposed As Boolean

        Public Sub New(broker As IOutlookBroker, ui As Dispatcher,
                       observe As Action(Of Task, String), saveFile As ISaveFileService)
            _broker = broker
            _ui = ui
            _observe = observe
            _saveFile = saveFile
            SaveAttachmentCommand = New AsyncRelayCommand(Of AttachmentInfo)(AddressOf SalvarAnexoAsync)

            _debounce = New DispatcherTimer(DispatcherPriority.Normal, ui) With {
                .Interval = TimeSpan.FromMilliseconds(SelecaoDebounceMs)
            }
            AddHandler _debounce.Tick, AddressOf OnDebounceTick

            _marcarLida = New DispatcherTimer(DispatcherPriority.Background, ui) With {
                .Interval = TimeSpan.FromMilliseconds(MarcarLidaAposMs)
            }
            AddHandler _marcarLida.Tick, AddressOf OnMarcarLidaTick
        End Sub

        Public ReadOnly Property Attachments As New ObservableCollection(Of AttachmentInfo)()
        Public ReadOnly Property SaveAttachmentCommand As IAsyncRelayCommand(Of AttachmentInfo)

        Private _attachmentStatus As String = ""

        ''' <summary>
        ''' Resultado da última tentativa de salvar. Fica visível porque
        ''' salvar um anexo pode falhar de três maneiras diferentes, e
        ''' silêncio faria as três parecerem sucesso.
        ''' </summary>
        Public Property AttachmentStatus As String
            Get
                Return _attachmentStatus
            End Get
            Private Set(value As String)
                If SetProperty(_attachmentStatus, value) Then
                    OnPropertyChanged(NameOf(HasAttachmentStatus))
                End If
            End Set
        End Property

        Public ReadOnly Property HasAttachmentStatus As Boolean
            Get
                Return Not String.IsNullOrEmpty(_attachmentStatus)
            End Get
        End Property

        Public Property Detail As MessageDetail
            Get
                Return _detail
            End Get
            Private Set(value As MessageDetail)
                If SetProperty(_detail, value) Then
                    OnPropertyChanged(NameOf(HasMessage))
                    OnPropertyChanged(NameOf(Subject))
                    OnPropertyChanged(NameOf(SenderLine))
                    OnPropertyChanged(NameOf(RecipientLine))
                    OnPropertyChanged(NameOf(Body))
                    OnPropertyChanged(NameOf(BodyNotice))
                    OnPropertyChanged(NameOf(HasBodyNotice))
                    OnPropertyChanged(NameOf(HasAttachments))
                    OnPropertyChanged(NameOf(CanReply))
                    OnPropertyChanged(NameOf(CanForward))
                    OnPropertyChanged(NameOf(PartialReadNotice))
                    OnPropertyChanged(NameOf(HasPartialRead))
                End If
            End Set
        End Property

        Public ReadOnly Property HasMessage As Boolean
            Get
                Return _detail IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property Subject As String
            Get
                If _detail Is Nothing Then Return ""
                Return If(String.IsNullOrWhiteSpace(_detail.Subject), "(sem assunto)", _detail.Subject)
            End Get
        End Property

        Public ReadOnly Property SenderLine As String
            Get
                If _detail Is Nothing Then Return ""
                Dim nome = _detail.SenderName
                If String.IsNullOrWhiteSpace(nome) Then nome = _detail.SenderAddress
                If _detail.ReceivedTime.HasValue Then
                    Return $"{nome} · {_detail.ReceivedTime.Value.LocalDateTime:dd/MM/yyyy HH:mm}"
                End If
                Return nome
            End Get
        End Property

        Public ReadOnly Property RecipientLine As String
            Get
                If _detail Is Nothing OrElse _detail.Recipients.Count = 0 Then Return ""
                Dim para = _detail.Recipients.
                    Where(Function(r) r.Kind <> RecipientKind.Bcc).
                    Select(Function(r) If(String.IsNullOrWhiteSpace(r.DisplayName), r.Address, r.DisplayName))
                Return "Para: " & String.Join(", ", para)
            End Get
        End Property

        Public ReadOnly Property Body As String
            Get
                Return If(_detail Is Nothing, "", _detail.TextBody)
            End Get
        End Property

        ''' <summary>
        ''' A UI traduz o <see cref="ErrorKind"/>. Nunca exibe o
        ''' <c>Detail</c> do resultado, que é diagnóstico.
        ''' </summary>
        Public ReadOnly Property BodyNotice As String
            Get
                If _detail Is Nothing Then Return ""
                Select Case _detail.BodyError
                    Case ErrorKind.Denied
                        Return "Esta mensagem é protegida. O Iris não exibe o conteúdo dela."
                    Case ErrorKind.NotDownloaded
                        Return "O conteúdo ainda não foi baixado pelo Outlook."
                    Case Else
                        Return ""
                End Select
            End Get
        End Property

        Public ReadOnly Property HasBodyNotice As Boolean
            Get
                Return Not String.IsNullOrEmpty(BodyNotice)
            End Get
        End Property

        ''' <summary>
        ''' Dá para responder a esta mensagem?
        '''
        ''' Com a lista de destinatários incompleta, "Responder a todos"
        ''' responde a MENOS gente do que a mensagem tinha — e ninguém
        ''' percebe, porque o resultado parece normal: uma resposta foi
        ''' enviada, para pessoas reais. O que falta é invisível.
        ''' </summary>
        Public ReadOnly Property CanReply As Boolean
            Get
                Return _detail IsNot Nothing AndAlso
                       ReplyReadiness.CanReply(_detail.RecipientsStatus)
            End Get
        End Property

        ''' <summary>
        ''' Encaminhar leva os ANEXOS junto, e anexo que não foi lido não
        ''' deixa rastro na tela — ao contrário de um corpo truncado.
        ''' </summary>
        Public ReadOnly Property CanForward As Boolean
            Get
                Return _detail IsNot Nothing AndAlso
                       ReplyReadiness.CanForward(_detail.AttachmentsStatus)
            End Get
        End Property

        ''' <summary>Explica o que ficou faltando, e por isso bloqueou.</summary>
        Public ReadOnly Property PartialReadNotice As String
            Get
                If _detail Is Nothing Then Return ""

                Dim destinatarios = ReplyReadiness.DescribeBlock("os destinatários",
                                                                 _detail.RecipientsStatus)
                Dim anexos = ReplyReadiness.DescribeBlock("os anexos",
                                                          _detail.AttachmentsStatus)

                Dim partes = New List(Of String)()
                If destinatarios.Length > 0 Then partes.Add(destinatarios)
                If anexos.Length > 0 Then partes.Add(anexos)
                If partes.Count = 0 Then Return ""

                Return String.Join(" ", partes) &
                       " Responder e encaminhar ficam bloqueados até uma releitura completa."
            End Get
        End Property

        Public ReadOnly Property HasPartialRead As Boolean
            Get
                Return PartialReadNotice.Length > 0
            End Get
        End Property

        Public ReadOnly Property HasAttachments As Boolean
            Get
                Return Attachments.Count > 0
            End Get
        End Property

        Public Property IsLoading As Boolean
            Get
                Return _isLoading
            End Get
            Private Set(value As Boolean)
                SetProperty(_isLoading, value)
            End Set
        End Property

        Public Property ErrorMessage As String
            Get
                Return _errorMessage
            End Get
            Private Set(value As String)
                If SetProperty(_errorMessage, value) Then OnPropertyChanged(NameOf(HasError))
            End Set
        End Property

        Public ReadOnly Property HasError As Boolean
            Get
                Return Not String.IsNullOrEmpty(_errorMessage)
            End Get
        End Property

        ' ===================================================================

        ''' <summary>
        ''' A seleção mudou. Nada é pedido ainda — o debounce decide.
        ''' </summary>
        Public Sub Show(linha As MessageRowViewModel)
            ' A MESMA mensagem, num objeto de linha novo.
            '
            ' Acontece toda vez que a lista recarrega e restaura a seleção
            ' pela chave — inclusive na recarga que a própria marcação como
            ' lida provoca, via ItemChange. Sem esta guarda, abrir uma
            ' mensagem não lida custava DOIS getMessageDetail: um por abrir,
            ' outro pela reconciliação. Uma ida ao COM à toa na fila única.
            If linha IsNot Nothing AndAlso _detail IsNot Nothing AndAlso
               linha.Key.Equals(_detail.Key) Then
                _linhaAtual = linha
                Return
            End If

            Interlocked.Increment(_generation)
            _marcarLida.Stop()
            _debounce.Stop()

            _pendente = linha
            _linhaAtual = Nothing
            AttachmentStatus = ""

            If linha Is Nothing Then
                Detail = Nothing
                Attachments.Clear()
                ErrorMessage = ""
                Return
            End If

            _debounce.Start()
        End Sub

        Public Sub Clear()
            Show(Nothing)
        End Sub

        Private Sub OnDebounceTick(sender As Object, e As EventArgs)
            _debounce.Stop()
            Dim linha = _pendente
            If linha Is Nothing Then Return
            _observe(CarregarAsync(linha, Volatile.Read(_generation)), "detail.load")
        End Sub

        Private Async Function CarregarAsync(linha As MessageRowViewModel, geracao As Long) As Task
            IsLoading = True
            ErrorMessage = ""
            Try
                Dim resultado = Await _broker.GetMessageDetailAsync(linha.Key, CancellationToken.None)

                ' A seleção pode ter mudado enquanto o corpo era lido.
                If Volatile.Read(_generation) <> geracao Then Return

                Await _ui.InvokeAsync(
                    Sub()
                        If Volatile.Read(_generation) <> geracao Then Return

                        If Not resultado.Succeeded Then
                            Detail = Nothing
                            Attachments.Clear()
                            ErrorMessage = Traduzir(resultado.Kind)
                            Return
                        End If

                        Detail = resultado.Value
                        Attachments.Clear()
                        For Each a In resultado.Value.Attachments
                            Attachments.Add(a)
                        Next
                        OnPropertyChanged(NameOf(HasAttachments))

                        _linhaAtual = linha
                        ' Só agora começa a contar o tempo de leitura.
                        If linha.IsUnread Then _marcarLida.Start()
                    End Sub).Task
            Finally
                ' So a geracao CORRENTE pode desligar o indicador. Uma leitura
                ' obsoleta terminando depois de o usuario ter selecionado
                ' outra mensagem apagaria o "carregando" da leitura nova.
                If Volatile.Read(_generation) = geracao Then IsLoading = False
            End Try
        End Function

        ''' <summary>
        ''' Marca como lida depois de a mensagem ficar exibida por um tempo.
        '''
        ''' Passar por uma mensagem com a seta do teclado não é lê-la, e
        ''' marcar na seleção transformaria navegação em leitura.
        ''' </summary>
        Private Sub OnMarcarLidaTick(sender As Object, e As EventArgs)
            _marcarLida.Stop()

            Dim linha = _linhaAtual
            If linha Is Nothing OrElse Not linha.IsUnread Then Return

            ' Otimista: a linha perde o negrito na hora. O ItemChange que o
            ' Outlook vai disparar cai no debounce normal do watcher e vira
            ' uma reconciliação só, em vez de um laço (F1-G).
            linha.IsUnread = False

            _observe(MarcarAsync(linha), "detail.markRead")
        End Sub

        Private Async Function MarcarAsync(linha As MessageRowViewModel) As Task
            Dim resultado = Await _broker.MarkReadAsync(linha.Key, True, CancellationToken.None)
            ' A janela fechou enquanto o Outlook marcava. A marcacao pode ter
            ' valido -- efeito no mundo nao se desfaz por a tela sumir --, mas
            ' reverter a linha na tela agora seria escrever numa lista que ja
            ' nao esta em lugar nenhum.
            If _disposed Then Return
            If resultado.Succeeded Then Return

            ' Em falha AMBÍGUA não se reverte: a marcação pode ter sido
            ' aplicada, e desfazer na tela mentiria tanto quanto manter.
            ' A reconciliação seguinte resolve.
            If resultado.IsAmbiguous Then Return

            Await _ui.InvokeAsync(
                Sub()
                    ' De novo DENTRO do delegate: entre o agendamento e a
                    ' execucao no dispatcher cabe um Dispose.
                    If _disposed Then Return
                    linha.IsUnread = True
                End Sub).Task
        End Function

        ''' <summary>
        ''' Salva um anexo onde o usuário escolher.
        '''
        ''' ABRIR o anexo continua fora: abrir é executar conteúdo não
        ''' confiável, e o Iris não vai fazer isso por conta própria (F1-J).
        ''' </summary>
        Private Async Function SalvarAnexoAsync(anexo As AttachmentInfo) As Task
            If anexo Is Nothing Then Return

            Dim destino = _saveFile.AskWhereToSave(anexo.FileName)
            If String.IsNullOrEmpty(destino) Then Return

            ' A GERACAO DESTE SALVAMENTO.
            '
            ' _disposed sozinho fecha so o fechamento da janela. Sem geracao,
            ' trocar de mensagem durante a gravacao publicava "Salvo em ..." --
            ' ou a falha do anexo da mensagem A -- no leitor da B, com o
            ' ViewModel vivo e ninguem para desconfiar.
            Dim geracao = Volatile.Read(_generation)

            AttachmentStatus = $"Salvando {anexo.FileName}…"

            ' O diálogo já confirmou a sobrescrita com o usuário; aqui o
            ' overwrite é consequência daquela confirmação, não decisão
            ' silenciosa nossa.
            Dim resultado = Await _broker.SaveAttachmentAsync(
                anexo.Key, destino, overwrite:=True, cancel:=CancellationToken.None)

            ' O ARQUIVO PODE TER SIDO GRAVADO, e isso vale: cancelar nao desfaz
            ' escrita ja comecada. O que nao vale e anunciar o desfecho num
            ' leitor que ja saiu da tela -- ou que agora mostra outra mensagem.
            If _disposed OrElse Volatile.Read(_generation) <> geracao Then Return

            Await _ui.InvokeAsync(
                Sub()
                    ' De novo dentro do delegate: entre agendar e o dispatcher
                    ' executar cabe um Dispose e cabe uma troca de mensagem.
                    If _disposed OrElse Volatile.Read(_generation) <> geracao Then Return
                    If resultado.Succeeded Then
                        AttachmentStatus = $"Salvo em {destino}"
                    ElseIf resultado.IsAmbiguous Then
                        ' Gravação em disco que falha depois de começar pode
                        ' ter deixado arquivo completo, parcial ou nenhum.
                        AttachmentStatus =
                            $"Não foi possível confirmar se {anexo.FileName} foi salvo. " &
                            "Verifique a pasta escolhida antes de tentar de novo."
                    Else
                        AttachmentStatus = TraduzirAnexo(resultado.Kind, anexo.FileName)
                    End If
                End Sub).Task
        End Function

        Private Shared Function TraduzirAnexo(kind As ErrorKind, nome As String) As String
            Select Case kind
                Case ErrorKind.Stale
                    Return $"{nome} mudou desde que a mensagem foi aberta. Reabra a mensagem."
                Case ErrorKind.NotFound
                    Return $"{nome} não está mais nesta mensagem."
                Case ErrorKind.Denied
                    Return $"Sem permissão para gravar {nome}, ou o arquivo está em uso."
                Case ErrorKind.NotConnected
                    Return "Sem conexão com o Outlook."
                Case ErrorKind.Busy
                    Return "O Outlook está ocupado. Tente de novo em instantes."
                Case Else
                    Return $"Não foi possível salvar {nome}."
            End Select
        End Function

        Private Shared Function Traduzir(kind As ErrorKind) As String
            Select Case kind
                Case ErrorKind.NotFound : Return "Esta mensagem não está mais aqui."
                Case ErrorKind.NotConnected : Return "Sem conexão com o Outlook."
                Case ErrorKind.Busy : Return "O Outlook está ocupado."
                Case ErrorKind.Denied : Return "Acesso negado pela política."
                Case Else : Return "Não foi possível abrir a mensagem."
            End Select
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            ' Invalida leitura em voo: sem isto, uma carga ja iniciada
            ' concluiria depois do descarte e ainda tocaria propriedades.
            Interlocked.Increment(_generation)
            _debounce.Stop()
            _marcarLida.Stop()
            RemoveHandler _debounce.Tick, AddressOf OnDebounceTick
            RemoveHandler _marcarLida.Tick, AddressOf OnMarcarLidaTick
        End Sub

    End Class

End Namespace

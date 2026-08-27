Imports System.IO
Imports System.Windows.Threading
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Cache
Imports Iris.Core
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.Assist
Imports Iris.Integration
Imports Iris.Integration.Outlook
Imports Iris.Model
Imports Iris.Sync

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' O acervo — o que o Iris guardou do que já viu — e a ressalva que ele
    ''' obriga a mostrar junto.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTE VIEWMODEL NÃO É</b>
    '''
    ''' Não é a lista de mensagens. A lista continua lendo <b>ao vivo</b> do
    ''' Outlook, pelo broker, e é isso que o usuário opera.
    '''
    ''' Este aqui mostra o <b>acervo</b>: o que a varredura publicou no cache.
    ''' São coisas diferentes e a §23 explica por quê — em modo cached o cache é
    ''' um arquivo histórico conservador, não o estado corrente da caixa. Pode
    ''' faltar mensagem que existe no servidor, e pode conter mensagem que o
    ''' usuário já apagou.
    '''
    ''' Trocar a lista para ler daqui é a fase seguinte, e exige que o cache
    ''' ganhe busca, ordenação e reconciliação com o que está ao vivo (§26.3).
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O DRENO</b>
    '''
    ''' A §26.2 exige o consumidor ligado ao <see cref="PublicationDrain"/> na
    ''' inicialização <b>e</b> durante a execução, e proíbe a UI de contornar o
    ''' dreno lendo o manifesto direto como substituto da dívida registrada.
    '''
    ''' Por isso este ViewModel <b>nunca</b> chama <c>ManifestReader</c>: ele
    ''' observa o <see cref="AcervoService"/>, e quem atualiza o serviço é o
    ''' dreno entregando a geração. A leitura direta acontece uma vez só, na
    ''' construção do serviço, para a tela não abrir vazia.
    ''' </summary>
    Public NotInheritable Class AcervoViewModel
        Inherits ObservableObject
        Implements IDisposable

        Private ReadOnly _ui As Dispatcher
        Private ReadOnly _db As CacheDatabase
        Private ReadOnly _servico As AcervoService
        Private ReadOnly _dreno As PublicationDrain
        Private ReadOnly _varredura As VarreduraDaPasta
        Private _alvoDoOutlook As FolderKey
        Private _nomeDoAlvo As String = ""
        Private _storeDoAlvo As StoreInfo
        Private ReadOnly _relogio As DispatcherTimer
        Private _disposed As Boolean

        ''' <summary>
        ''' O diário do egress, sobre o mesmo banco.
        '''
        ''' Mora aqui porque é aqui que o cache está aberto — e não porque o
        ''' acervo tenha algo a ver com a IA. Sem cache não há diário, e sem
        ''' diário a IA fica desligada: transmitir sem poder registrar seria
        ''' pior que não transmitir.
        ''' </summary>
        Public ReadOnly Property Diario As IDisclosureJournal

        ''' <summary>
        ''' Onde o cache mora. Em <c>%LOCALAPPDATA%</c> e não ao lado do
        ''' executável: o executável pode estar em Program Files, onde escrever
        ''' exige elevação, e um cache que só funciona com privilégio não é um
        ''' cache.
        ''' </summary>
        Public Shared Function CaminhoPadrao() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "cache.db")
        End Function

        ''' <summary>
        ''' Abre o acervo, ou devolve <c>Nothing</c> com o motivo em
        ''' <paramref name="motivoDaFalha"/>.
        '''
        ''' FALHA FECHADO e VISÍVEL: cache que não abre não pode virar tela
        ''' vazia silenciosa, porque tela vazia é indistinguível de "não há
        ''' nada guardado".
        ''' </summary>
        Public Shared Function Abrir(ui As Dispatcher, folderKey As Long,
                                     ByRef motivoDaFalha As String,
                                     Optional broker As IOutlookBroker = Nothing) As AcervoViewModel
            motivoDaFalha = Nothing
            Try
                Dim caminho = CaminhoPadrao()
                Directory.CreateDirectory(Path.GetDirectoryName(caminho))

                Dim falha As OpenFailure = Nothing
                Dim db = CacheDatabase.Open(caminho, CacheSchema.Intended(), falha)
                If db Is Nothing Then
                    motivoDaFalha = $"o cache não abriu ({falha})"
                    Return Nothing
                End If
                Return New AcervoViewModel(ui, db, folderKey, broker)
            Catch ex As Exception
                motivoDaFalha = $"o cache não abriu ({ex.GetType().Name}: {ex.Message})"
                Return Nothing
            End Try
        End Function

        Private Sub New(ui As Dispatcher, db As CacheDatabase, folderKey As Long,
                        broker As IOutlookBroker)
            _ui = ui
            _db = db
            ' Sem broker nao ha varredura, e o botao fica desabilitado. E o
            ' caso dos testes que so olham o lado de leitura.
            _varredura = If(broker Is Nothing, Nothing,
                            New VarreduraDaPasta(broker, db))
            _servico = New AcervoService(db, folderKey)
            _dreno = New PublicationDrain(db)
            Diario = New SqliteDisclosureJournal(db)

            AddHandler _servico.Mudou, AddressOf AoMudar

            DrenarCommand = New RelayCommand(AddressOf Drenar)
            VarrerCommand = New AsyncRelayCommand(AddressOf VarrerAsync, Function() PodeVarrer)

            ' Na INICIALIZACAO.
            Drenar()

            ' E DURANTE A EXECUCAO. O intervalo e folgado de proposito: nada no
            ' app publica ainda, entao o dreno so tem trabalho quando uma
            ' varredura rodar. Bater no banco a cada segundo para nao achar
            ' nada seria custo sem informacao.
            _relogio = New DispatcherTimer(DispatcherPriority.Background, ui) With {
                .Interval = TimeSpan.FromSeconds(30)}
            AddHandler _relogio.Tick, Sub() Drenar()
            _relogio.Start()

            Atualizar()
        End Sub

        Public ReadOnly Property DrenarCommand As RelayCommand

        ''' <summary>
        ''' <b>Varre a pasta selecionada.</b> Botão, e não automático.
        '''
        ''' ------------------------------------------------------------------
        ''' Varrer é caro — a §D5 mediu o custo por lote — e <b>escreve no
        ''' cache</b>. Disparar a cada clique numa pasta gastaria COM sem
        ''' ninguém pedir, e numa caixa grande travaria a troca de pasta.
        ''' O momento e o custo ficam na mão de quem opera.
        ''' </summary>
        Public ReadOnly Property VarrerCommand As AsyncRelayCommand

        Private _varrendo As Boolean
        Public Property Varrendo As Boolean
            Get
                Return _varrendo
            End Get
            Private Set(valor As Boolean)
                SetProperty(_varrendo, valor)
                OnPropertyChanged(NameOf(PodeVarrer))
                VarrerCommand.NotifyCanExecuteChanged()
            End Set
        End Property

        ''' <summary>
        ''' Dá para varrer? Exige pasta escolhida, broker, e não estar varrendo.
        '''
        ''' <b>Não</b> exige ambiente autorizado: quem recusa por isso é a
        ''' <see cref="VarreduraDaPasta"/>, e a recusa dela <b>explica</b>. Um
        ''' botão desabilitado sem motivo na tela é o defeito que a faixa da IA
        ''' já teve — o usuário clica e nada acontece.
        ''' </summary>
        Public ReadOnly Property PodeVarrer As Boolean
            Get
                Return _varredura IsNot Nothing AndAlso
                       _alvoDoOutlook IsNot Nothing AndAlso Not Varrendo
            End Get
        End Property

        Private _itens As Integer
        Public Property Itens As Integer
            Get
                Return _itens
            End Get
            Private Set(value As Integer)
                SetProperty(_itens, value)
            End Set
        End Property

        Private _ressalva As String
        ''' <summary>
        ''' O que a UI é obrigada a mostrar junto do acervo. <c>Nothing</c>
        ''' quando não há ressalva — que hoje nunca acontece em modo cached.
        ''' </summary>
        Public Property Ressalva As String
            Get
                Return _ressalva
            End Get
            Private Set(value As String)
                SetProperty(_ressalva, value)
            End Set
        End Property

        Private _temAlgoADizer As Boolean
        Public Property TemAlgoADizer As Boolean
            Get
                Return _temAlgoADizer
            End Get
            Private Set(value As Boolean)
                SetProperty(_temAlgoADizer, value)
            End Set
        End Property

        Private _travado As String
        ''' <summary>
        ''' Preenchido quando a fila de publicação travou na cabeça. A §26.2
        ''' exige que falha persistente do consumidor apareça — não fique
        ''' bloqueando em silêncio.
        ''' </summary>
        Public Property Travado As String
            Get
                Return _travado
            End Get
            Private Set(value As String)
                SetProperty(_travado, value)
            End Set
        End Property

        ' ==============================================================

        Private Sub Drenar()
            If _disposed Then Return
            Try
                _dreno.Drenar(_servico)
                Travado = Nothing
            Catch ex As Exception
                ' Consumidor que falha trava a cabeca da fila DE PROPOSITO —
                ' marcar como drenada uma geracao nao recebida seria perder em
                ' silencio. O que nao pode e o bloqueio ser invisivel.
                Dim g = _dreno.TravadoEm()
                Travado = If(g.HasValue,
                    $"A atualização do acervo parou na geração {g.Value}: {_dreno.UltimoErro(g.Value)}",
                    $"A atualização do acervo falhou: {ex.Message}")
            End Try
            Atualizar()
        End Sub

        Private Sub AoMudar(sender As Object, e As EventArgs)
            If _ui.CheckAccess() Then Atualizar() Else _ui.BeginInvoke(CType(AddressOf Atualizar, Action))
        End Sub

        ''' <summary>
        ''' <b>Passa a mostrar a pasta que o usuário selecionou.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' Era a constante 1 — a pasta que uma importação manual de teste tinha
        ''' criado. A lista ao lado mostrava uma pasta e o acervo mostrava
        ''' outra, sem nada dizendo isso.
        '''
        ''' Resolver a pasta <b>cria</b> a linha no cache se ela não existir, e
        ''' isso é de propósito: sem linha não há para onde a varredura
        ''' publicar, e o acervo de uma pasta nunca vista tem de poder dizer
        ''' "nada guardado" em vez de falhar.
        ''' </summary>
        Public Sub Apontar(pasta As FolderKey, nome As String, store As StoreInfo)
            _alvoDoOutlook = pasta
            _nomeDoAlvo = If(nome, "")
            _storeDoAlvo = store

            If pasta Is Nothing OrElse String.IsNullOrWhiteSpace(pasta.EntryId) Then
                OnPropertyChanged(NameOf(PodeVarrer))
                VarrerCommand.NotifyCanExecuteChanged()
                Return
            End If

            Try
                _servico.Apontar(New ResolvedorDoAcervo(_db).Pasta(
                    pasta.StoreId, pasta.EntryId, _nomeDoAlvo))
            Catch ex As Exception
                ' Resolver e escrita: disco cheio e banco travado chegam aqui.
                ' Nao pode derrubar a troca de pasta -- a lista ao lado
                ' continua funcionando, e o acervo e o painel secundario.
                Travado = "Nao consegui apontar o acervo para esta pasta."
            End Try

            OnPropertyChanged(NameOf(PodeVarrer))
            VarrerCommand.NotifyCanExecuteChanged()
            Atualizar()
        End Sub

        Private Async Function VarrerAsync() As Task
            If Not PodeVarrer Then Return

            Dim pasta = _alvoDoOutlook
            Dim nome = _nomeDoAlvo
            Dim store = _storeDoAlvo

            Varrendo = True
            Travado = Nothing
            Try
                ' Task.Run: varrer bloqueia, e bloquear o dispatcher congelaria
                ' a janela inteira pelo tempo da varredura.
                Dim r = Await Task.Run(
                    Function() _varredura.Executar(pasta, nome, store, CancellationToken.None))

                Travado = EmPortugues(r)
                _servico.Recarregar()
                Atualizar()
            Catch ex As Exception
                Travado = "A varredura falhou."
            Finally
                Varrendo = False
            End Try
        End Function

        ''' <summary>
        ''' O desfecho da varredura em português.
        '''
        ''' A recusa por ambiente não autorizado <b>diz o que fazer</b>: sem
        ''' isso o usuário clica, nada acontece, e a única pista mora num banco
        ''' SQLite. Foi exatamente o que a faixa da IA já fez com o botão
        ''' desabilitado.
        ''' </summary>
        Private Shared Function EmPortugues(r As ResultadoDaVarredura) As String
            Select Case r.Recusa
                Case RecusaDeVarredura.AmbienteNaoAutorizado
                    Return "Este ambiente ainda não foi autorizado para varredura " &
                           $"({r.Ambiente}). Autorize com:  " &
                           $"dotnet run --project tools\Iris.CrashHarness -- ambiente --autorizar {r.ChaveDoAmbiente}"
                Case RecusaDeVarredura.SemPasta
                    Return "Escolha uma pasta antes de varrer."
                Case RecusaDeVarredura.StoreDesconhecido
                    Return "Não sei a qual conta esta pasta pertence."
                Case RecusaDeVarredura.Falhou
                    Return "A varredura não terminou, e a tentativa foi descartada."
            End Select

            If r.Varredura Is Nothing Then Return "A varredura não produziu desfecho."
            If r.Varredura.Conclusion = SweepConclusion.Publicada Then Return Nothing

            ' Sem o Motivo: ele e texto, e pode carregar assunto de mensagem.
            Return $"A varredura terminou como {r.Varredura.Conclusion} e nada foi publicado."
        End Function

        Private Sub Atualizar()
            Dim m = _servico.Atual
            Itens = m.Items.Count
            Ressalva = m.Ressalva
            TemAlgoADizer = (m.Ressalva IsNot Nothing) OrElse (Travado IsNot Nothing)
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            _relogio?.Stop()
            RemoveHandler _servico.Mudou, AddressOf AoMudar
            _db?.Dispose()
        End Sub

    End Class

End Namespace

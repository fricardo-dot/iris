Imports System.IO
Imports System.Windows.Threading
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration

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
        Private ReadOnly _relogio As DispatcherTimer
        Private _disposed As Boolean

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
                                     ByRef motivoDaFalha As String) As AcervoViewModel
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
                Return New AcervoViewModel(ui, db, folderKey)
            Catch ex As Exception
                motivoDaFalha = $"o cache não abriu ({ex.GetType().Name}: {ex.Message})"
                Return Nothing
            End Try
        End Function

        Private Sub New(ui As Dispatcher, db As CacheDatabase, folderKey As Long)
            _ui = ui
            _db = db
            _servico = New AcervoService(db, folderKey)
            _dreno = New PublicationDrain(db)

            AddHandler _servico.Mudou, AddressOf AoMudar

            DrenarCommand = New RelayCommand(AddressOf Drenar)

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

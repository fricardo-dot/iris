Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' Observa a pasta exibida e avisa, com calma, que ela mudou.
    '''
    ''' Por que existe um atraso deliberado entre o evento e a recarga:
    '''
    '''   • Uma sincronização do Exchange dispara dezenas de eventos em
    '''     sequência. Recarregar a cada um faria a lista piscar e encheria
    '''     a fila única da STA — o F1-F.
    '''   • A Fase 0 provou que os eventos não têm ordem causal confiável e
    '''     que o estado lido depois pode não ser o que causou o evento. Não
    '''     há o que interpretar em ItemAdd/Change/Remove; só há "esta pasta
    '''     está suja".
    ''' </summary>
    Public NotInheritable Class FolderWatcher
        Implements IDisposable

        Private Const DebounceMs As Integer = 450
        Private Const TetoMs As Integer = 2000

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _ui As Dispatcher
        Private ReadOnly _observe As Action(Of Task, String)
        Private ReadOnly _onDirty As Action(Of FolderKey)
        Private ReadOnly _timer As DispatcherTimer
        Private ReadOnly _gate As New Object()

        Private _token As SubscriptionToken
        Private _watched As FolderKey
        Private _generation As Integer = 0

        ''' <summary>
        ''' A decisão de QUANDO recarregar mora no Core, como máquina pura.
        ''' Aqui fica o que é do WPF: timer, dispatcher e lock.
        ''' </summary>
        Private ReadOnly _debounce As New DirtyDebounce(DebounceMs, TetoMs)

        Private _disposed As Boolean

        Public Sub New(broker As IOutlookBroker, ui As Dispatcher,
                       observe As Action(Of Task, String), onDirty As Action(Of FolderKey))
            _broker = broker
            _ui = ui
            _observe = observe
            _onDirty = onDirty

            _timer = New DispatcherTimer(DispatcherPriority.Background, ui) With {
                .Interval = TimeSpan.FromMilliseconds(150)
            }
            AddHandler _timer.Tick, AddressOf OnTick

            AddHandler _broker.FolderInvalidated, AddressOf OnInvalidated
        End Sub

        ''' <summary>
        ''' Passa a observar outra pasta.
        '''
        ''' A geração existe porque estas chamadas NÃO são serializadas por
        ''' quem chama: trocar de pasta rápido disparava duas assinaturas
        ''' concorrentes, as respostas chegavam fora de ordem, e uma delas
        ''' ficava viva no broker sem token para removê-la — assinatura
        ''' órfã, com o event sink conectado para sempre.
        ''' </summary>
        Public Async Function WatchAsync(folder As FolderKey) As Task
            Dim geracao As Integer
            SyncLock _gate
                _generation += 1
                geracao = _generation
                _watched = folder
                _debounce.Clear()
            End SyncLock

            Await UnwatchCoreAsync()
            If folder Is Nothing Then Return

            Dim resultado = Await _broker.SubscribeFolderAsync(folder, CancellationToken.None)
            If Not resultado.Succeeded Then Return

            Dim aceita As Boolean
            SyncLock _gate
                aceita = _generation = geracao
                If aceita Then _token = resultado.Value
            End SyncLock

            If aceita Then
                _timer.Start()
            Else
                ' A resposta chegou tarde: outra pasta já foi pedida. Cancela
                ' ESTA assinatura, senão ela fica viva sem ninguém para
                ' desligá-la.
                Await _broker.UnsubscribeFolderAsync(resultado.Value, CancellationToken.None)
            End If
        End Function

        Public Async Function UnwatchAsync() As Task
            SyncLock _gate
                _generation += 1
                _watched = Nothing
                _debounce.Clear()
            End SyncLock
            Await UnwatchCoreAsync()
        End Function

        Private Async Function UnwatchCoreAsync() As Task
            _timer.Stop()

            Dim anterior As SubscriptionToken
            SyncLock _gate
                anterior = _token
                _token = Nothing
            End SyncLock

            If anterior Is Nothing Then Return
            Await _broker.UnsubscribeFolderAsync(anterior, CancellationToken.None)
        End Function

        ''' <summary>
        ''' Chega numa thread MTA do pool. Só marca o bit — nada de tocar em
        ''' UI, coleção ou COM daqui.
        ''' </summary>
        ''' <summary>
        ''' A sessão COM foi substituída.
        '''
        ''' As assinaturas morreram junto com ela, do lado do broker. O token
        ''' guardado aqui aponta para uma que não existe mais — e chamar
        ''' Unsubscribe com ele só produziria erro. Larga o token, apaga o
        ''' estado de debounce, e para de escutar a pasta antiga.
        '''
        ''' Sem isto, o watcher continuava achando que estava assinando: era
        ''' o segundo tempo da falha silenciosa, em que nenhum evento chegava
        ''' mais e ninguém tinha como perceber.
        ''' </summary>
        Public Sub OnSessionReplaced()
            SyncLock _gate
                _token = Nothing
                _watched = Nothing
                _debounce.Clear()
                _generation += 1
            End SyncLock

            _timer.Stop()
        End Sub

        Private Sub OnInvalidated(sender As Object, invalidation As FolderInvalidation)
            If _disposed Then Return

            ' Uma leitura só, sob lock: ler _token duas vezes permitia a UI
            ' anulá-lo entre a checagem e o uso.
            Dim idAtual As Integer
            SyncLock _gate
                If _token Is Nothing Then Return
                idAtual = _token.Id
            End SyncLock

            If invalidation.SubscriptionId <> idAtual Then Return
            Dim idDoEvento = invalidation.SubscriptionId

            _ui.InvokeAsync(
                Sub()
                    If _disposed Then Return
                    SyncLock _gate
                        ' Revalida JÁ na UI: o evento pode ter esperado na
                        ' fila do dispatcher enquanto o usuário trocava de
                        ' pasta, e marcaria a pasta nova como suja por conta
                        ' de uma mudança na antiga.
                        If _token Is Nothing OrElse _token.Id <> idDoEvento Then Return

                        _debounce.Mark(DateTimeOffset.UtcNow)
                    End SyncLock
                End Sub)
        End Sub

        Private Sub OnTick(sender As Object, e As EventArgs)
            Dim pasta As FolderKey = Nothing

            SyncLock _gate
                If _watched Is Nothing Then Return
                If Not _debounce.ShouldFlush(DateTimeOffset.UtcNow) Then Return

                _debounce.Clear()
                pasta = _watched
            End SyncLock

            _onDirty(pasta)
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            RemoveHandler _broker.FolderInvalidated, AddressOf OnInvalidated
            _timer.Stop()
            RemoveHandler _timer.Tick, AddressOf OnTick
            _observe(UnwatchAsync(), "watcher.dispose")
        End Sub

    End Class

End Namespace

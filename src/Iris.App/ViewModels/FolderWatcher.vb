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
    '''
    ''' Por isso: um bit de sujeira por pasta, debounce de 450 ms, e um teto
    ''' de 2 s para que uma sincronização contínua não adie a recarga para
    ''' sempre.
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

        Private _token As SubscriptionToken
        Private _watched As FolderKey
        Private _dirty As Boolean
        Private _primeiroEventoUtc As DateTimeOffset
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
        ''' Passa a observar outra pasta. A assinatura anterior é cancelada
        ''' ANTES, para que um evento em voo da pasta antiga não seja
        ''' contado como sujeira da nova.
        ''' </summary>
        Public Async Function WatchAsync(folder As FolderKey) As Task
            Await UnwatchAsync()
            If folder Is Nothing Then Return

            _watched = folder
            Dim resultado = Await _broker.SubscribeFolderAsync(folder, CancellationToken.None)
            If resultado.Succeeded Then
                _token = resultado.Value
                _timer.Start()
            End If
        End Function

        Public Async Function UnwatchAsync() As Task
            _timer.Stop()
            _dirty = False
            _watched = Nothing

            Dim anterior = _token
            _token = Nothing
            If anterior Is Nothing Then Return

            Await _broker.UnsubscribeFolderAsync(anterior, CancellationToken.None)
        End Function

        ''' <summary>
        ''' Chega numa thread MTA do pool. Só marca o bit — nada de tocar em
        ''' UI, coleção ou COM daqui.
        ''' </summary>
        Private Sub OnInvalidated(sender As Object, invalidation As FolderInvalidation)
            If _disposed OrElse _token Is Nothing Then Return
            If invalidation.SubscriptionId <> _token.Id Then Return

            _ui.InvokeAsync(
                Sub()
                    If _disposed Then Return
                    If Not _dirty Then _primeiroEventoUtc = DateTimeOffset.UtcNow
                    _dirty = True
                End Sub)
        End Sub

        Private Sub OnTick(sender As Object, e As EventArgs)
            If Not _dirty OrElse _watched Is Nothing Then Return

            Dim desde = (DateTimeOffset.UtcNow - _primeiroEventoUtc).TotalMilliseconds

            ' O teto existe para o caso de eventos chegarem sem parar: sem
            ' ele, uma sincronização longa adiaria a recarga indefinidamente.
            If desde < DebounceMs AndAlso desde < TetoMs Then Return

            _dirty = False
            _onDirty(_watched)
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

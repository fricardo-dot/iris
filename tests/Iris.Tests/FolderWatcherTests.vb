Imports System.Diagnostics
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports Iris.App.ViewModels
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Os dois invariantes do watcher que a máquina pura NÃO cobre, porque não
''' são sobre tempo — são sobre IDENTIDADE de assinatura.
'''
''' O <c>DirtyDebounce</c> prova quando recarregar. Aqui se prova de QUEM é
''' o evento: um aviso da pasta A que ficou na fila do dispatcher não pode
''' sujar a pasta B, e uma assinatura de A que só termina depois de o
''' usuário já ter trocado para B tem de ser desfeita.
'''
''' Os dois já foram defeito neste código, e os dois voltam calados: a lista
''' recarrega quando não devia, ou uma assinatura fica pendurada no Outlook
''' sem ninguém que a use.
''' </summary>
<TestClass>
Public Class FolderWatcherTests

    Private Shared ReadOnly PastaA As New FolderKey("pasta-a", "store-1")
    Private Shared ReadOnly PastaB As New FolderKey("pasta-b", "store-1")

    Private ReadOnly _criados As New List(Of FolderWatcher)()

    <TestCleanup>
    Public Sub Limpar()
        For Each w In _criados
            Try
                w.Dispose()
            Catch
            End Try
        Next
        _criados.Clear()
    End Sub

    Private Function Montar(broker As WatcherBroker, sujas As List(Of FolderKey)) As FolderWatcher
        SynchronizationContext.SetSynchronizationContext(
            New DispatcherSynchronizationContext(Dispatcher.CurrentDispatcher))

        Dim w As New FolderWatcher(broker, Dispatcher.CurrentDispatcher,
                                   Sub(t, nome) Aguardar(t),
                                   Sub(pasta) sujas.Add(pasta))
        _criados.Add(w)
        Return w
    End Function

    ''' <summary>
    ''' PEGA: despacho sem geração.
    '''
    ''' O evento da pasta A é postado no dispatcher e fica na fila. Antes de
    ''' ele rodar, o usuário troca para B. Se o handler não revalidar a
    ''' assinatura JÁ na UI, ele marca B como suja por causa de uma mudança
    ''' que aconteceu em A — e a lista de B recarrega do nada.
    ''' </summary>
    <STATestMethod>
    Public Sub Evento_da_pasta_anterior_nao_suja_a_pasta_nova()
        Dim broker As New WatcherBroker()
        Dim sujas As New List(Of FolderKey)()
        Dim w = Montar(broker, sujas)

        Aguardar(w.WatchAsync(PastaA))
        Dim tokenDeA = broker.UltimoToken()

        ' Evento de A entra na fila do dispatcher...
        broker.Invalidar(tokenDeA, PastaA)

        ' ...e o usuário troca de pasta antes de ele ser processado.
        Aguardar(w.WatchAsync(PastaB))

        ' Agora o dispatcher gira, e o evento velho é processado.
        BombearPor(700)

        CollectionAssert.DoesNotContain(sujas, PastaB,
            "Um evento da pasta anterior marcou a pasta nova como suja.")
        CollectionAssert.DoesNotContain(sujas, PastaA,
            "A pasta anterior já não está sendo observada.")
    End Sub

    ''' <summary>
    ''' Controle positivo: o mesmo caminho, sem troca de pasta, RECARREGA.
    ''' Sem isto, um watcher que ignorasse todo evento passaria no teste de
    ''' cima sem provar nada.
    ''' </summary>
    <STATestMethod>
    Public Sub Evento_da_pasta_observada_suja_a_pasta()
        Dim broker As New WatcherBroker()
        Dim sujas As New List(Of FolderKey)()
        Dim w = Montar(broker, sujas)

        Aguardar(w.WatchAsync(PastaA))
        broker.Invalidar(broker.UltimoToken(), PastaA)

        AguardarSujeira(sujas, 1)
        Assert.AreEqual(PastaA, sujas(0))
    End Sub

    ''' <summary>
    ''' PEGA: assinatura órfã.
    '''
    ''' A assinatura de A demora e só termina depois de o usuário já ter
    ''' pedido B. O token que chega é de uma pasta que ninguém observa mais,
    ''' e se ele for simplesmente guardado, fica uma assinatura viva no
    ''' Outlook sem dono — que continua entregando evento e segurando
    ''' referência COM.
    ''' </summary>
    <STATestMethod>
    Public Sub Assinatura_que_chega_atrasada_e_desfeita()
        Dim broker As New WatcherBroker()
        Dim sujas As New List(Of FolderKey)()
        Dim w = Montar(broker, sujas)

        broker.TravaDoSubscribe = New TaskCompletionSource(Of Boolean)()
        Dim assinandoA = w.WatchAsync(PastaA)
        BombearPor(80)

        ' O usuário troca para B enquanto a assinatura de A ainda não voltou.
        broker.TravaDoSubscribe.SetResult(True)
        broker.TravaDoSubscribe = Nothing
        Dim assinandoB = w.WatchAsync(PastaB)

        Aguardar(assinandoA)
        Aguardar(assinandoB)
        BombearPor(200)

        Assert.IsTrue(broker.Desassinadas.Count > 0,
            "A assinatura que chegou atrasada ficou pendurada no Outlook.")
    End Sub

    ''' <summary>
    ''' Depois da substituição de sessão, o token guardado aponta para uma
    ''' assinatura que já morreu junto com a sessão. Um evento com aquele id
    ''' não pode sujar mais nada.
    ''' </summary>
    <STATestMethod>
    Public Sub Depois_da_troca_de_sessao_o_token_velho_nao_suja_nada()
        Dim broker As New WatcherBroker()
        Dim sujas As New List(Of FolderKey)()
        Dim w = Montar(broker, sujas)

        Aguardar(w.WatchAsync(PastaA))
        Dim tokenVelho = broker.UltimoToken()

        w.OnSessionReplaced()

        broker.Invalidar(tokenVelho, PastaA)
        BombearPor(700)

        Assert.AreEqual(0, sujas.Count,
            "Token de uma sessão que já morreu não pode disparar recarga.")
    End Sub

    ' ================================================================

    Private Shared Sub AguardarSujeira(sujas As List(Of FolderKey), quantas As Integer,
                                       Optional limiteMs As Integer = 5000)
        Dim relogio = Stopwatch.StartNew()
        While sujas.Count < quantas
            If relogio.ElapsedMilliseconds > limiteMs Then
                Assert.Fail($"Esperava {quantas} recarga(s); vieram {sujas.Count}.")
            End If
            Bombear()
        End While
    End Sub

    Private Shared Sub Aguardar(t As Task, Optional limiteMs As Integer = 5000)
        If t Is Nothing Then Return
        Dim relogio = Stopwatch.StartNew()
        While Not t.IsCompleted
            If relogio.ElapsedMilliseconds > limiteMs Then Assert.Fail("A operação não terminou.")
            Bombear()
        End While
        t.GetAwaiter().GetResult()
    End Sub

    Private Shared Sub BombearPor(ms As Integer)
        Dim relogio = Stopwatch.StartNew()
        While relogio.ElapsedMilliseconds < ms
            Bombear()
        End While
    End Sub

    Private Shared Sub Bombear()
        Dim quadro As New DispatcherFrame()
        Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.SystemIdle,
            New Action(Sub() quadro.Continue = False))
        Dispatcher.PushFrame(quadro)
        Thread.Sleep(1)
    End Sub

End Class

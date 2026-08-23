Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Prova que a lista virtualiza de verdade.
'''
''' Configuração não é prova. O XAML pode declarar
''' <c>VirtualizingStackPanel</c>, <c>Recycling</c> e
''' <c>CanContentScroll</c> e ainda assim materializar tudo, porque basta um
''' ScrollViewer externo, um painel errado ou um template que meça altura
''' infinita para anular o conjunto.
'''
''' O teste conta CONTAINERS REALIZADOS. Com 5.000 itens, o número precisa
''' ficar na casa das linhas visíveis mais cache — não na casa dos milhares.
'''
''' Roda sem Outlook: 5.000 DTOs em memória. Isso prova a virtualização do
''' WPF e NÃO prova nada sobre o custo do OOM, que é medição separada.
''' </summary>
<TestClass>
Public Class VirtualizationTests

    Private Const TotalItens As Integer = 5000
    Private Const AlturaJanela As Double = 700

    <STATestMethod>
    Public Sub Lista_de_5000_itens_realiza_poucos_containers()
        Dim realizados = 0
        Dim total = 0

        ExecutarNaUi(
            Sub()
                Dim lista = MontarLista()
                Dim janela = MostrarEm(lista)
                Try
                    BombearAte(DispatcherPriority.ContextIdle)
                    total = lista.Items.Count
                    realizados = ContarRealizados(lista)
                Finally
                    janela.Close()
                End Try
            End Sub)

        Assert.AreEqual(TotalItens, total, "A lista não recebeu os 5.000 itens.")

        ' Uma linha ocupa ~56 DIP; numa janela de 700 cabem ~13. Somando o
        ' cache do WPF, algumas dezenas é o esperado. Centenas significaria
        ' virtualização quebrada, e é isso que este número detecta.
        Assert.IsTrue(realizados < 200,
            $"{realizados} containers realizados de {TotalItens}. " &
            "A virtualização está quebrada: algo no template ou no layout " &
            "está forçando a materialização da lista inteira.")
    End Sub

    ''' <summary>
    ''' CONTROLE NEGATIVO: o mesmo teste, com a virtualização desligada,
    ''' PRECISA reprovar. Sem isto, um contador quebrado que sempre devolve
    ''' zero faria o teste acima passar para sempre sem valer nada.
    ''' </summary>
    <STATestMethod>
    Public Sub Contador_detecta_lista_sem_virtualizacao()
        Dim realizados = 0

        ExecutarNaUi(
            Sub()
                Dim lista = MontarLista()
                ' Desliga a virtualização de propósito.
                VirtualizingPanel.SetIsVirtualizing(lista, False)
                ScrollViewer.SetCanContentScroll(lista, False)

                Dim janela = MostrarEm(lista)
                Try
                    BombearAte(DispatcherPriority.ContextIdle)
                    realizados = ContarRealizados(lista)
                Finally
                    janela.Close()
                End Try
            End Sub)

        Assert.IsTrue(realizados > 1000,
            $"Com a virtualização DESLIGADA só {realizados} containers foram " &
            "realizados. O contador não está medindo o que promete, e o teste " &
            "principal não teria valor.")
    End Sub

    ' ===================================================================

    Private Shared Function MontarLista() As ListBox
        Dim lista As New ListBox()

        ' As mesmas configurações do estilo do produto. Declaradas aqui em
        ' vez de carregar o dicionário para o teste não depender de um
        ' pack URI e continuar valendo se o arquivo mudar de lugar — o que
        ' ele mede é o COMPORTAMENTO, não a origem do estilo.
        ScrollViewer.SetCanContentScroll(lista, True)
        VirtualizingPanel.SetIsVirtualizing(lista, True)
        VirtualizingPanel.SetVirtualizationMode(lista, VirtualizationMode.Recycling)
        VirtualizingPanel.SetScrollUnit(lista, ScrollUnit.Pixel)

        lista.ItemsSource = Enumerable.Range(0, TotalItens).
            Select(Function(i) New MailSummary With {
                .Key = New ItemKey($"entry-{i}", "store"),
                .Subject = $"Mensagem sintética número {i}",
                .SenderName = $"Remetente {i Mod 50}",
                .ReceivedTime = DateTimeOffset.Now.AddMinutes(-i),
                .IsUnread = (i Mod 3 = 0)
            }).ToList()

        Return lista
    End Function

    Private Shared Function MostrarEm(conteudo As UIElement) As Window
        Dim janela As New Window With {
            .Width = 900,
            .Height = AlturaJanela,
            .Content = conteudo,
            .ShowActivated = False,
            .WindowStartupLocation = WindowStartupLocation.Manual,
            .Left = -10000,
            .Top = -10000
        }
        janela.Show()
        janela.UpdateLayout()
        Return janela
    End Function

    Private Shared Function ContarRealizados(lista As ListBox) As Integer
        Dim gerador = lista.ItemContainerGenerator
        Return Enumerable.Range(0, lista.Items.Count).
            Count(Function(i) gerador.ContainerFromIndex(i) IsNot Nothing)
    End Function

    Private Shared Sub BombearAte(prioridade As DispatcherPriority)
        Dim quadro As New DispatcherFrame()
        Dispatcher.CurrentDispatcher.BeginInvoke(prioridade,
            New Action(Sub() quadro.Continue = False))
        Dispatcher.PushFrame(quadro)
    End Sub

    Private Shared Sub ExecutarNaUi(acao As Action)
        acao()
    End Sub

End Class

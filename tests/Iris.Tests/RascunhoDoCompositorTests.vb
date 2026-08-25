Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports Iris.App.ViewModels
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace Global.Iris.Tests

    ''' <summary>
    ''' <b>O elo que faltava: <c>ComposerViewModel</c> → <c>RascunhoDoCompositor</c>.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ESTE ARQUIVO EXISTE</b>
    '''
    ''' A notificação de que o rascunho mudou atravessa três saltos:
    '''
    ''' <code>
    ''' Composer.UserText → RascunhoDoCompositor.Mudou → AssistenteViewModel → botão
    ''' </code>
    '''
    ''' Os dois últimos já tinham prova — inclusive uma que lê o
    ''' <c>Button.IsEnabled</c> do botão real. O <b>primeiro</b> não tinha: todos
    ''' aqueles testes injetam um rascunho de mentira, e por isso remover a
    ''' assinatura de <c>PropertyChanged</c> dentro do adaptador de produção
    ''' deixaria a suíte inteira verde com a ligação quebrada.
    '''
    ''' Aqui o compositor é o de verdade e o adaptador é o de produção.
    ''' </summary>
    <TestClass>
    Public Class RascunhoDoCompositorTests

        Private ReadOnly _criados As New List(Of ComposerViewModel)()

        <TestCleanup>
        Public Sub Limpar()
            For Each vm In _criados
                vm.Dispose()
            Next
            _criados.Clear()
        End Sub

        ''' <summary>
        ''' Um compositor de verdade, sem Outlook.
        '''
        ''' Sem contexto de sincronização o <c>DispatcherTimer</c> explode ao ser
        ''' tocado de fora da STA — o WPF instala isso sozinho no app, e aqui é
        ''' na mão. Mesmo motivo do <c>ComposerTests</c>.
        ''' </summary>
        Private Function Compositor() As ComposerViewModel
            SynchronizationContext.SetSynchronizationContext(
                New DispatcherSynchronizationContext(Dispatcher.CurrentDispatcher))

            Dim vm As New ComposerViewModel(New FakeBroker(), Dispatcher.CurrentDispatcher,
                                            Sub(t, nome)
                                            End Sub,
                                            New FakePickFile(), 40)
            _criados.Add(vm)
            Return vm
        End Function

        ''' <summary>
        ''' <b>Digitar no compositor avisa o rascunho.</b>
        '''
        ''' É este salto que faltava. Sem ele, o botão "Desfazer" continuaria
        ''' habilitado depois de o usuário digitar por cima da redação — o estado
        ''' estaria certo e invisível, que é o defeito da §38.10 inteiro.
        ''' </summary>
        <STATestMethod>
        Public Sub Digitar_no_compositor_AVISA_o_rascunho()
            Dim c = Compositor()
            Dim r As IRascunho = New RascunhoDoCompositor(c)

            Dim avisos = 0
            AddHandler r.Mudou, Sub(remetente As Object, arg As EventArgs) avisos += 1

            c.UserText = "o usuario digitou isto"

            Assert.IsTrue(avisos > 0,
                "o adaptador nao escuta o compositor: a notificacao morre no primeiro salto")
            Assert.AreEqual("o usuario digitou isto", r.Texto)
        End Sub

        ''' <summary>
        ''' Controle: o compositor <b>realmente</b> notificou.
        '''
        ''' Sem isto, um <c>UserText</c> que deixasse de levantar
        ''' <c>PropertyChanged</c> faria o teste de cima falhar apontando para o
        ''' adaptador, que estaria certo. Removendo só a assinatura do
        ''' adaptador, este controle continua verde e o de cima fica vermelho —
        ''' é assim que se sabe qual dos dois quebrou.
        ''' </summary>
        <STATestMethod>
        Public Sub Controle_o_compositor_notifica_UserText()
            Dim c = Compositor()

            Dim vistos As New List(Of String)()
            AddHandler c.PropertyChanged,
                Sub(remetente As Object, arg As ComponentModel.PropertyChangedEventArgs)
                    vistos.Add(arg.PropertyName)
                End Sub

            c.UserText = "o usuario digitou isto"

            CollectionAssert.Contains(vistos, NameOf(ComposerViewModel.UserText))
        End Sub

        ''' <summary>
        ''' <b>O adaptador escreve no compositor, e não numa cópia.</b>
        '''
        ''' A outra direção do mesmo elo: é por ela que a redação da IA chega ao
        ''' texto que o usuário vê.
        ''' </summary>
        <STATestMethod>
        Public Sub Escrever_no_rascunho_chega_ao_compositor()
            Dim c = Compositor()
            Dim r As IRascunho = New RascunhoDoCompositor(c)

            r.Texto = "a redacao da IA"

            Assert.AreEqual("a redacao da IA", c.UserText)
        End Sub

        ''' <summary>
        ''' <b>A sessão e a editabilidade vêm do compositor de verdade.</b>
        '''
        ''' Com o compositor fechado não se escreve nele: <c>PodeEditar</c> é
        ''' falso, e é isso que fecha o botão "Redigir resposta" e o "Desfazer".
        ''' </summary>
        <STATestMethod>
        Public Sub A_sessao_e_a_editabilidade_vem_do_compositor()
            Dim c = Compositor()
            Dim r As IRascunho = New RascunhoDoCompositor(c)

            Assert.IsFalse(c.IsOpen, "o compositor comeca fechado")
            Assert.IsFalse(r.PodeEditar,
                "compositor fechado nao aceita escrita, e o rascunho tem de dizer isso")
            Assert.AreEqual(c.Geracao, r.Sessao)
        End Sub

    End Class

End Namespace

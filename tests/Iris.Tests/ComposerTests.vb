Imports System.Diagnostics
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports Iris.App.ViewModels
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' O compositor é a única parte do Iris que escreve no mundo real. Estes
''' testes cobrem as quatro maneiras de errar que custam caro:
''' perder texto do usuário, gravar com chave vencida, enviar para quem não
''' devia e reenviar uma mensagem que talvez já tenha saído.
'''
''' Rodam sem Outlook, contra um broker de mentira. Não provam que o Object
''' Model se comporta assim — isso é verificação separada, com o Outlook
''' aberto. Provam a lógica do compositor, que é o que dá para provar aqui.
'''
''' Cada teste que afirma um bloqueio vem com o controle negativo do lado:
''' sem ele, um compositor que simplesmente nunca envia passaria em tudo.
''' </summary>
<TestClass>
Public Class ComposerTests

    ''' <summary>
    ''' Debounce curto. O de produção é 1,5 s, e esperar isso a cada teste
    ''' tornaria a suíte lenta e intermitente.
    ''' </summary>
    Private Const DebounceDeTeste As Integer = 40

    Private Shared Function Montar(broker As FakeBroker,
                                   Optional escolha As String = Nothing) As ComposerViewModel
        ' Sem contexto de sincronização, as continuações de Await caem no
        ' pool e o DispatcherTimer explode ao ser tocado de fora da STA. O
        ' WPF faz isto sozinho no app; no teste é na mão.
        SynchronizationContext.SetSynchronizationContext(
            New DispatcherSynchronizationContext(Dispatcher.CurrentDispatcher))

        Dim pick As New FakePickFile With {.Escolha = escolha}
        Return New ComposerViewModel(broker, Dispatcher.CurrentDispatcher,
                                     Sub(t, nome) Aguardar(t), pick, DebounceDeTeste)
    End Function

    ' ================================================================
    ' Abertura
    ' ================================================================

    <STATestMethod>
    Public Sub Abrir_cria_o_rascunho_antes_de_qualquer_digitacao()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)

        Aguardar(vm.NewMessageAsync())

        Assert.AreEqual(ComposerState.Editing, vm.State)
        CollectionAssert.Contains(broker.Chamadas, "create",
            "O rascunho tem de existir no store antes de o usuário digitar.")
        Assert.IsFalse(vm.IsDirty, "Um rascunho recém-criado não tem nada por salvar.")
    End Sub

    ''' <summary>
    ''' Controle negativo da abertura: se criar falha, o compositor NÃO
    ''' abre. Abrir mesmo assim daria uma tela onde tudo parece funcionar e
    ''' nada é gravado.
    ''' </summary>
    <STATestMethod>
    Public Sub Falha_ao_criar_nao_abre_o_compositor()
        Dim broker As New FakeBroker With {.FalhaAoCriar = ErrorKind.NotConnected}
        Dim vm = Montar(broker)

        Aguardar(vm.NewMessageAsync())

        Assert.AreEqual(ComposerState.Closed, vm.State)
        Assert.IsFalse(vm.IsOpen)
        Assert.IsTrue(vm.HasStatus, "O usuário precisa saber por que não abriu.")
        CollectionAssert.DoesNotContain(broker.Chamadas, "update")
    End Sub

    <STATestMethod>
    Public Sub Responder_traz_destinatario_e_citacao_do_Outlook()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)

        Aguardar(vm.ReplyAsync(New ItemKey("msg-1", "store-1"), replyAll:=False))

        Assert.AreEqual("alguem@exemplo.com", vm.ToLine)
        Assert.AreEqual("RE: original", vm.Subject)
        Assert.IsTrue(vm.HasQuoted, "A citação do Outlook tem de aparecer.")

        ' Carregar o rascunho na tela NÃO é o usuário editando. Se contasse
        ' como edição, abrir uma resposta já gravaria de volta o que
        ' acabou de ler.
        Assert.IsFalse(vm.IsDirty)
        CollectionAssert.DoesNotContain(broker.Chamadas, "update")
    End Sub

    ' ================================================================
    ' Autosave
    ' ================================================================

    <STATestMethod>
    Public Sub Digitar_varias_teclas_grava_uma_vez_so()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        ' Três "teclas" sem deixar o dispatcher respirar entre elas: é o
        ' que o debounce existe para colapsar.
        vm.UserText = "o"
        vm.UserText = "ol"
        vm.UserText = "olá"

        AguardarChamadas(broker, "update", 1)
        Assert.AreEqual(1, ContarChamadas(broker, "update"),
            "Uma rajada de teclas tem de virar uma gravação, não três.")
        Assert.AreEqual("olá", broker.Gravacoes.Last().UserText)
    End Sub

    ''' <summary>
    ''' Controle negativo do debounce: pausas de verdade entre as edições
    ''' PRECISAM gerar gravações separadas. Sem isto, um compositor que
    ''' nunca grava passaria no teste de cima.
    ''' </summary>
    <STATestMethod>
    Public Sub Edicoes_espacadas_geram_gravacoes_separadas()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.UserText = "um"
        AguardarChamadas(broker, "update", 1)

        vm.UserText = "dois"
        AguardarChamadas(broker, "update", 2)

        Assert.AreEqual(2, ContarChamadas(broker, "update"))
        Assert.AreEqual("dois", broker.Gravacoes.Last().UserText)
    End Sub

    ''' <summary>
    ''' O caso que um Boolean de "sujo" erraria: o usuário digita ENQUANTO a
    ''' gravação anterior está em voo. O sucesso dessa gravação não pode
    ''' limpar a marca — se limpasse, o texto novo nunca seria gravado.
    ''' </summary>
    <STATestMethod>
    Public Sub Texto_digitado_durante_a_gravacao_nao_se_perde()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        broker.TravaDoUpdate = New TaskCompletionSource(Of Boolean)()

        vm.UserText = "primeiro"
        Dim fechando = vm.SaveAndCloseCommand.ExecuteAsync(Nothing)
        AguardarChamadas(broker, "update", 1)

        ' Gravação presa; o usuário continua digitando.
        vm.UserText = "segundo"

        broker.TravaDoUpdate.SetResult(True)
        broker.TravaDoUpdate = Nothing
        Aguardar(fechando)

        Assert.AreEqual("segundo", broker.Gravacoes.Last().UserText,
            "A última gravação tem de conter o que estava na tela, não a versão antiga.")
        Assert.IsFalse(vm.IsDirty)
        Assert.AreEqual(ComposerState.Closed, vm.State)
    End Sub

    ''' <summary>
    ''' O EntryID muda a cada Save. Guardar a chave do começo daria NotFound
    ''' na hora de enviar — e o usuário descobriria isso já tendo confirmado
    ''' o envio.
    ''' </summary>
    <STATestMethod>
    Public Sub A_chave_do_rascunho_e_relida_a_cada_gravacao()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        Dim chaveInicial = broker.ChaveAtual()

        vm.UserText = "texto"
        AguardarChamadas(broker, "update", 1)

        ' Controle negativo embutido: o duplo REALMENTE troca a chave a cada
        ' gravação. Sem esta asserção, o teste passaria contra um broker que
        ' devolvesse sempre a mesma chave, provando nada.
        Assert.AreNotEqual(chaveInicial, broker.ChaveAtual(),
            "O duplo precisa girar a chave, senão o teste não prova nada.")

        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))

        Dim usadaNoPrepare = broker.ChavesRecebidas.Last()
        Assert.AreEqual(broker.ChaveAtual(), usadaNoPrepare,
            "Conferir o envio tem de usar a chave mais recente.")
        Assert.AreNotEqual(chaveInicial, usadaNoPrepare,
            "Usar a chave do começo daria NotFound no envio.")
    End Sub

    ' ================================================================
    ' Envio
    ' ================================================================

    ''' <summary>
    ''' A confirmação precisa descrever o que vai sair AGORA. Se o
    ''' compositor conferisse sem gravar antes, o usuário aprovaria uma
    ''' versão e o Outlook mandaria outra.
    ''' </summary>
    <STATestMethod>
    Public Sub Conferir_o_envio_grava_o_que_esta_na_tela_antes_de_perguntar()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.UserText = "texto que ainda não foi gravado"
        ' Sem esperar o debounce de propósito: é justamente o caso em que o
        ' usuário digita e clica em enviar na sequência.
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))

        Dim ultimoUpdate = broker.Chamadas.LastIndexOf("update")
        Dim prepare = broker.Chamadas.IndexOf("prepare")

        Assert.IsTrue(ultimoUpdate >= 0, "Tinha texto por gravar; a gravação tem de acontecer.")
        Assert.IsTrue(prepare > ultimoUpdate,
            "Gravar tem de vir ANTES de conferir, senão a confirmação descreve outra versão.")
        Assert.AreEqual("texto que ainda não foi gravado", broker.Gravacoes.Last().UserText)
    End Sub

    <STATestMethod>
    Public Sub Destinatario_resolvido_leva_a_confirmacao()
        Dim broker As New FakeBroker With {.TodosResolvidos = True}
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))

        Assert.AreEqual(ComposerState.ConfirmingSend, vm.State)
        Assert.AreEqual("eu@empresa.com", vm.PreviewAccount,
            "A confirmação tem de dizer por qual conta vai sair.")
        Assert.AreEqual(1, vm.PreviewRecipients.Count())
        CollectionAssert.DoesNotContain(broker.Chamadas, "send",
            "Conferir não envia. Enviar só depois do segundo clique.")
    End Sub

    ''' <summary>
    ''' Um nome que o Outlook não reconheceu pode virar endereço errado, e
    ''' não existe desfazer. Bloqueia.
    ''' </summary>
    <STATestMethod>
    Public Sub Destinatario_nao_resolvido_bloqueia_o_envio()
        Dim broker As New FakeBroker With {.TodosResolvidos = False}
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))

        Assert.AreEqual(ComposerState.Editing, vm.State,
            "Sem todos resolvidos, não chega nem à tela de confirmação.")
        Assert.IsTrue(vm.HasStatus, "O usuário precisa saber quem não foi reconhecido.")
        CollectionAssert.DoesNotContain(broker.Chamadas, "send")
    End Sub

    <STATestMethod>
    Public Sub Envio_bem_sucedido_fecha_o_compositor()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))
        Aguardar(vm.ConfirmSendCommand.ExecuteAsync(Nothing))

        Assert.AreEqual(ComposerState.Closed, vm.State)
        Assert.AreEqual(1, ContarChamadas(broker, "send"))
    End Sub

    ''' <summary>
    ''' O erro irreversível deste projeto. Envio ambíguo é terminal: o
    ''' compositor guarda o texto, mostra o aviso e NÃO oferece enviar de
    ''' novo. Repetir poderia mandar a mesma mensagem duas vezes.
    ''' </summary>
    <STATestMethod>
    Public Sub Envio_ambiguo_nao_permite_reenviar()
        Dim broker As New FakeBroker With {.ResultadoDoEnvio = ErrorKind.Ambiguous}
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        vm.UserText = "conteúdo importante"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))
        Aguardar(vm.ConfirmSendCommand.ExecuteAsync(Nothing))

        Assert.AreEqual(ComposerState.SendUnknown, vm.State)
        Assert.IsTrue(vm.IsOpen, "O texto do usuário não pode sumir junto com o erro.")
        Assert.AreEqual("conteúdo importante", vm.UserText)
        Assert.IsFalse(vm.RequestSendCommand.CanExecute(Nothing),
            "O botão de enviar não pode voltar depois de um envio ambíguo.")

        ' E se alguém chamar o comando por fora assim mesmo, nada sai.
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))
        Assert.AreEqual(1, ContarChamadas(broker, "send"),
            "Nem por caminho indireto o envio pode acontecer duas vezes.")
    End Sub

    ''' <summary>
    ''' Controle negativo do bloqueio acima: uma falha CONHECIDA — em que se
    ''' sabe que nada saiu — devolve o compositor à edição, com o botão de
    ''' enviar de volta. Sem este teste, um compositor que travasse em
    ''' qualquer erro passaria como se estivesse correto.
    ''' </summary>
    <STATestMethod>
    Public Sub Falha_conhecida_no_envio_devolve_o_compositor_a_edicao()
        Dim broker As New FakeBroker With {.ResultadoDoEnvio = ErrorKind.NotConnected}
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))
        Aguardar(vm.ConfirmSendCommand.ExecuteAsync(Nothing))

        Assert.AreEqual(ComposerState.Editing, vm.State)
        Assert.IsTrue(vm.RequestSendCommand.CanExecute(Nothing),
            "Falha em que se sabe que nada saiu pode ser tentada de novo.")
    End Sub

    ''' <summary>
    ''' Se a gravação falha, conferir não pode seguir. Seguir mostraria uma
    ''' confirmação da versão ANTIGA e enviaria essa versão — o usuário
    ''' aprovaria um texto e sairia outro.
    ''' </summary>
    <STATestMethod>
    Public Sub Gravacao_falhando_nao_chega_a_confirmacao_de_envio()
        Dim broker As New FakeBroker With {.FalhaAoGravar = ErrorKind.NotConnected}
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        vm.UserText = "texto"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))

        Assert.AreEqual(ComposerState.Editing, vm.State)
        Assert.IsTrue(vm.IsDirty)
        Assert.IsTrue(vm.HasStatus, "Botão que não faz nada e não explica é pior que botão desabilitado.")
        CollectionAssert.DoesNotContain(broker.Chamadas, "prepare",
            "Nem conferir: a confirmação descreveria a versão errada.")
        CollectionAssert.DoesNotContain(broker.Chamadas, "send")
    End Sub

    ' ================================================================
    ' Fechamento
    ' ================================================================

    <STATestMethod>
    Public Sub Fechar_com_alteracao_pendente_pergunta_em_vez_de_descartar()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.UserText = "algo que o usuário escreveu"
        vm.CloseCommand.Execute(Nothing)

        Assert.AreEqual(ComposerState.ConfirmingClose, vm.State)
        Assert.IsTrue(vm.IsOpen)
        CollectionAssert.DoesNotContain(broker.Chamadas, "delete",
            "Fechar não pode apagar o rascunho por conta própria.")
    End Sub

    ''' <summary>
    ''' Controle negativo: sem nada pendente, fechar fecha. Sem isto, um
    ''' compositor que perguntasse SEMPRE passaria no teste de cima.
    ''' </summary>
    <STATestMethod>
    Public Sub Fechar_sem_alteracao_pendente_fecha_direto()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.CloseCommand.Execute(Nothing)

        Assert.AreEqual(ComposerState.Closed, vm.State)
        Assert.IsFalse(vm.IsOpen)
    End Sub

    <STATestMethod>
    Public Sub Descartar_apaga_o_rascunho_e_fecha()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.UserText = "texto"
        vm.CloseCommand.Execute(Nothing)
        Aguardar(vm.DiscardCommand.ExecuteAsync(Nothing))

        Assert.AreEqual(ComposerState.Closed, vm.State)
        Assert.AreEqual(1, ContarChamadas(broker, "delete"))
    End Sub

    ''' <summary>
    ''' Se a gravação falha, fechar seria perder o texto justamente no caso
    ''' em que ele não está guardado em lugar nenhum.
    ''' </summary>
    <STATestMethod>
    Public Sub Salvar_e_fechar_nao_fecha_quando_a_gravacao_falha()
        Dim broker As New FakeBroker With {.FalhaAoGravar = ErrorKind.NotConnected}
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.UserText = "texto que não pode sumir"
        vm.CloseCommand.Execute(Nothing)
        Aguardar(vm.SaveAndCloseCommand.ExecuteAsync(Nothing))

        Assert.IsTrue(vm.IsOpen, "Com a gravação falhando, fechar apagaria o texto.")
        Assert.AreEqual("texto que não pode sumir", vm.UserText)
        Assert.IsTrue(vm.IsDirty)
    End Sub

    ' ================================================================
    ' Bombeamento
    ' ================================================================

    Private Shared Function ContarChamadas(broker As FakeBroker, nome As String) As Integer
        ' Where().Count() e nao Count(predicado): List(Of T) tem uma
        ' PROPRIEDADE Count, e ela eclipsa a extensao do LINQ de mesmo nome.
        Return broker.Chamadas.Where(Function(c) c = nome).Count()
    End Function

    ''' <summary>
    ''' Bombeia o dispatcher até a operação terminar. Bloquear a STA com
    ''' Wait() prenderia justamente a fila que precisa girar para a
    ''' continuação rodar — e o teste travaria em vez de falhar.
    ''' </summary>
    Private Shared Sub Aguardar(t As Task, Optional limiteMs As Integer = 5000)
        If t Is Nothing Then Return

        Dim relogio = Stopwatch.StartNew()
        While Not t.IsCompleted
            If relogio.ElapsedMilliseconds > limiteMs Then
                Assert.Fail($"A operação não terminou em {limiteMs} ms.")
            End If
            Bombear()
        End While

        ' Propaga a exceção, se houve. Engoli-la faria um teste verde em
        ' cima de um compositor que estourou.
        t.GetAwaiter().GetResult()
    End Sub

    Private Shared Sub AguardarChamadas(broker As FakeBroker, nome As String,
                                        quantas As Integer, Optional limiteMs As Integer = 5000)
        Dim relogio = Stopwatch.StartNew()
        While ContarChamadas(broker, nome) < quantas
            If relogio.ElapsedMilliseconds > limiteMs Then
                Assert.Fail($"Esperava {quantas} chamada(s) de '{nome}' em {limiteMs} ms; " &
                            $"vieram {ContarChamadas(broker, nome)}.")
            End If
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

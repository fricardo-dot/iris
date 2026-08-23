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

    Private ReadOnly _criados As New List(Of ComposerViewModel)()
    Private ReadOnly _temporarios As New List(Of String)()

    ''' <summary>
    ''' Descarta os compositores do teste.
    '''
    ''' Sem isto, o DispatcherTimer de um compositor sobrevive ao teste que
    ''' o criou. O STA e o dispatcher sao os MESMOS entre os testes da
    ''' classe, entao o Bombear() de um teste seguinte dispara o autosave
    ''' do compositor anterior — e um teste passa sozinho e falha na suite,
    ''' que e a pior forma de falhar.
    ''' </summary>
    <TestCleanup>
    Public Sub Limpar()
        For Each vm In _criados
            vm.Dispose()
        Next
        _criados.Clear()

        For Each caminho In _temporarios
            Try
                IO.File.Delete(caminho)
            Catch
                ' Nao e problema do teste se o arquivo continuar preso.
            End Try
        Next
        _temporarios.Clear()
    End Sub

    Private Function Montar(broker As FakeBroker,
                            Optional escolha As String = Nothing) As ComposerViewModel
        ' Sem contexto de sincronização, as continuações de Await caem no
        ' pool e o DispatcherTimer explode ao ser tocado de fora da STA. O
        ' WPF faz isto sozinho no app; no teste é na mão.
        SynchronizationContext.SetSynchronizationContext(
            New DispatcherSynchronizationContext(Dispatcher.CurrentDispatcher))

        Dim pick As New FakePickFile With {.Escolha = escolha}
        Dim vm As New ComposerViewModel(broker, Dispatcher.CurrentDispatcher,
                                        Sub(t, nome) Aguardar(t), pick, DebounceDeTeste)
        _criados.Add(vm)
        Return vm
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

        vm.ToLine = "fulano@empresa.com"
        vm.Subject = "assunto conferível"
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

        ' Ordem de chamadas não basta como prova. O duplo monta a prévia a
        ' partir do que está GRAVADO, então conferir o assunto exibido prova
        ' que a confirmação descreve a versão salva, e não uma anterior.
        Assert.AreEqual(ComposerState.ConfirmingSend, vm.State)
        Assert.AreEqual("assunto conferível", vm.Preview.Subject)
    End Sub

    <STATestMethod>
    Public Sub Destinatario_resolvido_leva_a_confirmacao()
        Dim broker As New FakeBroker With {.Modo = ModoDeDestinatario.Smtp}
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
        Dim broker As New FakeBroker With {.Modo = ModoDeDestinatario.NaoResolvido}
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


    ''' <summary>
    ''' A janela que sobrou depois de eu corrigir a descarga: o usuário
    ''' digita DEPOIS da gravação e ANTES de a prévia ficar pronta. A prévia
    ''' descreve a versão salva; a tela já mostra outra. Aprovar ali seria
    ''' aprovar um texto e mandar outro — e, pior, o texto novo morreria no
    ''' Encerrar() depois do envio.
    ''' </summary>
    <STATestMethod>
    Public Sub Digitar_enquanto_o_envio_e_conferido_invalida_a_confirmacao()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        vm.UserText = "versão aprovada"

        broker.TravaDoPrepare = New TaskCompletionSource(Of Boolean)()
        Dim conferindo = vm.RequestSendCommand.ExecuteAsync(Nothing)
        AguardarChamadas(broker, "prepare", 1)

        ' Prévia presa; o usuário continua digitando.
        vm.UserText = "versão nova, que a prévia não viu"

        broker.TravaDoPrepare.SetResult(True)
        broker.TravaDoPrepare = Nothing
        Aguardar(conferindo)

        Assert.AreNotEqual(ComposerState.ConfirmingSend, vm.State,
            "A prévia está vencida; mostrá-la faria o usuário aprovar outra versão.")
        Assert.IsTrue(vm.HasStatus)
        Assert.AreEqual("versão nova, que a prévia não viu", vm.UserText,
            "O texto digitado durante a conferência não pode sumir.")
        CollectionAssert.DoesNotContain(broker.Chamadas, "send")
    End Sub

    ''' <summary>
    ''' Anexar SALVA, e todo Save gira o EntryID. O compositor precisa
    ''' instalar a chave nova: com a velha, o próximo autosave ou o envio
    ''' devolveria NotFound.
    ''' </summary>
    <STATestMethod>
    Public Sub Anexar_instala_a_chave_nova_do_rascunho()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker, escolha:=CaminhoDeArquivoReal())
        Aguardar(vm.NewMessageAsync())

        Dim chaveAntesDoAnexo = broker.ChaveAtual()
        Aguardar(vm.AttachCommand.ExecuteAsync(Nothing))
        Assert.AreEqual(1, vm.Attachments.Count)

        ' O anexo salvou, então o duplo já girou a chave. Sem isso o teste
        ' não provaria nada — provaria só que a chave nunca muda.
        Dim chaveDepoisDoAnexo = broker.ChaveAtual()
        Assert.AreNotEqual(chaveAntesDoAnexo, chaveDepoisDoAnexo)

        ' Se o compositor tivesse ficado com a chave velha, esta gravação
        ' voltaria NotFound e o rascunho continuaria sujo.
        vm.UserText = "depois de anexar"
        AguardarChamadas(broker, "update", 1)

        Assert.AreEqual(chaveDepoisDoAnexo, broker.ChavesRecebidas.Last(),
            "A gravação seguinte usou a chave anterior ao anexo.")
        Assert.IsFalse(vm.IsDirty,
            "A gravação depois de anexar falhou — sinal de chave vencida.")
    End Sub

    ''' <summary>
    ''' O caso que engana: o Outlook diz RESOLVIDO e entrega um /O=..., que
    ''' ninguém consegue conferir. A tela existe para ser conferida, então
    ''' isso bloqueia — mesmo com Resolved = True.
    ''' </summary>
    <STATestMethod>
    Public Sub Endereco_Exchange_legado_bloqueia_mesmo_dizendo_resolvido()
        Dim broker As New FakeBroker With {.Modo = ModoDeDestinatario.ExchangeLegado}
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))

        Assert.AreEqual(ComposerState.Editing, vm.State,
            "/O=... não é endereço que o usuário possa conferir.")
        Assert.IsTrue(vm.HasStatus)
        CollectionAssert.DoesNotContain(broker.Chamadas, "send")
    End Sub

    ''' <summary>
    ''' Duas confirmações quase simultâneas. É o controle negativo que
    ''' faltava: o teste de ambiguidade provava que não se REABRE o fluxo,
    ''' não que dois cliques rápidos mandam uma vez só.
    ''' </summary>
    <STATestMethod>
    Public Sub Duas_confirmacoes_simultaneas_enviam_uma_vez_so()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))
        Assert.AreEqual(ComposerState.ConfirmingSend, vm.State)

        ' Segura o envio para que a segunda chamada caia dentro da janela em
        ' que a primeira ainda não terminou.
        broker.TravaDoSend = New TaskCompletionSource(Of Boolean)()

        Dim primeiro = vm.ConfirmSendCommand.ExecuteAsync(Nothing)
        AguardarChamadas(broker, "send", 1)
        Dim segundo = vm.ConfirmSendCommand.ExecuteAsync(Nothing)

        broker.TravaDoSend.SetResult(True)
        broker.TravaDoSend = Nothing
        Aguardar(primeiro)
        Aguardar(segundo)

        Assert.AreEqual(1, ContarChamadas(broker, "send"),
            "Dois cliques não podem virar duas mensagens.")
    End Sub

    ''' <summary>
    ''' A decisão da seção 12 promete anexos na confirmação. Mandar o anexo
    ''' errado para fora é tão irreversível quanto mandar para a pessoa
    ''' errada.
    ''' </summary>
    <STATestMethod>
    Public Sub A_confirmacao_mostra_os_anexos()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker, escolha:=CaminhoDeArquivoReal())
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        Aguardar(vm.AttachCommand.ExecuteAsync(Nothing))
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))

        Assert.AreEqual(ComposerState.ConfirmingSend, vm.State)
        Assert.IsTrue(vm.HasPreviewAttachments)
        Assert.AreEqual(1, vm.PreviewAttachments.Count())
    End Sub

    ''' <summary>
    ''' Controle negativo do teste acima: sem anexo, a confirmação não
    ''' inventa nenhum. Sem isto, uma propriedade que devolvesse sempre algo
    ''' passaria.
    ''' </summary>
    <STATestMethod>
    Public Sub Sem_anexo_a_confirmacao_nao_mostra_nenhum()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))

        Assert.AreEqual(ComposerState.ConfirmingSend, vm.State)
        Assert.IsFalse(vm.HasPreviewAttachments)
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
    ' Ciclo de vida e estados de envio
    ' ================================================================

    ''' <summary>
    ''' Depois de um envio AMBIGUO o rascunho e a evidencia: o usuario vai
    ''' compara-lo com os Itens Enviados para decidir se a mensagem saiu.
    ''' Gravar por cima destroi exatamente o que ele precisa.
    '''
    ''' A edicao continua sendo REGISTRADA — o texto esta na tela e o
    ''' compositor sabe que ha coisa por salvar. So nao grava sozinho.
    ''' </summary>
    <STATestMethod>
    Public Sub Depois_de_envio_ambiguo_o_rascunho_nao_e_regravado()
        Dim broker As New FakeBroker With {.ResultadoDoEnvio = ErrorKind.Ambiguous}
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))
        Aguardar(vm.ConfirmSendCommand.ExecuteAsync(Nothing))
        Assert.AreEqual(ComposerState.SendUnknown, vm.State)

        Dim gravacoesAntes = ContarChamadas(broker, "update")
        vm.UserText = "mexido depois do envio ambiguo"

        ' Bem mais que o debounce: se fosse armar, ja teria gravado.
        BombearPor(DebounceDeTeste * 6)

        Assert.AreEqual(gravacoesAntes, ContarChamadas(broker, "update"),
            "Gravar por cima do rascunho ambiguo apaga a evidencia da reconciliacao.")
        Assert.IsTrue(vm.IsDirty, "A edicao tem de ser registrada mesmo sem gravar.")
        Assert.AreEqual("mexido depois do envio ambiguo", vm.UserText)
    End Sub

    ''' <summary>
    ''' Controle negativo do teste acima: no estado de edicao, a MESMA
    ''' sequencia grava. Sem isto, um autosave quebrado passaria.
    ''' </summary>
    <STATestMethod>
    Public Sub Editando_a_mesma_edicao_grava()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        Dim gravacoesAntes = ContarChamadas(broker, "update")
        vm.UserText = "mexido durante a edicao"
        BombearPor(DebounceDeTeste * 6)

        Assert.IsTrue(ContarChamadas(broker, "update") > gravacoesAntes)
        Assert.IsFalse(vm.IsDirty)
    End Sub

    ''' <summary>
    ''' Fechar a janela durante o envio nao pode desmontar o compositor: a
    ''' continuacao do Send voltaria para um objeto sem chave e sem estado,
    ''' e um resultado AMBIGUO nao teria para quem ser contado. E o unico
    ''' momento do Iris em que a resposta certa e "espere".
    ''' </summary>
    <STATestMethod>
    Public Sub Fechar_a_janela_durante_o_envio_e_recusado()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))

        broker.TravaDoSend = New TaskCompletionSource(Of Boolean)()
        Dim enviando = vm.ConfirmSendCommand.ExecuteAsync(Nothing)
        AguardarChamadas(broker, "send", 1)
        Assert.AreEqual(ComposerState.Sending, vm.State)

        Assert.IsFalse(vm.RequestCloseFromWindow(), "Fechar durante o envio desmontaria o compositor.")
        Assert.IsTrue(vm.IsOpen)

        broker.TravaDoSend.SetResult(True)
        broker.TravaDoSend = Nothing
        Aguardar(enviando)
    End Sub

    ''' <summary>
    ''' Controle negativo: parado e sem sujeira, a janela fecha.
    ''' </summary>
    <STATestMethod>
    Public Sub Fechar_a_janela_com_compositor_limpo_e_permitido()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        Assert.IsTrue(vm.RequestCloseFromWindow())
        Assert.IsFalse(vm.IsOpen)
    End Sub

    ''' <summary>
    ''' O compositor foi encerrado enquanto uma operacao estava em voo. A
    ''' continuacao NAO pode escrever o resultado: instalaria previa, chave
    ''' e estado num rascunho que ja nao existe — o compositor voltaria
    ''' sozinho para a tela de confirmacao depois de fechado.
    ''' </summary>
    <STATestMethod>
    Public Sub Resultado_que_volta_depois_do_fechamento_e_descartado()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"

        broker.TravaDoPrepare = New TaskCompletionSource(Of Boolean)()
        Dim conferindo = vm.RequestSendCommand.ExecuteAsync(Nothing)
        AguardarChamadas(broker, "prepare", 1)

        ' Fecha por baixo, com a previa ainda em voo.
        vm.CloseCommand.Execute(Nothing)
        Assert.AreEqual(ComposerState.Closed, vm.State)

        broker.TravaDoPrepare.SetResult(True)
        broker.TravaDoPrepare = Nothing
        Aguardar(conferindo)

        Assert.AreEqual(ComposerState.Closed, vm.State,
            "O compositor nao pode reabrir sozinho depois de fechado.")
        Assert.IsNull(vm.Preview)
    End Sub


    ''' <summary>
    ''' Anexar DESCARREGA e depois anexa. Se a trava cobrisse so a descarga,
    ''' outro comando entraria nessa fresta, pegaria a chave, e o anexo
    ''' giraria o EntryID debaixo dele — o duplo devolve NotFound para chave
    ''' vencida, entao a corrida vira falha visivel.
    '''
    ''' Este teste dispara anexar e conferir o envio ao mesmo tempo, com a
    ''' gravacao presa no meio para garantir a sobreposicao.
    ''' </summary>
    <STATestMethod>
    Public Sub Anexar_e_conferir_ao_mesmo_tempo_nao_disputam_a_chave()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker, escolha:=CaminhoDeArquivoReal())
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        vm.UserText = "texto por gravar"

        ' Prende a gravacao para que a segunda operacao chegue enquanto a
        ' primeira ainda esta no meio do caminho.
        broker.TravaDoUpdate = New TaskCompletionSource(Of Boolean)()

        Dim anexando = vm.AttachCommand.ExecuteAsync(Nothing)
        AguardarChamadas(broker, "update", 1)

        Dim conferindo = vm.RequestSendCommand.ExecuteAsync(Nothing)

        broker.TravaDoUpdate.SetResult(True)
        broker.TravaDoUpdate = Nothing
        Aguardar(anexando)
        Aguardar(conferindo)

        ' Com a chave disputada, o duplo teria devolvido NotFound e nada
        ' disto valeria.
        Assert.AreEqual(1, vm.Attachments.Count, "O anexo nao entrou.")
        Assert.AreEqual(ComposerState.ConfirmingSend, vm.State,
            "A conferencia falhou — sinal de chave vencida entre as duas operacoes.")
        Assert.AreEqual(broker.ChaveAtual(), broker.ChavesRecebidas.Last(),
            "A ultima operacao usou uma chave que ja nao valia.")
        Assert.IsTrue(vm.HasPreviewAttachments, "O anexo tem de aparecer na confirmacao.")
    End Sub

    ''' <summary>
    ''' Uma tecla que escapou para dentro da confirmacao deixa o rascunho
    ''' sujo. Voltar para a edicao sem rearmar o timer faria esse texto
    ''' esperar a proxima tecla para ser gravado — e ela pode nao vir.
    ''' </summary>
    <STATestMethod>
    Public Sub Voltar_da_confirmacao_com_texto_pendente_volta_a_gravar()
        Dim broker As New FakeBroker()
        Dim vm = Montar(broker)
        Aguardar(vm.NewMessageAsync())

        vm.ToLine = "fulano@empresa.com"
        Aguardar(vm.RequestSendCommand.ExecuteAsync(Nothing))
        Assert.AreEqual(ComposerState.ConfirmingSend, vm.State)

        ' Tecla que escapou: registra, mas nao arma o timer neste estado.
        vm.UserText = "escapou para dentro da confirmacao"
        Assert.IsTrue(vm.IsDirty)

        Dim gravacoesAntes = ContarChamadas(broker, "update")
        vm.CancelSendCommand.Execute(Nothing)
        BombearPor(DebounceDeTeste * 6)

        Assert.IsTrue(ContarChamadas(broker, "update") > gravacoesAntes,
            "Voltar a editar tem de rearmar o autosave do que ficou pendente.")
        Assert.IsFalse(vm.IsDirty)
    End Sub

    ' ================================================================
    ' Bombeamento
    ' ================================================================

    ''' <summary>
    ''' Um arquivo que existe de verdade. O compositor entrega o caminho ao
    ''' broker sem conferir; quem confere é a camada COM. Aqui basta ser um
    ''' caminho plausível.
    '''
    ''' Nome ÚNICO por chamada. Um nome fixo compartilhado entre os testes
    ''' dava IOException intermitente — "o arquivo está sendo usado por
    ''' outro processo" — em cerca de uma execução a cada dezesseis. Bastava
    ''' o indexador ou o antivírus estar com o handle aberto do teste
    ''' anterior. Um teste que falha de vez em quando acaba sendo ignorado,
    ''' que é o mesmo que não ter teste.
    ''' </summary>
    Private Function CaminhoDeArquivoReal() As String
        Dim caminho = IO.Path.Combine(IO.Path.GetTempPath(),
                                      $"iris-teste-anexo-{Guid.NewGuid():N}.txt")
        IO.File.WriteAllText(caminho, "anexo de teste")
        _temporarios.Add(caminho)
        Return caminho
    End Function

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

    ''' <summary>
    ''' Bombeia por um tempo, sem esperar condicao nenhuma. Serve para
    ''' provar que algo NAO acontece: dar tempo de sobra e ver que nada
    ''' veio.
    ''' </summary>
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

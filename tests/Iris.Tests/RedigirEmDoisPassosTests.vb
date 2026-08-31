Imports System.Collections.Generic
Imports System.Threading.Tasks
Imports Iris.App.ViewModels
Imports Iris.Assist
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>REDIGIR MOSTRA; ENVIAR PARA RASCUNHO APLICA.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTAVA ERRADO</b>
'''
''' O botão dizia "Redigir resposta" e fazia <b>duas</b> coisas: pedia o texto
''' à IA e o escrevia dentro do rascunho. Ninguém tinha pedido a segunda.
'''
''' E ele escrevia o resultado no <b>mesmo lugar</b> do resumo — então redigir
''' apagava da tela o resumo que o usuário acabou de pagar.
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTES TESTES PRENDEM</b>
'''
''' <list type="number">
''' <item><b>Redigir não toca no rascunho.</b> É o teste que impede o botão de
''' voltar a fazer duas coisas.</item>
''' <item><b>Redigir não apaga o resumo.</b> Os dois quadros coexistem.</item>
''' <item><b>Redigir exige o resumo</b>, e leva o resumo no pedido — é o que
''' torna a segunda chamada diferente da primeira em vez de repetida.</item>
''' <item><b>Enviar para rascunho aplica</b>, e o desfazer continua
''' funcionando depois dele.</item>
''' </list>
''' </summary>
<TestClass>
Public Class RedigirEmDoisPassosTests

    ''' <summary>
    ''' Guarda a instrução que chegou ao envelope, para o teste poder conferir
    ''' que o resumo viajou junto.
    ''' </summary>
    Private NotInheritable Class ContextoQueLembra
        Implements IAssistContext

        Friend UltimaInstrucao As String = ""

        Public Function Pedido(operacao As AssistOperation) As PreflightRequest _
                               Implements IAssistContext.Pedido
            Return AssistenteViewModelTests.Voo()
        End Function

        Public Function Classificar() As IReadOnlyList(Of MessageClassification) _
                                        Implements IAssistContext.Classificar
            Return {AssistenteViewModelTests.Classificada(1)}
        End Function

        Public Function Montar(operacao As AssistOperation, instrucao As String) _
                               As EnvelopeResult Implements IAssistContext.Montar
            UltimaInstrucao = instrucao
            Return New EnvelopeBuilder().Montar(operacao, instrucao,
                                                {AssistenteViewModelTests.Preparada(1)})
        End Function
    End Class

    ' ==================================================================

    ''' <summary>
    ''' <b>Redigir NÃO toca no rascunho.</b>
    '''
    ''' O teste que impede o botão de voltar a fazer duas coisas.
    '''
    ''' <b>Controle negativo:</b> pondo o <c>_rascunho.Texto = Resposta</c> de
    ''' volta no fim do <c>Executar</c>, este teste cai.
    ''' </summary>
    <TestMethod>
    Public Async Function Redigir_NAO_escreve_no_rascunho() As Task
        Dim rascunho As New AssistenteViewModelTests.RascunhoFalso()
        rascunho.Texto = "o que eu digitei"
        Dim vm = AssistenteViewModelTests.Montar(
            AssistenteViewModelTests.Ativacao(),
            New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "texto da IA"},
            AssistenteViewModelTests.Pronta(), rascunho:=rascunho)

        Await vm.Resumir()
        Assert.IsTrue(vm.TemResultado, "controle: o resumo tinha de sair")

        Await vm.Redigir()

        Assert.IsTrue(vm.TemResposta, "a redação não produziu resposta")
        Assert.AreEqual("o que eu digitei", rascunho.Texto,
            "redigir escreveu no rascunho por conta própria. O botão diz " &
            "'redigir'; aplicar é outro ato.")
    End Function

    ''' <summary>
    ''' <b>Redigir NÃO apaga o resumo.</b>
    '''
    ''' As duas operações publicam pelo mesmo caminho interno, e ele escreve em
    ''' <c>Resultado</c>. Sem o cuidado de devolver, a redação apagava da tela o
    ''' resumo que o usuário acabou de pagar.
    ''' </summary>
    <TestMethod>
    Public Async Function Redigir_NAO_apaga_o_resumo() As Task
        ' TEXTOS DIFERENTES nas duas chamadas, e isso e o teste.
        '
        ' O primeiro corte usava o mesmo texto para resumo e resposta, e o
        ' controle negativo passou: com o resumo sendo apagado pela redacao,
        ' Resultado continuava igual ao que o teste guardara -- porque os dois
        ' textos eram iguais. Um teste cujo dado nao distingue os dois estados
        ' nao distingue nada.
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "O RESUMO"}
        Dim vm = AssistenteViewModelTests.Montar(
            AssistenteViewModelTests.Ativacao(), p,
            AssistenteViewModelTests.Pronta())

        Await vm.Resumir()
        Dim resumo = vm.Resultado
        Assert.AreEqual("O RESUMO", resumo, "controle: houve resumo")

        p.Texto = "A RESPOSTA"
        Await vm.Redigir()

        Assert.AreEqual(resumo, vm.Resultado,
            "a redação apagou o resumo. São dois quadros, e o de cima não é " &
            "rascunho do de baixo.")
        Assert.AreEqual("A RESPOSTA", vm.Resposta)
    End Function

    ''' <summary>
    ''' <b>Redigir exige o resumo, e o LEVA no pedido.</b>
    '''
    ''' A exigência é de produto: a redação sozinha já receberia o e-mail
    ''' inteiro, e sem o resumo as duas chamadas fariam o mesmo trabalho duas
    ''' vezes. O que o resumo acrescenta é o que a primeira concluiu.
    ''' </summary>
    <TestMethod>
    Public Async Function Redigir_exige_o_resumo_e_o_leva_junto() As Task
        Dim ctx As New ContextoQueLembra()
        Dim vm = AssistenteViewModelTests.Montar(
            AssistenteViewModelTests.Ativacao(),
            New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "SUMARIO"},
            AssistenteViewModelTests.Pronta(), contexto:=ctx)

        Assert.IsFalse(vm.PodeRedigir, "ofereceu redigir sem resumo nenhum")
        StringAssert.Contains(vm.PorQueNaoRedige, "Resuma primeiro",
            "o botão apagado não diz o que falta")

        Await vm.Resumir()
        Assert.IsTrue(vm.PodeRedigir, "não liberou redigir depois do resumo")

        Await vm.Redigir()

        StringAssert.Contains(ctx.UltimaInstrucao, "SUMARIO",
            "o resumo não viajou no pedido da redação, e a segunda chamada " &
            "refaz o trabalho da primeira")
    End Function

    ''' <summary>
    ''' <b>Enviar para rascunho aplica — e o desfazer devolve.</b>
    '''
    ''' As guardas do rascunho vieram inteiras do fim da redação para este ato;
    ''' este teste prova que vieram funcionando, e não só copiadas.
    ''' </summary>
    <TestMethod>
    Public Async Function Enviar_para_rascunho_aplica_e_desfazer_devolve() As Task
        Dim rascunho As New AssistenteViewModelTests.RascunhoFalso()
        rascunho.Texto = "o que eu digitei"
        Dim vm = AssistenteViewModelTests.Montar(
            AssistenteViewModelTests.Ativacao(),
            New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "texto da IA"},
            AssistenteViewModelTests.Pronta(), rascunho:=rascunho)

        Await vm.Resumir()
        Await vm.Redigir()

        Assert.IsTrue(vm.PodeEnviarParaRascunho, "não ofereceu o envio havendo rascunho")
        Await vm.EnviarParaRascunho()

        Assert.AreEqual(vm.Resposta, rascunho.Texto, "não aplicou no rascunho")
        Assert.IsTrue(vm.PodeDesfazer, "aplicou e não deixou desfazer")

        vm.Desfazer()
        Assert.AreEqual("o que eu digitei", rascunho.Texto,
            "o desfazer não devolveu o que estava lá")
    End Function

    ''' <summary>
    ''' <b>Sem lugar nenhum para escrever, o envio se explica.</b>
    '''
    ''' Não é mais "sem compositor aberto": compositor fechado deixou de
    ''' ser recusa, porque o botão abre a resposta ele mesmo. O que sobra
    ''' aqui é o caso em que não há <b>nada</b> a responder — nenhuma
    ''' mensagem selecionada, ou o compositor travado na confirmação de
    ''' envio.
    '''
    ''' A resposta continua na tela para copiar: ela já custou e já saiu da
    ''' máquina, e esconder o texto não desfaz divulgação nenhuma.
    ''' </summary>
    <TestMethod>
    Public Async Function Sem_lugar_nenhum_o_envio_diz_o_que_falta() As Task
        Dim vm = AssistenteViewModelTests.Montar(
            AssistenteViewModelTests.Ativacao(),
            New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "texto da IA"},
            AssistenteViewModelTests.Pronta(),
            rascunho:=New AssistenteViewModelTests.RascunhoFalso() With {
                .PodeEditar = False, .PodeAbrirResposta = False})

        Await vm.Resumir()
        Await vm.Redigir()

        Assert.IsTrue(vm.TemResposta, "a resposta sumiu porque não havia onde aplicá-la")
        Assert.IsFalse(vm.PodeEnviarParaRascunho)
        Assert.IsTrue(vm.TemMotivoParaNaoEnviar,
            "botão apagado e nenhuma explicação: promete e não diz o que falta")
    End Function

    ''' <summary>
    ''' <b>Compositor fechado NÃO é recusa: o botão abre a resposta.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A PRÉ-CONDIÇÃO ERA MINHA, E NÃO DO PEDIDO</b>
    '''
    ''' O botão exigia que o usuário clicasse em Responder antes, e essa
    ''' exigência não vinha do que ele pediu — vinha da forma do código: a
    ''' porta do rascunho só sabia <i>escrever</i> em compositor aberto.
    ''' Gastei três rodadas de "ainda não funciona" consertando a plumbagem
    ''' em volta dela: a mensagem que mentia, e depois o botão que não era
    ''' avisado. Os dois eram defeitos de verdade, e nenhum era <b>o</b>
    ''' defeito.
    '''
    ''' Quem pede "manda esta resposta para um rascunho" está pedindo o
    ''' rascunho também.
    ''' </summary>
    <TestMethod>
    Public Async Function Compositor_fechado_o_envio_ABRE_a_resposta() As Task
        Dim r As New AssistenteViewModelTests.RascunhoFalso() With {.PodeEditar = False}
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "O RESUMO"}
        Dim vm = AssistenteViewModelTests.Montar(AssistenteViewModelTests.Ativacao(), p,
                                                 AssistenteViewModelTests.Pronta(),
                                                 Nothing, r)

        Await vm.Resumir()
        p.Texto = "A RESPOSTA"
        Await vm.Redigir()

        Assert.IsTrue(vm.EnviarParaRascunhoCommand.CanExecute(Nothing),
            "com mensagem para responder, o botao tem de estar aceso mesmo " &
            "sem compositor aberto")
        Assert.AreEqual("", vm.PorQueNaoEnvia,
            "e nao ha o que explicar: nao falta nada ao usuario")

        Await vm.EnviarParaRascunho()

        Assert.AreEqual(1, r.Aberturas, "o envio tinha de abrir a resposta")
        Assert.AreEqual("A RESPOSTA", r.Texto, "e escrever nela")
        Assert.IsTrue(vm.PodeDesfazer, "o desfazer vale igual pelo caminho novo")
    End Function

    ''' <summary>
    ''' <b>Abrir falhou: não escreve, e diz.</b>
    '''
    ''' Abrir é assíncrono — pode falhar, pode demorar, e o usuário pode
    ''' mexer no meio. Escrever assumindo que deu certo poria a resposta em
    ''' lugar nenhum, ou no rascunho errado.
    '''
    ''' <b>É o controle negativo da abertura:</b> sem a segunda conferência,
    ''' este é o único teste que quebra.
    ''' </summary>
    <TestMethod>
    Public Async Function Se_a_resposta_NAO_abre_nada_e_escrito() As Task
        Dim r As New AssistenteViewModelTests.RascunhoFalso() With {
            .PodeEditar = False, .AbrirFunciona = False, .Texto = "nao me toque"}
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "O RESUMO"}
        Dim vm = AssistenteViewModelTests.Montar(AssistenteViewModelTests.Ativacao(), p,
                                                 AssistenteViewModelTests.Pronta(),
                                                 Nothing, r)

        Await vm.Resumir()
        p.Texto = "A RESPOSTA"
        Await vm.Redigir()
        Await vm.EnviarParaRascunho()

        Assert.AreEqual(1, r.Aberturas, "tentou abrir")
        Assert.AreEqual("nao me toque", r.Texto,
            "escreveu num rascunho que nao abriu")
        Assert.IsFalse(vm.PodeDesfazer,
            "nada foi aplicado, entao nao pode haver desfazer armado")
        Assert.AreNotEqual("", vm.Aviso, "falhou calado")
    End Function

    ''' <summary>
    ''' <b>O botão é avisado quando o compositor muda.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ESTE TESTE OUVE O EVENTO, E NÃO PERGUNTA A PROPRIEDADE</b>
    '''
    ''' <c>PodeEnviarParaRascunho</c> lê o compositor na hora, e
    ''' <c>CanExecute</c> a chama na hora. Perguntar qualquer uma das duas
    ''' responde certo mesmo com o botão errado na tela: <b>escrevi essa
    ''' versão primeiro, e ela passou com o defeito no lugar.</b>
    '''
    ''' Quem fica desatualizado é o <b>botão</b>. O WPF guarda a última
    ''' resposta e só reconsulta quando <c>CanExecuteChanged</c> chega — e o
    ''' compositor é a única coisa da lista que muda por fora do
    ''' assistente. O evento do rascunho já chegava ao <c>Avisar</c>; ele é
    ''' que não repassava.
    '''
    ''' O caso aqui é o inverso do que era: o compositor <b>trava</b> na
    ''' confirmação de envio, e o botão tem de apagar — e a explicação
    ''' aparecer — sem ninguém tocar no assistente.
    ''' </summary>
    <TestMethod>
    Public Async Function Travar_o_compositor_APAGA_o_botao_de_enviar() As Task
        Dim r As New AssistenteViewModelTests.RascunhoFalso()
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "O RESUMO"}
        Dim vm = AssistenteViewModelTests.Montar(AssistenteViewModelTests.Ativacao(), p,
                                                 AssistenteViewModelTests.Pronta(),
                                                 Nothing, r)

        Await vm.Resumir()
        p.Texto = "A RESPOSTA"
        Await vm.Redigir()
        Assert.IsTrue(vm.EnviarParaRascunhoCommand.CanExecute(Nothing))

        Dim cutucadas = 0
        AddHandler vm.EnviarParaRascunhoCommand.CanExecuteChanged,
            Sub(remetente As Object, arg As EventArgs) cutucadas += 1

        Dim frases = 0
        AddHandler vm.PropertyChanged,
            Sub(remetente As Object, arg As ComponentModel.PropertyChangedEventArgs)
                If arg.PropertyName = NameOf(AssistenteViewModel.PorQueNaoEnvia) Then
                    frases += 1
                End If
            End Sub

        ' O compositor TRAVA: confirmacao de envio, campos bloqueados.
        r.PodeEditar = False
        r.PodeAbrirResposta = False

        Assert.AreNotEqual(0, cutucadas,
            "o BOTAO nao foi avisado. Ele guarda a ultima resposta e fica " &
            "aceso ate alguem mandar reconsultar")
        Assert.AreNotEqual(0, frases,
            "a explicacao nao foi avisada: o motivo novo nao chega a tela")

        Assert.IsFalse(vm.EnviarParaRascunhoCommand.CanExecute(Nothing))
        Assert.IsTrue(vm.TemMotivoParaNaoEnviar)
    End Function

    ''' <summary>
    ''' <b>Trocar de mensagem enquanto a resposta abre NÃO escreve nada.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A CORRIDA QUE A ABERTURA CRIOU</b>
    '''
    ''' Enquanto o envio exigia um compositor já aberto, ele era síncrono e
    ''' não havia janela. Fazer o botão <i>abrir</i> a resposta pôs um
    ''' <c>Await</c> no meio — e tudo que atravessa um <c>Await</c> precisa
    ''' perguntar, do outro lado, se ainda está falando da mesma coisa.
    '''
    ''' A sequência era: clica em enviar na mensagem A; o compositor de A
    ''' começa a abrir; o usuário clica na B; o <c>Trocou</c> restaura a
    ''' resposta de B; e o envio, que lia <c>Resposta</c> do outro lado do
    ''' <c>Await</c>, escrevia <b>a resposta de B dentro do rascunho de
    ''' A</b>.
    '''
    ''' Achado por revisão externa, e não pela suíte: o teste que existia
    ''' encenava a abertura como instantânea.
    ''' </summary>
    <TestMethod>
    Public Async Function Trocar_de_mensagem_enquanto_abre_NAO_escreve() As Task
        Dim r As New AssistenteViewModelTests.RascunhoFalso() With {
            .PodeEditar = False, .Texto = "o que ja estava no rascunho de A"}
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "O RESUMO"}
        Dim vm = AssistenteViewModelTests.Montar(AssistenteViewModelTests.Ativacao(), p,
                                                 AssistenteViewModelTests.Pronta(),
                                                 Nothing, r)
        vm.Trocou(New ItemKey("E-A", "store-1"))

        Await vm.Resumir()
        p.Texto = "A RESPOSTA DE A"
        Await vm.Redigir()

        ' O usuario clica noutra mensagem EXATAMENTE enquanto a resposta abre.
        r.NoMeioDaAbertura = Sub() vm.Trocou(New ItemKey("E-B", "store-1"))

        Await vm.EnviarParaRascunho()

        Assert.AreEqual("o que ja estava no rascunho de A", r.Texto,
            "a resposta foi escrita num rascunho que e de outra mensagem")
        Assert.IsFalse(vm.PodeDesfazer,
            "nada foi aplicado, entao nao pode haver desfazer armado")
        Assert.AreNotEqual("", vm.Aviso,
            "recusou calado: o usuario clicou e nada aconteceu sem explicacao")
    End Function

End Class

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
        vm.EnviarParaRascunho()

        Assert.AreEqual(vm.Resposta, rascunho.Texto, "não aplicou no rascunho")
        Assert.IsTrue(vm.PodeDesfazer, "aplicou e não deixou desfazer")

        vm.Desfazer()
        Assert.AreEqual("o que eu digitei", rascunho.Texto,
            "o desfazer não devolveu o que estava lá")
    End Function

    ''' <summary>
    ''' <b>Sem rascunho, o envio se explica em vez de sumir.</b>
    '''
    ''' A resposta continua na tela para copiar — ela já custou e já saiu da
    ''' máquina; esconder o texto não desfaz divulgação nenhuma.
    ''' </summary>
    <TestMethod>
    Public Async Function Sem_rascunho_o_envio_diz_o_que_falta() As Task
        Dim vm = AssistenteViewModelTests.Montar(
            AssistenteViewModelTests.Ativacao(),
            New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "texto da IA"},
            AssistenteViewModelTests.Pronta(),
            rascunho:=New AssistenteViewModelTests.RascunhoFalso() With {.PodeEditar = False})

        Await vm.Resumir()
        Await vm.Redigir()

        Assert.IsTrue(vm.TemResposta, "a resposta sumiu porque não havia onde aplicá-la")
        Assert.IsFalse(vm.PodeEnviarParaRascunho)
        Assert.IsTrue(vm.TemMotivoParaNaoEnviar,
            "botão apagado e nenhuma explicação: promete e não diz o que falta")
    End Function

    ''' <summary>
    ''' <b>O botão ACORDA quando a resposta é aberta.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ESTE TESTE OUVE O EVENTO, E NÃO PERGUNTA A PROPRIEDADE</b>
    '''
    ''' <c>PodeEnviarParaRascunho</c> estava <b>certa</b> o tempo todo — ela
    ''' lê o compositor na hora, e <c>CanExecute</c> a chama na hora
    ''' também. Perguntar qualquer uma das duas responde <c>True</c> mesmo
    ''' com o botão apagado na tela: <b>escrevi essa versão primeiro, e ela
    ''' passou com o defeito no lugar.</b>
    '''
    ''' Quem fica desatualizado é o <b>botão</b>. O WPF guarda a última
    ''' resposta e só reconsulta quando <c>CanExecuteChanged</c> chega — e o
    ''' compositor é a única coisa da lista que muda por fora do
    ''' assistente. O evento do rascunho já chegava ao <c>Avisar</c>; ele é
    ''' que não repassava, e quem clicava em Responder via o botão continuar
    ''' apagado e a frase embaixo continuar mandando abrir uma resposta que
    ''' já estava aberta.
    '''
    ''' É a terceira vez nesta base, e o parágrafo que descreve a armadilha
    ''' está <b>duas telas acima da linha que faltava</b>.
    ''' </summary>
    <TestMethod>
    Public Async Function Abrir_a_resposta_ACORDA_o_botao_de_enviar() As Task
        Dim r As New AssistenteViewModelTests.RascunhoFalso() With {.PodeEditar = False}
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "O RESUMO"}
        Dim vm = AssistenteViewModelTests.Montar(AssistenteViewModelTests.Ativacao(), p,
                                                 AssistenteViewModelTests.Pronta(),
                                                 Nothing, r)

        Await vm.Resumir()
        p.Texto = "A RESPOSTA"
        Await vm.Redigir()

        Assert.IsFalse(vm.EnviarParaRascunhoCommand.CanExecute(Nothing),
            "sem compositor aberto o botao tem de estar apagado")
        StringAssert.Contains(vm.PorQueNaoEnvia, "Responder",
            "e a frase tem de dizer o que fazer")

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

        ' O usuario clica em Responder: o compositor entra em edicao.
        r.PodeEditar = True

        Assert.AreNotEqual(0, cutucadas,
            "o BOTAO nao foi avisado. Ele guarda a ultima resposta e fica " &
            "apagado ate alguem mandar reconsultar")
        Assert.AreNotEqual(0, frases,
            "a frase que manda abrir uma resposta ficou na tela depois de a " &
            "resposta ter sido aberta")

        Assert.AreEqual("", vm.PorQueNaoEnvia)
        vm.EnviarParaRascunho()
        Assert.AreEqual("A RESPOSTA", r.Texto, "e ai ele funciona")
    End Function

End Class

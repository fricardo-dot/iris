Imports System.Threading.Tasks
Imports Iris.App.ViewModels
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O TRABALHO DA IA SOBREVIVE A UM CLIQUE NOUTRA MENSAGEM.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTAVA ERRADO</b>
'''
''' <c>Trocou()</c> zerava o resumo, e nada o trazia de volta. Ler o resumo,
''' conferir uma coisa na mensagem de cima e voltar apagava o que o usuário
''' <b>pagou</b> — em dinheiro e em espera. E a redação nem era zerada: a
''' resposta redigida para uma mensagem <b>ficava na tela</b> sob a seguinte,
''' que é o defeito oposto e pior.
'''
''' ------------------------------------------------------------------
''' <b>AS DUAS METADES, E A SEGUNDA É QUE IMPORTA</b>
'''
''' Guardar é conveniência. <b>Não mostrar o guardado no lugar errado é
''' correção</b>: um resumo da mensagem A embaixo da mensagem B tem cara de
''' certo, e ninguém desconfia de um texto plausível.
'''
''' Por isso a memória é indexada por <see cref="ItemKey"/>, chave ausente não
''' restaura nada, e sessão nova apaga tudo — <c>EntryID</c> só identifica
''' dentro da ligação em que foi lido.
'''
''' ------------------------------------------------------------------
''' <b>CONTROLE NEGATIVO</b>
'''
''' <see cref="O_resumo_de_uma_mensagem_NAO_aparece_sob_outra"/> falha se a
''' restauração ignorar a chave. É ele que separa "guardei" de "guardei no
''' lugar certo", e sem ele uma memória que devolvesse sempre o último resumo
''' passaria em todos os outros testes deste arquivo.
''' </summary>
' NAO PARALELIZAR: usa AssistenteViewModelTests.Montar, que mexe em estado
' Shared (o relogio e a ultima copia). Mesmo motivo declarado la.
<TestClass>
<DoNotParallelize>
Public Class MemoriaPorMensagemTests

    Private Shared Function Chave(n As Integer) As ItemKey
        Return New ItemKey($"E-{n}", "store-1")
    End Function

    ''' <summary>Um assistente já apontado para a mensagem <paramref name="n"/>.</summary>
    Private Shared Function Apontado(p As AssistenteViewModelTests.ProvedorControlado,
                                     n As Integer) As AssistenteViewModel
        Dim vm = AssistenteViewModelTests.Montar(AssistenteViewModelTests.Ativacao(), p,
                                                 AssistenteViewModelTests.Pronta())
        vm.Trocou(Chave(n))
        Return vm
    End Function

    <TestMethod>
    Public Async Function Voltar_para_a_mensagem_ja_resumida_TRAZ_o_resumo_de_volta() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo da UM"}
        Dim vm = Apontado(p, 1)

        Await vm.Resumir()
        Assert.AreEqual("o resumo da UM", vm.Resultado)

        vm.Trocou(Chave(2))
        Assert.AreEqual("", vm.Resultado, "a mensagem DOIS nunca foi resumida")

        vm.Trocou(Chave(1))
        Assert.AreEqual("o resumo da UM", vm.Resultado,
            "voltar tem de trazer o resumo de volta, e nao cobrar outro")
        Assert.AreEqual(1, p.Chamadas, "voltar nao pode pedir de novo ao provedor")
    End Function

    ''' <summary>
    ''' <b>O CONTROLE NEGATIVO DESTE ARQUIVO.</b>
    '''
    ''' Se a restauração deixar de olhar a chave — devolvendo "o último
    ''' guardado" — este é o único teste que quebra.
    ''' </summary>
    <TestMethod>
    Public Async Function O_resumo_de_uma_mensagem_NAO_aparece_sob_outra() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo da UM"}
        Dim vm = Apontado(p, 1)
        Await vm.Resumir()

        vm.Trocou(Chave(2))

        Assert.AreEqual("", vm.Resultado,
            "resumo de outra mensagem embaixo desta e pior que resumo nenhum")
        Assert.IsFalse(vm.TemResultado)
        Assert.AreEqual("", vm.Ficha, "a ficha do voo alheio tambem nao fica")
    End Function

    ''' <summary>
    ''' A resposta redigida volta junto — e some junto. Ela era o caso pior:
    ''' <c>Trocou</c> zerava <c>Resultado</c> e <b>esquecia</b> de zerar
    ''' <c>Resposta</c>, então uma resposta redigida para a mensagem A
    ''' continuava na tela sob a mensagem B.
    ''' </summary>
    <TestMethod>
    Public Async Function A_resposta_redigida_acompanha_a_mensagem() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "O RESUMO"}
        Dim vm = Apontado(p, 1)
        Await vm.Resumir()
        p.Texto = "A RESPOSTA"
        Await vm.Redigir()
        Assert.AreEqual("A RESPOSTA", vm.Resposta)

        vm.Trocou(Chave(2))
        Assert.AreEqual("", vm.Resposta,
            "a resposta redigida para outra mensagem NAO pode ficar na tela")

        vm.Trocou(Chave(1))
        Assert.AreEqual("A RESPOSTA", vm.Resposta, "e volta com a mensagem dela")
        Assert.AreEqual("O RESUMO", vm.Resultado, "junto com o resumo")
    End Function

    ''' <summary>
    ''' Trocar de PASTA, ou desmarcar, chega aqui sem chave. Sem saber de que
    ''' mensagem se fala, restaurar seria adivinhar.
    ''' </summary>
    <TestMethod>
    Public Async Function Sem_chave_nao_restaura_nada() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo da UM"}
        Dim vm = Apontado(p, 1)
        Await vm.Resumir()

        vm.Trocou()
        Assert.AreEqual("", vm.Resultado, "sem chave, a tela fica limpa")

        vm.Trocou(Chave(1))
        Assert.AreEqual("o resumo da UM", vm.Resultado,
            "mas o guardado continua guardado: passar por lugar nenhum nao apaga")
    End Function

    ''' <summary>
    ''' Chave vazia é <b>chave nenhuma</b>. <see cref="ItemKey.IsEmpty"/> existe
    ''' porque <c>EntryID</c> ausente vira string vazia, e duas mensagens sem
    ''' identidade seriam a "mesma" chave — o vazamento que o controle negativo
    ''' proíbe, entrando pela porta dos fundos.
    ''' </summary>
    <TestMethod>
    Public Async Function Chave_VAZIA_nao_guarda_e_nao_restaura() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "sem dono"}
        Dim vm = AssistenteViewModelTests.Montar(AssistenteViewModelTests.Ativacao(), p,
                                                 AssistenteViewModelTests.Pronta())
        vm.Trocou(New ItemKey("", "store-1"))
        Await vm.Resumir()

        vm.Trocou(New ItemKey("", "store-1"))
        Assert.AreEqual("", vm.Resultado,
            "duas mensagens sem EntryID nao sao a mesma mensagem")
    End Function

    ''' <summary>
    ''' Sessão nova é outra ligação com o Outlook: o mesmo <c>EntryID</c> pode
    ''' apontar para outra coisa, ou para nada.
    ''' </summary>
    <TestMethod>
    Public Async Function Sessao_nova_apaga_a_memoria_inteira() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo da UM"}
        Dim vm = Apontado(p, 1)
        Await vm.Resumir()

        vm.EsquecerASessao()
        Assert.AreEqual("", vm.Resultado, "a tela limpa")

        vm.Trocou(Chave(1))
        Assert.AreEqual("", vm.Resultado,
            "e a memoria tambem: o EntryID de ontem nao identifica nada hoje")
    End Function

    ''' <summary>
    ''' Vinte mensagens cabem; a vigésima primeira empurra a primeira para
    ''' fora. Sem teto, uma varredura de pasta grande acumularia o corpo de
    ''' cada resumo pelo tempo que o programa ficasse aberto.
    ''' </summary>
    <TestMethod>
    Public Async Function A_memoria_tem_TETO_e_esquece_a_mais_antiga() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado()
        Dim vm = AssistenteViewModelTests.Montar(AssistenteViewModelTests.Ativacao(), p,
                                                 AssistenteViewModelTests.Pronta())

        For n = 1 To 21
            vm.Trocou(Chave(n))
            p.Texto = $"resumo {n}"
            Await vm.Resumir()
        Next

        vm.Trocou(Chave(21))
        Assert.AreEqual("resumo 21", vm.Resultado, "a ultima continua la")
        vm.Trocou(Chave(2))
        Assert.AreEqual("resumo 2", vm.Resultado, "e a segunda tambem")

        vm.Trocou(Chave(1))
        Assert.AreEqual("", vm.Resultado, "a primeira saiu, e e a primeira que sai")
    End Function

    ''' <summary>
    ''' Passar por uma mensagem <b>sem</b> pedir nada não apaga o que ela já
    ''' tinha — e é diferente de nunca ter pedido. Um <c>Guardar</c> que
    ''' gravasse o vazio por cima transformaria a memória em algo que se apaga
    ''' sozinho ao ser usada.
    ''' </summary>
    <TestMethod>
    Public Async Function Visitar_sem_pedir_nada_NAO_apaga_o_que_havia() As Task
        Dim p As New AssistenteViewModelTests.ProvedorControlado() With {.Texto = "o resumo da UM"}
        Dim vm = Apontado(p, 1)
        Await vm.Resumir()

        vm.Trocou(Chave(2))
        vm.Trocou(Chave(3))
        vm.Trocou(Chave(1))

        Assert.AreEqual("o resumo da UM", vm.Resultado)
    End Function

End Class

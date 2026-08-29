Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.App.ViewModels
Imports Iris.Core
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A TELA DE TAREFAS — e as duas etapas que a Fase 5 exige.</b>
'''
''' O ESCOPO escreveu a fase assim: <i>a IA sugere, você confirma, o Iris
''' cria</i>. A parte que importa é o meio — <b>nunca criação silenciosa em
''' massa</b>. Uma sugestão que virasse tarefa sozinha encheria de lixo
''' justamente a lista que existe para dizer o que falta fazer.
'''
''' Por isso o primeiro teste daqui é o que prova que <b>propor não cria</b>.
''' </summary>
<TestClass>
Public Class TarefasViewModelTests

    Private Shared ReadOnly Pasta As New FolderKey("tarefas-1", "store-1")

    Private NotInheritable Class BrokerDeTarefas
        Implements ITarefasBroker

        Friend ReadOnly Chamadas As New List(Of String)()
        Friend Itens As New List(Of TaskInfo)()
        Friend Recusados As Integer? = 0
        Friend Truncada As Boolean
        Friend Recusa As String
        Friend UltimoRascunho As TaskDraft
        Friend UltimaPasta As FolderKey
        Friend UltimaChave As TaskKey
        Friend Trava As TaskCompletionSource(Of Boolean)
        Friend SemPasta As Boolean

        Public Function GetDefaultTasksFolderAsync(cancel As CancellationToken) _
            As Task(Of OperationResult(Of FolderKey)) _
            Implements ITarefasBroker.GetDefaultTasksFolderAsync

            Chamadas.Add("pasta")
            If SemPasta Then
                Return Task.FromResult(OperationResult(Of FolderKey).Fail(ErrorKind.NotFound, "sem pasta"))
            End If
            Return Task.FromResult(OperationResult(Of FolderKey).Ok(Pasta))
        End Function

        Public Function GetTasksAsync(folder As FolderKey, teto As Integer,
                                      cancel As CancellationToken) _
            As Task(Of OperationResult(Of TaskList)) _
            Implements ITarefasBroker.GetTasksAsync

            Chamadas.Add("ler")
            Dim lista As New TaskList With {.Skipped = Recusados, .Truncada = Truncada}
            lista.Items.AddRange(Itens)
            Return Task.FromResult(OperationResult(Of TaskList).Ok(lista))
        End Function

        Public Async Function CreateTaskAsync(folder As FolderKey, rascunho As TaskDraft,
                                              cancel As CancellationToken) _
            As Task(Of OperationResult(Of TaskInfo)) _
            Implements ITarefasBroker.CreateTaskAsync

            Chamadas.Add("criar")
            UltimaPasta = folder
            UltimoRascunho = rascunho
            If Trava IsNot Nothing Then Await Trava.Task
            If Recusa IsNot Nothing Then
                Return OperationResult(Of TaskInfo).Fail(ErrorKind.Denied, Recusa)
            End If
            Return OperationResult(Of TaskInfo).Ok(
                New TaskInfo With {.Key = New ItemKey("nova", "store-1"),
                                   .Subject = rascunho.Subject})
        End Function

        Public Function CompleteTaskAsync(chave As TaskKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of TaskInfo)) _
            Implements ITarefasBroker.CompleteTaskAsync

            Chamadas.Add("concluir")
            UltimaChave = chave
            If Recusa IsNot Nothing Then
                Return Task.FromResult(OperationResult(Of TaskInfo).Fail(ErrorKind.Denied, Recusa))
            End If
            Return Task.FromResult(OperationResult(Of TaskInfo).Ok(New TaskInfo()))
        End Function
    End Class

    Private Shared Function Tarefa(nome As String,
                                   Optional atribuida As Boolean = False,
                                   Optional concluida As Boolean = False) As TaskInfo
        Return New TaskInfo With {
            .Key = New ItemKey(nome, "store-1"),
            .Subject = nome,
            .Atribuida = atribuida,
            .Concluida = concluida}
    End Function

    Private Shared Async Function Aberta(b As BrokerDeTarefas) As Task(Of TarefasViewModel)
        Dim vm As New TarefasViewModel(b)
        Await vm.AbrirCommand.ExecuteAsync(Nothing)
        Return vm
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>PROPOR NÃO CRIA — é a invariante da fase.</b>
    '''
    ''' "A IA sugere, você confirma, o Iris cria": três coisas, e a do meio é a
    ''' que impede a lista de encher de tarefa que ninguém pediu. Propor
    ''' preenche o formulário e <b>para</b>.
    '''
    ''' <b>Controle negativo:</b> fazendo o <c>ProporDaMensagem</c> chamar o
    ''' criar, a asserção do meio cai.
    ''' </summary>
    <TestMethod>
    Public Async Function Propor_NAO_cria() As Task
        Dim b As New BrokerDeTarefas()
        Dim vm = Await Aberta(b)

        vm.ProporDaMensagem("Revisar o aditivo do contrato")

        Assert.AreEqual("Revisar o aditivo do contrato", vm.NovoAssunto,
                        "a proposta nao chegou ao formulario")
        CollectionAssert.DoesNotContain(b.Chamadas, "criar",
            "propor criou a tarefa sozinho -- e a fase inteira existe para isso nao acontecer")
        Assert.IsTrue(vm.TemProposta)
    End Function

    ''' <summary>
    ''' <b>Mensagem sem assunto não vira tarefa anônima.</b>
    '''
    ''' O rascunho precisa de assunto — o <c>TaskWriting</c> recusa vazio. Se a
    ''' proposta viesse vazia, o usuário clicaria em criar e levaria uma recusa
    ''' que ele não causou. Melhor propor um texto que ele veja e troque.
    ''' </summary>
    <TestMethod>
    Public Async Function Mensagem_sem_assunto_propoe_algo_editavel() As Task
        Dim b As New BrokerDeTarefas()
        Dim vm = Await Aberta(b)

        vm.ProporDaMensagem("   ")

        Assert.IsTrue(vm.TemProposta, "a proposta veio vazia, e criar seria recusado")
        StringAssert.Contains(vm.NovoAssunto, "sem assunto")
    End Function

    ''' <summary>Criar manda para a pasta que a abertura descobriu, e recarrega.</summary>
    <TestMethod>
    Public Async Function Criar_usa_a_pasta_de_tarefas_e_recarrega() As Task
        Dim b As New BrokerDeTarefas()
        Dim vm = Await Aberta(b)

        vm.ProporDaMensagem("Ligar para o fornecedor")
        Await vm.CriarCommand.ExecuteAsync(Nothing)

        Assert.AreEqual(Pasta, b.UltimaPasta, "criou noutra pasta")
        Assert.AreEqual("Ligar para o fornecedor", b.UltimoRascunho.Subject)
        Assert.IsNull(b.UltimoRascunho.Vence, "inventou um prazo que ninguem marcou")
        Assert.AreEqual(2, b.Chamadas.Where(Function(c) c = "ler").Count(),
                        "nao recarregou depois de criar")
        Assert.AreEqual("", vm.NovoAssunto, "o formulario nao limpou")
    End Function

    ''' <summary>
    ''' <b>TAREFA ATRIBUÍDA NÃO OFERECE O BOTÃO — e a tela diz por quê.</b>
    '''
    ''' Mexer numa tarefa atribuída manda atualização por e-mail. O
    ''' <c>TaskWriting</c> recusa de qualquer jeito; o que a tela faz é não
    ''' prometer. E prometer menos sem explicar seria pior — por isso o aviso.
    ''' </summary>
    <TestMethod>
    Public Async Function Tarefa_atribuida_nao_oferece_concluir_e_explica() As Task
        Dim b As New BrokerDeTarefas()
        b.Itens.Add(Tarefa("minha"))
        b.Itens.Add(Tarefa("do outro", atribuida:=True))
        Dim vm = Await Aberta(b)

        ' CONTROLE POSITIVO: a minha pode ser concluida.
        vm.Selecionada = vm.Tarefas.First(Function(t) t.Assunto = "minha")
        Assert.IsTrue(vm.PodeConcluir, "controle: tarefa propria tinha de poder concluir")
        Assert.AreEqual("", vm.AvisoDaSelecionada)

        vm.Selecionada = vm.Tarefas.First(Function(t) t.Assunto = "do outro")

        Assert.IsFalse(vm.PodeConcluir, "ofereceu concluir numa tarefa atribuida")
        Assert.IsFalse(vm.ConcluirCommand.CanExecute(Nothing))
        StringAssert.Contains(vm.AvisoDaSelecionada, "e-mail",
            "desabilitou o botao e nao disse por que: " & vm.AvisoDaSelecionada)
    End Function

    ''' <summary>A recusa do escritor aparece com as palavras dele.</summary>
    <TestMethod>
    Public Async Function A_recusa_do_escritor_aparece_na_tela() As Task
        Dim b As New BrokerDeTarefas() With {.Recusa = "esta tarefa está atribuída a alguém"}
        Dim vm = Await Aberta(b)

        vm.ProporDaMensagem("qualquer coisa")
        Await vm.CriarCommand.ExecuteAsync(Nothing)

        Assert.IsTrue(vm.TemErro)
        StringAssert.Contains(vm.Erro, "atribuída a alguém",
            "a tela engoliu o motivo: " & vm.Erro)
    End Function

    ''' <summary>
    ''' <b>O segundo clique não cria a segunda tarefa.</b>
    '''
    ''' Mutação não tem retry porque criar não é idempotente; uma tela que deixe
    ''' clicar duas vezes desfaz a garantia por fora. É a mesma lição da agenda,
    ''' e aqui o estrago é uma tarefa duplicada na lista de verdade.
    ''' </summary>
    <TestMethod>
    Public Async Function O_segundo_clique_nao_cria_de_novo() As Task
        Dim b As New BrokerDeTarefas() With {.Trava = New TaskCompletionSource(Of Boolean)()}
        Dim vm = Await Aberta(b)

        vm.ProporDaMensagem("uma vez só")
        Dim emVoo = vm.CriarCommand.ExecuteAsync(Nothing)

        Assert.IsFalse(vm.PodeCriar, "o botao continuou habilitado com gravacao em voo")

        b.Trava.SetResult(True)
        Await emVoo

        Assert.AreEqual(1, b.Chamadas.Where(Function(c) c = "criar").Count())
    End Function

    ''' <summary>
    ''' <b>Lista vazia não afirma ausência.</b>
    '''
    ''' "Nenhuma tarefa" seria afirmação sobre o que o Outlook expõe. A mesma
    ''' regra da agenda e da lista de mensagens, que esta base corrigiu em
    ''' quatro superfícies: o que se sabe é o que foi <i>lido</i>.
    ''' </summary>
    <TestMethod>
    Public Async Function Lista_vazia_NAO_afirma_ausencia() As Task
        Dim b As New BrokerDeTarefas()
        Dim vm = Await Aberta(b)

        StringAssert.Contains(vm.Resumo, "nenhuma tarefa LIDA")
        StringAssert.Contains(vm.Resumo, "não é o mesmo que não haver")
    End Function

    ''' <summary>
    ''' <b>Recusados desconhecidos não viram zero.</b>
    '''
    ''' <c>Nothing</c> em <c>Skipped</c> é "não contei", e zero é "contei e não
    ''' houve". Colapsar os dois é a família de defeito que esta base passou
    ''' uma série de revisões corrigindo.
    ''' </summary>
    <TestMethod>
    Public Async Function Recusados_desconhecidos_nao_viram_zero() As Task
        Dim b As New BrokerDeTarefas() With {.Recusados = Nothing}
        b.Itens.Add(Tarefa("uma"))
        Dim vm = Await Aberta(b)

        StringAssert.Contains(vm.Resumo, "não sei quantos itens foram recusados")

        ' CONTROLE POSITIVO: zero conhecido fica calado.
        Dim b2 As New BrokerDeTarefas() With {.Recusados = 0}
        b2.Itens.Add(Tarefa("uma"))
        Dim vm2 = Await Aberta(b2)
        Assert.IsFalse(vm2.Resumo.Contains("não sei quantos"))
    End Function

End Class

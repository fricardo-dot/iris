Imports System.IO
Imports System.Linq
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>TAREFAS — e as duas armadilhas que a Fase 5 tinha.</b>
'''
''' ------------------------------------------------------------------
''' <b>A PRIMEIRA: TAREFA ATRIBUÍDA CONVERSA POR E-MAIL</b>
'''
''' É a mesma do calendário, num objeto diferente. <c>TaskItem.Assign()</c>
''' seguido de <c>Send()</c> manda um pedido de tarefa <b>por e-mail</b>, e
''' depois disso cada mudança de status vai e volta pela caixa. Salvar uma
''' tarefa atribuída não é escrita local.
'''
''' O desenho é o mesmo que passou na Fase 6: <see cref="TaskDraft"/> não tem
''' responsável, então não existe caminho para o Iris atribuir; e concluir
''' confere <c>DelegationState</c> antes de tocar no item.
'''
''' ------------------------------------------------------------------
''' <b>A SEGUNDA: O "SEM PRAZO" DO OUTLOOK É UMA DATA</b>
'''
''' <c>TaskItem.DueDate</c> nunca é nulo. Sem prazo, ele vale
''' <c>4501-01-01</c> — um sentinela. Deixar isso virar um
''' <c>DateTimeOffset</c> comum transformaria "não tem prazo" em "vence em
''' 4501", que é <b>ausência virando fato</b>: a família de defeito que esta
''' base passou uma série inteira de revisões corrigindo, em cinco lugares
''' diferentes.
'''
''' Por isso o vencimento é anulável, e por isso a tradução tem teste.
''' </summary>
<TestClass>
Public Class TarefasTests

    ' ==================================================================
    ' O sentinela

    ''' <summary>
    ''' <b>Sem prazo continua sem prazo.</b>
    '''
    ''' <b>Controle negativo:</b> devolvendo o sentinela como data comum, a
    ''' primeira asserção cai — e a tela passa a mostrar um vencimento que
    ''' ninguém marcou.
    ''' </summary>
    <TestMethod>
    Public Sub O_sentinela_de_sem_data_vira_Nothing()
        Assert.IsNull(TaskWriting.Vencimento(New Date(4501, 1, 1)),
                      "o sentinela do Outlook virou um vencimento de verdade")

        ' E qualquer coisa depois dele também é sentinela: o Outlook usa
        ' 4501-01-01, mas uma leitura torta pode devolver adiante.
        Assert.IsNull(TaskWriting.Vencimento(New Date(4501, 6, 30)))
    End Sub

    ''' <summary>
    ''' <b>CONTROLE POSITIVO: uma data de verdade atravessa.</b>
    '''
    ''' Sem ele, uma tradução que devolvesse <c>Nothing</c> sempre passaria no
    ''' teste acima — o bloqueio sem controle negativo que o CLAUDE.md
    ''' descreve.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_uma_data_de_verdade_atravessa()
        Dim quando = New Date(2026, 9, 15)
        Dim r = TaskWriting.Vencimento(quando)

        Assert.IsTrue(r.HasValue, "perdeu um vencimento de verdade")
        Assert.AreEqual(quando, r.Value.LocalDateTime.Date)
    End Sub

    ' ==================================================================
    ' A guarda de criação

    <TestMethod>
    Public Sub Controle_um_rascunho_comum_e_aceito()
        Assert.IsNull(TaskWriting.RecusarRascunho(
            New TaskDraft With {.Subject = "Responder o contrato"}))
    End Sub

    ''' <summary>
    ''' <b>Tarefa sem assunto não entra.</b>
    '''
    ''' Uma linha sem texto na lista de tarefas não diz o que é para fazer — e
    ''' quem a encontrar depois não tem como saber, porque o Iris não estava
    ''' lá quando ela foi criada.
    ''' </summary>
    <DataTestMethod>
    <DataRow("")>
    <DataRow("   ")>
    Public Sub Sem_assunto_RECUSA(assunto As String)
        Dim motivo = TaskWriting.RecusarRascunho(New TaskDraft With {.Subject = assunto})

        Assert.IsNotNull(motivo, "aceitou tarefa sem assunto")
        StringAssert.Contains(motivo, "assunto")
    End Sub

    <TestMethod>
    Public Sub Rascunho_nulo_RECUSA()
        Assert.IsNotNull(TaskWriting.RecusarRascunho(Nothing))
    End Sub

    ' ==================================================================
    ' A invariante, presa pelo tipo

    ''' <summary>
    ''' <b>O RASCUNHO NÃO TEM ONDE PÔR RESPONSÁVEL — e é essa a garantia.</b>
    '''
    ''' Este teste não exercita comportamento: ele prende o <i>desenho</i>. A
    ''' invariante "o Iris não atribui tarefa" está sustentada por não existir
    ''' campo, e não por alguém lembrar de não preencher.
    '''
    ''' Se um dia alguém acrescentar <c>Owner</c>, <c>Delegator</c> ou
    ''' <c>AssignedTo</c>, este teste cai — e quem o acrescentar tem de ler o
    ''' comentário do <c>TaskWriting</c> antes de apagá-lo.
    ''' </summary>
    <TestMethod>
    Public Sub O_rascunho_NAO_tem_campo_de_responsavel()
        Dim campos = GetType(TaskDraft).GetProperties().
                     Select(Function(p) p.Name.ToLowerInvariant()).ToList()

        For Each proibido In {"owner", "delegator", "assignedto", "recipients",
                              "to", "responsavel", "delegatedto"}
            Assert.IsFalse(campos.Contains(proibido),
                $"TaskDraft ganhou '{proibido}': tarefa atribuída é pedido " &
                "enviado por e-mail, e o Iris não envia. Leia o comentário do " &
                "TaskWriting antes de mexer nisto.")
        Next
    End Sub

    ''' <summary>
    ''' <b>A tarefa lida DIZ se está atribuída.</b>
    '''
    ''' Não é detalhe de apresentação: é o aviso de que aquela linha pertence a
    ''' uma negociação da qual o Iris não participa — e é o campo que a tela usa
    ''' para não oferecer um botão que vai ser recusado.
    ''' </summary>
    ''' <summary>
    ''' <b>O CÓDIGO QUE ESCREVE NÃO CHAMA Assign NEM Send.</b>
    '''
    ''' Este teste lê o <b>fonte</b> do <c>TaskWriting</c>, e é feio de
    ''' propósito. A revisão externa apontou o buraco exato: os testes por
    ''' reflexão provam que o <c>TaskDraft</c> não tem campo de responsável, e
    ''' não provam nada sobre o que o escritor faz. Alguém que acrescentasse
    ''' <c>t.Assign()</c> e <c>t.Send()</c> direto, sem tocar no rascunho,
    ''' passaria em toda a suíte e mandaria um pedido de tarefa por e-mail.
    '''
    ''' A alternativa honesta seria um duplo de <c>TaskItem</c>, e não existe:
    ''' <c>OL.TaskItem</c> é uma interface COM que o teste não instancia. Entre
    ''' uma varredura de texto e nenhuma prova, a varredura é melhor — e este
    ''' comentário é a ressalva de que ela é um proxy.
    ''' </summary>
    <TestMethod>
    Public Sub O_escritor_de_tarefas_NAO_chama_Assign_nem_Send()
        Dim fonte = LerFonteDoTaskWriting()

        Assert.IsFalse(fonte.Contains(".Assign("),
            "TaskWriting chama Assign: atribuir tarefa manda pedido por e-mail, " &
            "e o Iris não envia. Leia o comentário do módulo antes de mexer nisto.")
        Assert.IsFalse(fonte.Contains(".Send("),
            "TaskWriting chama Send: nada sai por e-mail sem o usuário mandar.")

        ' CONTROLE: a varredura acha o que existe. Sem isto, um caminho de
        ' arquivo errado deixaria as duas asserções passarem por lerem vazio --
        ' o bloqueio que nunca bloqueia.
        StringAssert.Contains(fonte, ".Save()",
            "a varredura não está lendo o TaskWriting: as asserções acima " &
            "estariam passando por não olhar nada")
    End Sub

    ''' <summary>
    ''' <b>A guarda de atribuição vem ANTES do Save, e não depois.</b>
    '''
    ''' Ordem é a coisa inteira. Conferir <c>DelegationState</c> depois de
    ''' <c>Complete = True</c> e <c>Save()</c> seria conferir depois de o
    ''' e-mail já ter saído — e o teste de recusa continuaria verde, porque a
    ''' função ainda devolveria <c>Denied</c>.
    ''' </summary>
    <TestMethod>
    Public Sub A_guarda_de_atribuicao_vem_antes_do_Save()
        Dim fonte = LerFonteDoTaskWriting()
        Dim concluir = fonte.Substring(fonte.IndexOf("Public Function Concluir"))
        concluir = concluir.Substring(0, concluir.IndexOf("End Function"))

        Dim ondeConfere = concluir.IndexOf("EhAtribuida")
        Dim ondeGrava = concluir.IndexOf(".Save()")

        Assert.IsTrue(ondeConfere >= 0, "Concluir parou de conferir EhAtribuida")
        Assert.IsTrue(ondeGrava >= 0, "não achei o Save em Concluir: a varredura errou o recorte")
        Assert.IsTrue(ondeConfere < ondeGrava,
            "a conferência de atribuição ficou DEPOIS do Save -- conferir " &
            "depois de o e-mail sair não é conferir")
    End Sub

    ''' <summary>
    ''' O fonte do <c>TaskWriting</c> <b>sem os comentários</b>.
    '''
    ''' O primeiro corte destes testes falhou aqui, e a falha ensinou: o
    ''' comentário do módulo <i>menciona</i> <c>Assign</c> e <c>Send</c>,
    ''' justamente para explicar por que eles não podem aparecer. Uma
    ''' varredura que lê o comentário acusa o texto que a protege.
    '''
    ''' É a mesma lição do recorte de linha do XAML: um teste que acusa onde
    ''' não há é um teste que alguém desliga na primeira vez que atrapalha.
    ''' </summary>
    Private Shared Function LerFonteDoTaskWriting() As String
        Dim inteiro = LerArquivoDoTaskWriting()
        Dim linhas = inteiro.Split(CChar(vbLf)).
                     Where(Function(l) Not l.TrimStart().StartsWith("'"))
        Return String.Join(vbLf, linhas)
    End Function

    Private Shared Function LerArquivoDoTaskWriting() As String
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
            d = d.Parent
        End While
        Assert.IsNotNull(d, "não achei a raiz do repositório")
        Dim caminho = Path.Combine(d.FullName, "src", "Iris.Outlook", "TaskWriting.vb")
        Assert.IsTrue(File.Exists(caminho), "TaskWriting.vb não encontrado em " & caminho)
        Return File.ReadAllText(caminho)
    End Function

    <TestMethod>
    Public Sub A_tarefa_lida_carrega_o_estado_de_atribuicao()
        Dim campos = GetType(TaskInfo).GetProperties().
                     Select(Function(p) p.Name).ToList()

        CollectionAssert.Contains(campos, "Atribuida",
            "sem este campo a tela não tem como avisar que mexer manda e-mail")
    End Sub

End Class

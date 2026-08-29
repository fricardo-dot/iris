Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.App.ViewModels
Imports Iris.Core
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A AGENDA ESCREVENDO — e a pergunta antes de apagar.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE APAGAR PEDE CONFIRMAÇÃO E CRIAR NÃO</b>
'''
''' Criar um compromisso errado é visível e reversível: ele aparece na agenda e
''' você apaga. Apagar é o contrário — o item some da tela, vai para os Itens
''' Excluídos, e ninguém olha lá. <b>Não há desfazer visível</b>, então um
''' clique não pode bastar.
'''
''' A confirmação guarda <i>a linha</i>, e não um booleano, para a pergunta
''' poder citar o assunto. E isso tem uma consequência que estes testes
''' prendem: trocar de compromisso, ou trocar de pasta, <b>cancela</b> a
''' confirmação pendente. Sem isso, confirmar apagaria um item que a pergunta
''' não citou — que é o mesmo defeito do "número com dono errado" que a lista
''' de mensagens teve com o total da pasta anterior.
''' </summary>
<TestClass>
Public Class AgendaEscritaTests

    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 9, 1, 9, 10, 0, TimeSpan.FromHours(-3))
    Private Shared ReadOnly Cal As New FolderKey("cal-1", "store-1")
    Private Shared ReadOnly Outra As New FolderKey("cal-2", "store-1")

    ''' <summary>Fonte de leitura com dois compromissos, sempre a mesma.</summary>
    Private NotInheritable Class FonteFixa
        Implements IAgendaSource

        Friend Leituras As Integer

        Public Function GetAppointmentsAsync(folder As FolderKey,
                                             de As DateTimeOffset, ate As DateTimeOffset,
                                             cancel As CancellationToken) _
            As Task(Of OperationResult(Of AppointmentWindow)) _
            Implements IAgendaSource.GetAppointmentsAsync

            Interlocked.Increment(Leituras)

            Dim j As New AppointmentWindow With {
                .De = de, .Ate = ate, .Skipped = 0, .FromRecurrence = 0}
            For i = 1 To 2
                j.Items.Add(New AppointmentInfo With {
                    .Key = New ItemKey($"C-{i}", "store-1"),
                    .Subject = $"compromisso {i}",
                    .Start = Agora.AddHours(i),
                    .End = Agora.AddHours(i + 1)})
            Next
            Return Task.FromResult(OperationResult(Of AppointmentWindow).Ok(j))
        End Function
    End Class

    ''' <summary>Escritor de teste: conta, guarda o que recebeu, e recusa se pedirem.</summary>
    Private NotInheritable Class EscritorDeTeste
        Implements IAgendaWriter

        Friend ReadOnly Chamadas As New List(Of String)()
        Friend UltimaPasta As FolderKey
        Friend UltimoRascunho As AppointmentDraft
        Friend UltimaChave As AppointmentKey
        Friend Recusa As String
        Friend Trava As TaskCompletionSource(Of Boolean)

        Public Async Function CreateAppointmentAsync(folder As FolderKey, rascunho As AppointmentDraft,
                                                     cancel As CancellationToken) _
            As Task(Of OperationResult(Of AppointmentInfo)) _
            Implements IAgendaWriter.CreateAppointmentAsync

            Chamadas.Add("create")
            UltimaPasta = folder
            UltimoRascunho = rascunho
            If Trava IsNot Nothing Then Await Trava.Task
            If Recusa IsNot Nothing Then
                Return OperationResult(Of AppointmentInfo).Fail(ErrorKind.Denied, Recusa)
            End If
            Return OperationResult(Of AppointmentInfo).Ok(
                New AppointmentInfo With {.Key = New ItemKey("novo", "store-1"),
                                          .Subject = rascunho.Subject,
                                          .Start = rascunho.De, .End = rascunho.Ate})
        End Function

        Public Function UpdateAppointmentAsync(chave As AppointmentKey, rascunho As AppointmentDraft,
                                               cancel As CancellationToken) _
            As Task(Of OperationResult(Of AppointmentInfo)) _
            Implements IAgendaWriter.UpdateAppointmentAsync
            Chamadas.Add("update")
            Throw New NotSupportedException("nao usado neste teste")
        End Function

        Public Function DeleteAppointmentAsync(chave As AppointmentKey, cancel As CancellationToken) _
            As Task(Of OperationResult(Of Boolean)) _
            Implements IAgendaWriter.DeleteAppointmentAsync

            Chamadas.Add("delete")
            UltimaChave = chave
            If Recusa IsNot Nothing Then
                Return Task.FromResult(OperationResult(Of Boolean).Fail(ErrorKind.Denied, Recusa))
            End If
            Return Task.FromResult(OperationResult(Of Boolean).Ok(True))
        End Function
    End Class

    Private Shared Function Montar(fonte As FonteFixa, escritor As IAgendaWriter) As AgendaViewModel
        Return New AgendaViewModel(fonte, Function() Agora, escritor)
    End Function

    Private Shared Async Function Carregada(fonte As FonteFixa, escritor As IAgendaWriter) As Task(Of AgendaViewModel)
        Dim vm = Montar(fonte, escritor)
        vm.Apontar(Cal)
        Await vm.CarregarAsync()
        Return vm
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>CONTROLE: sem escritor a agenda continua inteira, e só de leitura.</b>
    '''
    ''' A Fase 6 entregou a leitura primeiro, e ela tem de continuar montável
    ''' sozinha — senão todo duplo de leitura passaria a precisar implementar
    ''' escrita, que é exatamente o que a porta estreita evita.
    ''' </summary>
    <TestMethod>
    Public Async Function Sem_escritor_a_agenda_le_e_nao_escreve() As Task
        Dim fonte As New FonteFixa()
        Dim vm As New AgendaViewModel(fonte, Function() Agora)
        vm.Apontar(Cal)
        Await vm.CarregarAsync()

        Assert.AreEqual(2, vm.Compromissos.Count, "controle: a leitura funciona sem escritor")
        Assert.IsFalse(vm.PodeEscrever)
        Assert.IsFalse(vm.PodeCriar)
        Assert.IsFalse(vm.CriarCommand.CanExecute(Nothing))
    End Function

    ''' <summary>
    ''' <b>Criar manda o rascunho para a PASTA ABERTA, e recarrega.</b>
    '''
    ''' Recarregar em vez de enfiar na lista o que eu acho que foi gravado: o
    ''' calendário é quem manda na tela, e reler é mais barato que confiar.
    ''' </summary>
    <TestMethod>
    Public Async Function Criar_manda_para_a_pasta_aberta_e_recarrega() As Task
        Dim fonte As New FonteFixa()
        Dim escritor As New EscritorDeTeste()
        Dim vm = Await Carregada(fonte, escritor)

        Dim leiturasAntes = fonte.Leituras
        vm.NovoAssunto = "Conversa com o Marcos"
        vm.NovaDuracaoMinutos = 45
        Await vm.CriarCommand.ExecuteAsync(Nothing)

        CollectionAssert.Contains(escritor.Chamadas, "create")
        Assert.AreEqual(Cal, escritor.UltimaPasta, "criou em outra pasta")
        Assert.AreEqual("Conversa com o Marcos", escritor.UltimoRascunho.Subject)
        Assert.AreEqual(45, (escritor.UltimoRascunho.Ate - escritor.UltimoRascunho.De).TotalMinutes)

        Assert.AreEqual(leiturasAntes + 1, fonte.Leituras, "nao recarregou depois de criar")
        Assert.AreEqual("", vm.NovoAssunto, "o formulario nao limpou")
        Assert.IsFalse(vm.TemErro)
    End Function

    ''' <summary>
    ''' <b>A recusa do escritor aparece com as palavras dele.</b>
    '''
    ''' A validação — assunto vazio, fim antes do início, e a recusa de reunião
    ''' — mora no <c>CalendarWriting</c>. A tela <b>não</b> a repete: guarda
    ''' duplicada é guarda que ninguém prova. Então o que ela precisa fazer é
    ''' mostrar o motivo que veio, e não um genérico.
    ''' </summary>
    <TestMethod>
    Public Async Function A_recusa_do_escritor_aparece_na_tela() As Task
        Dim fonte As New FonteFixa()
        Dim escritor As New EscritorDeTeste() With {
            .Recusa = "este compromisso é uma reunião com participantes"}
        Dim vm = Await Carregada(fonte, escritor)

        vm.NovoAssunto = "qualquer coisa"
        Await vm.CriarCommand.ExecuteAsync(Nothing)

        Assert.IsTrue(vm.TemErro)
        StringAssert.Contains(vm.Erro, "reunião com participantes",
            "a tela engoliu o motivo e mostrou um generico: " & vm.Erro)
    End Function

    ''' <summary>
    ''' <b>APAGAR PEDE CONFIRMAÇÃO — e o comando não executa antes dela.</b>
    '''
    ''' <b>Controle negativo:</b> deixando o <c>ApagarCommand</c> executável
    ''' sem confirmação, a primeira asserção cai — e um clique passaria a
    ''' apagar.
    ''' </summary>
    <TestMethod>
    Public Async Function Apagar_exige_confirmacao() As Task
        Dim fonte As New FonteFixa()
        Dim escritor As New EscritorDeTeste()
        Dim vm = Await Carregada(fonte, escritor)

        vm.Selecionado = vm.Compromissos(0)

        Assert.IsFalse(vm.ApagarCommand.CanExecute(Nothing),
            "dava para apagar sem confirmar")
        CollectionAssert.DoesNotContain(escritor.Chamadas, "delete")

        vm.PedirExclusaoCommand.Execute(Nothing)

        Assert.IsTrue(vm.EstaConfirmando)
        StringAssert.Contains(vm.PerguntaDaExclusao, "compromisso 1",
            "a pergunta nao cita o compromisso: " & vm.PerguntaDaExclusao)
        Assert.IsTrue(vm.ApagarCommand.CanExecute(Nothing))

        Await vm.ApagarCommand.ExecuteAsync(Nothing)

        CollectionAssert.Contains(escritor.Chamadas, "delete")
        Assert.AreEqual("C-1", escritor.UltimaChave.Item.EntryId, "apagou o compromisso errado")
        Assert.IsFalse(vm.EstaConfirmando)
    End Function

    ''' <summary>
    ''' <b>TROCAR DE COMPROMISSO CANCELA A CONFIRMAÇÃO.</b>
    '''
    ''' Sem isto, a sequência "pergunta sobre o 1, clico no 2, confirmo"
    ''' apagaria o 2 — enquanto a pergunta na tela citava o 1. É o dano do
    ''' número com dono errado, numa operação que não desfaz.
    ''' </summary>
    <TestMethod>
    Public Async Function Trocar_de_compromisso_cancela_a_confirmacao() As Task
        Dim fonte As New FonteFixa()
        Dim escritor As New EscritorDeTeste()
        Dim vm = Await Carregada(fonte, escritor)

        vm.Selecionado = vm.Compromissos(0)
        vm.PedirExclusaoCommand.Execute(Nothing)
        Assert.IsTrue(vm.EstaConfirmando, "controle: a confirmacao comecou")

        vm.Selecionado = vm.Compromissos(1)

        Assert.IsFalse(vm.EstaConfirmando,
            "a confirmacao sobreviveu a troca, e apagaria o item que a pergunta nao citou")
        Assert.IsFalse(vm.ApagarCommand.CanExecute(Nothing))
    End Function

    ''' <summary>
    ''' <b>TROCAR DE PASTA CANCELA A CONFIRMAÇÃO.</b>
    '''
    ''' Pior que a troca de compromisso: aqui o item confirmado nem está mais
    ''' na tela, e apagá-lo seria mexer noutro calendário.
    ''' </summary>
    <TestMethod>
    Public Async Function Trocar_de_pasta_cancela_a_confirmacao() As Task
        Dim fonte As New FonteFixa()
        Dim escritor As New EscritorDeTeste()
        Dim vm = Await Carregada(fonte, escritor)

        vm.Selecionado = vm.Compromissos(0)
        vm.PedirExclusaoCommand.Execute(Nothing)
        Assert.IsTrue(vm.EstaConfirmando, "controle")

        vm.Apontar(Outra)

        Assert.IsFalse(vm.EstaConfirmando, "a confirmacao atravessou a troca de calendario")
        Assert.IsNull(vm.Selecionado)
    End Function

    ''' <summary>
    ''' <b>O SEGUNDO CLIQUE NÃO VIRA O SEGUNDO COMPROMISSO.</b>
    '''
    ''' Mutação não tem retry neste projeto porque criar não é idempotente. Uma
    ''' tela que deixe clicar duas vezes desfaz essa garantia por fora — e o
    ''' resultado é um compromisso duplicado na agenda de verdade.
    '''
    ''' <b>Controle negativo:</b> tirando o <c>_gravando</c> do
    ''' <c>PodeCriar</c>, a asserção do final cai.
    ''' </summary>
    <TestMethod>
    Public Async Function O_segundo_clique_nao_cria_de_novo() As Task
        Dim fonte As New FonteFixa()
        Dim escritor As New EscritorDeTeste() With {
            .Trava = New TaskCompletionSource(Of Boolean)()}
        Dim vm = Await Carregada(fonte, escritor)

        vm.NovoAssunto = "uma vez só"
        Dim emVoo = vm.CriarCommand.ExecuteAsync(Nothing)

        Assert.IsFalse(vm.PodeCriar, "o botao continuou habilitado com gravacao em voo")
        Assert.IsFalse(vm.CriarCommand.CanExecute(Nothing))

        escritor.Trava.SetResult(True)
        Await emVoo

        Assert.AreEqual(1, escritor.Chamadas.Where(Function(c) c = "create").Count(),
            "criou mais de uma vez")
    End Function

End Class

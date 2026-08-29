Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.App.ViewModels
Imports Iris.Core
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A AGENDA NA TELA.</b>
'''
''' ------------------------------------------------------------------
''' <c>CalendarioRealTests</c> prova a leitura contra o Outlook. Este arquivo
''' prova o <b>ciclo de vida</b> — e ele existe porque a revisão externa de
''' 28/08/2026 achou dois furos que só um teste destes pega:
'''
'''   • o <c>Catch</c> conferia só <c>_disposed</c>, então uma leitura VELHA
'''     que falhasse podia limpar a lista boa de uma leitura nova e publicar
'''     erro sobre ela;
'''   • o <c>Finally</c> apagava <c>Carregando</c> em qualquer geração, então
'''     uma leitura velha desligava o indicador da nova — e reabilitava o
'''     comando, permitindo uma terceira leitura concorrente.
'''
''' É a mesma família que o assistente já tinha ensinado: <b>quem limpa o
''' estado do voo é o dono do voo</b>, e não a geração corrente.
''' </summary>
<TestClass>
Public Class AgendaViewModelTests

    Private Shared ReadOnly Cal As New FolderKey("cal-1", "store-1")
    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 28, 9, 0, 0, TimeSpan.Zero)

    ''' <summary>
    ''' Broker que devolve a janela pedida, e dá para segurar no meio.
    '''
    ''' O <c>FakeBroker</c> responde "fora da alçada" para calendário, de
    ''' propósito. Um duplo dedicado aqui é mais honesto que afrouxar aquele:
    ''' os testes que não falam de agenda continuam quebrando se pedirem
    ''' agenda.
    ''' </summary>
    Private NotInheritable Class BrokerDeAgenda
        Implements IAgendaSource

        Friend Trava As TaskCompletionSource(Of Boolean)
        Friend Property Resposta As OperationResult(Of AppointmentWindow)
        Friend Property Explodir As Boolean
        Friend Chamadas As Integer

        Public Async Function GetAppointmentsAsync(folder As FolderKey,
                                                   de As DateTimeOffset, ate As DateTimeOffset,
                                                   cancel As CancellationToken) _
                                                   As Task(Of OperationResult(Of AppointmentWindow)) _
                                                   Implements IAgendaSource.GetAppointmentsAsync
            Interlocked.Increment(Chamadas)
            If Trava IsNot Nothing Then Await Trava.Task
            If Explodir Then Throw New InvalidOperationException("broker explodiu")
            Return Resposta
        End Function
    End Class

    Private Shared Function Janela(quantos As Integer,
                                   Optional series As Integer = 0,
                                   Optional truncada As Boolean = False) _
                                   As OperationResult(Of AppointmentWindow)
        Dim j As New AppointmentWindow With {
            .De = Agora, .Ate = Agora.AddDays(7),
            .FromRecurrence = series, .Skipped = 0,
            .Truncada = truncada,
            .MotivoDoCorte = If(truncada, "a leitura foi interrompida no meio", "")}
        For i = 1 To quantos
            j.Items.Add(New AppointmentInfo With {
                .Key = New ItemKey($"C-{i}", "store-1"),
                .Subject = $"compromisso {i}",
                .Start = Agora.AddHours(i),
                .End = Agora.AddHours(i + 1)})
        Next
        Return OperationResult(Of AppointmentWindow).Ok(j)
    End Function

    Private Shared Function Montar(b As BrokerDeAgenda) As AgendaViewModel
        Return New AgendaViewModel(b, Function() Agora)
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>CONTROLE POSITIVO: a agenda mostra o que leu.</b>
    ''' </summary>
    <TestMethod>
    Public Async Function Controle_a_agenda_mostra_os_compromissos() As Task
        Dim b As New BrokerDeAgenda() With {.Resposta = Janela(3, series:=1)}
        Dim vm = Montar(b)
        vm.Apontar(Cal)

        Await vm.CarregarAsync()

        Assert.AreEqual(3, vm.Compromissos.Count)
        StringAssert.Contains(vm.Resumo, "3 compromisso(s)")
        StringAssert.Contains(vm.Resumo, "1 de séries")

        ' E O NUMERO POSITIVO TAMBEM E "LIDO". A primeira correcao qualificou
        ' so o zero, e sobrou "12 compromissos" numa pasta cuja cobertura
        ' ninguem mediu -- a mesma afirmacao, mais dificil de notar.
        '
        ' A SEQUENCIA, e nao as duas peças soltas: exigir "3 compromisso(s)" e
        ' "lido(s)" em asserções separadas passa com o total sem qualificacao e
        ' o "lido(s)" em outra clausula da frase.
        StringAssert.Contains(vm.Resumo, "3 compromisso(s) lido(s)")
        Assert.IsFalse(vm.TemErro)
        Assert.IsFalse(vm.Carregando)
    End Function

    ''' <summary>
    ''' <b>ZERO COMPROMISSOS NÃO É "NÃO HÁ COMPROMISSOS".</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O CAMINHO INTEIRO QUE A REVISÃO EXTERNA DESENROLOU</b>
    '''
    ''' O comentário da classe já reconhecia que a medição de cobertura só
    ''' alcança o <b>calendário padrão local</b>, e esta agenda abre qualquer
    ''' pasta classificada como calendário — caixa compartilhada, outro store,
    ''' ninguém mediu. O XAML, ao lado, continuava dizendo <i>"por isso ela não
    ''' tem ressalva de cobertura"</i>. E a tela mostrava <c>0 compromisso(s)</c>.
    '''
    ''' Ou seja: numa pasta que ninguém mediu, o aplicativo <b>afirmava
    ''' ausência</b> — que é exatamente o que este projeto proíbe em todo lugar
    ''' menos aqui, porque aqui ninguém tinha olhado.
    '''
    ''' <b>Controle negativo:</b> devolvendo o <c>$"{j.Items.Count}
    ''' compromisso(s)"</c> incondicional, a asserção do <c>0 compromisso(s)</c>
    ''' cai.
    ''' </summary>
    <TestMethod>
    Public Async Function Zero_compromissos_NAO_afirma_ausencia() As Task
        Dim b As New BrokerDeAgenda() With {.Resposta = Janela(0)}
        Dim vm = Montar(b)
        vm.Apontar(Cal)

        Await vm.CarregarAsync()

        Assert.AreEqual(0, vm.Compromissos.Count, "controle: a janela veio vazia")
        StringAssert.Contains(vm.Resumo, "nenhum compromisso LIDO")
        StringAssert.Contains(vm.Resumo, "não é o mesmo que não haver")
        Assert.IsFalse(vm.Resumo.Contains("0 compromisso(s)"),
            "a tela afirma ausencia numa pasta cuja cobertura ninguem mediu")
    End Function

    ''' <summary>
    ''' <b>Lista truncada avisa que está truncada.</b>
    '''
    ''' O campo nasceu em 28/08 porque dois caminhos devolviam sucesso com
    ''' lista incompleta. Um campo que ninguém mostra na tela é o mesmo defeito
    ''' num lugar diferente.
    ''' </summary>
    <TestMethod>
    Public Async Function Lista_truncada_APARECE_no_resumo() As Task
        Dim b As New BrokerDeAgenda() With {.Resposta = Janela(2, truncada:=True)}
        Dim vm = Montar(b)
        vm.Apontar(Cal)

        Await vm.CarregarAsync()

        StringAssert.Contains(vm.Resumo, "LISTA INCOMPLETA")
        StringAssert.Contains(vm.Resumo, "interrompida")
    End Function

    ''' <summary>
    ''' <b>Leitura VELHA que falha não apaga a lista da nova.</b>
    '''
    ''' O furo que a revisão achou. O <c>Catch</c> conferia só o descarte, e
    ''' não a geração — então a falha de um voo vencido limpava
    ''' <c>Compromissos</c> e publicava erro sobre um resultado que estava
    ''' certo na tela.
    '''
    ''' <b>Controle negativo:</b> tirando a conferência de geração do
    ''' <c>Catch</c>, este teste falha.
    ''' </summary>
    <TestMethod>
    Public Async Function Falha_de_leitura_VENCIDA_nao_apaga_a_lista_nova() As Task
        Dim presa As New TaskCompletionSource(Of Boolean)()
        Dim b As New BrokerDeAgenda() With {
            .Resposta = Janela(3),
            .Trava = presa,
            .Explodir = True
        }
        Dim vm = Montar(b)
        vm.Apontar(Cal)

        ' A leitura VELHA comeca e fica presa no broker, programada para
        ' explodir quando for solta.
        Dim velha = vm.CarregarAsync()
        Assert.AreEqual(1, b.Chamadas, "controle: a primeira leitura tinha de ter comecado")

        ' O usuario troca de pasta e volta: geracao NOVA, sem trava e sem
        ' explosao. Ela termina primeiro e enche a lista.
        vm.Apontar(Cal)
        b.Explodir = False
        b.Trava = Nothing
        Await vm.CarregarAsync()
        Assert.AreEqual(3, vm.Compromissos.Count, "controle: a leitura nova encheu a lista")

        ' AGORA a velha e solta, e ela explode. Soltar e completar a MESMA
        ' TaskCompletionSource que ela esta esperando -- trocar o campo do
        ' broker por Nothing nao solta ninguem, e foi assim que a primeira
        ' versao deste teste travou a suite.
        b.Explodir = True
        presa.SetResult(True)
        Await velha

        Assert.AreEqual(3, vm.Compromissos.Count,
            "a falha de uma leitura vencida apagou a lista da leitura nova")
        Assert.IsFalse(vm.TemErro,
            "a falha de uma leitura vencida publicou erro sobre um resultado bom")
    End Function

    ''' <summary>
    ''' <b>Leitura VELHA não desliga o "carregando" da nova.</b>
    '''
    ''' O segundo furo. O <c>Finally</c> apagava <c>Carregando</c> em qualquer
    ''' geração — a tela parava de mostrar progresso com trabalho em voo, e o
    ''' comando voltava a habilitar, permitindo uma terceira leitura
    ''' concorrente.
    '''
    ''' É a mesma lição do assistente: quem limpa o estado do voo é o
    ''' <b>dono do voo</b>.
    '''
    ''' <b>Controle negativo:</b> tirando a conferência de geração do
    ''' <c>Finally</c>, este teste falha.
    ''' </summary>
    <TestMethod>
    Public Async Function Leitura_VENCIDA_nao_desliga_o_carregando_da_nova() As Task
        Dim b As New BrokerDeAgenda() With {
            .Resposta = Janela(2),
            .Trava = New TaskCompletionSource(Of Boolean)()
        }
        Dim vm = Montar(b)
        vm.Apontar(Cal)

        Dim velha = vm.CarregarAsync()
        Assert.IsTrue(vm.Carregando, "controle: a primeira leitura ligou o indicador")

        ' Geração nova, ainda em voo — a segunda também fica presa.
        vm.Apontar(Cal)
        Dim nova = vm.CarregarAsync()
        Assert.IsTrue(vm.Carregando, "controle: a segunda leitura mantem o indicador")

        ' A velha termina primeiro.
        b.Trava.SetResult(True)
        Await velha

        Assert.IsTrue(vm.Carregando,
            "a leitura vencida desligou o indicador enquanto a nova ainda estava em voo")

        Await nova
        Assert.IsFalse(vm.Carregando, "terminada a leitura corrente, o indicador desliga")
    End Function

    ''' <summary>
    ''' <b>Descartar durante a leitura não escreve na tela que saiu.</b>
    ''' </summary>
    <TestMethod>
    Public Async Function Descarte_durante_a_leitura_NAO_publica() As Task
        Dim b As New BrokerDeAgenda() With {
            .Resposta = Janela(4),
            .Trava = New TaskCompletionSource(Of Boolean)()
        }
        Dim vm = Montar(b)
        vm.Apontar(Cal)

        Dim voo = vm.CarregarAsync()
        vm.Dispose()
        b.Trava.SetResult(True)
        Await voo

        Assert.AreEqual(0, vm.Compromissos.Count,
            "a leitura publicou numa agenda ja descartada")
    End Function

    ''' <summary>
    ''' <b>Apontar para nada esvazia.</b>
    '''
    ''' Números sem dono, descrevendo um calendário que ninguém está olhando,
    ''' é o mesmo defeito que o acervo já teve.
    ''' </summary>
    <TestMethod>
    Public Async Function Apontar_para_nada_esvazia() As Task
        Dim b As New BrokerDeAgenda() With {.Resposta = Janela(3)}
        Dim vm = Montar(b)
        vm.Apontar(Cal)
        Await vm.CarregarAsync()
        Assert.AreEqual(3, vm.Compromissos.Count, "controle")

        vm.Apontar(Nothing)

        Assert.AreEqual(0, vm.Compromissos.Count)
        Assert.IsFalse(vm.TemPasta)
        Assert.AreEqual("", vm.Resumo)
    End Function

End Class

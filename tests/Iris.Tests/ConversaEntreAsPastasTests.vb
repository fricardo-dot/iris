Imports System.Linq
Imports System.Threading
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>AS DUAS PONTAS SE ENCONTRAM? — a medição que decide a Fase 3.</b>
'''
''' ------------------------------------------------------------------
''' <b>DUAS SUPOSIÇÕES QUE A FILA INTEIRA APOIA, E NENHUMA MEDIDA</b>
'''
''' A fila responde "quem falou por último nesta conversa, e há quantos dias".
''' Para isso ela junta a Caixa de Entrada com Itens Enviados pela conversa, e
''' ordena por data. Duas coisas precisam ser verdade, e <b>supor</b> qualquer
''' uma delas já me custou um dia inteiro nas colunas da conversa:
'''
''' <list type="number">
''' <item><b>A data existe na mensagem enviada.</b> O cache guarda
''' <c>received_at</c>, e numa mensagem que saiu daqui isso deveria ser a hora
''' do envio. Se vier vazia, a conta de dias sai errada — e sai
''' <i>plausível</i>, que é pior.</item>
''' <item><b>O <c>ConversationID</c> casa entre as duas pastas.</b> Se a sua
''' resposta em Itens Enviados não cair na mesma conversa da mensagem recebida,
''' a fila mostra como pendente <b>tudo o que você já respondeu</b>. É o pior
''' defeito possível para esta tela: ela existe justamente para dizer o que
''' falta.</item>
''' </list>
'''
''' ------------------------------------------------------------------
''' <b>O QUE FALHA AQUI QUER DIZER</b>
'''
''' Que a fila não pode ser construída sobre a junção por conversa como está, e
''' precisa de outra chave — ou de ler o corpo, que é caro e passa pelo portão.
''' Descobrir isso agora custa uma medição; descobrir depois custa uma tela
''' inteira que mente com confiança.
'''
''' Requer Outlook clássico aberto.
''' </summary>
<TestClass>
Public Class ConversaEntreAsPastasTests

    Private Const Amostra As Integer = 120

    <TestMethod, TestCategory("Integracao")>
    Public Async Function Medir_o_encontro_das_duas_pastas() As Task
        Dim broker = Await PagingIntegrationTests.AbrirBrokerAsync()
        If broker Is Nothing Then Return

        Try
            Dim enviados = Await Achar(broker, {"Itens Enviados", "Sent Items"})
            Dim entrada = Await Achar(broker, {"Caixa de Entrada", "Inbox"})
            If enviados Is Nothing OrElse entrada Is Nothing Then
                Assert.Inconclusive("preciso das duas pastas para medir o encontro")
                Return
            End If

            Dim saiu = Await Ler(broker, enviados)
            Dim chegou = Await Ler(broker, entrada)
            Assert.IsTrue(saiu.Count > 0 AndAlso chegou.Count > 0,
                "uma das pastas veio vazia; sem as duas nao ha encontro a medir")

            ' ---- 1. A DATA NA MENSAGEM ENVIADA ----------------------------
            Dim semData = saiu.Where(Function(m) Not m.ReceivedTime.HasValue).Count()

            ' ---- 2. O ENCONTRO PELA CONVERSA ------------------------------
            Dim conversasQueSairam = saiu.Where(Function(m) Not String.IsNullOrEmpty(m.ConversationId)).
                                     Select(Function(m) m.ConversationId).
                                     Distinct(StringComparer.Ordinal).ToList()
            Dim conversasQueChegaram = chegou.Where(Function(m) Not String.IsNullOrEmpty(m.ConversationId)).
                                       Select(Function(m) m.ConversationId).
                                       Distinct(StringComparer.Ordinal).ToHashSet(StringComparer.Ordinal)

            Dim emComum = conversasQueSairam.Where(Function(c) conversasQueChegaram.Contains(c)).Count()

            Dim medida =
                $"Enviados: {saiu.Count} mensagens, {conversasQueSairam.Count} conversas, " &
                $"{semData} sem data. " &
                $"Entrada: {chegou.Count} mensagens, {conversasQueChegaram.Count} conversas. " &
                $"Em comum: {emComum}."

            Assert.AreEqual(0, semData,
                "mensagem enviada sem data: a conta de dias da fila sairia errada " &
                "e sairia plausivel. " & medida)

            ' O NUMERO EM COMUM NAO PODE SER ZERO -- e a prova de que a
            ' juncao por conversa funciona atraves das pastas. Zero aqui quer
            ' dizer que a fila mostraria como pendente tudo o que ja foi
            ' respondido.
            Assert.IsTrue(emComum > 0,
                "NENHUMA conversa aparece nas duas pastas. A juncao por " &
                "ConversationID nao atravessa a fronteira das pastas nesta " &
                "caixa, e a fila nao pode ser construida sobre ela. " & medida)

            Console.WriteLine(medida)
        Finally
            broker.Dispose()
        End Try
    End Function

    Private Shared Async Function Ler(broker As OutlookBroker,
                                      pasta As FolderKey) As Task(Of List(Of MailSummary))
        Dim consulta As New MessageQuery(pasta, MessageSort.ReceivedDesc, 1)
        Dim r = Await broker.GetMessagePageAsync(consulta, Nothing, Amostra, CancellationToken.None)
        Assert.IsTrue(r.Succeeded, $"a leitura da pagina falhou: {r.Kind}")
        Return r.Value.Items.ToList()
    End Function

    ''' <summary>A primeira pasta de correio com um destes nomes, em qualquer store.</summary>
    Private Shared Async Function Achar(broker As OutlookBroker,
                                        nomes As String()) As Task(Of FolderKey)
        Dim stores = Await broker.GetStoresAsync(CancellationToken.None)
        Assert.IsTrue(stores.Succeeded AndAlso stores.Value.Count > 0, "nenhum store")

        For Each store In stores.Value
            Dim filhas = Await broker.GetFolderChildrenAsync(store.RootFolder,
                                                             CancellationToken.None)
            If Not filhas.Succeeded Then Continue For

            For Each f In filhas.Value
                If f.ContentKind <> FolderContentKind.Mail Then Continue For
                For Each nome In nomes
                    If f.Name.StartsWith(nome, StringComparison.OrdinalIgnoreCase) Then
                        Return f.Key
                    End If
                Next
            Next
        Next
        Return Nothing
    End Function

End Class

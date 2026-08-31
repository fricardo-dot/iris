Imports System.Linq
Imports System.Threading
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A DIREÇÃO FUNCIONA NA CAIXA REAL? — a medição que decide a Fase 3.</b>
'''
''' ------------------------------------------------------------------
''' <b>A PERGUNTA</b>
'''
''' Toda a fila de respostas pendentes se apoia numa comparação: o remetente
''' desta mensagem está no conjunto de identidades do dono? Se a resposta for
''' errada, a fila não fica vazia — ela fica <b>cheia de mentira</b>, cobrando
''' do dono respostas a mensagens que ele mesmo escreveu.
'''
''' Três coisas precisam se encontrar, e cada uma foi construída num dia
''' diferente: o <c>SenderEmailAddress</c> lido pela <c>Table</c>, as
''' identidades lidas das contas e do usuário da sessão, e a normalização do
''' modelo. Cada uma está testada sozinha. <b>Nenhum teste as via juntas.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ITENS ENVIADOS</b>
'''
''' É a única pasta onde a resposta certa é conhecida de antemão: tudo ali foi
''' escrito pelo dono. Uma mensagem de Itens Enviados classificada como
''' <c>DoOutro</c> é um erro <b>demonstrável</b>, e não uma opinião.
'''
''' A Caixa de Entrada serve de contraprova: se ela também vier toda
''' <c>Minha</c>, a comparação está casando com qualquer coisa.
'''
''' Requer Outlook clássico aberto.
''' </summary>
<TestClass>
Public Class DirecaoNaCaixaRealTests

    Private Const Amostra As Integer = 40

    <TestMethod>
    Public Async Function A_direcao_acerta_nas_duas_pontas() As Task
        Dim broker = Await PagingIntegrationTests.AbrirBrokerAsync()
        If broker Is Nothing Then Return

        Try
            Dim r = Await broker.GetIdentidadesAsync(CancellationToken.None)
            Assert.IsTrue(r.Succeeded AndAlso r.Value.Count > 0,
                "sem identidades nao ha direcao nenhuma a medir")
            Dim eu As New MinhasIdentidades(r.Value)

            Dim enviados = Await AcharAsync(broker, {"Itens Enviados", "Sent Items"})
            If enviados Is Nothing Then
                Assert.Inconclusive("nao achei Itens Enviados no store padrao")
                Return
            End If

            Dim minhas = Await ContarAsync(broker, enviados, eu)
            Dim entrada = Await AcharAsync(broker, {"Caixa de Entrada", "Inbox"})
            Dim recebidas = Await ContarAsync(broker, entrada, eu)

            Dim medida =
                $"Itens Enviados: {minhas.Minha} minhas, {minhas.DoOutro} de outros, " &
                $"{minhas.Desconhecida} nao sei (de {minhas.Total}). " &
                $"Caixa de Entrada: {recebidas.Minha} minhas, {recebidas.DoOutro} de outros, " &
                $"{recebidas.Desconhecida} nao sei (de {recebidas.Total})."

            Assert.IsTrue(minhas.Total > 0, "Itens Enviados vazia nao mede nada. " & medida)

            ' A PONTA QUE IMPORTA. Tudo em Itens Enviados foi escrito pelo
            ' dono, entao "de outros" ali e erro demonstravel -- e cada um
            ' desses vira uma linha na fila cobrando dele uma resposta que ele
            ' ja deu.
            Assert.AreEqual(0, minhas.DoOutro,
                "mensagem escrita pelo dono classificada como sendo de outra " &
                "pessoa: falta uma forma do endereco dele no conjunto de " &
                "identidades. " & medida)

            Assert.IsTrue(minhas.Minha > 0,
                "NENHUMA mensagem de Itens Enviados foi reconhecida como do " &
                "dono. As tres pecas -- endereco lido, identidades semeadas e " &
                "normalizacao -- nao se encontram. " & medida)

            ' A CONTRAPROVA. Sem ela, uma comparacao que respondesse Minha para
            ' tudo passaria nas duas assercoes acima.
            If recebidas.Total > 0 Then
                Assert.IsTrue(recebidas.DoOutro > 0,
                    "a Caixa de Entrada tambem veio toda 'minha': a comparacao " &
                    "esta casando com qualquer coisa. " & medida)
            End If

            Console.WriteLine(medida)
        Finally
            broker.Dispose()
        End Try
    End Function

    Private NotInheritable Class Contagem
        Public Property Total As Integer
        Public Property Minha As Integer
        Public Property DoOutro As Integer
        Public Property Desconhecida As Integer
    End Class

    Private Shared Async Function ContarAsync(broker As OutlookBroker, pasta As FolderKey,
                                              eu As MinhasIdentidades) As Task(Of Contagem)
        Dim c As New Contagem()
        If pasta Is Nothing Then Return c

        Dim consulta As New MessageQuery(pasta, MessageSort.ReceivedDesc, 1)
        Dim r = Await broker.GetMessagePageAsync(consulta, Nothing, Amostra, CancellationToken.None)
        Assert.IsTrue(r.Succeeded, $"a leitura da pagina falhou: {r.Kind}")

        For Each m In r.Value.Items
            c.Total += 1
            Select Case eu.DirecaoDe(m.SenderAddress)
                Case Direcao.Minha : c.Minha += 1
                Case Direcao.DoOutro : c.DoOutro += 1
                Case Else : c.Desconhecida += 1
            End Select
        Next
        Return c
    End Function

    ''' <summary>A primeira pasta de correio do store padrão com um destes nomes.</summary>
    Private Shared Async Function AcharAsync(broker As OutlookBroker,
                                             nomes As String()) As Task(Of FolderKey)
        Dim stores = Await broker.GetStoresAsync(CancellationToken.None)
        Assert.IsTrue(stores.Succeeded AndAlso stores.Value.Count > 0, "nenhum store")

        Dim filhas = Await broker.GetFolderChildrenAsync(stores.Value(0).RootFolder,
                                                         CancellationToken.None)
        Assert.IsTrue(filhas.Succeeded, "GetFolderChildrenAsync falhou")

        For Each f In filhas.Value
            If f.ContentKind <> FolderContentKind.Mail Then Continue For
            For Each nome In nomes
                If f.Name.StartsWith(nome, StringComparison.OrdinalIgnoreCase) Then Return f.Key
            Next
        Next
        Return Nothing
    End Function

End Class

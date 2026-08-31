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
''' É a única pasta onde a resposta certa é conhecida de antemão: tudo ali saiu
''' desta caixa. Uma mensagem de Itens Enviados classificada como
''' <c>DoOutro</c> é um erro <b>demonstrável</b>, e não uma opinião.
'''
''' <b>Com uma ressalva que este teste não cobre:</b> numa caixa compartilhada,
''' ou com permissão de <i>enviar como</i>, uma mensagem de Itens Enviados pode
''' legitimamente ter sido escrita em nome de outra pessoa — e aí
''' <c>DoOutro</c> seria a resposta certa. São exatamente os casos que
''' <c>IdentidadesDoDono</c> declara não saber cobrir. Nesta caixa a asserção
''' vale; noutra ela pode falhar por um motivo que não é defeito, e a mensagem
''' de falha diz onde procurar.
'''
''' A Caixa de Entrada serve de contraprova: se ela também vier toda
''' <c>Minha</c>, a comparação está casando com qualquer coisa.
'''
''' Requer Outlook clássico aberto.
''' </summary>
<TestClass>
Public Class DirecaoNaCaixaRealTests

    Private Const Amostra As Integer = 40

    <TestMethod, TestCategory("Integracao")>
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
                Assert.Inconclusive("nao achei Itens Enviados em store nenhum")
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

            ' A CONTRAPROVA, E ELA NAO PODE SER PULADA EM SILENCIO.
            '
            ' Sem ela, uma comparacao que respondesse Minha para tudo passaria
            ' nas duas assercoes acima. Ela estava condicionada a "se a Caixa
            ' de Entrada tiver mensagens" -- e caixa vazia, ou pasta com outro
            ' nome, desligava a contraprova sem uma linha dizendo isso. E o
            ' bloqueio que nunca bloqueia, de novo.
            If recebidas.Total = 0 Then
                Assert.Inconclusive(
                    "sem Caixa de Entrada legivel nao ha contraprova, e sem " &
                    "contraprova este teste nao vale: uma comparacao que " &
                    "respondesse Minha para tudo passaria. " & medida)
                Return
            End If

            Assert.IsTrue(recebidas.DoOutro > 0,
                "a Caixa de Entrada tambem veio toda minha: a comparacao esta " &
                "casando com qualquer coisa. " & medida)

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

    ''' <summary>
    ''' A primeira pasta de correio com um destes nomes, <b>em qualquer
    ''' store</b>.
    '''
    ''' A versão anterior olhava só <c>stores.Value(0)</c> e chamava aquilo de
    ''' "store padrão". Nada prova que o primeiro seja o padrão: num perfil com
    ''' PST, arquivo morto ou caixa compartilhada a ordem é outra, e a medição
    ''' sairia da caixa errada — ou não acharia pasta nenhuma e o teste ficaria
    ''' inconclusivo por um motivo falso.
    ''' </summary>
    Private Shared Async Function AcharAsync(broker As OutlookBroker,
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

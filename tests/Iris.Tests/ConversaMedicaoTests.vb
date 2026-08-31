Imports System.Linq
Imports System.Threading
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A CONVERSA E O ENDEREÇO VÊM MESMO? — medição, e não suposição.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ISTO PRECISA SER MEDIDO</b>
'''
''' O caminho rápido de listagem lê por <c>Table</c>, e a <c>Table</c> do
''' Outlook aceita uma coluna ou recusa — não há documento que valha mais que a
''' resposta da máquina. E a recusa aqui é <b>silenciosa por desenho</b>: as
''' três colunas somem juntas e a listagem continua funcionando, porque derrubar
''' a lista de e-mails inteira para não perder um campo de fila seria a troca
''' errada.
'''
''' Silêncio por desenho precisa de alguém perguntando. Sem este teste, o
''' cenário provável não é erro: é o cache ganhando três colunas que nunca se
''' preenchem, a fila da Fase 3 saindo vazia, e a explicação estando a quatro
''' camadas de distância.
'''
''' Foi exatamente assim que <c>MessageClass</c> passou meses gravando uma
''' constante: 1.123 linhas, <b>uma</b> classe distinta, um número bonito demais
''' para ser medição.
'''
''' ------------------------------------------------------------------
''' <b>O QUE FALHA AQUI QUER DIZER</b>
'''
''' Não quer dizer defeito no código: quer dizer que <b>nesta caixa</b> o
''' caminho por <c>Table</c> não entrega conversa, e a Fase 3 vai precisar da
''' leitura por item — que é mais cara e existe. Saber disso antes de construir
''' a fila vale mais do que descobrir depois que ela está sempre vazia.
'''
''' Requer Outlook clássico aberto, como os outros testes contra a caixa real.
''' </summary>
<TestClass>
Public Class ConversaMedicaoTests

    <TestMethod>
    Public Async Function Medir_a_conversa_e_o_endereco_na_caixa_real() As Task
        Dim broker = Await PagingIntegrationTests.AbrirBrokerAsync()
        If broker Is Nothing Then Return

        Try
            Dim pasta = Await PagingIntegrationTests.AcharEntradaAsync(broker)
            Dim consulta As New MessageQuery(pasta, MessageSort.ReceivedDesc, 1)

            Dim r = Await broker.GetMessagePageAsync(consulta, Nothing, 40, CancellationToken.None)
            Assert.IsTrue(r.Succeeded, $"a leitura da pagina falhou: {r.Kind}")

            Dim linhas = r.Value.Items
            Assert.IsTrue(linhas.Count > 0,
                "caixa de entrada vazia nao mede nada; a medicao precisa de amostra")

            Dim comConversa = linhas.Where(Function(m) Not String.IsNullOrEmpty(m.ConversationId)).Count()
            Dim comIndice = linhas.Where(Function(m) Not String.IsNullOrEmpty(m.ConversationIndex)).Count()
            Dim comEndereco = linhas.Where(Function(m) Not String.IsNullOrEmpty(m.SenderAddress)).Count()
            Dim conversasDistintas = linhas.Where(Function(m) Not String.IsNullOrEmpty(m.ConversationId)).
                                     Select(Function(m) m.ConversationId).Distinct().Count()

            Dim medida =
                $"{linhas.Count} linhas: {comConversa} com conversa, {comIndice} com indice, " &
                $"{comEndereco} com endereco, {conversasDistintas} conversas distintas."

            Assert.IsTrue(comConversa > 0,
                "NENHUMA linha trouxe ConversationID. O caminho por Table recusou as " &
                "tres colunas nesta caixa, e caiu no conjunto de oito — que e o " &
                "comportamento desenhado, e nao um defeito. Consequencia: a fila da " &
                "Fase 3 precisa da leitura por item. " & medida)

            Assert.IsTrue(comEndereco > 0,
                "nenhuma linha trouxe SenderEmailAddress, e sem ele a direcao da " &
                "mensagem e sempre 'nao sei'. " & medida)

            ' A conversa NAO PODE SER UMA SO. Uma pagina inteira caindo numa
            ' conversa unica e o sintoma de a coluna estar vindo constante --
            ' foi assim que o MessageClass enganou por meses, com 1.123 linhas
            ' e uma classe distinta.
            Assert.IsTrue(conversasDistintas > 1,
                "todas as linhas caem na MESMA conversa: isso e coluna constante, " &
                "e nao conversa. " & medida)

            ' A medida vai para a saida do teste mesmo quando ele passa: e o
            ' numero que decide se a fila da Fase 3 pode confiar na varredura.
            Console.WriteLine(medida)
        Finally
            broker.Dispose()
        End Try
    End Function

End Class

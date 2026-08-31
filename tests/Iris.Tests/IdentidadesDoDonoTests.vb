Imports System.Linq
Imports System.Threading
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>AS IDENTIDADES DO DONO, NA CAIXA REAL — medição.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTA MEDIÇÃO DECIDE</b>
'''
''' A fila de respostas pendentes compara o remetente de cada mensagem com o
''' conjunto de identidades do dono. Se o conjunto não contiver a forma em que
''' o remetente chega, <b>as mensagens do próprio dono aparecem como sendo de
''' terceiros</b> — e a fila cobra dele respostas que ele já deu.
'''
''' Numa organização Exchange essa forma é X.500, e não SMTP. Semear só o SMTP
''' e descobrir depois seria o mesmo erro das colunas da conversa, que passaram
''' um dia inteiro vindo vazias porque eu pedi pelo nome errado.
'''
''' ------------------------------------------------------------------
''' <b>O QUE FALHA AQUI QUER DIZER</b>
'''
''' Que a semeadura automática não basta nesta caixa, e o arquivo
''' <c>identidades.txt</c> vai precisar de correção à mão — que é justamente o
''' motivo de ele ser um arquivo editável, e não uma decisão do programa.
'''
''' Requer Outlook clássico aberto, como os outros testes contra a caixa real.
''' </summary>
<TestClass>
Public Class IdentidadesDoDonoTests

    <TestMethod, TestCategory("Integracao")>
    Public Async Function Medir_as_identidades_na_caixa_real() As Task
        Dim broker = Await PagingIntegrationTests.AbrirBrokerAsync()
        If broker Is Nothing Then Return

        Try
            Dim r = Await broker.GetIdentidadesAsync(CancellationToken.None)
            Assert.IsTrue(r.Succeeded, $"a leitura das identidades falhou: {r.Kind}")

            Dim achadas = r.Value
            Dim comArroba = achadas.Where(Function(e) e.Contains("@"c)).Count()
            Dim comX500 = achadas.Where(
                Function(e) e.Contains("/o=", StringComparison.OrdinalIgnoreCase)).Count()

            ' O ENDERECO INTEIRO NAO VAI PARA A SAIDA. Ele identifica uma
            ' pessoa real, e a mesma regra do EntryID vale aqui: so a forma e
            ' a contagem, nunca o valor.
            Dim medida = $"{achadas.Count} identidades: {comArroba} com arroba, " &
                         $"{comX500} em X.500."

            Assert.IsTrue(achadas.Count > 0,
                "nenhuma identidade veio do Outlook. A semeadura nao vai " &
                "acontecer, o conjunto fica vazio, e a fila responde 'nao sei' " &
                "para toda mensagem. " & medida)

            Assert.IsTrue(comArroba > 0,
                "nenhum endereco SMTP: mensagem externa do proprio dono nao " &
                "seria reconhecida. " & medida)

            ' O X.500 NAO E EXIGIDO. Caixa que nao e Exchange nao tem essa
            ' forma, e cobra-la faria este teste falhar por ambiente em vez de
            ' por defeito. O que importa e que ele apareca QUANDO existe -- e a
            ' medida abaixo e o que diz se apareceu.
            Console.WriteLine(medida)
        Finally
            broker.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' <b>O que o Outlook devolve entra no conjunto e volta a casar.</b>
    '''
    ''' Não basta ler as identidades: elas passam por <c>Normalizar</c> e pela
    ''' exigência de forma antes de valerem. Uma forma que o Outlook entrega e
    ''' que o modelo rejeita seria um conjunto cheio e inútil — e o defeito
    ''' apareceria longe, como uma fila inteira dizendo "não sei".
    '''
    ''' <b>É reflexivo, e só isso.</b> Ele pega a lista, monta o conjunto com a
    ''' mesma lista, e pergunta se cada item se reconhece. Detecta normalização
    ''' destrutiva e nada mais: <b>não</b> prova que o Outlook coletou as formas
    ''' necessárias, e passaria com a coleta do X.500 removida. Quem prende isso
    ''' é <c>DirecaoNaCaixaRealTests</c>, contra mensagens de verdade.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Async Function O_que_o_Outlook_devolve_CASA_no_modelo() As Task
        Dim broker = Await PagingIntegrationTests.AbrirBrokerAsync()
        If broker Is Nothing Then Return

        Try
            Dim r = Await broker.GetIdentidadesAsync(CancellationToken.None)
            Assert.IsTrue(r.Succeeded, $"a leitura das identidades falhou: {r.Kind}")
            Assert.IsTrue(r.Value.Count > 0, "sem identidade nao ha o que conferir")

            Dim eu As New MinhasIdentidades(r.Value)

            For Each bruta In r.Value
                Assert.AreEqual(Direcao.Minha, eu.DirecaoDe(bruta),
                    "uma identidade que o Outlook devolveu nao se reconhece a si " &
                    "mesma depois de passar pelo modelo: a exigencia de forma " &
                    "esta recusando algo legitimo, e a fila inteira responderia " &
                    "'nao sei'")
            Next
        Finally
            broker.Dispose()
        End Try
    End Function

End Class

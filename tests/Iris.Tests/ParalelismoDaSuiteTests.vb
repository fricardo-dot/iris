Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A FALHA RARA DA SUITE, ou como fechar o que nunca reproduziu.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ACONTECEU</b>
'''
''' Em 25/08/2026 a suite falhou <b>1 vez em 10 execucoes</b>, com
''' <c>Cannot access a disposed object: SQLitePCL.sqlite3</c>, sempre na
''' primeira execucao depois de uma compilacao. A execucao que falhou foi
''' justamente a unica sem <c>trx</c>: o nome do teste, a mensagem e a pilha
''' se perderam.
'''
''' A causa PROVAVEL foi <c>SqliteConnection.ClearAllPools()</c>, que e
''' <b>global</b>: chamada no <c>TestCleanup</c> de uma classe, ela derruba a
''' conexao de uma classe vizinha rodando em paralelo. As dez classes que
''' tocavam SQLite receberam <c>&lt;DoNotParallelize&gt;</c>, e a duracao subiu
''' de 38 s para ~58 s. Nao houve nova ocorrencia em mais de trinta execucoes.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE TRINTA EXECUCOES LIMPAS NAO FECHAM NADA</b>
'''
''' Ausencia de sintoma nao e prova de correcao — e o projeto inteiro e
''' construido sobre essa frase. Uma falha de 1 em 10 que some depois de uma
''' mudanca pode ter sido corrigida, ou pode ter ficado mais rara.
'''
''' O que da para fazer, e e o que este arquivo faz, e transformar a hipotese
''' numa <b>regra imposta</b>: se a causa provavel era paralelismo entre
''' classes que tocam SQLite, entao nenhuma classe que toca SQLite pode voltar
''' a rodar em paralelo — e isso e verificavel a cada execucao, em vez de
''' depender de alguem lembrar.
'''
''' A dívida deixa de ser "explicar a falha" e passa a ser "a causa provavel
''' esta fechada por construcao; se ela voltar, a hipotese estava errada".
''' Isso e menos que uma explicacao, e e honesto sobre ser menos.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE LER O CODIGO-FONTE, E NAO REFLEXAO</b>
'''
''' <c>&lt;DoNotParallelize&gt;</c> e visivel por reflexao; o que NAO e visivel e
''' "esta classe toca SQLite". Isso mora no corpo dos metodos, e descobri-lo
''' por IL seria construir um analisador para responder uma pergunta que o
''' texto responde direto.
'''
''' <c>CallerFilePath</c> da o caminho absoluto DESTE arquivo em tempo de
''' compilacao, e dele sai a pasta da suite. Nao depende de diretorio de
''' trabalho nem de copiar fonte para a saida.
''' </summary>
<TestClass>
Public Class ParalelismoDaSuiteTests

    ''' <summary>
    ''' O que conta como "toca SQLite". Ampla de proposito: um falso positivo
    ''' custa um atributo a mais numa classe, e um falso negativo custa a
    ''' falha rara de volta.
    ''' </summary>
    Private Shared ReadOnly Marcas As String() = {
        "SqliteConnection", "CacheDatabase.Open", "AcervoViewModel.Abrir",
        "New MainViewModel", "CacheWriter", "ClearAllPools"
    }

    Private Shared Function PastaDaSuite(<CallerFilePath> Optional aqui As String = Nothing) As String
        Return Path.GetDirectoryName(aqui)
    End Function

    ''' <summary>
    ''' <b>Toda classe de teste que toca SQLite roda sozinha.</b>
    '''
    ''' <b>Controle negativo:</b> tirando <c>&lt;DoNotParallelize&gt;</c> de
    ''' qualquer uma das classes listadas, este teste falha nomeando o arquivo.
    ''' </summary>
    <TestMethod>
    Public Sub Classe_que_toca_SQLite_NAO_roda_em_paralelo()
        Dim pasta = PastaDaSuite()
        Assert.IsTrue(Directory.Exists(pasta), $"nao achei a pasta da suite: {pasta}")

        Dim arquivos = Directory.GetFiles(pasta, "*.vb").
                       Where(Function(f) Not Path.GetFileName(f).StartsWith("obj")).ToList()
        Assert.IsTrue(arquivos.Count > 20,
            $"controle: esperava dezenas de arquivos na suite, achei {arquivos.Count}")

        Dim faltando As New List(Of String)()
        Dim conferidos = 0

        For Each f In arquivos
            Dim texto = File.ReadAllText(f)
            If Not texto.Contains("<TestClass>") Then Continue For
            If Not Marcas.Any(Function(m) texto.Contains(m)) Then Continue For

            conferidos += 1
            If Not texto.Contains("<DoNotParallelize>") Then
                faltando.Add(Path.GetFileName(f))
            End If
        Next

        ' CONTROLE POSITIVO. Sem ele, um erro no caminho ou nas marcas faria
        ' zero classes serem conferidas e o teste passaria dizendo nada.
        Assert.IsTrue(conferidos >= 10,
            $"controle: esperava ao menos 10 classes tocando SQLite, conferi {conferidos}")

        Assert.AreEqual(0, faltando.Count,
            "classes que tocam SQLite e podem rodar em paralelo — e foi assim que a " &
            "falha rara de 25/08/2026 apareceu: " & String.Join(", ", faltando))
    End Sub

    ''' <summary>
    ''' <b>E o assembly continua paralelizando o resto.</b>
    '''
    ''' O irmao do teste acima, e ele existe para que a regra nao seja
    ''' satisfeita da maneira preguicosa. Desligar o paralelismo do assembly
    ''' inteiro faria o teste acima passar para sempre e custaria minutos a
    ''' cada execucao — trocar um problema real por um imposto permanente.
    ''' </summary>
    <TestMethod>
    Public Sub O_assembly_continua_paralelizando_o_resto()
        Dim pasta = PastaDaSuite()
        Dim settings = Path.Combine(pasta, "MSTestSettings.vb")
        Assert.IsTrue(File.Exists(settings), "MSTestSettings.vb sumiu")

        Dim texto = File.ReadAllText(settings)
        StringAssert.Contains(texto, "Parallelize",
            "o assembly deixou de declarar paralelismo, e a regra acima virou decoracao")
    End Sub

End Class

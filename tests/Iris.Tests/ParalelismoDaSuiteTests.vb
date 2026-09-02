Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
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
''' conexao de uma classe vizinha rodando em paralelo. As classes que tocavam
''' SQLite receberam <c>&lt;DoNotParallelize&gt;</c>, e a duracao subiu de
''' 38 s para ~58 s. Nao houve nova ocorrencia em mais de trinta execucoes.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE TRINTA EXECUCOES LIMPAS NAO FECHAM NADA</b>
'''
''' Ausencia de sintoma nao e prova de correcao. Uma falha de 1 em 10 que
''' some depois de uma mudanca pode ter sido corrigida, ou pode ter ficado
''' mais rara.
'''
''' O que da para fazer, e e o que este arquivo faz, e transformar a hipotese
''' numa <b>regra imposta</b>. A divida deixa de ser "explicar a falha" e
''' passa a ser "a causa provavel esta fechada por construcao; se ela voltar,
''' a hipotese estava errada". Isso e menos que uma explicacao, e esta dito
''' como sendo menos.
'''
''' ------------------------------------------------------------------
''' <b>A PRIMEIRA VERSAO DESTA REGRA TINHA CINCO FUROS</b>
'''
''' A revisao externa de 28/08 listou: nao percorria subpastas; comparava
''' texto sensivel a maiusculas; contava ocorrencia dentro de comentario e de
''' literal; aceitava <c>&lt;DoNotParallelize&gt;</c> em <b>qualquer</b>
''' classe do arquivo como se cobrisse todas; e o teste irmao procurava a
''' substring <c>"Parallelize"</c>, que <c>DoNotParallelize</c> tambem
''' contem — trocar a configuracao por <c>&lt;Assembly: DoNotParallelize&gt;</c>
''' passaria.
'''
''' Uma regra com furo e pior que regra nenhuma, porque quem a le acredita
''' nela. Esta versao trabalha <b>por classe</b>, remove comentarios antes de
''' procurar, percorre subpastas e casa o atributo do assembly por expressao
''' ancorada.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE LER O CODIGO-FONTE, E NAO REFLEXAO</b>
'''
''' <c>&lt;DoNotParallelize&gt;</c> e visivel por reflexao; o que NAO e visivel
''' e "esta classe toca SQLite". Isso mora no corpo dos metodos, e descobri-lo
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
    ''' O que conta como "toca SQLite".
    '''
    ''' As marcas casam <b>construcao ou abertura</b>, e nao mencao. A primeira
    ''' versao usava so o nome do tipo, e cobrava tres classes que nunca abrem
    ''' banco nenhum: <c>ArchitectureTests</c> le
    ''' <c>GetType(...MainViewModel).Assembly</c> por reflexao,
    ''' <c>BindingsDaJanelaTests</c> usa <c>GetType(AcervoViewModel)</c> para
    ''' conferir nomes de propriedade, e <c>ContextoDoOutlookTests</c> chama um
    ''' <c>Shared</c> e LE o arquivo-fonte como texto.
    '''
    ''' Marcar essas tres custaria minutos de serializacao por execucao para
    ''' proteger de um risco que elas nao correm — e regra que cobra onde nao
    ''' faz falta e regra que se aprende a ignorar.
    '''
    ''' <b>O erro que sobra e o oposto, e ele e conhecido:</b> uma classe que
    ''' chegue ao banco por um auxiliar novo, com outro nome, escapa. Quem
    ''' segura isso e o controle positivo la embaixo — se a contagem de classes
    ''' conferidas cair, alguem mexeu nas marcas ou nas classes, e o teste
    ''' para. Nao e garantia; e um alarme.
    ''' </summary>
    ' O NOME DO TIPO, e nao "new " + o nome.
    '
    ' As marcas eram construcoes -- "new resolvedordoacervo" --, e uma classe
    ' que chamasse um HELPER seu para abrir o banco nao casava com nenhuma: o
    ' "new" mora no helper, e o helper pode estar noutra classe do mesmo
    ' arquivo. Pelo TIPO, a declaracao do helper ja denuncia ("As
    ' AcervoDeTodasAsPastas"), e um "Dim x = Abrir()" com inferencia continua
    ' sendo pego pela assinatura de Abrir.
    '
    ' Nao e prova -- nada aqui e. E um alarme com a rede mais fechada, e o que
    ' ele ainda deixa passar e uma classe que so receba o banco pronto por
    ' parametro de um tipo que nao esta nesta lista. Achado por revisao externa
    ' em 01/09/2026.
    '
    ' EXCECAO: CacheDatabase continua sendo "cachedatabase.open", e nao o tipo
    ' nu. Ele aparece em ASSINATURA sem ser usado -- ArchitectureTests tem uma
    ' isca "IscaQueUsaOManifestReader(db As CacheDatabase)" que nunca e chamada,
    ' e que existe so para um teste de reflexao ter um uso plantado. Marca-la
    ' obrigaria a por o atributo onde ele nao faz falta, e isso ensina a por o
    ' atributo em toda parte -- que e como uma regra deixa de significar algo.
    '
    ' (Comentario AQUI e nao dentro das chaves: a virgula continua a expressao,
    ' e um comentario no meio quebra a continuacao. Oitava vez nesta base.)
    Private Shared ReadOnly Marcas As String() = {
        "sqliteconnection", "sqlitecommand", "clearallpools",
        "acervoviewmodel.abrir", "new mainviewmodel",
        "cachedatabase.open", "cachewriter", "sqlitesweepsink",
        "sqlitedisclosurejournal", "resolvedordoacervo", "varreduradapasta",
        "rotulosnocache", "acervodetodasaspastas", "publicationdrain",
        "acervoservice", "classificarumapasta", "buscanoacervo"
    }

    Private Shared Function PastaDaSuite(<CallerFilePath> Optional aqui As String = Nothing) As String
        Return Path.GetDirectoryName(aqui)
    End Function

    ''' <summary>
    ''' Tira comentarios de linha antes da busca.
    '''
    ''' Sem isto, esta propria classe se acusaria: ela CITA
    ''' <c>SqliteConnection</c> e <c>ClearAllPools</c> na documentacao, e nao
    ''' toca em banco nenhum. Uma regra que nao distingue codigo de prosa
    ''' obriga a colocar atributo onde ele nao faz falta, e isso ensina a
    ''' ignorar a regra.
    '''
    ''' Aproximacao deliberada: nao entende literal de string com aspas. Uma
    ''' classe que so mencione "sqlite" dentro de um literal vai ser cobrada,
    ''' e cobrar demais e o lado seguro deste erro.
    ''' </summary>
    Private Shared Function SemComentario(texto As String) As String
        Dim sb As New Text.StringBuilder()
        For Each linha In texto.Split(ChrW(10))
            Dim limpa = SemTexto(linha)
            Dim corte = limpa.IndexOf("'"c)
            sb.AppendLine(If(corte >= 0, limpa.Substring(0, corte), limpa))
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' <b>Tira o conteúdo das aspas antes de procurar as marcas.</b>
    '''
    ''' Pelo mesmo motivo que os comentários saem: <c>ArchitectureTests</c> cita
    ''' <c>AcervoService</c> dentro de uma <i>mensagem de asserção</i> — prosa que
    ''' o compilador guarda como texto — e não toca em banco nenhum. Uma regra que
    ''' não distingue código de prosa obriga a pôr o atributo onde ele não faz
    ''' falta, e isso ensina a pô-lo em toda parte, que é como uma regra deixa de
    ''' significar alguma coisa.
    '''
    ''' As aspas ficam no lugar do texto: tirar as duas juntas colaria o que vem
    ''' antes no que vem depois e poderia <i>fabricar</i> uma marca.
    ''' </summary>
    Private Shared Function SemTexto(linha As String) As String
        Dim sb As New Text.StringBuilder()
        Dim dentro = False
        For Each c In linha
            If c = """"c Then
                dentro = Not dentro
                sb.Append(c)
            ElseIf Not dentro Then
                sb.Append(c)
            End If
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Parte o arquivo em blocos de classe, para que o atributo de uma nao
    ''' seja creditado a outra. Arquivos da suite tem uma ou duas classes, e
    ''' <c>FakeBroker.vb</c> tem tres — creditar por arquivo era o quarto
    ''' furo da versao anterior.
    ''' </summary>
    Private Shared Iterator Function Classes(texto As String) _
                                             As IEnumerable(Of (Nome As String, Corpo As String))
        Dim marcas = Regex.Matches(texto, "^\s*<TestClass>",
                                   RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        For i = 0 To marcas.Count - 1
            Dim ini = marcas(i).Index
            Dim fim = If(i + 1 < marcas.Count, marcas(i + 1).Index, texto.Length)
            Dim corpo = texto.Substring(ini, fim - ini)
            Dim nome = Regex.Match(corpo, "(?:Class|Module)\s+(\w+)", RegexOptions.IgnoreCase)
            Yield (If(nome.Success, nome.Groups(1).Value, "?"), corpo)
        Next
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>Toda classe de teste que toca SQLite roda sozinha.</b>
    '''
    ''' <b>Controle negativo:</b> tirando <c>&lt;DoNotParallelize&gt;</c> de
    ''' qualquer uma das classes, este teste falha nomeando arquivo e classe.
    ''' </summary>
    <TestMethod>
    Public Sub Classe_que_toca_SQLite_NAO_roda_em_paralelo()
        Dim pasta = PastaDaSuite()
        Assert.IsTrue(Directory.Exists(pasta), $"nao achei a pasta da suite: {pasta}")

        ' SUBPASTAS TAMBEM. A versao anterior so olhava o primeiro nivel, e
        ' bastava mover um arquivo para uma pasta para sair da regra.
        Dim arquivos = Directory.GetFiles(pasta, "*.vb", SearchOption.AllDirectories).
                       Where(Function(f) Not f.Contains("\obj\") AndAlso Not f.Contains("\bin\")).
                       ToList()
        Assert.IsTrue(arquivos.Count > 20,
            $"controle: esperava dezenas de arquivos na suite, achei {arquivos.Count}")

        Dim faltando As New List(Of String)()
        Dim conferidas = 0

        For Each f In arquivos
            Dim codigo = SemComentario(File.ReadAllText(f)).ToLowerInvariant()
            For Each c In Classes(codigo)
                If Not Marcas.Any(Function(m) c.Corpo.Contains(m)) Then Continue For
                conferidas += 1
                If Not c.Corpo.Contains("<donotparallelize>") Then
                    faltando.Add($"{Path.GetFileName(f)}:{c.Nome}")
                End If
            Next
        Next

        ' CONTROLE POSITIVO. Sem ele, um erro no caminho, nas marcas ou na
        ' quebra por classe faria zero classes serem conferidas e o teste
        ' passaria dizendo nada.
        Assert.IsTrue(conferidas >= 10,
            $"controle: esperava ao menos 10 classes tocando SQLite, conferi {conferidas}")

        Assert.AreEqual(0, faltando.Count,
            "classes que tocam SQLite e podem rodar em paralelo — e foi assim que a " &
            "falha rara de 25/08/2026 apareceu: " & String.Join(", ", faltando))
    End Sub

    ''' <summary>
    ''' <b>E o assembly continua paralelizando o resto.</b>
    '''
    ''' O irmao do teste acima, e ele existe para que a regra nao seja
    ''' satisfeita da maneira preguicosa: desligar o paralelismo do assembly
    ''' inteiro faria o teste acima passar para sempre e custaria minutos a
    ''' cada execucao.
    '''
    ''' A busca e <b>ancorada</b>, e nao por substring. A versao anterior
    ''' procurava <c>"Parallelize"</c>, e <c>DoNotParallelize</c> contem essa
    ''' substring — trocar a configuracao por
    ''' <c>&lt;Assembly: DoNotParallelize&gt;</c> passaria no teste que existe
    ''' justamente para impedir isso.
    ''' </summary>
    <TestMethod>
    Public Sub O_assembly_continua_paralelizando_o_resto()
        Dim pasta = PastaDaSuite()
        Dim settings = Path.Combine(pasta, "MSTestSettings.vb")
        Assert.IsTrue(File.Exists(settings), "MSTestSettings.vb sumiu")

        Dim codigo = SemComentario(File.ReadAllText(settings))

        Assert.IsTrue(Regex.IsMatch(codigo, "<\s*Assembly\s*:\s*Parallelize\s*\("),
            "o assembly deixou de declarar Parallelize, e a regra irma virou decoracao")
        Assert.IsFalse(Regex.IsMatch(codigo, "<\s*Assembly\s*:\s*DoNotParallelize"),
            "o assembly inteiro foi serializado: a regra irma passa a valer de graca " &
            "e a suite paga minutos por execucao")
    End Sub

End Class

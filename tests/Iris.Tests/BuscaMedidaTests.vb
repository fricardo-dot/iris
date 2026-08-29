Imports System.IO
Imports System.Linq
Imports System.Text.Json
Imports Iris.Integration
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A GARANTIA QUE UM COMENTÁRIO MEU AFIRMOU ANTES DE EXISTIR.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ACONTECEU</b>
'''
''' A medição da busca (<c>tools/medir-busca.py</c>) reimplementa em Python a
''' regra do <see cref="TermoDeBusca"/>: normalização, radical, distância,
''' grau. A duplicação é inevitável — a medição precisa rodar sobre o SQLite
''' sem subir a aplicação.
'''
''' O comentário daquele arquivo dizia que existia um teste comparando as duas
''' implementações, e que <i>"ele quebra no dia em que divergirem"</i>. <b>Não
''' existia.</b> A revisão externa de 29/08 procurou e não achou.
'''
''' Comentário que promete proteção inexistente é pior que comentário nenhum:
''' quem lê para de procurar. É a mesma família dos quatro erros que este
''' projeto já cometeu — afirmar mais do que se sabe —, desta vez sobre o
''' próprio código.
'''
''' ------------------------------------------------------------------
''' <b>COMO A GARANTIA FUNCIONA AGORA</b>
'''
''' <c>tools/casos-de-busca.json</c> é lido pelos <b>dois</b> lados:
'''
''' <list type="bullet">
''' <item>este teste;</item>
''' <item><c>python tools/medir-busca.py --conferir</c>.</item>
''' </list>
'''
''' Não é o VB conferindo o Python — é cada um conferindo a mesma tabela. Quem
''' divergir falha sozinho, e nenhum dos dois precisa executar o outro.
'''
''' <b>O que isto NÃO garante:</b> que alguém rode o lado Python. Se a regra do
''' VB mudar e o JSON for atualizado junto, este teste passa e o Python passa a
''' divergir em silêncio até a próxima medição. A proteção é real e é parcial —
''' e dizer qual das duas coisas ela é foi exatamente o que faltou da primeira
''' vez.
''' </summary>
<TestClass>
Public Class BuscaMedidaTests

    Private NotInheritable Class Caso
        Public Property consulta As String
        Public Property assunto As String
        Public Property remetente As String
        Public Property grau As Integer
    End Class

    ''' <summary>
    ''' <b>A tabela compartilhada vale para esta implementação.</b>
    '''
    ''' <b>Controle negativo:</b> mudar qualquer regra do <c>TermoDeBusca</c>
    ''' sem mexer no JSON derruba este teste — foi confirmado desfazendo o piso
    ''' de tamanho do radical.
    ''' </summary>
    <TestMethod>
    Public Sub As_duas_implementacoes_concordam()
        Dim casos = LerCasos()

        Assert.IsTrue(casos.Count >= 15,
            "a tabela compartilhada encolheu para " & casos.Count & " casos: " &
            "uma tabela pequena demais concorda por não perguntar nada")

        Dim divergiram As New List(Of String)()
        For Each c In casos
            Dim item = New ManifestItem("E-1", c.assunto, c.remetente,
                                        "2026-08-29T00:00:00Z", False,
                                        Iris.Sync.PresenceState.Presente)
            Dim obtido = CInt(New TermoDeBusca(c.consulta).Grau(item))
            If obtido <> c.grau Then
                divergiram.Add($"'{c.consulta}' vs '{c.assunto}': " &
                               $"esperado {c.grau}, obtido {obtido}")
            End If
        Next

        Assert.AreEqual(0, divergiram.Count,
            "a regra do TermoDeBusca divergiu de tools/casos-de-busca.json. " &
            "Se a mudança foi de propósito, atualize o JSON E rode " &
            "'python tools/medir-busca.py --conferir', senão a medição passa a " &
            "medir outra coisa. Divergências: " & String.Join(" | ", divergiram))
    End Sub

    ''' <summary>
    ''' <b>A tabela cobre os três graus.</b>
    '''
    ''' Sem isto ela poderia encolher para só exatos e continuar "concordando"
    ''' — o bloqueio que nunca bloqueia, aplicado a uma tabela de casos.
    ''' </summary>
    <TestMethod>
    Public Sub A_tabela_cobre_os_tres_graus()
        Dim casos = LerCasos()
        For Each g In {0, 1, 2}
            Assert.IsTrue(casos.Any(Function(c) c.grau = g),
                $"a tabela compartilhada não tem nenhum caso de grau {g}")
        Next
    End Sub

    ''' <summary>
    ''' <b>E ela documenta o que o conserto NÃO alcança.</b>
    '''
    ''' Transposição ("conttaro") é distância <b>2</b> para este algoritmo, e
    ''' diminutivo não cabe num radical de nove linhas. Os dois foram medidos em
    ''' 29/08 e deram <b>0%</b> — 237 e 140 consultas.
    '''
    ''' Estão na tabela como <c>grau 0</c> de propósito. Um limite conhecido e
    ''' escrito é um limite; um limite que ninguém anotou é uma surpresa
    ''' esperando quem for mexer.
    ''' </summary>
    <TestMethod>
    Public Sub A_tabela_prende_os_limites_conhecidos()
        Dim casos = LerCasos()

        Assert.IsTrue(casos.Any(Function(c) c.consulta = "conttaro" AndAlso c.grau = 0),
            "o caso da transposição sumiu da tabela: ele é o que documenta " &
            "que trocar duas letras de lugar é distância 2 e não é achado")
        Assert.IsTrue(casos.Any(Function(c) c.consulta = "contratinho" AndAlso c.grau = 0),
            "o caso do diminutivo sumiu: ele é o que documenta que o radical " &
            "cobre número e não cobre morfologia")
    End Sub

    Private Shared Function LerCasos() As List(Of Caso)
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
            d = d.Parent
        End While
        Assert.IsNotNull(d, "não achei a raiz do repositório")

        Dim caminho = Path.Combine(d.FullName, "tools", "casos-de-busca.json")
        Assert.IsTrue(File.Exists(caminho), "casos-de-busca.json não está em " & caminho)

        Using doc = JsonDocument.Parse(File.ReadAllText(caminho))
            Dim saida As New List(Of Caso)()
            For Each e In doc.RootElement.GetProperty("casos").EnumerateArray()
                saida.Add(New Caso With {
                    .consulta = e.GetProperty("consulta").GetString(),
                    .assunto = e.GetProperty("assunto").GetString(),
                    .remetente = e.GetProperty("remetente").GetString(),
                    .grau = e.GetProperty("grau").GetInt32()})
            Next
            Return saida
        End Using
    End Function

End Class

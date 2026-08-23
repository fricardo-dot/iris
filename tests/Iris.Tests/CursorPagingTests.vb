Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Core
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' O algoritmo de paginação por cursor, e os DOIS defeitos que ele já teve.
'''
''' Este arquivo existe na forma que tem por causa de um erro da Q1: o teste
''' sintético de lá verificava um algoritmo DIFERENTE do que rodava contra o
''' Outlook — o teste avançava com <c>&lt; T</c> e o real com <c>&lt;= T</c>.
''' Os cenários passavam, provando outra coisa. Foi a terceira vez no projeto
''' que um teste prometeu mais do que verificava.
'''
''' A correção é estrutural: existe UM algoritmo, em Iris.Core, e quem chama
''' é o adaptador COM e este teste. A fonte é diferente; o algoritmo não.
'''
''' A caixa real tem no máximo 16 itens no mesmo segundo, então o caso que
''' quebra — grupo empatado MAIOR que a página — só dá para exercitar aqui.
''' </summary>
<TestClass>
Public Class CursorPagingTests

    ' ==================================================================
    ' Fonte sintética
    ' ==================================================================

    Private NotInheritable Class LinhaFake
        Public Property Id As String = ""
        Public Property Quando As DateTimeOffset
    End Class

    Private NotInheritable Class FonteFake
        Implements IRowSource(Of LinhaFake)

        Private ReadOnly _todas As List(Of LinhaFake)
        Private ReadOnly _ordemInstavel As Boolean
        Private _janela As List(Of LinhaFake)
        Private _pos As Integer
        Private _aberta As Boolean

        ''' <summary>Quantas vezes Abrir e Fechar foram chamados.</summary>
        Public Property Aberturas As Integer
        Public Property Fechamentos As Integer

        Public Sub New(todas As IEnumerable(Of LinhaFake), ordemInstavel As Boolean)
            _todas = New List(Of LinhaFake)(todas)
            _ordemInstavel = ordemInstavel
        End Sub

        Public Sub Abrir(fronteira As DateTimeOffset?, inclusiva As Boolean) _
            Implements IRowSource(Of LinhaFake).Abrir

            Aberturas += 1
            _aberta = True

            Dim conjunto As IEnumerable(Of LinhaFake) = _todas
            If fronteira.HasValue Then
                If inclusiva Then
                    conjunto = _todas.Where(Function(x) x.Quando <= fronteira.Value)
                Else
                    conjunto = _todas.Where(Function(x) x.Quando < fronteira.Value)
                End If
            End If

            Dim ordenada = conjunto.OrderByDescending(Function(x) x.Quando).ToList()

            ' Dentro do empate a ordem é EMBARALHADA a cada abertura: o OOM
            ' não promete ordem estável ali, e algoritmo que dependa dela
            ' está errado mesmo passando.
            If _ordemInstavel Then
                Dim sorteio As New Random(_todas.Count * 7919 + Aberturas)
                ordenada = ordenada.
                    GroupBy(Function(x) x.Quando).
                    OrderByDescending(Function(g) g.Key).
                    SelectMany(Function(g) g.OrderBy(Function(x) sorteio.Next())).
                    ToList()
            End If

            _janela = ordenada
            _pos = 0
        End Sub

        Public Function Ler(quantas As Integer) As IReadOnlyList(Of LinhaFake) _
            Implements IRowSource(Of LinhaFake).Ler

            Assert.IsTrue(_aberta, "Ler chamado com o cursor fechado")
            Dim fim = Math.Min(_pos + quantas, _janela.Count)
            Dim saida As New List(Of LinhaFake)()
            For i = _pos To fim - 1
                saida.Add(_janela(i))
            Next
            _pos = fim
            Return saida
        End Function

        Public Sub Fechar() Implements IRowSource(Of LinhaFake).Fechar
            Fechamentos += 1
            _aberta = False
        End Sub

        Public Function InstanteDe(linha As LinhaFake) As DateTimeOffset _
            Implements IRowSource(Of LinhaFake).InstanteDe
            Return linha.Quando
        End Function

        Public Function ChaveDe(linha As LinhaFake) As String _
            Implements IRowSource(Of LinhaFake).ChaveDe
            Return linha.Id
        End Function
    End Class

    Private Shared Function Montar(antes As Integer, empatados As Integer,
                                   depois As Integer) As List(Of LinhaFake)
        Dim baseHora = New DateTimeOffset(2026, 8, 1, 12, 0, 0, TimeSpan.Zero)
        Dim saida As New List(Of LinhaFake)()
        For i = 0 To antes - 1
            saida.Add(New LinhaFake With {.Id = $"A{i}", .Quando = baseHora.AddSeconds(-i)})
        Next
        Dim t = baseHora.AddSeconds(-antes)
        For i = 0 To empatados - 1
            saida.Add(New LinhaFake With {.Id = $"E{i}", .Quando = t})
        Next
        For i = 0 To depois - 1
            saida.Add(New LinhaFake With {.Id = $"B{i}", .Quando = t.AddSeconds(-1 - i)})
        Next
        Return saida
    End Function

    ''' <summary>
    ''' Devolve a travessia INTEIRA, nao so a contagem. Descartar
    ''' Exhausted deixava o teste aceitar uma travessia que parou por
    ''' travamento desde que tivesse lido o numero certo de linhas.
    ''' </summary>
    Private Shared Function Percorrer(linhas As List(Of LinhaFake), pagina As Integer,
                                      instavel As Boolean,
                                      defeitos As PagingDefects) As TraverseOutcome(Of LinhaFake)
        Dim fonte As New FonteFake(linhas, instavel)
        Dim saida = CursorPaging.Traverse(fonte, pagina, defeitos)
        Assert.AreEqual(fonte.Aberturas, fonte.Fechamentos,
                        "todo cursor aberto tem de ser fechado")
        Return saida
    End Function

    Private Shared Function Cenarios() As List(Of (Nome As String, Linhas As List(Of LinhaFake), Instavel As Boolean))
        Return New List(Of (String, List(Of LinhaFake), Boolean)) From {
            ("sem empate", Montar(300, 0, 0), False),
            ("empate de 3 (como a caixa real)", Montar(150, 3, 150), False),
            ("empate de 16 (o maior medido)", Montar(100, 16, 100), False),
            ("empate de 50 (= pagina) + antigos", Montar(100, 50, 200), False),
            ("empate de 100 (2x) + antigos", Montar(100, 100, 200), False),
            ("empate de 500 (10x) + antigos", Montar(100, 500, 300), False),
            ("tudo no mesmo segundo", Montar(0, 200, 0), False),
            ("empate no FIM da pasta", Montar(200, 100, 0), False),
            ("empate de 200, ordem INSTAVEL", Montar(100, 200, 100), True)
        }
    End Function

    ' ==================================================================

    ''' <summary>
    ''' Conjunto EXATO, nao contagem. Contagem certa com item trocado
    ''' passaria — e a ordem instavel dentro do empate e exatamente o
    ''' cenario onde isso poderia acontecer.
    ''' </summary>
    <TestMethod>
    Public Sub Travessia_le_tudo_em_todos_os_cenarios()
        For Each c In Cenarios()
            Dim saida = Percorrer(c.Linhas, 50, c.Instavel, PagingDefects.None)

            Assert.IsTrue(saida.Exhausted,
                          $"cenario '{c.Nome}' parou sem chegar ao fim")

            Dim esperadas = New HashSet(Of String)(c.Linhas.Select(Function(x) x.Id))
            Dim obtidas = New HashSet(Of String)(saida.Rows.Select(Function(x) x.Id))
            Assert.AreEqual(saida.Rows.Count, obtidas.Count,
                            $"cenario '{c.Nome}' devolveu linha REPETIDA")
            Assert.IsTrue(esperadas.SetEquals(obtidas),
                          $"cenario '{c.Nome}': faltaram " &
                          $"{esperadas.Except(obtidas).Count()} e sobraram " &
                          $"{obtidas.Except(esperadas).Count()}")
        Next
    End Sub

    ''' <summary>
    ''' Os DOIS defeitos precisam ser discriminados, CADA UM POR SI.
    '''
    ''' Um guarda que aceitasse "algum dos dois perdeu item" deixaria passar
    ''' o dia em que um deles parasse de perder — e metade do controle
    ''' negativo viraria decoração sem ninguém notar. Isso foi um achado da
    ''' revisão da Q1, não uma precaução teórica.
    ''' </summary>
    <TestMethod>
    Public Sub Cada_defeito_perde_item_por_si()
        Dim pegouSemDrenagem = False
        Dim pegouInclusiva = False

        For Each c In Cenarios()
            Dim total = c.Linhas.Count
            Dim semDrenar = Percorrer(c.Linhas, 50, c.Instavel,
                                      New PagingDefects With {.SkipDrain = True})
            Dim inclusiva = Percorrer(c.Linhas, 50, c.Instavel,
                                      New PagingDefects With {.InclusiveBoundary = True})

            If semDrenar.Rows.Count <> total Then pegouSemDrenagem = True
            If inclusiva.Rows.Count <> total Then pegouInclusiva = True

            ' A fronteira inclusiva perde item TRAVANDO: ela para sem
            ' chegar ao fim. Exigir isso separa o defeito de um simples
            ' "leu menos".
            If inclusiva.Rows.Count <> total Then
                Assert.IsFalse(inclusiva.Exhausted,
                               $"'{c.Nome}': a fronteira inclusiva perdeu item mas " &
                               "declarou fim normal — nao e o defeito que eu quis reproduzir")
            End If

            Assert.IsTrue(semDrenar.Rows.Count <= total, "defeito nao pode inventar item")
            Assert.IsTrue(inclusiva.Rows.Count <= total, "defeito nao pode inventar item")
        Next

        Assert.IsTrue(pegouSemDrenagem,
                      "nenhum cenario discriminou SkipDrain: controle negativo que nao perde item nao controla nada")
        Assert.IsTrue(pegouInclusiva,
                      "nenhum cenario discriminou InclusiveBoundary: idem")
    End Sub

    <TestMethod>
    Public Sub Sem_drenagem_perde_exatamente_o_resto_do_grupo()
        ' Grupo de 500 com página de 50: sem drenar, a fronteira avança
        ' estrita depois dos 50 primeiros e os outros 450 somem.
        Dim linhas = Montar(100, 500, 300)
        Dim saida = Percorrer(linhas, 50, False, New PagingDefects With {.SkipDrain = True})
        Assert.AreEqual(linhas.Count - 450, saida.Rows.Count)
    End Sub

    <TestMethod>
    Public Sub Fronteira_inclusiva_trava_quando_o_grupo_e_maior_que_a_pagina()
        ' Reabrir com <= depois de ler o grupo recomeça no mesmo grupo:
        ' nada é novo, e a travessia para com itens mais antigos por ler.
        Dim linhas = Montar(100, 100, 200)
        Dim saida = Percorrer(linhas, 50, False, New PagingDefects With {.InclusiveBoundary = True})
        Assert.IsTrue(saida.Rows.Count < linhas.Count, "deveria ter perdido item")
        Assert.IsFalse(saida.Exhausted, "e por TRAVAMENTO, nao por fim normal")
    End Sub

    <TestMethod>
    Public Sub Pagina_drenada_pode_passar_do_alvo_e_diz_quanto()
        ' O alvo é 10 e o grupo tem 16: a página vai até o fim do grupo,
        ' senão os empatados que ficaram para trás seriam pulados. Isso é
        ' contrato, não defeito — por isso DrainedExtra existe.
        Dim fonte As New FonteFake(Montar(0, 16, 50), False)
        Dim saida = CursorPaging.ReadPage(fonte, Nothing, 10)

        Assert.AreEqual(16, saida.Rows.Count, "a pagina tem de conter o grupo INTEIRO")
        Assert.AreEqual(6, saida.DrainedExtra)
        Assert.IsFalse(saida.Ended)
    End Sub

    ''' <summary>
    ''' Grupo unico, sem nada mais antigo. A pagina drena tudo e declara
    ''' fim numa abertura SO — sem uma consulta extra que voltaria vazia.
    ''' Um algoritmo que devolvesse fronteira aqui passaria nos outros
    ''' testes e gastaria uma ida ao Outlook por travessia.
    ''' </summary>
    <TestMethod>
    Public Sub Grupo_unico_termina_em_uma_abertura()
        Dim fonte As New FonteFake(Montar(0, 16, 0), False)
        Dim saida = CursorPaging.ReadPage(fonte, Nothing, 10)

        Assert.AreEqual(16, saida.Rows.Count)
        Assert.AreEqual(6, saida.DrainedExtra)
        Assert.IsTrue(saida.Ended, "nao ha nada mais antigo: tem de terminar aqui")
        Assert.AreEqual(1, fonte.Aberturas)
        Assert.AreEqual(1, fonte.Fechamentos)
    End Sub

    ''' <summary>
    ''' Falha ao ABRIR, nao ao ler. O adaptador COM adquire a Table e so
    ''' depois configura coluna e ordenacao — se qualquer uma lancar, o
    ''' recurso ja existe. Antes, Abrir ficava FORA do Try e esse caso
    ''' vazava a Table.
    ''' </summary>
    <TestMethod>
    Public Sub Cursor_fecha_quando_ABRIR_lanca()
        Dim fonte As New FonteQueExplodeAoAbrir()
        Assert.ThrowsException(Of InvalidOperationException)(
            Sub() CursorPaging.ReadPage(fonte, Nothing, 30))
        Assert.AreEqual(1, fonte.Fechamentos,
                        "Abrir tem de estar DENTRO do Try que garante o Fechar")
    End Sub

    Private NotInheritable Class FonteQueExplodeAoAbrir
        Implements IRowSource(Of LinhaFake)

        Public Property Fechamentos As Integer

        Public Sub Abrir(fronteira As DateTimeOffset?, inclusiva As Boolean) _
            Implements IRowSource(Of LinhaFake).Abrir
            ' Como o adaptador real: adquire e depois falha.
            Throw New InvalidOperationException("falha ao configurar a coluna")
        End Sub

        Public Function Ler(quantas As Integer) As IReadOnlyList(Of LinhaFake) _
            Implements IRowSource(Of LinhaFake).Ler
            Throw New InvalidOperationException("nao deveria chegar aqui")
        End Function

        Public Sub Fechar() Implements IRowSource(Of LinhaFake).Fechar
            Fechamentos += 1
        End Sub

        Public Function InstanteDe(linha As LinhaFake) As DateTimeOffset _
            Implements IRowSource(Of LinhaFake).InstanteDe
            Return linha.Quando
        End Function

        Public Function ChaveDe(linha As LinhaFake) As String _
            Implements IRowSource(Of LinhaFake).ChaveDe
            Return linha.Id
        End Function
    End Class

    <TestMethod>
    Public Sub Fonte_vazia_termina_sem_cursor()
        Dim fonte As New FonteFake(New List(Of LinhaFake)(), False)
        Dim saida = CursorPaging.ReadPage(fonte, Nothing, 30)
        Assert.AreEqual(0, saida.Rows.Count)
        Assert.IsTrue(saida.Ended)
        Assert.AreEqual(1, fonte.Fechamentos, "cursor precisa fechar mesmo vazio")
    End Sub

    <TestMethod>
    Public Sub Cursor_fecha_mesmo_quando_a_leitura_lanca()
        Dim fonte As New FonteQueExplode()
        Assert.ThrowsException(Of InvalidOperationException)(
            Sub() CursorPaging.ReadPage(fonte, Nothing, 30))
        Assert.AreEqual(1, fonte.Fechamentos, "Fechar tem de estar no Finally")
    End Sub

    Private NotInheritable Class FonteQueExplode
        Implements IRowSource(Of LinhaFake)

        Public Property Fechamentos As Integer

        Public Sub Abrir(fronteira As DateTimeOffset?, inclusiva As Boolean) _
            Implements IRowSource(Of LinhaFake).Abrir
        End Sub

        Public Function Ler(quantas As Integer) As IReadOnlyList(Of LinhaFake) _
            Implements IRowSource(Of LinhaFake).Ler
            Throw New InvalidOperationException("falha simulada da fonte")
        End Function

        Public Sub Fechar() Implements IRowSource(Of LinhaFake).Fechar
            Fechamentos += 1
        End Sub

        Public Function InstanteDe(linha As LinhaFake) As DateTimeOffset _
            Implements IRowSource(Of LinhaFake).InstanteDe
            Return linha.Quando
        End Function

        Public Function ChaveDe(linha As LinhaFake) As String _
            Implements IRowSource(Of LinhaFake).ChaveDe
            Return linha.Id
        End Function
    End Class

    ' ==================================================================
    ' O cursor
    ' ==================================================================

    Private Shared Function Consulta(Optional pasta As String = "P1",
                                     Optional sort As MessageSort = MessageSort.ReceivedDesc,
                                     Optional geracao As Long = 7) As MessageQuery
        Return New MessageQuery(New FolderKey(pasta, "S1"), sort, geracao)
    End Function

    <TestMethod>
    Public Sub Cursor_de_fronteira_faz_ida_e_volta()
        Dim q = Consulta()
        Dim quando = New DateTimeOffset(2026, 8, 22, 19, 42, 32, TimeSpan.FromHours(-3))
        Dim texto = MessageCursor.ForBoundary(q, quando).Encode()

        Dim volta As MessageCursor = Nothing
        Assert.IsTrue(MessageCursor.TryDecode(texto, q, volta))
        Assert.AreEqual(CursorMode.ReceivedDesc, volta.Mode)
        Assert.AreEqual(quando.UtcTicks, volta.Boundary.Value.UtcTicks)
    End Sub

    <TestMethod>
    Public Sub Cursor_de_offset_faz_ida_e_volta()
        Dim q = Consulta(sort:=MessageSort.SubjectAsc)
        Dim texto = MessageCursor.ForOffset(q, 120).Encode()

        Dim volta As MessageCursor = Nothing
        Assert.IsTrue(MessageCursor.TryDecode(texto, q, volta))
        Assert.AreEqual(CursorMode.LegacyOffset, volta.Mode)
        Assert.AreEqual(120, volta.Offset)
    End Sub

    ''' <summary>
    ''' Cursor de outra consulta produziria página de OUTRA PASTA sem a UI
    ''' ter como perceber. Tem de ser recusado, não reinterpretado.
    ''' </summary>
    <TestMethod>
    Public Sub Cursor_de_outra_consulta_e_recusado()
        Dim original = Consulta()
        Dim texto = MessageCursor.ForOffset(original, 10).Encode()
        Dim volta As MessageCursor = Nothing

        Assert.IsFalse(MessageCursor.TryDecode(texto, Consulta(pasta:="P2"), volta),
                       "outra pasta")
        Assert.IsFalse(MessageCursor.TryDecode(texto, Consulta(sort:=MessageSort.SenderAsc), volta),
                       "outra ordenacao")
        Assert.IsFalse(MessageCursor.TryDecode(texto, Consulta(geracao:=8), volta),
                       "outra geracao")
        Assert.IsNull(volta)
    End Sub

    <TestMethod>
    Public Sub Cursor_corrompido_e_recusado_sem_lancar()
        Dim q = Consulta()
        Dim volta As MessageCursor = Nothing

        For Each lixo In {"", "nao é base64!!", "YWJj", New String("A"c, 2000)}
            Assert.IsFalse(MessageCursor.TryDecode(lixo, q, volta), $"deveria recusar: {lixo.Length} chars")
            Assert.IsNull(volta)
        Next
    End Sub

End Class

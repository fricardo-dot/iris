Imports System.Collections.Generic
Imports System.Threading
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' O caminho rápido contra o Outlook DE VERDADE.
'''
''' Os testes de <see cref="CursorPagingTests"/> exercitam o algoritmo com
''' uma fonte sintética. Eles não tocam em COM, então não provam nada sobre
''' o adaptador: nome de coluna, ordem dos índices, formato do filtro DASL
''' e — a pior de todas — o fuso.
'''
''' A armadilha que este arquivo existe para pegar: o filtro DASL interpreta
''' a data como UTC e a Table devolve hora local. Na Q1 essa confusão
''' perdeu 803 de 1.003 mensagens E A PAGINAÇÃO TERMINOU PARECENDO
''' COMPLETA. Nenhum teste de unidade veria isso.
'''
''' A prova é o CRUZAMENTO: o caminho por Table e o caminho por iteração são
''' implementações independentes. Se os dois devolvem o mesmo CONJUNTO de
''' chaves, nenhum dos dois perdeu — e isso não depende de eu confiar em
''' nenhum deles.
'''
''' Requer Outlook clássico aberto. Sem ele, o teste é INCONCLUSIVO, nunca
''' verde: "não deu para verificar" não pode virar "verificado".
''' </summary>
<TestClass>
Public Class PagingIntegrationTests

    Private Const Tamanho As Integer = 30
    Private Const MaxPaginas As Integer = 500

    Private Shared Async Function AbrirBrokerAsync() As Task(Of OutlookBroker)
        Dim broker As New OutlookBroker(New NullLog())
        broker.Start()
        Dim estado = Await broker.ConnectAsync(CancellationToken.None)
        If estado <> SessionState.Connected Then
            broker.Dispose()
            Assert.Inconclusive($"Outlook nao disponivel (estado: {estado}).")
        End If
        Return broker
    End Function

    ''' <summary>Caixa de Entrada do store padrão.</summary>
    Private Shared Async Function AcharEntradaAsync(broker As OutlookBroker) As Task(Of FolderKey)
        Dim stores = Await broker.GetStoresAsync(CancellationToken.None)
        Assert.IsTrue(stores.Succeeded, "GetStoresAsync falhou")
        Assert.IsTrue(stores.Value.Count > 0, "nenhum store")

        Dim raiz = stores.Value(0).RootFolder
        Dim filhas = Await broker.GetFolderChildrenAsync(raiz, CancellationToken.None)
        Assert.IsTrue(filhas.Succeeded, "GetFolderChildrenAsync falhou")

        For Each f In filhas.Value
            If f.ContentKind = FolderContentKind.Mail AndAlso
               (f.Name.StartsWith("Caixa de Entrada", StringComparison.OrdinalIgnoreCase) OrElse
                f.Name.Equals("Inbox", StringComparison.OrdinalIgnoreCase)) Then
                Return f.Key
            End If
        Next
        Assert.Inconclusive("nao achei a Caixa de Entrada")
        Return Nothing
    End Function

    Private NotInheritable Class Travessia
        Public Property Chaves As New List(Of ItemKey)()
        Public Property Paginas As Integer
        Public Property Ignorados As Integer
        Public Property TotalNoInicio As Integer?
        Public Property ExtraDrenado As Integer
    End Class

    Private Shared Async Function PercorrerAsync(broker As OutlookBroker, pasta As FolderKey,
                                                 sort As MessageSort) As Task(Of Travessia)
        Dim saida As New Travessia()
        Dim consulta As New MessageQuery(pasta, sort, 1)
        Dim cursor As String = Nothing

        Do
            Dim r = Await broker.GetMessagePageAsync(consulta, cursor, Tamanho, CancellationToken.None)
            Assert.IsTrue(r.Succeeded, $"pagina {saida.Paginas + 1} falhou: {r.Kind}")

            Dim p = r.Value
            saida.Paginas += 1
            saida.Ignorados += p.SkippedCount
            saida.ExtraDrenado += p.DrainedExtra
            If p.TotalAtStart.HasValue AndAlso Not saida.TotalNoInicio.HasValue Then
                saida.TotalNoInicio = p.TotalAtStart
            End If
            For Each m In p.Items
                saida.Chaves.Add(m.Key)
            Next

            cursor = p.NextCursor
            Assert.IsTrue(saida.Paginas < MaxPaginas,
                          "paginacao nao terminou: passou de " & MaxPaginas & " paginas")
        Loop While Not String.IsNullOrEmpty(cursor)

        Return saida
    End Function

    ''' <summary>
    ''' O cruzamento. Se a Table perder mensagem em silêncio, os conjuntos
    ''' divergem — que é exatamente o sintoma que a Q1 levou horas para ver.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Async Function Table_e_iteracao_leem_o_MESMO_conjunto() As Task
        Dim broker = Await AbrirBrokerAsync()
        Try
            Dim entrada = Await AcharEntradaAsync(broker)

            ' ReceivedDesc usa Table + cursor; SubjectAsc usa iteracao. Duas
            ' implementacoes independentes sobre a mesma pasta.
            Dim rapido = Await PercorrerAsync(broker, entrada, MessageSort.ReceivedDesc)
            Dim lento = Await PercorrerAsync(broker, entrada, MessageSort.SubjectAsc)

            Dim porTabela = New HashSet(Of ItemKey)(rapido.Chaves)
            Dim porIteracao = New HashSet(Of ItemKey)(lento.Chaves)

            Assert.AreEqual(rapido.Chaves.Count, porTabela.Count,
                            "a Table devolveu chave REPETIDA")
            Assert.AreEqual(lento.Chaves.Count, porIteracao.Count,
                            "a iteracao devolveu chave REPETIDA")

            Console.WriteLine($"por Table   : {porTabela.Count} chaves, {rapido.Paginas} paginas")
            Console.WriteLine($"por iteracao: {porIteracao.Count} chaves, {lento.Paginas} paginas")
            ' So o COMPRIMENTO da chave. A do curto prazo tem 48 chars e a
            ' duravel tem 140 — foi assim que o defeito apareceu. O valor
            ' inteiro carrega o endereco do usuario em hex e nao precisa
            ' aparecer em log de teste.
            If rapido.Chaves.Count > 0 AndAlso lento.Chaves.Count > 0 Then
                Console.WriteLine($"comprimento da chave: Table={rapido.Chaves(0).EntryId.Length} " &
                                  $"iteracao={lento.Chaves(0).EntryId.Length}")
                Assert.AreEqual(lento.Chaves(0).EntryId.Length, rapido.Chaves(0).EntryId.Length,
                                "as duas chaves tem de ter o MESMO formato")
            End If

            Dim soNaTabela = porTabela.Except(porIteracao).Count()
            Dim soNaIteracao = porIteracao.Except(porTabela).Count()

            Assert.AreEqual(0, soNaIteracao,
                $"a Table PERDEU {soNaIteracao} de {porIteracao.Count} " &
                "(sintoma classico de fuso no filtro DASL)")
            Assert.AreEqual(0, soNaTabela,
                $"a Table inventou {soNaTabela} itens que a iteracao nao viu")

            Assert.IsTrue(porTabela.Count > 0, "pasta vazia nao prova nada")
        Finally
            broker.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' A travessia tem de TERMINAR, e o cursor tem de recusar consulta
    ''' trocada.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Async Function Cursor_termina_e_recusa_consulta_trocada() As Task
        Dim broker = Await AbrirBrokerAsync()
        Try
            Dim entrada = Await AcharEntradaAsync(broker)
            Dim consulta As New MessageQuery(entrada, MessageSort.ReceivedDesc, 1)

            Dim primeira = Await broker.GetMessagePageAsync(consulta, Nothing, Tamanho,
                                                            CancellationToken.None)
            Assert.IsTrue(primeira.Succeeded)
            Assert.IsTrue(primeira.Value.TotalAtStart.HasValue,
                          "a primeira pagina precisa trazer TotalAtStart")

            If String.IsNullOrEmpty(primeira.Value.NextCursor) Then
                Assert.Inconclusive("pasta com uma pagina so; o cruzamento nao seria util")
            End If

            ' Mesma pasta, GERACAO diferente: o cursor nao serve.
            Dim outra As New MessageQuery(entrada, MessageSort.ReceivedDesc, 2)
            Dim recusada = Await broker.GetMessagePageAsync(
                outra, primeira.Value.NextCursor, Tamanho, CancellationToken.None)

            Assert.IsFalse(recusada.Succeeded, "cursor de outra geracao deveria ser recusado")
            Assert.AreEqual(ErrorKind.Stale, recusada.Kind)

            ' E a segunda pagina da consulta CERTA funciona.
            Dim segunda = Await broker.GetMessagePageAsync(
                consulta, primeira.Value.NextCursor, Tamanho, CancellationToken.None)
            Assert.IsTrue(segunda.Succeeded)
            Assert.IsFalse(segunda.Value.TotalAtStart.HasValue,
                           "TotalAtStart e so da primeira pagina")
        Finally
            broker.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' Não é benchmark rigoroso — é confirmação de ordem de grandeza. A
    ''' Q1 mediu ~18x; se a diferença sumir, alguma coisa mudou de caminho
    ''' sem ninguém notar.
    ''' </summary>
    <TestMethod, TestCategory("Integracao")>
    Public Async Function O_caminho_por_Table_e_bem_mais_rapido() As Task
        Dim broker = Await AbrirBrokerAsync()
        Try
            Dim entrada = Await AcharEntradaAsync(broker)

            Dim crono = Diagnostics.Stopwatch.StartNew()
            Dim rapido = Await PercorrerAsync(broker, entrada, MessageSort.ReceivedDesc)
            crono.Stop()
            Dim msTabela = crono.Elapsed.TotalMilliseconds

            crono.Restart()
            Dim lento = Await PercorrerAsync(broker, entrada, MessageSort.SubjectAsc)
            crono.Stop()
            Dim msIteracao = crono.Elapsed.TotalMilliseconds

            Console.WriteLine($"itens             : {rapido.Chaves.Count}")
            Console.WriteLine($"Table + cursor    : {msTabela:N0} ms em {rapido.Paginas} paginas")
            Console.WriteLine($"iteracao + offset : {msIteracao:N0} ms em {lento.Paginas} paginas")
            Console.WriteLine($"ganho             : {msIteracao / Math.Max(msTabela, 1):N1}x")
            Console.WriteLine($"drenagem extra    : {rapido.ExtraDrenado} linhas alem do alvo")

            Assert.IsTrue(msTabela < msIteracao,
                          $"Table ({msTabela:N0} ms) deveria ganhar da iteracao ({msIteracao:N0} ms)")
        Finally
            broker.Dispose()
        End Try
    End Function

End Class

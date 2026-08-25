Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports Iris.Model
Imports Iris.Sync

''' <summary>
''' Uma fonte que MUTA DURANTE A LEITURA.
'''
''' A §22.10 deixou o requisito escrito, e ele não é preciosismo: um falso que
''' devolve uma lista fixa testa o algoritmo contra um mundo que não existe. A
''' Fase 2 inteira foi sobre o mundo mudando embaixo da varredura — mensagem
''' chegando, mensagem sendo movida, pasta encolhendo, o Outlook morrendo. Um
''' falso estático passa verde em tudo isso e não prova nada.
'''
''' Por isso a mutação aqui é <b>agendada por página</b>: some um item depois
''' da página 1, chega outro durante a contagem final, o universo troca no
''' meio. Cada agendamento é um cenário que aconteceu de verdade nesta caixa.
''' </summary>
Friend Class FonteFalsaMutavel
    Implements ISweepSource

    Private ReadOnly _chaves As List(Of String)
    Private _universo As SweepUniverse

    ''' <summary>Página em que cada mutação dispara. 0 = na contagem inicial.</summary>
    Friend ReadOnly Property Agenda As New Dictionary(Of Integer, Action(Of FonteFalsaMutavel))()

    Friend Property ContagemDeclarada As Integer? = Nothing
    Friend Property TruncarApos As Integer? = Nothing
    Friend Property LancarNaPagina As Integer? = Nothing
    Friend Property LancarNaContagem As Integer? = Nothing
    ''' <summary>
    ''' A fonte RECUSA a pagina 1 com este <c>ErrorKind</c> — recusa
    ''' classificada, como o adaptador real emite, e nao excecao qualquer.
    ''' </summary>
    Friend Property RecusarNaPagina As ErrorKind? = Nothing
    ''' <summary>A fonte RECUSA a contagem com este <c>ErrorKind</c>.</summary>
    Friend Property RecusarNaContagem As ErrorKind? = Nothing
    ''' <summary>
    ''' A fonte CANCELA nesta página — que é o que o adaptador real faz quando o
    ''' broker devolve <c>ErrorKind.Cancelled</c>.
    ''' </summary>
    Friend Property LancarCancelamentoNaPagina As Integer? = Nothing
    Friend Property CursorTravado As Boolean = False
    Friend Property UniversoNulo As Boolean = False
    Friend Property PaginaNula As Boolean = False
    Friend Property ContagemNula As Boolean = False
    ''' <summary>Devolve mais linhas do que o pedido — fonte defeituosa.</summary>
    Friend Property EstourarLote As Integer? = Nothing
    ''' <summary>Página vazia, sem fim, cursor parado: laço sem progresso.</summary>
    Friend Property VaziaSemFim As Boolean = False
    ''' <summary>Linhas descartadas por pagina, DECLARADAS na pagina.</summary>
    Friend Property DescartarPorPagina As Integer = 0

    Friend Property PaginasLidas As Integer = 0
    Friend Property VezesQueContou As Integer = 0

    Friend Sub New(universo As SweepUniverse, ParamArray chaves As String())
        _universo = universo
        _chaves = chaves.ToList()
    End Sub

    ' ---- o mundo, mutável ----

    Friend Sub Remover(chave As String)
        _chaves.Remove(chave)
    End Sub

    Friend Sub Acrescentar(chave As String)
        _chaves.Add(chave)
    End Sub

    ''' <summary>Troca um item por outro: a contagem não se mexe.</summary>
    Friend Sub Trocar(sai As String, entra As String)
        _chaves.Remove(sai)
        _chaves.Add(entra)
    End Sub

    Friend Sub TrocarUniverso(novo As SweepUniverse)
        _universo = novo
    End Sub

    Friend ReadOnly Property Quantos As Integer
        Get
            Return _chaves.Count
        End Get
    End Property

    Friend ReadOnly Property Estado As List(Of String)
        Get
            Return _chaves.ToList()
        End Get
    End Property

    ' ---- ISweepSource ----

    Public Function Contar(ct As CancellationToken) As SourceCount Implements ISweepSource.Contar
        ct.ThrowIfCancellationRequested()
        VezesQueContou += 1
        If LancarNaContagem.HasValue AndAlso VezesQueContou = LancarNaContagem.Value Then
            Throw New InvalidOperationException($"fonte falhou na contagem {VezesQueContou}")
        End If
        If RecusarNaContagem.HasValue Then
            Throw New SourceUnavailableException(RecusarNaContagem.Value, "contagem")
        End If
        If VezesQueContou = 1 Then Disparar(0)
        If ContagemNula Then Return Nothing
        Return New SourceCount(If(ContagemDeclarada, _chaves.Count), UniversoDaLeitura())
    End Function

    Private Function UniversoDaLeitura() As SweepUniverse
        Return If(UniversoNulo, Nothing, _universo)
    End Function

    Public Function LerPagina(cursor As String, tamanho As Integer,
                              ct As CancellationToken) As SourcePage _
                              Implements ISweepSource.LerPagina
        ct.ThrowIfCancellationRequested()
        PaginasLidas += 1

        If LancarNaPagina.HasValue AndAlso PaginasLidas = LancarNaPagina.Value Then
            Throw New InvalidOperationException($"fonte falhou na pagina {PaginasLidas}")
        End If
        If RecusarNaPagina.HasValue Then
            Throw New SourceUnavailableException(RecusarNaPagina.Value, $"pagina {PaginasLidas}")
        End If
        If LancarCancelamentoNaPagina.HasValue AndAlso
           PaginasLidas = LancarCancelamentoNaPagina.Value Then
            Throw New OperationCanceledException()
        End If
        If PaginaNula Then Return Nothing

        If VaziaSemFim Then
            Return New SourcePage(Enumerable.Empty(Of SourceRow)(), cursor, False, UniversoDaLeitura())
        End If

        Dim de = If(cursor Is Nothing, 0, Integer.Parse(cursor))
        Dim quanto = If(EstourarLote, tamanho)
        Dim lote = _chaves.Skip(de).Take(quanto).ToList()
        Dim ate = de + lote.Count

        ' Truncar: a fonte simplesmente para de devolver, e diz que acabou.
        ' E o caso mais perigoso porque PARECE sucesso.
        Dim truncou = TruncarApos.HasValue AndAlso PaginasLidas >= TruncarApos.Value

        Dim proximo = If(CursorTravado, cursor, ate.ToString())
        Dim fim = truncou OrElse ate >= _chaves.Count

        Dim p As New SourcePage(lote.Select(AddressOf Linha), proximo, fim, UniversoDaLeitura(),
                                drenadoAlem:=0, descartadas:=DescartarPorPagina)

        ' A mutacao acontece DEPOIS de a pagina ser montada e ANTES de a
        ' proxima ser pedida — que e exatamente a janela onde o mundo muda
        ' numa varredura de verdade.
        Disparar(PaginasLidas)
        Return p
    End Function

    Private Shared Function Linha(chave As String) As SourceRow
        Return New SourceRow With {
            .Key = chave,
            .Subject = "assunto " & chave,
            .SenderName = "Fulano",
            .ReceivedAt = New DateTime(2026, 1, 1, 12, 0, 0, DateTimeKind.Utc).ToString("o"),
            .MessageClass = "IPM.Note",
            .IsUnread = False}
    End Function

    Private Sub Disparar(pagina As Integer)
        Dim acao As Action(Of FonteFalsaMutavel) = Nothing
        If Agenda.TryGetValue(pagina, acao) Then
            Agenda.Remove(pagina)
            acao(Me)
        End If
    End Sub

End Class

''' <summary>Destino falso: registra o que recebeu, e pode falhar de propósito.</summary>
Friend Class DestinoFalso
    Implements ISweepSink

    Friend ReadOnly Paginas As New List(Of (Numero As Integer, Chaves As List(Of String), Cursor As String))()
    Friend ReadOnly Descartadas As New List(Of String)()
    Friend Property Publicadas As Integer = 0
    Friend Property CoberturaPublicada As FolderCoverage = FolderCoverage.Desconhecida
    Friend Property Epoca As Long = 0
    Friend Property Abertas As Integer = 0
    Friend Property LancarAoGravar As Boolean = False
    ''' <summary>
    ''' O DESTINO lança <c>SourceUnavailableException</c> — o mesmo tipo que a
    ''' fonte usa. Existe para provar que o tipo sozinho não classifica origem.
    ''' </summary>
    Friend Property RecusarAoGravar As ErrorKind? = Nothing
    Friend Property LancarAoPublicar As Boolean = False
    Friend Property LancarNaEpoca As Boolean = False
    Friend Property RespostaAoPublicar As SinkPublishResult = SinkPublishResult.Publicada

    Friend ReadOnly Property ChavesGravadas As List(Of String)
        Get
            Return Paginas.SelectMany(Function(p) p.Chaves).ToList()
        End Get
    End Property

    Public Function AbrirTentativa(universo As SweepUniverse, epoca As Long,
                                   numero As Integer) As Long Implements ISweepSink.AbrirTentativa
        Abertas += 1
        Return Abertas
    End Function

    Public Sub GravarPagina(tentativa As Long, pagina As Integer,
                            linhas As IReadOnlyList(Of SourceRow),
                            cursorDepois As String) Implements ISweepSink.GravarPagina
        If LancarAoGravar Then Throw New InvalidOperationException("destino falhou ao gravar")
        If RecusarAoGravar.HasValue Then
            Throw New SourceUnavailableException(RecusarAoGravar.Value, "destino, nao fonte")
        End If
        Paginas.Add((pagina, linhas.Select(Function(l) l.Key).ToList(), cursorDepois))
    End Sub

    Public Function Publicar(tentativa As Long, cobertura As FolderCoverage,
                             antes As Integer, depois As Integer) As SinkPublishResult _
                             Implements ISweepSink.Publicar
        If LancarAoPublicar Then Throw New InvalidOperationException("destino falhou ao publicar")
        If RespostaAoPublicar <> SinkPublishResult.Publicada Then Return RespostaAoPublicar
        Publicadas += 1
        CoberturaPublicada = cobertura
        Return SinkPublishResult.Publicada
    End Function

    Public Sub Descartar(tentativa As Long, motivo As String) Implements ISweepSink.Descartar
        Descartadas.Add(motivo)
    End Sub

    Public Function EpocaCorrente() As Long Implements ISweepSink.EpocaCorrente
        If LancarNaEpoca Then Throw New InvalidOperationException("destino falhou ao ler a epoca")
        Return Epoca
    End Function

End Class

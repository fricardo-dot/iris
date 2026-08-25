Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
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
    Friend Property CursorTravado As Boolean = False

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

    ' ---- ISweepSource ----

    Public Function Contar(ct As CancellationToken) As Integer Implements ISweepSource.Contar
        VezesQueContou += 1
        If VezesQueContou = 1 Then Disparar(0)
        Return If(ContagemDeclarada, _chaves.Count)
    End Function

    Public Function UniversoAgora() As SweepUniverse Implements ISweepSource.UniversoAgora
        Return _universo
    End Function

    Public Function LerPagina(cursor As String, tamanho As Integer,
                              ct As CancellationToken) As SourcePage _
                              Implements ISweepSource.LerPagina
        ct.ThrowIfCancellationRequested()
        PaginasLidas += 1

        If LancarNaPagina.HasValue AndAlso PaginasLidas = LancarNaPagina.Value Then
            Throw New InvalidOperationException($"fonte falhou na pagina {PaginasLidas}")
        End If

        Dim de = If(cursor Is Nothing, 0, Integer.Parse(cursor))
        Dim lote = _chaves.Skip(de).Take(tamanho).ToList()
        Dim ate = de + lote.Count

        ' Truncar: a fonte simplesmente para de devolver, e diz que acabou.
        ' E o caso mais perigoso porque PARECE sucesso.
        Dim truncou = TruncarApos.HasValue AndAlso PaginasLidas >= TruncarApos.Value

        Dim proximo = If(CursorTravado, cursor, ate.ToString())
        Dim fim = truncou OrElse ate >= _chaves.Count

        Dim p As New SourcePage(lote, proximo, fim, _universo)

        ' A mutacao acontece DEPOIS de a pagina ser montada e ANTES de a
        ' proxima ser pedida — que e exatamente a janela onde o mundo muda
        ' numa varredura de verdade.
        Disparar(PaginasLidas)
        Return p
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
    Friend Property Publicadas As Integer = 0
    Friend Property Epoca As Long = 0
    Friend Property LancarAoGravar As Boolean = False
    Friend Property LancarAoPublicar As Boolean = False

    Friend ReadOnly Property ChavesGravadas As List(Of String)
        Get
            Return Paginas.SelectMany(Function(p) p.Chaves).ToList()
        End Get
    End Property

    Public Sub GravarPagina(pagina As Integer, chaves As IReadOnlyList(Of String),
                            cursorDepois As String) Implements ISweepSink.GravarPagina
        If LancarAoGravar Then Throw New InvalidOperationException("destino falhou ao gravar")
        Paginas.Add((pagina, chaves.ToList(), cursorDepois))
    End Sub

    Public Sub Publicar(a As SweepAttempt) Implements ISweepSink.Publicar
        If LancarAoPublicar Then Throw New InvalidOperationException("destino falhou ao publicar")
        Publicadas += 1
    End Sub

    Public Function EpocaCorrente() As Long Implements ISweepSink.EpocaCorrente
        Return Epoca
    End Function

End Class

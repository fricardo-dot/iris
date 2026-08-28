Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports Iris.Sync

''' <summary>
''' Uma fonte de varredura que devolve <b>as linhas que o teste escreveu</b>.
'''
''' ------------------------------------------------------------------
''' <c>FonteFalsaMutavel</c> carrega só chaves e fabrica assunto e remetente
''' a partir delas — "assunto E-1", "Fulano". Isso serve para provar
''' paginação, corridas e a guarda S6, que é para o que ela foi feita.
'''
''' Não serve para provar busca. Uma busca testada contra assuntos gerados
''' por concatenação estaria sendo testada contra um vocabulário que não
''' existe: sem acento, sem caixa alta, sem prefixo de conversa, e com um
''' remetente só. Os três casos que mais importam numa caixa em português
''' nunca apareceriam.
'''
''' Esta fonte é deliberadamente burra: ela não muta, não falha, não trunca.
''' Quem precisa disso continua usando a irmã.
''' </summary>
Friend NotInheritable Class FonteDeLinhas
    Implements ISweepSource

    Private ReadOnly _linhas As List(Of SourceRow)
    Private ReadOnly _universo As SweepUniverse

    Friend Sub New(universo As SweepUniverse, linhas As IEnumerable(Of SourceRow))
        _universo = universo
        _linhas = linhas.ToList()
    End Sub

    Public Function Contar(ct As CancellationToken) As SourceCount _
                           Implements ISweepSource.Contar
        ct.ThrowIfCancellationRequested()
        Return New SourceCount(_linhas.Count, _universo)
    End Function

    Public Function LerPagina(cursor As String, tamanho As Integer,
                              ct As CancellationToken) As SourcePage _
                              Implements ISweepSource.LerPagina
        ct.ThrowIfCancellationRequested()
        Dim de = If(cursor Is Nothing, 0, Integer.Parse(cursor))
        Dim lote = _linhas.Skip(de).Take(tamanho).ToList()
        Dim ate = de + lote.Count
        Return New SourcePage(lote, ate.ToString(), ate >= _linhas.Count, _universo)
    End Function

End Class

Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Cache
Imports Iris.Sync

Namespace Global.Iris.Integration

    ''' <summary>
    ''' O <see cref="ISweepSink"/> real, sobre o <see cref="CacheWriter"/>.
    '''
    ''' É onde a porta que o <c>Iris.Sync</c> declara encontra a persistência
    ''' que o 2.1 construiu. Fino de propósito: toda decisão já foi tomada no
    ''' <c>SweepModel</c> e no <c>SweepRunner</c>; aqui só há tradução.
    '''
    ''' A parte que não é tradução, e que importa: <b>a persistência tem
    ''' fencing próprio</b>. O <c>CacheWriter.Publicar</c> revalida época e
    ''' ordem de tentativa dentro da mesma transação em que escreve, e pode
    ''' recusar <i>depois</i> de o modelo ter aprovado. Por isso
    ''' <see cref="ISweepSink.Publicar"/> devolve resultado em vez de ser um
    ''' <c>Sub</c> — engolir essa recusa faria o runner reportar publicado o
    ''' que o banco rejeitou.
    ''' </summary>
    Public NotInheritable Class SqliteSweepSink
        Implements ISweepSink

        Private ReadOnly _writer As CacheWriter
        Private ReadOnly _folderKey As Long
        Private ReadOnly _environmentKey As Long
        Private ReadOnly _db As CacheDatabase

        Public Sub New(db As CacheDatabase, folderKey As Long, environmentKey As Long)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _db = db
            _writer = New CacheWriter(db)
            _folderKey = folderKey
            _environmentKey = environmentKey
        End Sub

        Public Function AbrirTentativa(universo As SweepUniverse, epoca As Long,
                                       numero As Integer) As Long _
                                       Implements ISweepSink.AbrirTentativa
            Return _writer.AbrirTentativa(_folderKey, _environmentKey,
                                          universo.Fingerprint(), epoca, numero,
                                          universo.AlgorithmVersion,
                                          NuloSeVazio(universo.RetentionCutoff))
        End Function

        Public Sub GravarPagina(tentativa As Long, pagina As Integer,
                                linhas As IReadOnlyList(Of SourceRow),
                                cursorDepois As String) Implements ISweepSink.GravarPagina
            _writer.GravarPagina(tentativa, _folderKey, pagina,
                                 linhas.Select(AddressOf Traduzir).ToList(), cursorDepois)
        End Sub

        Public Function Publicar(tentativa As Long, cobertura As FolderCoverage,
                                 antes As Integer, depois As Integer) As SinkPublishResult _
                                 Implements ISweepSink.Publicar
            Dim g As Long = 0
            ' Tipo de varredura e ALCANCE sao eixos diferentes: esta varredura e
            ' sempre de tipo COMPLETO (percorreu a pasta inteira), e o quanto
            ' dela se alcancou vai na observacao de cobertura.
            Dim r = _writer.Publicar(tentativa, _folderKey, "completa",
                                     antes, depois, g,
                                     alcance:=TextoDoAlcance(cobertura),
                                     environmentKey:=_environmentKey)
            Select Case r
                Case PublishOutcome.Publicada : Return SinkPublishResult.Publicada
                Case PublishOutcome.RecusadaPorEpoca : Return SinkPublishResult.RecusadaPorEpoca
                Case PublishOutcome.RecusadaPorOrdem : Return SinkPublishResult.RecusadaPorOrdem
                Case Else : Return SinkPublishResult.RecusadaPorEstado
            End Select
        End Function

        Public Sub Descartar(tentativa As Long, motivo As String) Implements ISweepSink.Descartar
            _writer.Descartar(tentativa, motivo)
        End Sub

        Public Function EpocaCorrente() As Long Implements ISweepSink.EpocaCorrente
            Return _writer.EpocaDaPasta(_folderKey)
        End Function

        ' ==============================================================

        Private Shared Function Traduzir(l As SourceRow) As StagedRow
            Return New StagedRow With {
                .ProviderEntryId = l.Key,
                .SearchKey = l.SearchKey,
                .InternetMessageId = l.InternetMessageId,
                .Subject = l.Subject,
                .SenderName = l.SenderName,
                .ReceivedAt = l.ReceivedAt,
                .LastModifiedAt = l.LastModifiedAt,
                .SizeBytes = l.SizeBytes,
                .HasAttachments = l.HasAttachments,
                .IsUnread = l.IsUnread,
                .MessageClass = l.MessageClass}
        End Function

        ''' <summary>
        ''' O ALCANCE, que vai para <c>coverage_observation.coverage</c>.
        '''
        ''' Não confundir com <c>generation.coverage_kind</c>: aquele diz que
        ''' TIPO de varredura foi — completa ou incremental —, este diz QUANTO
        ''' dela se alcançou. Uma varredura pode ser de tipo completo e alcance
        ''' parcial: ela percorreu integralmente <b>o conjunto que o provider
        ''' expôs</b>, que não é a pasta inteira. Foi a §19.2, com pastas cheias
        ''' reportando zero — e essa diferença entre o exposto e o real é
        ''' exatamente o motivo de a cobertura ser parcial.
        '''
        ''' A primeira versão desta classe traduzia parcial como 'completa' e
        ''' justificava no comentário. Era mentir na coluna errada: o alcance
        ''' parcial sumia, e uma geração sem alcance conhecido é pior que não
        ''' ter geração.
        ''' </summary>
        Private Shared Function TextoDoAlcance(c As FolderCoverage) As String
            Select Case c
                Case FolderCoverage.Completa : Return "completa"
                Case FolderCoverage.Parcial : Return "parcial"
                Case Else : Return "desconhecida"
            End Select
        End Function

        Private Shared Function NuloSeVazio(s As String) As String
            Return If(String.IsNullOrEmpty(s), Nothing, s)
        End Function

    End Class

End Namespace

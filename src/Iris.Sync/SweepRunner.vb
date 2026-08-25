Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading

Namespace Global.Iris.Sync

    ''' <summary>
    ''' Uma linha lida da fonte, já reduzida a metadado (D1) — sem corpo, sem
    ''' anexo.
    '''
    ''' A fonte devolve isto, e não só a chave, porque quem persiste precisa do
    ''' metadado junto: buscar de novo por chave seria uma segunda ida ao
    ''' provider por linha, e o custo de ida ao OOM é o que a Q1 mediu.
    ''' </summary>
    Public NotInheritable Class SourceRow
        Public Property Key As String
        Public Property SearchKey As String
        Public Property InternetMessageId As String
        Public Property Subject As String
        Public Property SenderName As String
        Public Property ReceivedAt As String
        Public Property LastModifiedAt As String
        Public Property SizeBytes As Long?
        Public Property HasAttachments As Boolean?
        Public Property IsUnread As Boolean?
        Public Property MessageClass As String
    End Class

    ''' <summary>
    ''' Uma contagem, com o universo <b>da mesma leitura</b>.
    '''
    ''' Os dois vêm juntos pelo mesmo motivo que na página: perguntar a
    ''' contagem e depois perguntar o universo são duas chamadas, e entre elas
    ''' cabe uma mudança. Era assim antes — <c>Contar()</c> seguido de
    ''' <c>UniversoAgora()</c> — e a corrida ficava justamente na fronteira que
    ''' o S6 usa para decidir.
    ''' </summary>
    Public NotInheritable Class SourceCount
        Public ReadOnly Property Count As Integer
        Public ReadOnly Property Universo As SweepUniverse

        Public Sub New(count As Integer, universo As SweepUniverse)
            Me.Count = count
            Me.Universo = universo
        End Sub
    End Class

    ''' <summary>
    ''' Uma página lida da fonte, com o universo <b>no momento da leitura</b>.
    ''' </summary>
    Public NotInheritable Class SourcePage
        Public ReadOnly Property Rows As IReadOnlyList(Of SourceRow)
        Public ReadOnly Property NextCursor As String
        Public ReadOnly Property Fim As Boolean
        Public ReadOnly Property Universo As SweepUniverse

        Public Sub New(rows As IEnumerable(Of SourceRow), nextCursor As String,
                       fim As Boolean, universo As SweepUniverse)
            Me.Rows = If(rows, Enumerable.Empty(Of SourceRow)()).ToList()
            Me.NextCursor = nextCursor
            Me.Fim = fim
            Me.Universo = universo
        End Sub

        Public Function Keys() As IReadOnlyList(Of String)
            Return Rows.Select(Function(r) r.Key).ToList()
        End Function
    End Class

    ''' <summary>
    ''' A fonte, do ponto de vista da orquestração. O adaptador do Outlook
    ''' implementa isto no 2.2b; os testes implementam com um falso que muta
    ''' <b>durante</b> a leitura.
    ''' </summary>
    Public Interface ISweepSource
        ''' <summary>Contagem declarada pela fonte. Pode mentir — o S6 existe por isso.</summary>
        Function Contar(ct As CancellationToken) As SourceCount
        Function LerPagina(cursor As String, tamanho As Integer, ct As CancellationToken) As SourcePage
    End Interface

    Public Enum SinkPublishResult
        Publicada
        RecusadaPorEpoca
        RecusadaPorOrdem
        RecusadaPorEstado
    End Enum

    ''' <summary>
    ''' Onde a varredura grava. Existe para a orquestração não depender de
    ''' SQLite: o 2.2a inteiro se prova com um destino falso, e o real é um
    ''' adaptador num projeto de infraestrutura.
    '''
    ''' A porta é definida aqui e implementada lá fora — a dependência aponta
    ''' <b>para dentro</b>. Fazer <c>Iris.Sync</c> referenciar
    ''' <c>Iris.Cache</c> inverteria a fronteira e amarraria a orquestração ao
    ''' SQLite.
    ''' </summary>
    Public Interface ISweepSink
        ''' <summary>Registra a tentativa e devolve a chave dela.</summary>
        Function AbrirTentativa(universo As SweepUniverse, epoca As Long, numero As Integer) As Long

        ''' <summary>Página e checkpoint, atomicamente. Idempotente.</summary>
        Sub GravarPagina(tentativa As Long, pagina As Integer,
                         linhas As IReadOnlyList(Of SourceRow), cursorDepois As String)

        ''' <summary>
        ''' Geração, cabeça e dívida para a UI, atomicamente. Devolve o
        ''' resultado — a persistência tem o seu próprio fencing, e ele pode
        ''' recusar depois de o modelo ter aprovado.
        ''' </summary>
        Function Publicar(tentativa As Long, cobertura As FolderCoverage,
                          antes As Integer, depois As Integer) As SinkPublishResult

        Sub Descartar(tentativa As Long, motivo As String)

        ''' <summary>Época corrente da pasta, lida no momento da publicação (fencing).</summary>
        Function EpocaCorrente() As Long
    End Interface

    Public Enum SweepConclusion
        Publicada
        Rejeitada
        Cancelada
        Falhou
    End Enum

    Public NotInheritable Class SweepResult
        Public ReadOnly Property Conclusion As SweepConclusion
        Public ReadOnly Property Attempt As SweepAttempt
        Public ReadOnly Property Motivo As String
        Public ReadOnly Property Paginas As Integer
        Public ReadOnly Property Cobertura As FolderCoverage

        Friend Sub New(c As SweepConclusion, a As SweepAttempt, motivo As String,
                       paginas As Integer,
                       Optional cobertura As FolderCoverage = FolderCoverage.Desconhecida)
            Conclusion = c
            Attempt = a
            Me.Motivo = motivo
            Me.Paginas = paginas
            Me.Cobertura = cobertura
        End Sub

        Public ReadOnly Property Publicou As Boolean
            Get
                Return Conclusion = SweepConclusion.Publicada
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Conduz uma varredura: abre, conta, pagina, conta de novo, publica.
    '''
    ''' O <see cref="SweepModel"/> é a máquina de estados e decide tudo o que é
    ''' decidível sem tocar no mundo. Esta classe é o que <b>toca o mundo</b>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ESTADO PROPOSTO NÃO É ESTADO CONFIRMADO</b>
    '''
    ''' A regra "o modelo transiciona primeiro, o efeito depois" está certa,
    ''' mas a primeira versão a implementou instalando o estado novo antes de
    ''' o efeito acontecer. Com isso, destino falhando ao publicar produzia
    ''' este absurdo:
    '''
    '''   1. o modelo devolve estado <c>Publicada</c>;
    '''   2. a variável corrente recebe esse estado;
    '''   3. o destino lança;
    '''   4. <c>Falhar</c> transforma uma tentativa <b>já publicada</b> em
    '''      descartada.
    '''
    ''' Publicação é imutável — foi o que o 2.1 inteiro construiu — e a
    ''' orquestração desfazia isso em três linhas.
    '''
    ''' Agora o estado só é instalado <b>depois</b> de o efeito confirmar. Se o
    ''' efeito falhar, o corrente continua sendo o anterior, que é a verdade:
    ''' aquilo não aconteceu.
    ''' </summary>
    Public NotInheritable Class SweepRunner

        ''' <summary>
        ''' Teto de páginas. Cursor que não avança é laço infinito, e laço
        ''' infinito na fila da STA trava a UI do usuário — não é defesa
        ''' hipotética, é o modo de falha desta arquitetura.
        ''' </summary>
        Public Const MaxPaginas As Integer = 100000

        Private ReadOnly _fonte As ISweepSource
        Private ReadOnly _destino As ISweepSink
        Private ReadOnly _tamanhoPagina As Integer

        Public Sub New(fonte As ISweepSource, destino As ISweepSink,
                       Optional tamanhoPagina As Integer = 200)
            If fonte Is Nothing Then Throw New ArgumentNullException(NameOf(fonte))
            If destino Is Nothing Then Throw New ArgumentNullException(NameOf(destino))
            If tamanhoPagina < 1 Then Throw New ArgumentOutOfRangeException(NameOf(tamanhoPagina))
            _fonte = fonte
            _destino = destino
            _tamanhoPagina = tamanhoPagina
        End Sub

        Public Function Executar(universo As SweepUniverse, epoca As Long,
                                 numeroDaTentativa As Integer,
                                 capacidades As EnvironmentCapabilities,
                                 ct As CancellationToken) As SweepResult

            ' O gate de ABERTURA e estreito: so recusa ambiente NAO
            ' IDENTIFICADO. Ambiente limitado varre normalmente e publica
            ' parcial — §23, e ver o comentario de SweepModel.Abrir para por
            ' que isto ja foi diferente e o que custou.
            Dim identificado = capacidades IsNot Nothing

            Dim r = SweepModel.Abrir(universo, epoca, numeroDaTentativa, identificado)
            If r.Rejected Then
                Return New SweepResult(SweepConclusion.Rejeitada, Nothing, r.Rejection, 0)
            End If

            Dim a = r.State
            Dim paginas = 0
            Dim tentativa As Long = 0
            Dim abriuNoDestino = False

            Try
                tentativa = _destino.AbrirTentativa(universo, epoca, numeroDaTentativa)
                abriuNoDestino = True

                ' --- contagem inicial ---
                If ct.IsCancellationRequested Then Return Cancelar(a, paginas, tentativa, abriuNoDestino)
                Dim antes = _fonte.Contar(ct)
                If antes Is Nothing Then
                    Return Falhar(a, "fonte devolveu contagem nula", paginas, tentativa, abriuNoDestino)
                End If
                Dim proposto = SweepModel.ContagemInicial(a, antes.Count, antes.Universo)
                If proposto.Rejected Then Return Rejeitar(proposto, paginas, tentativa, abriuNoDestino)
                a = proposto.State

                ' --- paginacao ---
                Dim cursor As String = Nothing
                Do
                    If ct.IsCancellationRequested Then Return Cancelar(a, paginas, tentativa, abriuNoDestino)

                    paginas += 1
                    If paginas > MaxPaginas Then
                        Return Falhar(a, $"passou de {MaxPaginas} paginas: cursor nao avanca",
                                      paginas, tentativa, abriuNoDestino)
                    End If

                    Dim p = _fonte.LerPagina(cursor, _tamanhoPagina, ct)
                    If p Is Nothing Then
                        Return Falhar(a, "fonte devolveu pagina nula", paginas, tentativa, abriuNoDestino)
                    End If

                    ' A fonte nao pode estourar o lote pedido. O tamanho da
                    ' pagina E o orcamento de tempo da fila da STA (D5); uma
                    ' fonte que devolve milhares de linhas de uma vez trava a
                    ' UI mesmo sem nenhum laco infinito.
                    If p.Rows.Count > _tamanhoPagina Then
                        Return Falhar(a, $"fonte devolveu {p.Rows.Count} linhas, pedi {_tamanhoPagina}",
                                      paginas, tentativa, abriuNoDestino)
                    End If

                    ' Progresso e obrigatorio, com ou sem linhas. A guarda so
                    ' olhava paginas NAO VAZIAS, e uma fonte devolvendo pagina
                    ' vazia sem fim e com cursor parado rodava 100.001 vezes
                    ' antes de o teto pegar.
                    If Not p.Fim AndAlso String.Equals(cursor, p.NextCursor, StringComparison.Ordinal) Then
                        Return Falhar(a, "cursor nao avancou e a varredura nao terminou",
                                      paginas, tentativa, abriuNoDestino)
                    End If

                    proposto = SweepModel.Pagina(a, p.Keys(), p.NextCursor, p.Universo)
                    If proposto.Rejected Then Return Rejeitar(proposto, paginas, tentativa, abriuNoDestino)

                    ' O efeito SO acontece se o modelo mandou — e o estado so
                    ' e instalado se o efeito deu certo.
                    If proposto.Commands.Contains(SweepCommand.StagePagina) Then
                        _destino.GravarPagina(tentativa, paginas, p.Rows, p.NextCursor)
                    End If
                    a = proposto.State

                    If p.Fim Then Exit Do
                    cursor = p.NextCursor
                Loop

                ' --- contagem final ---
                If ct.IsCancellationRequested Then Return Cancelar(a, paginas, tentativa, abriuNoDestino)
                Dim depois = _fonte.Contar(ct)
                If depois Is Nothing Then
                    Return Falhar(a, "fonte devolveu contagem nula", paginas, tentativa, abriuNoDestino)
                End If
                proposto = SweepModel.ContagemFinal(a, depois.Count, depois.Universo)
                If proposto.Rejected Then Return Rejeitar(proposto, paginas, tentativa, abriuNoDestino)
                a = proposto.State

                ' --- publicacao ---
                If ct.IsCancellationRequested Then Return Cancelar(a, paginas, tentativa, abriuNoDestino)
                proposto = SweepModel.Publicar(a, _destino.EpocaCorrente(),
                                               capacidades.PodeAfirmarCoberturaCompleta)
                If proposto.Rejected Then Return Rejeitar(proposto, paginas, tentativa, abriuNoDestino)

                If proposto.Commands.Contains(SweepCommand.PublicarGeracao) Then
                    Dim res = _destino.Publicar(tentativa, proposto.Cobertura,
                                                a.CountBefore.Value, a.CountAfter.Value)
                    If res <> SinkPublishResult.Publicada Then
                        ' A persistencia tem fencing proprio e pode recusar
                        ' depois de o modelo aprovar. Nao instala o estado.
                        Return Rejeitar(SweepModel.Falhar(a, $"persistencia recusou: {res}"),
                                        paginas, tentativa, abriuNoDestino:=False)
                    End If
                End If

                ' So agora a tentativa E publicada.
                a = proposto.State
                Return New SweepResult(SweepConclusion.Publicada, a, Nothing, paginas, proposto.Cobertura)

            Catch ex As OperationCanceledException
                Return Cancelar(a, paginas, tentativa, abriuNoDestino)
            Catch ex As Exception
                ' Falha da fonte ou do destino. Em qualquer fronteira o efeito
                ' e o mesmo: descarta. Nunca publica metade.
                Return Falhar(a, $"{ex.GetType().Name}: {ex.Message}", paginas, tentativa, abriuNoDestino)
            End Try
        End Function

        ' ==============================================================

        Private Function Rejeitar(r As SweepOutcome, paginas As Integer,
                                  tentativa As Long, abriuNoDestino As Boolean) As SweepResult
            Descartar(tentativa, abriuNoDestino, r.Rejection)
            Return New SweepResult(SweepConclusion.Rejeitada, r.State, r.Rejection, paginas)
        End Function

        Private Function Cancelar(a As SweepAttempt, paginas As Integer,
                                  tentativa As Long, abriuNoDestino As Boolean) As SweepResult
            Dim r = SweepModel.Cancelar(a, "cancelado")
            Descartar(tentativa, abriuNoDestino, r.Rejection)
            Return New SweepResult(SweepConclusion.Cancelada, r.State, r.Rejection, paginas)
        End Function

        Private Function Falhar(a As SweepAttempt, motivo As String, paginas As Integer,
                                tentativa As Long, abriuNoDestino As Boolean) As SweepResult
            Dim r = SweepModel.Falhar(a, motivo)
            Descartar(tentativa, abriuNoDestino, r.Rejection)
            Return New SweepResult(SweepConclusion.Falhou, r.State, r.Rejection, paginas)
        End Function

        ''' <summary>
        ''' Marca a tentativa como descartada na persistência. Engole falha de
        ''' propósito: já estamos num caminho de erro, e uma exceção aqui
        ''' substituiria o motivo verdadeiro por um secundário — o diagnóstico
        ''' errado sobre uma rejeição certa que a §17 já custou caro.
        ''' </summary>
        Private Sub Descartar(tentativa As Long, abriuNoDestino As Boolean, motivo As String)
            If Not abriuNoDestino Then Return
            Try
                _destino.Descartar(tentativa, motivo)
            Catch
            End Try
        End Sub

    End Class

End Namespace

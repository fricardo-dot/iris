Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading

Namespace Global.Iris.Sync

    ''' <summary>
    ''' Uma página lida da fonte, com o universo <b>no momento da leitura</b>.
    '''
    ''' O universo vem junto de propósito. Ele é o que permite detectar que o
    ''' mundo mudou embaixo da varredura — e a Fase 2 inteira foi sobre isso.
    ''' Se a fonte devolvesse só as chaves, a orquestração teria de perguntar o
    ''' universo separado, e entre a pergunta e a leitura cabe uma mudança.
    ''' </summary>
    Public NotInheritable Class SourcePage
        Public ReadOnly Property Keys As IReadOnlyList(Of String)
        Public ReadOnly Property NextCursor As String
        Public ReadOnly Property Fim As Boolean
        Public ReadOnly Property Universo As SweepUniverse

        Public Sub New(keys As IEnumerable(Of String), nextCursor As String,
                       fim As Boolean, universo As SweepUniverse)
            Me.Keys = If(keys, Enumerable.Empty(Of String)()).ToList()
            Me.NextCursor = nextCursor
            Me.Fim = fim
            Me.Universo = universo
        End Sub
    End Class

    ''' <summary>
    ''' A fonte, do ponto de vista da orquestração. O adaptador do Outlook
    ''' implementa isto no 2.2b; os testes implementam com um falso que muta
    ''' <b>durante</b> a leitura.
    ''' </summary>
    Public Interface ISweepSource
        ''' <summary>Contagem declarada pela fonte. Pode mentir — o S6 existe por isso.</summary>
        Function Contar(ct As CancellationToken) As Integer
        Function LerPagina(cursor As String, tamanho As Integer, ct As CancellationToken) As SourcePage
        ''' <summary>O universo agora. Usado nas fronteiras que não leem página.</summary>
        Function UniversoAgora() As SweepUniverse
    End Interface

    ''' <summary>
    ''' Onde a varredura grava. Existe para a orquestração não depender de
    ''' SQLite: o 2.2a inteiro se prova com um destino falso, e o real é um
    ''' adaptador fino sobre o <c>CacheWriter</c>.
    ''' </summary>
    Public Interface ISweepSink
        ''' <summary>Página e checkpoint, atomicamente. Idempotente.</summary>
        Sub GravarPagina(pagina As Integer, chaves As IReadOnlyList(Of String), cursorDepois As String)
        ''' <summary>Geração, cabeça e dívida para a UI, atomicamente.</summary>
        Sub Publicar(a As SweepAttempt)
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

        Friend Sub New(c As SweepConclusion, a As SweepAttempt, motivo As String, paginas As Integer)
            Conclusion = c
            Attempt = a
            Me.Motivo = motivo
            Me.Paginas = paginas
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
    ''' O <see cref="SweepModel"/> é a máquina de estados e já decide tudo o
    ''' que é decidível sem tocar no mundo. Esta classe é o que <b>toca o
    ''' mundo</b>: chama a fonte, grava no destino, e traduz cada resposta do
    ''' modelo em ação. A separação importa porque o modelo pode ser testado
    ''' sem fonte nenhuma, e o runner pode ser testado com uma fonte que
    ''' mente.
    '''
    ''' <b>A regra que organiza o corpo todo:</b> o modelo transiciona
    ''' primeiro, o efeito acontece depois, e só se o modelo mandar. Nunca
    ''' gravar e depois perguntar se podia — gravar é o que deixa rastro, e
    ''' rastro de uma varredura que o modelo iria rejeitar é exatamente o
    ''' estado sujo que a Fase 2 passou inteira evitando.
    ''' </summary>
    Public NotInheritable Class SweepRunner

        ''' <summary>
        ''' Teto de páginas. Cursor que não avança é laço infinito, e laço
        ''' infinito na fila da STA trava a UI do usuário — não é hipótese
        ''' defensiva, é o modo de falha desta arquitetura.
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

            ' A EnvironmentPolicy entra AQUI, e nao mais fundo, porque e aqui
            ' que a decisao ainda e barata: recusar antes de abrir nao deixa
            ' tentativa pela metade para alguem retomar depois.
            Dim suportado = capacidades IsNot Nothing AndAlso
                            capacidades.PodeAfirmarCoberturaCompleta

            Dim r = SweepModel.Abrir(universo, epoca, numeroDaTentativa, suportado)
            If r.Rejected Then
                Return New SweepResult(SweepConclusion.Rejeitada, Nothing,
                    If(suportado, r.Rejection,
                       $"{r.Rejection} | capacidades: {If(capacidades?.Reason, "nenhuma")}"), 0)
            End If

            Dim a = r.State
            Dim paginas = 0

            Try
                ' --- contagem inicial ---
                If ct.IsCancellationRequested Then Return Cancelar(a, paginas)
                Dim antes = _fonte.Contar(ct)
                r = SweepModel.ContagemInicial(a, antes, _fonte.UniversoAgora())
                If r.Rejected Then Return Rejeitar(r, paginas)
                a = r.State

                ' --- paginacao ---
                Dim cursor As String = Nothing
                Do
                    If ct.IsCancellationRequested Then Return Cancelar(a, paginas)

                    paginas += 1
                    If paginas > MaxPaginas Then
                        Return Falhar(a, $"passou de {MaxPaginas} paginas: cursor nao avanca", paginas)
                    End If

                    Dim p = _fonte.LerPagina(cursor, _tamanhoPagina, ct)
                    If p Is Nothing Then Return Falhar(a, "fonte devolveu pagina nula", paginas)

                    r = SweepModel.Pagina(a, p.Keys, p.NextCursor, p.Universo)
                    If r.Rejected Then Return Rejeitar(r, paginas)
                    a = r.State

                    ' O efeito SO acontece se o modelo mandou.
                    If r.Commands.Contains(SweepCommand.StagePagina) Then
                        _destino.GravarPagina(paginas, p.Keys, p.NextCursor)
                    End If

                    If p.Fim Then Exit Do

                    ' Cursor que nao muda com pagina nao vazia e laco infinito
                    ' disfarcado de progresso.
                    If p.Keys.Count > 0 AndAlso String.Equals(cursor, p.NextCursor, StringComparison.Ordinal) Then
                        Return Falhar(a, "cursor nao avancou apos pagina nao vazia", paginas)
                    End If
                    cursor = p.NextCursor
                Loop

                ' --- contagem final ---
                If ct.IsCancellationRequested Then Return Cancelar(a, paginas)
                Dim depois = _fonte.Contar(ct)
                r = SweepModel.ContagemFinal(a, depois, _fonte.UniversoAgora())
                If r.Rejected Then Return Rejeitar(r, paginas)
                a = r.State

                ' --- publicacao ---
                If ct.IsCancellationRequested Then Return Cancelar(a, paginas)
                r = SweepModel.Publicar(a, _destino.EpocaCorrente())
                If r.Rejected Then Return Rejeitar(r, paginas)
                a = r.State

                If r.Commands.Contains(SweepCommand.PublicarGeracao) Then
                    _destino.Publicar(a)
                End If

                Return New SweepResult(SweepConclusion.Publicada, a, Nothing, paginas)

            Catch ex As OperationCanceledException
                Return Cancelar(a, paginas)
            Catch ex As Exception
                ' Falha da fonte ou do destino. Em qualquer fronteira o efeito
                ' e o mesmo: descarta. Nunca publica metade.
                Return Falhar(a, $"{ex.GetType().Name}: {ex.Message}", paginas)
            End Try
        End Function

        Private Shared Function Rejeitar(r As SweepOutcome, paginas As Integer) As SweepResult
            Return New SweepResult(SweepConclusion.Rejeitada, r.State, r.Rejection, paginas)
        End Function

        Private Shared Function Cancelar(a As SweepAttempt, paginas As Integer) As SweepResult
            Dim r = SweepModel.Cancelar(a, "cancelado")
            Return New SweepResult(SweepConclusion.Cancelada, r.State, r.Rejection, paginas)
        End Function

        Private Shared Function Falhar(a As SweepAttempt, motivo As String, paginas As Integer) As SweepResult
            Dim r = SweepModel.Falhar(a, motivo)
            Return New SweepResult(SweepConclusion.Falhou, r.State, r.Rejection, paginas)
        End Function

    End Class

End Namespace

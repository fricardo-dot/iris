Imports System.Collections.Generic

Namespace Global.Iris.Core

    ''' <summary>
    ''' Fonte de linhas ordenadas por instante, decrescente. Sem COM: quem
    ''' implementa é que fala com o Outlook.
    '''
    ''' Existe separada porque na Q1 eu escrevi um teste que verificava um
    ''' algoritmo DIFERENTE do que o script real rodava — o teste avançava
    ''' com <c>&lt; T</c> e o real com <c>&lt;= T</c>, e os cenários passavam
    ''' provando outra coisa. A correção estrutural foi ter UM algoritmo, e
    ''' o teste e o adaptador COM chamarem o mesmo.
    ''' </summary>
    Public Interface IRowSource(Of TLinha)

        ''' <summary>
        ''' Abre um cursor. <paramref name="fronteira"/> nulo é a primeira
        ''' página. <paramref name="inclusiva"/> só existe para o controle
        ''' negativo: em produção é sempre False.
        ''' </summary>
        Sub Abrir(fronteira As DateTimeOffset?, inclusiva As Boolean)

        ''' <summary>Lê até N linhas. Lista vazia significa fim.</summary>
        Function Ler(quantas As Integer) As IReadOnlyList(Of TLinha)

        ''' <summary>
        ''' Libera o cursor. Tem de ser IDEMPOTENTE e tem de funcionar
        ''' mesmo que <see cref="Abrir"/> tenha falhado no meio: a
        ''' implementacao COM adquire a Table antes de configurar colunas
        ''' e ordenacao, e qualquer uma dessas etapas pode lancar.
        ''' </summary>
        ''' <summary>
        ''' Libera o cursor. Tem de ser IDEMPOTENTE e tem de funcionar
        ''' mesmo que <see cref="Abrir"/> tenha falhado no meio: a
        ''' implementacao COM adquire a Table antes de configurar colunas
        ''' e ordenacao, e qualquer uma dessas etapas pode lancar.
        ''' </summary>
        ''' <summary>
        ''' Libera o cursor. Tem de ser IDEMPOTENTE e tem de funcionar
        ''' mesmo que <see cref="Abrir"/> tenha falhado no meio: a
        ''' implementacao COM adquire a Table antes de configurar colunas
        ''' e ordenacao, e qualquer uma dessas etapas pode lancar.
        ''' </summary>
        ''' <summary>
        ''' Libera o cursor. Tem de ser IDEMPOTENTE e tem de funcionar
        ''' mesmo que <see cref="Abrir"/> tenha falhado no meio: a
        ''' implementacao COM adquire a Table antes de configurar colunas
        ''' e ordenacao, e qualquer uma dessas etapas pode lancar.
        ''' </summary>
        Sub Fechar()

        Function InstanteDe(linha As TLinha) As DateTimeOffset

        ''' <summary>
        ''' Identidade da linha. O CURSOR não usa isto — ele não guarda
        ''' chave nenhuma. Quem usa é a travessia, para contar linhas
        ''' distintas e detectar paginação travada.
        ''' </summary>
        Function ChaveDe(linha As TLinha) As String

    End Interface

    ''' <summary>
    ''' Como a fonte se comportou ao entregar uma página.
    ''' </summary>
    Public NotInheritable Class PageOutcome(Of TLinha)
        Public ReadOnly Property Rows As IReadOnlyList(Of TLinha)

        ''' <summary>
        ''' Fronteira da próxima página, ou Nothing quando acabou.
        ''' </summary>
        Public ReadOnly Property NextBoundary As DateTimeOffset?

        ''' <summary>
        ''' Quantas linhas vieram da DRENAGEM, além do alvo pedido. Fica
        ''' visível porque a página drenada não tem teto, e uma página que
        ''' voltou com 45 quando pediram 30 não pode ser mistério.
        ''' </summary>
        Public ReadOnly Property DrainedExtra As Integer

        Public ReadOnly Property Ended As Boolean
            Get
                Return Not NextBoundary.HasValue
            End Get
        End Property

        Public Sub New(rows As IReadOnlyList(Of TLinha),
                       nextBoundary As DateTimeOffset?,
                       drainedExtra As Integer)
            Me.Rows = rows
            Me.NextBoundary = nextBoundary
            Me.DrainedExtra = drainedExtra
        End Sub
    End Class

    ''' <summary>
    ''' Defeitos que este algoritmo já teve, ligáveis por opção.
    '''
    ''' Existem para os controles negativos serem TESTE EXECUTÁVEL, e não um
    ''' número que alguém anotou depois de editar o arquivo à mão. Ligar
    ''' qualquer um deles em produção perde mensagem em silêncio — por isso
    ''' são Friend, e o adaptador COM não os expõe.
    ''' </summary>
    Friend NotInheritable Class PagingDefects
        Public Property SkipDrain As Boolean
        Public Property InclusiveBoundary As Boolean

        Public Shared ReadOnly Property None As PagingDefects
            Get
                Return New PagingDefects()
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Paginação por cursor sobre um campo que NÃO é ordem total.
    '''
    ''' ---------------------------------------------------------------
    ''' O PROBLEMA
    '''
    ''' <c>ReceivedTime</c> repete: vários itens compartilham o mesmo
    ''' instante, e o OOM não promete ordem estável dentro do empate.
    '''
    ''' Paginar com <c>&lt; T</c> pula os empatados que ficaram para trás.
    ''' Paginar com <c>&lt;= T</c> relê o grupo inteiro, e se ele for maior
    ''' que a página a paginação TRAVA: nenhum item novo, e a consulta
    ''' seguinte devolve os mesmos.
    '''
    ''' ---------------------------------------------------------------
    ''' O ALGORITMO
    '''
    '''   1. abre um cursor e lê uma página;
    '''   2. DRENA o resto do grupo do último instante NO MESMO cursor —
    '''      sem reabrir, então sem filtro envolvido nessa parte;
    '''   3. a próxima fronteira é aquele instante, com <c>&lt;</c> ESTRITO.
    '''
    ''' ---------------------------------------------------------------
    ''' CONSEQUÊNCIA QUE O CHAMADOR PRECISA ACEITAR
    '''
    ''' <c>tamanhoAlvo</c> é ALVO, não teto. A página vai até o fim do grupo
    ''' do último instante, então pode devolver mais. É isso que dispensa
    ''' guardar "chaves já vistas" no cursor: quando a página termina, o
    ''' grupo ficou para trás POR INTEIRO.
    '''
    ''' Um desenho que limitasse a página ao teto teria de reabrir de forma
    ''' inclusiva e carregar as chaves já emitidas — outro algoritmo, que
    ''' este teste não prova.
    ''' </summary>
    Public NotInheritable Class CursorPaging

        Public Const DefaultTargetSize As Integer = 30

        ''' <summary>Teto de aberturas numa travessia completa.</summary>
        Public Const MaxOpenings As Integer = 10000

        Public Shared Function ReadPage(Of TLinha)(
                source As IRowSource(Of TLinha),
                boundary As DateTimeOffset?,
                targetSize As Integer) As PageOutcome(Of TLinha)

            Return ReadPage(source, boundary, targetSize, PagingDefects.None)
        End Function

        Friend Shared Function ReadPage(Of TLinha)(
                source As IRowSource(Of TLinha),
                boundary As DateTimeOffset?,
                targetSize As Integer,
                defects As PagingDefects) As PageOutcome(Of TLinha)

            If source Is Nothing Then Throw New ArgumentNullException(NameOf(source))
            If targetSize <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(targetSize))
            If defects Is Nothing Then defects = PagingDefects.None

            Dim colhidas As New List(Of TLinha)()
            Dim extras = 0
            Dim proxima As DateTimeOffset? = Nothing

            ' Abrir fica DENTRO do Try. Fora dele, uma falha no meio da
            ' abertura — configurar coluna, ordenar — deixava o recurso ja
            ' adquirido sem ninguem para liberar.
            Try
                source.Abrir(boundary, defects.InclusiveBoundary)
                Dim primeira = source.Ler(targetSize)
                If primeira Is Nothing OrElse primeira.Count = 0 Then
                    ' Fim: nada nesta fronteira.
                    Return New PageOutcome(Of TLinha)(colhidas, Nothing, 0)
                End If

                colhidas.AddRange(primeira)
                Dim ultimo = source.InstanteDe(colhidas(colhidas.Count - 1))

                ' DRENA o grupo do último instante, no MESMO cursor. Para no
                ' primeiro instante diferente SEM consumi-lo: aquelas linhas
                ' voltam na consulta seguinte, porque a fronteira é estrita e
                ' elas são menores que ela.
                Dim grupoCompleto = defects.SkipDrain
                If Not defects.SkipDrain Then
                    Do
                        Dim extra = source.Ler(targetSize)
                        If extra Is Nothing OrElse extra.Count = 0 Then
                            ' A fonte acabou durante a drenagem: não existe
                            ' nada mais antigo que este instante.
                            Return New PageOutcome(Of TLinha)(colhidas, Nothing, extras)
                        End If

                        Dim saiu = False
                        For Each linha In extra
                            If source.InstanteDe(linha) <> ultimo Then
                                saiu = True
                                Exit For
                            End If
                            colhidas.Add(linha)
                            extras += 1
                        Next
                        If saiu Then
                            grupoCompleto = True
                            Exit Do
                        End If
                    Loop
                End If

                If grupoCompleto Then proxima = ultimo
                Return New PageOutcome(Of TLinha)(colhidas, proxima, extras)
            Finally
                ' Fechar SEMPRE. Um RCW que sobrevive à chamada é o defeito
                ' que a regra R7 do ESCOPO.md cobra.
                source.Fechar()
            End Try
        End Function

        ''' <summary>
        ''' Travessia completa, para teste e diagnóstico. A produção pagina
        ''' uma página por vez, sob demanda da UI.
        ''' </summary>
        Friend Shared Function Traverse(Of TLinha)(
                source As IRowSource(Of TLinha),
                targetSize As Integer,
                defects As PagingDefects) As TraverseOutcome(Of TLinha)

            Dim todas As New List(Of TLinha)()
            Dim vistas As New HashSet(Of String)()
            Dim fronteira As DateTimeOffset? = Nothing
            Dim aberturas = 0

            Do
                If aberturas >= MaxOpenings Then
                    ' Sair pelo teto DEVOLVENDO parcial em silêncio é o que o
                    ' PowerShell fazia. Aqui é falha declarada.
                    Return New TraverseOutcome(Of TLinha)(todas, aberturas, exhausted:=False)
                End If

                Dim pagina = ReadPage(source, fronteira, targetSize, defects)
                aberturas += 1

                Dim novas = 0
                For Each linha In pagina.Rows
                    If vistas.Add(source.ChaveDe(linha)) Then
                        todas.Add(linha)
                        novas += 1
                    End If
                Next

                If pagina.Ended Then
                    Return New TraverseOutcome(Of TLinha)(todas, aberturas, exhausted:=True)
                End If

                ' Página inteira repetida e ainda não é o fim: a fronteira
                ' não anda. É o sintoma exato do defeito da fronteira
                ' inclusiva quando o grupo empatado é maior que a página —
                ' nada é novo, e insistir só repete a mesma consulta.
                If novas = 0 Then
                    Return New TraverseOutcome(Of TLinha)(todas, aberturas, exhausted:=False)
                End If

                fronteira = pagina.NextBoundary
            Loop
        End Function

    End Class

    Friend NotInheritable Class TraverseOutcome(Of TLinha)
        Public ReadOnly Property Rows As IReadOnlyList(Of TLinha)
        Public ReadOnly Property Openings As Integer
        ''' <summary>False = parou pelo teto ou por travamento, não por fim.</summary>
        Public ReadOnly Property Exhausted As Boolean

        Public Sub New(rows As IReadOnlyList(Of TLinha), openings As Integer, exhausted As Boolean)
            Me.Rows = rows
            Me.Openings = openings
            Me.Exhausted = exhausted
        End Sub
    End Class

End Namespace

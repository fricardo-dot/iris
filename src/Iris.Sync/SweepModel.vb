Imports System.Collections.Generic
Imports System.Linq

Namespace Global.Iris.Sync

    ''' <summary>
    ''' Em que ponto uma TENTATIVA de varredura está.
    '''
    ''' Tentativa não é geração. A geração é o resultado PUBLICADO, imutável,
    ''' e só existe quando a tentativa passou por tudo. Misturar as duas foi
    ''' o bloqueador da 1ª versão do plano: "a geração só existe quando
    ''' válida" e "checkpoint por geração" não cabem juntos, porque
    ''' checkpoint é estado de trabalho incompleto.
    ''' </summary>
    Public Enum AttemptStage
        Aberta
        ContagemInicialLida
        Varrendo
        ContagemFinalLida
        Publicada
        Descartada
    End Enum

    ''' <summary>
    ''' O universo de uma varredura. Faz parte da IDENTIDADE da geração
    ''' (I7): gerações de universos diferentes não se comparam, e comparar
    ''' conclui que milhares de itens sumiram.
    ''' </summary>
    Public NotInheritable Class SweepUniverse
        Public ReadOnly Property StoreKey As String
        Public ReadOnly Property FolderKey As String
        Public ReadOnly Property Filter As String
        Public ReadOnly Property RetentionCutoff As String
        Public ReadOnly Property AlgorithmVersion As Integer
        Public ReadOnly Property EnvironmentFingerprint As String

        Public Sub New(storeKey As String, folderKey As String, filter As String,
                       retentionCutoff As String, algorithmVersion As Integer,
                       environmentFingerprint As String)
            Me.StoreKey = If(storeKey, "")
            Me.FolderKey = If(folderKey, "")
            Me.Filter = If(filter, "")
            Me.RetentionCutoff = If(retentionCutoff, "")
            Me.AlgorithmVersion = algorithmVersion
            Me.EnvironmentFingerprint = If(environmentFingerprint, "")
        End Sub

        Public Function Fingerprint() As String
            Return $"{StoreKey}|{FolderKey}|{Filter}|{RetentionCutoff}|{AlgorithmVersion}|{EnvironmentFingerprint}"
        End Function

        Public Function MesmoQue(outro As SweepUniverse) As Boolean
            Return outro IsNot Nothing AndAlso
                   String.Equals(Fingerprint(), outro.Fingerprint(), StringComparison.Ordinal)
        End Function
    End Class

    Public Enum SweepCommand
        StagePagina
        PublicarGeracao
        MarcarNaoVistosComoSuspeitos
        EmitirPublicationLog
        DescartarTentativa
        AgendarRetry
        MarcarPastaInstavel
    End Enum

    Public NotInheritable Class SweepOutcome
        Public ReadOnly Property State As SweepAttempt
        Public ReadOnly Property Commands As IReadOnlyList(Of SweepCommand)
        Public ReadOnly Property Rejection As String

        ''' <summary>
        ''' A cobertura que esta publicacao pode declarar. So faz sentido no
        ''' resultado de <see cref="SweepModel.Publicar"/>; nas outras
        ''' transicoes fica <c>Desconhecida</c>.
        ''' </summary>
        Public ReadOnly Property Cobertura As FolderCoverage

        Friend Sub New(state As SweepAttempt, commands As IEnumerable(Of SweepCommand),
                       Optional rejection As String = Nothing,
                       Optional cobertura As FolderCoverage = FolderCoverage.Desconhecida)
            Me.State = state
            Me.Commands = If(commands, Enumerable.Empty(Of SweepCommand)()).ToList()
            Me.Rejection = rejection
            Me.Cobertura = cobertura
        End Sub

        Public ReadOnly Property Rejected As Boolean
            Get
                Return Rejection IsNot Nothing
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Uma tentativa de varredura, imutável — cada transição devolve outra.
    ''' </summary>
    Public NotInheritable Class SweepAttempt
        Public ReadOnly Property Stage As AttemptStage
        Public ReadOnly Property Universe As SweepUniverse
        ''' <summary>Época capturada na abertura. É o fencing do CAS.</summary>
        Public ReadOnly Property ReconcileEpoch As Long
        Public ReadOnly Property CountBefore As Integer?
        Public ReadOnly Property CountAfter As Integer?
        ''' <summary>Chaves já vistas. Vive em scan_stage, não no cursor.</summary>
        Public ReadOnly Property StagedKeys As IReadOnlyCollection(Of String)
        Public ReadOnly Property RowsRead As Integer
        Public ReadOnly Property Cursor As String
        Public ReadOnly Property AttemptNumber As Integer

        Friend Sub New(stage As AttemptStage, universe As SweepUniverse, epoch As Long,
                       countBefore As Integer?, countAfter As Integer?,
                       staged As IEnumerable(Of String), rowsRead As Integer,
                       cursor As String, attemptNumber As Integer)
            Me.Stage = stage
            Me.Universe = universe
            Me.ReconcileEpoch = epoch
            Me.CountBefore = countBefore
            Me.CountAfter = countAfter
            Me.StagedKeys = New HashSet(Of String)(If(staged, Enumerable.Empty(Of String)()),
                                                   StringComparer.Ordinal)
            Me.RowsRead = rowsRead
            Me.Cursor = cursor
            Me.AttemptNumber = attemptNumber
        End Sub

        Public ReadOnly Property DistinctKeys As Integer
            Get
                Return StagedKeys.Count
            End Get
        End Property
    End Class

    ''' <summary>
    ''' A máquina de estados de uma varredura. Função pura: estado + evento
    ''' devolve estado novo e comandos. Relógio, IDs e época ENTRAM.
    '''
    ''' ------------------------------------------------------------------
    ''' A DIVISÃO DE AUTORIDADE, que é o coração disto
    '''
    ''' <b>S6 REJEITA, e não valida.</b> Contagens que concordam não provam
    ''' que a varredura viu o conjunto certo — a §17.1 lista os falsos
    ''' negativos, e o principal é a mutação balanceada: sai um, entra
    ''' outro, e as três contagens continuam iguais.
    '''
    ''' <b>S7 AUTORIZA transições negativas, e não valida varredura.</b> Sem
    ''' cobertura comprovada, nenhuma ausência é confirmada — a §19.2 mediu
    ''' pastas cheias reportando zero itens.
    '''
    ''' Nenhum caminho daqui leva a <c>AusenteDaPasta</c>. Publicar autoriza
    ''' presença e suspeita, só.
    ''' </summary>
    Public NotInheritable Class SweepModel

        ''' <summary>
        ''' Abre uma tentativa. Só recusa quando o ambiente é
        ''' <b>desconhecido</b> — não quando ele é limitado.
        '''
        ''' <b>Isto mudou, e a mudança corrige um conflito de contrato.</b> A
        ''' D2 original dizia "ambiente não medido RECUSA operar", e o
        ''' <c>SweepRunner</c> chegou a implementar isso passando
        ''' <c>PodeAfirmarCoberturaCompleta</c> como permissão para abrir. O
        ''' efeito: na caixa do usuário — cached, janela não legível — o Iris
        ''' não varreria <b>nada</b>. O produto não faria nada.
        '''
        ''' A §23 aceitou coisa diferente e mais precisa: cached não
        ''' <i>conclui</i> ausência, não <i>afirma</i> cobertura completa e não
        ''' <i>usa</i> incremental. Não aceitou "não opera". São duas perguntas
        ''' separadas, e misturá-las custou o produto inteiro:
        '''
        '''   <i>posso observar e guardar positivos?</i>
        '''   <i>posso declarar que observei tudo?</i>
        '''
        ''' Falta de cobertura não invalida observação positiva. Então varre,
        ''' encena, publica — como <b>parcial</b>.
        '''
        ''' O que ainda recusa é o ambiente <b>não identificado</b>: sem saber
        ''' onde se está, nem o positivo tem a quem ser atribuído.
        ''' </summary>
        Public Shared Function Abrir(universe As SweepUniverse, epoch As Long,
                                     attemptNumber As Integer,
                                     ambienteIdentificado As Boolean) As SweepOutcome
            If universe Is Nothing Then
                Return New SweepOutcome(Nothing, Nothing, "universo nulo")
            End If
            If Not ambienteIdentificado Then
                Return New SweepOutcome(Nothing, Nothing,
                    "ambiente nao identificado: " & universe.EnvironmentFingerprint)
            End If
            Return New SweepOutcome(
                New SweepAttempt(AttemptStage.Aberta, universe, epoch, Nothing, Nothing,
                                 Nothing, 0, Nothing, attemptNumber),
                Nothing)
        End Function

        Public Shared Function ContagemInicial(a As SweepAttempt, count As Integer,
                                               universoLido As SweepUniverse) As SweepOutcome
            Dim erro = Guardas(a, {AttemptStage.Aberta}, universoLido)
            If erro IsNot Nothing Then Return Descartar(a, erro)
            If count < 0 Then Return Descartar(a, "contagem inicial negativa")
            Return New SweepOutcome(
                New SweepAttempt(AttemptStage.ContagemInicialLida, a.Universe, a.ReconcileEpoch,
                                 count, Nothing, a.StagedKeys, 0, a.Cursor, a.AttemptNumber),
                Nothing)
        End Function

        ''' <summary>
        ''' Uma página lida. Página e cursor avançam JUNTOS — na
        ''' implementação, na mesma transação. Separá-los cria o estado
        ''' inválido "cursor à frente da página", que na retomada pula
        ''' linhas em silêncio.
        ''' </summary>
        Public Shared Function Pagina(a As SweepAttempt, chaves As IEnumerable(Of String),
                                      proximoCursor As String,
                                      universoLido As SweepUniverse) As SweepOutcome
            Dim erro = Guardas(a, {AttemptStage.ContagemInicialLida, AttemptStage.Varrendo},
                               universoLido)
            If erro IsNot Nothing Then Return Descartar(a, erro)

            Dim lista = If(chaves, Enumerable.Empty(Of String)()).ToList()
            For Each k In lista
                If String.IsNullOrWhiteSpace(k) Then
                    Return Descartar(a, "chave vazia ou invalida na pagina")
                End If
            Next

            Dim vistas = New HashSet(Of String)(a.StagedKeys, StringComparer.Ordinal)
            For Each k In lista
                If Not vistas.Add(k) Then
                    ' A mesma chave em paginas diferentes. Pode ser cursor
                    ' repetindo ou fonte instavel; das duas, a varredura nao
                    ' vale.
                    Return Descartar(a, "chave repetida entre paginas: a varredura nao e confiavel")
                End If
            Next

            Return New SweepOutcome(
                New SweepAttempt(AttemptStage.Varrendo, a.Universe, a.ReconcileEpoch,
                                 a.CountBefore, Nothing, vistas, a.RowsRead + lista.Count,
                                 proximoCursor, a.AttemptNumber),
                {SweepCommand.StagePagina})
        End Function

        Public Shared Function ContagemFinal(a As SweepAttempt, count As Integer,
                                             universoLido As SweepUniverse) As SweepOutcome
            Dim erro = Guardas(a, {AttemptStage.Varrendo, AttemptStage.ContagemInicialLida},
                               universoLido)
            If erro IsNot Nothing Then Return Descartar(a, erro)
            If count < 0 Then Return Descartar(a, "contagem final negativa")
            Return New SweepOutcome(
                New SweepAttempt(AttemptStage.ContagemFinalLida, a.Universe, a.ReconcileEpoch,
                                 a.CountBefore, count, a.StagedKeys, a.RowsRead,
                                 a.Cursor, a.AttemptNumber),
                Nothing)
        End Function

        ''' <summary>
        ''' Publica, se o S6 não rejeitar E o CAS de época passar.
        '''
        ''' <paramref name="epocaCorrenteDaPasta"/> é o fencing: se outra
        ''' tentativa publicou enquanto esta corria, esta perde. Sem isso,
        ''' uma varredura lenta terminando depois de uma rápida sobrescreve o
        ''' resultado mais novo com o mais velho — que é o item 10 do antigo
        ''' §8.
        ''' </summary>
        Public Shared Function Publicar(a As SweepAttempt,
                                        epocaCorrenteDaPasta As Long,
                                        podeAfirmarCoberturaCompleta As Boolean) As SweepOutcome
            If a Is Nothing Then Return New SweepOutcome(Nothing, Nothing, "tentativa nula")
            If a.Stage = AttemptStage.Publicada Then
                ' Republicar a MESMA tentativa e idempotente: nao cria
                ' segunda geracao nem segundo evento.
                Return New SweepOutcome(a, Nothing)
            End If
            If a.Stage <> AttemptStage.ContagemFinalLida Then
                Return Descartar(a, $"publicar em estagio {a.Stage}")
            End If

            ' --- S6, que REJEITA ---
            If Not a.CountBefore.HasValue OrElse Not a.CountAfter.HasValue Then
                Return Descartar(a, "faltou contagem")
            End If
            If a.RowsRead <> a.CountBefore.Value Then
                Return Descartar(a, $"lidas {a.RowsRead} <> antes {a.CountBefore.Value}")
            End If
            If a.RowsRead <> a.CountAfter.Value Then
                Return Descartar(a, $"lidas {a.RowsRead} <> depois {a.CountAfter.Value}")
            End If
            If a.DistinctKeys <> a.RowsRead Then
                Return Descartar(a, $"distintas {a.DistinctKeys} <> lidas {a.RowsRead}")
            End If

            ' --- fencing ---
            If a.ReconcileEpoch <> epocaCorrenteDaPasta Then
                Return Descartar(a,
                    $"epoca {a.ReconcileEpoch} <> corrente {epocaCorrenteDaPasta}: " &
                    "outra geracao publicou enquanto esta corria")
            End If

            ' A COBERTURA e decidida aqui, e a decisao e um degrau, nao uma
            ' recusa. Sem autorizacao para afirmar cobertura completa, a
            ' geracao sai como PARCIAL - e continua valendo. Observacao
            ' positiva nao fica invalida por faltar cobertura (§23).
            Dim cobertura = If(podeAfirmarCoberturaCompleta,
                               FolderCoverage.Completa, FolderCoverage.Parcial)

            ' MarcarNaoVistosComoSuspeitos sai SEMPRE, e nao depende da
            ' cobertura. Suspeito nao e conclusao negativa: e exatamente "nao
            ' vi e nao posso concluir por que". Condiciona-lo a cobertura
            ' completa bloquearia justamente a resposta honesta, deixando o
            ' item como Presente - que e uma afirmacao mais forte do que a
            ' evidencia sustenta.
            Return New SweepOutcome(
                New SweepAttempt(AttemptStage.Publicada, a.Universe, a.ReconcileEpoch,
                                 a.CountBefore, a.CountAfter, a.StagedKeys, a.RowsRead,
                                 a.Cursor, a.AttemptNumber),
                {SweepCommand.PublicarGeracao,
                 SweepCommand.MarcarNaoVistosComoSuspeitos,
                 SweepCommand.EmitirPublicationLog},
                Nothing, cobertura)
        End Function

        ''' <summary>
        ''' Cancelamento e falha da fonte. Em QUALQUER fronteira o efeito é o
        ''' mesmo: descarta. Nunca publica metade, nunca converte presença em
        ''' suspeita.
        ''' </summary>
        Public Shared Function Cancelar(a As SweepAttempt, motivo As String) As SweepOutcome
            Return Descartar(a, If(motivo, "cancelado"))
        End Function

        Public Shared Function Falhar(a As SweepAttempt, motivo As String) As SweepOutcome
            Return Descartar(a, If(motivo, "falha da fonte"))
        End Function

        ''' <summary>
        ''' Retomar uma tentativa interrompida.
        '''
        ''' Divergência em QUALQUER componente do universo abandona — versão
        ''' do algoritmo, pasta, store, filtro, cutoff, ambiente. E a época
        ''' também: retomar sob outra época é retomar noutro mundo.
        ''' </summary>
        Public Shared Function Retomar(a As SweepAttempt, universoAgora As SweepUniverse,
                                       epocaAgora As Long) As SweepOutcome
            If a Is Nothing Then Return New SweepOutcome(Nothing, Nothing, "tentativa nula")
            If a.Stage = AttemptStage.Publicada OrElse a.Stage = AttemptStage.Descartada Then
                Return Descartar(a, $"retomar tentativa {a.Stage}")
            End If
            If Not a.Universe.MesmoQue(universoAgora) Then
                Return Descartar(a, "universo mudou desde a interrupcao")
            End If
            If a.ReconcileEpoch <> epocaAgora Then
                Return Descartar(a, "epoca mudou desde a interrupcao")
            End If
            ' Estado invalido: cursor a frente sem pagina correspondente.
            If a.Cursor IsNot Nothing AndAlso a.RowsRead = 0 Then
                Return Descartar(a, "cursor avancado sem pagina staged")
            End If
            Return New SweepOutcome(a, Nothing)
        End Function

        ' ==============================================================

        ' Recebe TODOS os estagios aceitos de uma vez.
        '
        ' Antes eram duas chamadas em sequencia — tenta um estagio, se
        ' falhar tenta outro — e a segunda mensagem MASCARAVA a primeira: a
        ' rejeicao por universo trocado saia como "estagio errado". Um
        ' diagnostico errado sobre uma rejeicao certa e o tipo de coisa que
        ' custa uma hora quando o defeito for de verdade.
        Private Shared Function Guardas(a As SweepAttempt, aceitos As AttemptStage(),
                                        universoLido As SweepUniverse) As String
            If a Is Nothing Then Return "tentativa nula"
            If a.Stage = AttemptStage.Publicada Then Return "tentativa ja publicada e imutavel"
            If a.Stage = AttemptStage.Descartada Then Return "tentativa descartada nao volta"

            ' Universo ANTES do estagio: e a rejeicao mais informativa das duas.
            '
            ' E universo AUSENTE e REJEICAO, nao dispensa da guarda. Antes isto
            ' era "If universoLido IsNot Nothing AndAlso ..." — uma fonte que
            ' devolvesse Nothing DESLIGAVA a verificacao e passava como se nada
            ' tivesse mudado. E o formato de defeito que esta fase inteira
            ' persegue: a protecao some junto com o dado que ela protegia, e o
            ' resultado parece igual ao de um caso legitimo.
            If universoLido Is Nothing Then
                Return "fonte nao informou o universo: sem ele nao da para saber se mudou"
            End If
            If Not a.Universe.MesmoQue(universoLido) Then
                Return "universo mudou no meio da tentativa"
            End If
            If Not aceitos.Contains(a.Stage) Then
                Return $"estagio {a.Stage}, esperado {String.Join(" ou ", aceitos)}"
            End If
            Return Nothing
        End Function

        Private Shared Function Descartar(a As SweepAttempt, motivo As String) As SweepOutcome
            Dim novo As SweepAttempt = Nothing
            If a IsNot Nothing Then
                novo = New SweepAttempt(AttemptStage.Descartada, a.Universe, a.ReconcileEpoch,
                                        a.CountBefore, a.CountAfter, a.StagedKeys,
                                        a.RowsRead, a.Cursor, a.AttemptNumber)
            End If
            Return New SweepOutcome(novo,
                                    {SweepCommand.DescartarTentativa, SweepCommand.AgendarRetry},
                                    motivo)
        End Function

    End Class

End Namespace

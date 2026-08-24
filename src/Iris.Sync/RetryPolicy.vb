Imports System.Collections.Generic

Namespace Global.Iris.Sync

    Public Enum RetryDecision
        ''' <summary>Tenta agora.</summary>
        Tentar
        ''' <summary>Espera o backoff.</summary>
        Aguardar
        ''' <summary>Desistiu nesta ativação. A pasta degrada.</summary>
        Degradar
    End Enum

    Public NotInheritable Class RetryOutcome
        Public ReadOnly Property Decision As RetryDecision
        Public ReadOnly Property NotBefore As DateTimeOffset?
        Public ReadOnly Property Reason As String

        Friend Sub New(decision As RetryDecision, notBefore As DateTimeOffset?, reason As String)
            Me.Decision = decision
            Me.NotBefore = notBefore
            Me.Reason = reason
        End Sub
    End Class

    ''' <summary>
    ''' Quantas vezes repetir uma varredura que foi invalidada, e quando.
    '''
    ''' O ponto de partida é a §17.1: discordância entre as contagens
    ''' INVALIDA a geração, e a resposta é "descarte e repita". Mas "repita"
    ''' sem teto é um laço infinito numa pasta com tráfego — e a Caixa de
    ''' Entrada recebe mensagem sozinha, então a discordância pode ser
    ''' permanente.
    '''
    ''' Por isso o limite existe, e a degradação não é derrota: a pasta fica
    ''' <c>instavel</c>, a geração ANTERIOR é preservada, nenhuma ausência é
    ''' confirmada, e a fila é cedida para outra pasta. Um Iris que insiste
    ''' para sempre numa pasta movimentada trava a fila da STA e não entrega
    ''' nada — pior que um que admite não ter conseguido.
    ''' </summary>
    Public NotInheritable Class RetryPolicy

        ''' <summary>
        ''' Tentativas consecutivas por ativação. Três, e a quarta não
        ''' acontece.
        ''' </summary>
        Public Const MaxTentativas As Integer = 3

        ''' <summary>
        ''' 30 s, 2 min, 8 min, e depois teto de 15 min. A progressão é para
        ''' a pasta movimentada parar de ser sondada em rajada.
        ''' </summary>
        Public Shared ReadOnly Backoff As IReadOnlyList(Of TimeSpan) = New TimeSpan() {
            TimeSpan.FromSeconds(30),
            TimeSpan.FromMinutes(2),
            TimeSpan.FromMinutes(8)
        }

        Public Shared ReadOnly TetoDeBackoff As TimeSpan = TimeSpan.FromMinutes(15)

        Public Shared Function EsperaApos(tentativasFalhas As Integer) As TimeSpan
            If tentativasFalhas <= 0 Then Return TimeSpan.Zero
            If tentativasFalhas <= Backoff.Count Then Return Backoff(tentativasFalhas - 1)
            Return TetoDeBackoff
        End Function

        ''' <summary>
        ''' Decide o que fazer, dado o estado da pasta e o relógio.
        '''
        ''' O relógio ENTRA como parâmetro. Um relógio interno tornaria isto
        ''' intestável, e a §12 do FASE1 já cobrou esse desenho no
        ''' <c>DirtyDebounce</c>.
        ''' </summary>
        ''' <param name="emVoo">
        ''' Já há uma tentativa desta pasta em execução. Nunca pode haver
        ''' duas: duas varreduras concorrentes da mesma pasta produzem
        ''' exatamente o corte fraturado que a §16.1 mediu.
        ''' </param>
        Public Shared Function Decidir(tentativasFalhasNestaAtivacao As Integer,
                                       ultimaFalhaEm As DateTimeOffset?,
                                       agora As DateTimeOffset,
                                       emVoo As Boolean) As RetryOutcome

            If emVoo Then
                Return New RetryOutcome(RetryDecision.Aguardar, Nothing,
                                        "ja ha uma tentativa desta pasta em voo")
            End If

            If tentativasFalhasNestaAtivacao >= MaxTentativas Then
                Return New RetryOutcome(RetryDecision.Degradar, Nothing,
                                        $"{tentativasFalhasNestaAtivacao} tentativas sem convergir")
            End If

            If tentativasFalhasNestaAtivacao <= 0 Then
                Return New RetryOutcome(RetryDecision.Tentar, Nothing, "primeira tentativa")
            End If

            If Not ultimaFalhaEm.HasValue Then
                ' Falhou mas ninguem anotou quando. Fail-closed: espera.
                Return New RetryOutcome(RetryDecision.Aguardar, Nothing,
                                        "sem instante da ultima falha")
            End If

            Dim liberaEm = ultimaFalhaEm.Value + EsperaApos(tentativasFalhasNestaAtivacao)
            If agora < liberaEm Then
                Return New RetryOutcome(RetryDecision.Aguardar, liberaEm, "backoff")
            End If

            Return New RetryOutcome(RetryDecision.Tentar, Nothing,
                                    $"backoff cumprido, tentativa {tentativasFalhasNestaAtivacao + 1}")
        End Function

        ''' <summary>
        ''' Quantas tentativas o relógio autoriza AGORA.
        '''
        ''' Existe por um adversário específico: relógio avançado
        ''' arbitrariamente — máquina que hibernou, horário que mudou — não
        ''' pode liberar uma rajada de retries atrasados de uma vez. O
        ''' backoff é sobre espaçar, e um salto no relógio não desfaz isso.
        ''' </summary>
        Public Shared Function TentativasAutorizadasApos(saltoDoRelogio As TimeSpan) As Integer
            Return 1
        End Function

    End Class

End Namespace

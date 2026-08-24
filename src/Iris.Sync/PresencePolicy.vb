Imports Iris.Model

Namespace Global.Iris.Sync

    ''' <summary>
    ''' O que o Iris sabe sobre uma associação item–pasta.
    '''
    ''' <c>NaoVerificado</c> é o primeiro valor DE PROPÓSITO: uma associação
    ''' que ninguém preencheu não pode nascer afirmando presença. É a §3.8, e
    ''' o mesmo motivo pelo qual <c>ProtectionState.Unknown</c> é o primeiro
    ''' lá — default que mente é pior que default ausente.
    ''' </summary>
    Public Enum PresenceState
        NaoVerificado
        Presente
        ''' <summary>
        ''' Não foi visto numa geração publicada. NÃO é ausência: a §16.3
        ''' mediu que a pasta de origem não distingue movido de excluído.
        ''' </summary>
        Suspeito
        ''' <summary>
        ''' Confirmado ausente DAQUELA PASTA. Nunca significa "não existe" —
        ''' o item pode estar em outra pasta, e a §11.1 mediu que a chave
        ''' muda no Move, então procurar pela antiga não o encontra.
        ''' </summary>
        AusenteDaPasta
    End Enum

    ''' <summary>
    ''' Se o item é alcançável pelo OOM no universo observado.
    '''
    ''' Dimensão SEPARADA da presença, e a separação não é elegância: a
    ''' §19.2 mediu pastas cheias reportando ZERO itens porque o conteúdo
    ''' está fora da janela de cache. "Não encontrei" é compatível com
    ''' exclusão, movimento, saída da janela, indisponibilidade e mudança de
    ''' universo — cinco coisas diferentes que um enum só achataria numa.
    ''' </summary>
    Public Enum Observability
        Desconhecida
        Observavel
        ''' <summary>
        ''' Fora do universo observado — tipicamente a janela de cache. Só
        ''' pode ser afirmado com cutoff EXPLÍCITO e evidência do MESMO
        ''' universo. Data antiga sozinha não prova expulsão.
        ''' </summary>
        NaoObservavelNoUniverso
    End Enum

    ''' <summary>
    ''' O quanto da pasta o cache alcança. Sem isto, <c>Count = 0</c> seria
    ''' lido como "pasta vazia" — e na caixa medida dezenas de pastas cheias
    ''' reportam zero (§19.2). É o S7.
    ''' </summary>
    Public Enum FolderCoverage
        Desconhecida
        Parcial
        Completa
    End Enum

    ''' <summary>
    ''' O que uma verificação individual de um suspeito descobriu.
    ''' </summary>
    Public Enum ProbeResult
        ''' <summary>Não deu para decidir. É um resultado, não uma falha.</summary>
        Inconclusivo
        EncontradoNaMesmaPasta
        EncontradoEmOutraPasta
        NaoEncontrado
    End Enum

    ''' <summary>
    ''' Uma observação sobre uma associação, com a proveniência que o I4
    ''' exige — de qual geração veio, em qual universo, e com que cobertura.
    ''' </summary>
    Public NotInheritable Class PresenceObservation
        Public ReadOnly Property Seen As Boolean
        Public ReadOnly Property GenerationKey As Long
        Public ReadOnly Property UniverseFingerprint As String
        Public ReadOnly Property Coverage As FolderCoverage
        Public ReadOnly Property HasExplicitCutoff As Boolean

        Public Sub New(seen As Boolean, generationKey As Long,
                       universeFingerprint As String, coverage As FolderCoverage,
                       Optional hasExplicitCutoff As Boolean = False)
            Me.Seen = seen
            Me.GenerationKey = generationKey
            Me.UniverseFingerprint = If(universeFingerprint, "")
            Me.Coverage = coverage
            Me.HasExplicitCutoff = hasExplicitCutoff
        End Sub
    End Class

    ''' <summary>
    ''' O estado de uma associação, com a versão que o CAS usa.
    ''' </summary>
    Public NotInheritable Class AssociationState
        Public ReadOnly Property Presence As PresenceState
        Public ReadOnly Property Observability As Observability
        ''' <summary>Universo em que a última conclusão foi tirada.</summary>
        Public ReadOnly Property ConcludedInUniverse As String
        ''' <summary>Geração que produziu o estado atual. Fencing.</summary>
        Public ReadOnly Property GenerationKey As Long
        ''' <summary>Incrementa a cada transição aplicada. CAS.</summary>
        Public ReadOnly Property Version As Long

        Public Sub New(presence As PresenceState, observability As Observability,
                       concludedInUniverse As String, generationKey As Long, version As Long)
            Me.Presence = presence
            Me.Observability = observability
            Me.ConcludedInUniverse = If(concludedInUniverse, "")
            Me.GenerationKey = generationKey
            Me.Version = version
        End Sub

        Public Shared Function Nova() As AssociationState
            Return New AssociationState(PresenceState.NaoVerificado,
                                        Observability.Desconhecida, "", 0, 0)
        End Function

        Friend Function Com(presence As PresenceState, observability As Observability,
                            universo As String, generationKey As Long) As AssociationState
            Return New AssociationState(presence, observability, universo,
                                        generationKey, Version + 1)
        End Function
    End Class

    ''' <summary>
    ''' As transições de presença e observabilidade. Função pura.
    '''
    ''' A regra que mais importa está em <see cref="AplicarVerificacao"/>:
    ''' <c>NaoEncontrado</c> só vira <c>AusenteDaPasta</c> quando a cobertura
    ''' da pasta é <b>Completa</b> e no MESMO universo. Com cobertura parcial
    ''' ou desconhecida, o suspeito continua suspeito — para sempre, se for o
    ''' caso.
    '''
    ''' Isso significa que, num Outlook com janela de cache curta, muitas
    ''' associações ficarão suspeitas indefinidamente. <b>É o resultado
    ''' correto</b>, não uma falha do algoritmo: o Iris não tem como saber, e
    ''' inventar a resposta apagaria correspondência do usuário.
    ''' </summary>
    Public NotInheritable Class PresencePolicy

        ''' <summary>
        ''' O que uma geração PUBLICADA faz com uma associação.
        '''
        ''' Geração publicada autoriza positivos e suspeitas. Ela NÃO
        ''' certifica completude nem ausência — nenhum item chega a
        ''' <c>AusenteDaPasta</c> por aqui.
        ''' </summary>
        Public Shared Function AplicarGeracao(atual As AssociationState,
                                              obs As PresenceObservation) As AssociationState
            If atual Is Nothing Then atual = AssociationState.Nova()
            If obs Is Nothing Then Return atual

            ' Geração mais VELHA que o estado atual não retrocede nada.
            ' Fencing: a §17 S2 e o CAS de época.
            If obs.GenerationKey < atual.GenerationKey Then Return atual

            If obs.Seen Then
                ' Visto é visto: vira Presente venha de onde vier, inclusive
                ' de AusenteDaPasta — o item voltou.
                Return atual.Com(PresenceState.Presente, Observability.Observavel,
                                 obs.UniverseFingerprint, obs.GenerationKey)
            End If

            ' NÃO visto.
            Select Case atual.Presence
                Case PresenceState.Presente
                    ' Único caminho para Suspeito.
                    Return atual.Com(PresenceState.Suspeito,
                                     ObservabilidadeAoNaoVer(atual, obs),
                                     obs.UniverseFingerprint, obs.GenerationKey)

                Case PresenceState.NaoVerificado
                    ' Nunca foi visto e continua não sendo. Não há do que
                    ' suspeitar: suspeita pressupõe presença anterior.
                    Return atual

                Case PresenceState.Suspeito
                    ' Continua suspeito. Gerações que passam NÃO promovem
                    ' suspeita a ausência por contagem nem por tempo.
                    Return New AssociationState(PresenceState.Suspeito,
                                                ObservabilidadeAoNaoVer(atual, obs),
                                                atual.ConcludedInUniverse,
                                                Math.Max(atual.GenerationKey, obs.GenerationKey),
                                                atual.Version + 1)

                Case Else ' AusenteDaPasta
                    ' Já era a conclusão. Observação parcial não a rebaixa
                    ' nem a reafirma.
                    Return atual
            End Select
        End Function

        ''' <summary>
        ''' Não ter visto, sozinho, deixa a observabilidade DESCONHECIDA.
        '''
        ''' Só vira <c>NaoObservavelNoUniverso</c> com cutoff explícito E no
        ''' mesmo universo — evidência de outro universo não serve, porque
        ''' universos diferentes não se comparam (I7).
        ''' </summary>
        Private Shared Function ObservabilidadeAoNaoVer(atual As AssociationState,
                                                        obs As PresenceObservation) As Observability
            If obs.HasExplicitCutoff AndAlso obs.Coverage <> FolderCoverage.Desconhecida Then
                Return Observability.NaoObservavelNoUniverso
            End If
            ' Mudança de universo invalida conclusão anterior de
            ' observabilidade.
            If atual.Observability = Observability.NaoObservavelNoUniverso AndAlso
               Not String.Equals(atual.ConcludedInUniverse, obs.UniverseFingerprint,
                                 StringComparison.Ordinal) Then
                Return Observability.Desconhecida
            End If
            Return Observability.Desconhecida
        End Function

        ''' <summary>
        ''' O resultado de uma verificação individual de um suspeito.
        '''
        ''' FENCING: só aplica se a associação ainda estiver <c>Suspeito</c>
        ''' E na mesma versão em que a verificação foi pedida. Uma
        ''' verificação antiga chegando depois de uma geração nova ter
        ''' marcado o item presente seria sobrescrever o novo com o velho —
        ''' e esse é o pior caso, porque some com o que acabou de ser visto.
        ''' </summary>
        Public Shared Function AplicarVerificacao(atual As AssociationState,
                                                  baseadoNaVersao As Long,
                                                  baseadoNaGeracao As Long,
                                                  resultado As ProbeResult,
                                                  cobertura As FolderCoverage,
                                                  universo As String) As AssociationState
            If atual Is Nothing Then atual = AssociationState.Nova()

            ' CAS: estado mudou desde que a verificação foi pedida.
            If atual.Version <> baseadoNaVersao Then Return atual
            ' Só suspeito é verificável.
            If atual.Presence <> PresenceState.Suspeito Then Return atual
            ' Verificação de outra geração não vale.
            If baseadoNaGeracao <> atual.GenerationKey Then Return atual
            ' Cobertura de outro universo não autoriza nada.
            If Not String.Equals(atual.ConcludedInUniverse, universo, StringComparison.Ordinal) Then
                Return atual
            End If

            Select Case resultado
                Case ProbeResult.EncontradoNaMesmaPasta
                    Return atual.Com(PresenceState.Presente, Observability.Observavel,
                                     universo, atual.GenerationKey)

                Case ProbeResult.EncontradoEmOutraPasta
                    ' Evidência POSITIVA de que a encarnação está em outro
                    ' lugar. É o único jeito de confirmar ausência sem
                    ' depender da cobertura.
                    Return atual.Com(PresenceState.AusenteDaPasta, Observability.Observavel,
                                     universo, atual.GenerationKey)

                Case ProbeResult.NaoEncontrado
                    ' AQUI mora o S7. Sem cobertura COMPLETA, "não encontrei"
                    ' não prova nada — a §19.2 mediu pasta cheia com Count=0.
                    If cobertura = FolderCoverage.Completa Then
                        Return atual.Com(PresenceState.AusenteDaPasta,
                                         Observability.Observavel, universo, atual.GenerationKey)
                    End If
                    Return atual

                Case Else ' Inconclusivo
                    Return atual
            End Select
        End Function

    End Class

End Namespace

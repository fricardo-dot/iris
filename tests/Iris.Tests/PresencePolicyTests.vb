Imports Iris.Sync
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' As transições de presença e observabilidade — adversários 27 a 46 da
''' lista de revisão do Marco 2.1.
'''
''' A regra que estes testes existem para proteger é uma só, e ela é
''' contra-intuitiva: <b>num Outlook com janela de cache curta, muitas
''' associações vão ficar SUSPEITAS para sempre</b>. Isso é o resultado
''' certo. Um algoritmo que "resolvesse" a suspeita promovendo-a a ausência
''' por tempo ou por contagem apagaria correspondência do usuário — e a
''' §19.2 mediu dezenas de pastas cheias reportando ZERO itens, então a
''' oportunidade de errar assim é abundante.
''' </summary>
<TestClass>
Public Class PresencePolicyTests

    Private Const U As String = "store|pasta|filtro|cutoff|1|amb"
    Private Const OUTRO_U As String = "store|pasta|filtro|OUTRO|1|amb"

    Private Shared Function Obs(seen As Boolean, ger As Long,
                                Optional cob As FolderCoverage = FolderCoverage.Completa,
                                Optional universo As String = U,
                                Optional cutoff As Boolean = False) As PresenceObservation
        Return New PresenceObservation(seen, ger, universo, cob, cutoff)
    End Function

    Private Shared Function Presente(Optional ger As Long = 1) As AssociationState
        Return PresencePolicy.AplicarGeracao(AssociationState.Nova(), Obs(True, ger))
    End Function

    Private Shared Function Suspeito(Optional gerVista As Long = 1,
                                     Optional gerPerdida As Long = 2) As AssociationState
        Return PresencePolicy.AplicarGeracao(Presente(gerVista), Obs(False, gerPerdida))
    End Function

    ' ==================================================================
    ' Controle positivo (adversários 1–3)
    ' ==================================================================

    <TestMethod>
    Public Sub Visto_vira_presente_e_observavel()
        Dim s = Presente()
        Assert.AreEqual(PresenceState.Presente, s.Presence)
        Assert.AreEqual(Observability.Observavel, s.Observability)
    End Sub

    <TestMethod>
    Public Sub Presente_e_nao_visto_vira_suspeito()
        Dim s = Suspeito()
        Assert.AreEqual(PresenceState.Suspeito, s.Presence)
    End Sub

    ''' <summary>
    ''' O invariante central: NADA chega a ausência por publicação de
    ''' geração. Publicar autoriza presença e suspeita, e só.
    ''' </summary>
    <TestMethod>
    Public Sub Nenhuma_geracao_produz_ausencia_direta()
        For Each partida In {AssociationState.Nova(), Presente(), Suspeito()}
            For Each cob In {FolderCoverage.Completa, FolderCoverage.Parcial, FolderCoverage.Desconhecida}
                Dim s = PresencePolicy.AplicarGeracao(partida, Obs(False, 99, cob))
                Assert.AreNotEqual(PresenceState.AusenteDaPasta, s.Presence,
                    $"geracao publicada nunca confirma ausencia (cobertura {cob})")
            Next
        Next
    End Sub

    ' ==================================================================
    ' Matriz de presença (adversários 30–40)
    ' ==================================================================

    <TestMethod>
    Public Sub Nao_verificado_e_nao_visto_continua_nao_verificado()
        ' Suspeita pressupoe presenca anterior. Nunca visto nao da do que
        ' suspeitar.
        Dim s = PresencePolicy.AplicarGeracao(AssociationState.Nova(), Obs(False, 1))
        Assert.AreEqual(PresenceState.NaoVerificado, s.Presence)
    End Sub

    <TestMethod>
    Public Sub Suspeito_visto_de_novo_volta_a_presente()
        Dim s = PresencePolicy.AplicarGeracao(Suspeito(), Obs(True, 3))
        Assert.AreEqual(PresenceState.Presente, s.Presence)
    End Sub

    <TestMethod>
    Public Sub Ausente_visto_de_novo_volta_a_presente()
        Dim ausente = PresencePolicy.AplicarVerificacao(
            Suspeito(), Suspeito().Version, 2, ProbeResult.NaoEncontrado,
            FolderCoverage.Completa, U)
        Assert.AreEqual(PresenceState.AusenteDaPasta, ausente.Presence)

        Dim voltou = PresencePolicy.AplicarGeracao(ausente, Obs(True, 5))
        Assert.AreEqual(PresenceState.Presente, voltou.Presence, "o item voltou; presenca vence")
    End Sub

    ''' <summary>
    ''' Adversário 39, e o mais perigoso da matriz: passar muitas gerações
    ''' sem ver um suspeito NÃO o promove. Nem por contagem, nem por tempo.
    ''' </summary>
    <TestMethod>
    Public Sub Muitas_geracoes_sem_ver_NAO_promovem_suspeito_a_ausente()
        Dim s = Suspeito()
        For ger = 3L To 200L
            s = PresencePolicy.AplicarGeracao(s, Obs(False, ger))
            Assert.AreEqual(PresenceState.Suspeito, s.Presence,
                $"na geracao {ger} o suspeito virou {s.Presence}")
        Next
    End Sub

    <TestMethod>
    Public Sub Ausente_nao_visto_de_novo_permanece_ausente()
        Dim ausente = PresencePolicy.AplicarVerificacao(
            Suspeito(), Suspeito().Version, 2, ProbeResult.NaoEncontrado,
            FolderCoverage.Completa, U)
        Dim depois = PresencePolicy.AplicarGeracao(ausente, Obs(False, 9, FolderCoverage.Parcial))
        Assert.AreEqual(PresenceState.AusenteDaPasta, depois.Presence,
            "observacao parcial nao rebaixa nem reafirma a conclusao")
    End Sub

    ' ==================================================================
    ' S7 — o controle negativo que importa (adversários 21, 36, 37)
    ' ==================================================================

    ''' <summary>
    ''' O caso da §19.2: pasta cheia reportando zero itens.
    '''
    ''' Com cobertura PARCIAL ou DESCONHECIDA, "não encontrei" não prova
    ''' nada. Se este teste falhar, o Iris apaga anos de correspondência.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_cobertura_completa_NaoEncontrado_nao_confirma_ausencia()
        For Each cob In {FolderCoverage.Parcial, FolderCoverage.Desconhecida}
            Dim s = Suspeito()
            Dim r = PresencePolicy.AplicarVerificacao(s, s.Version, 2,
                        ProbeResult.NaoEncontrado, cob, U)
            Assert.AreEqual(PresenceState.Suspeito, r.Presence,
                $"cobertura {cob} nao autoriza ausencia — a §19.2 mediu pasta cheia com Count=0")
        Next
    End Sub

    <TestMethod>
    Public Sub Com_cobertura_completa_NaoEncontrado_confirma_ausencia()
        Dim s = Suspeito()
        Dim r = PresencePolicy.AplicarVerificacao(s, s.Version, 2,
                    ProbeResult.NaoEncontrado, FolderCoverage.Completa, U)
        Assert.AreEqual(PresenceState.AusenteDaPasta, r.Presence)
    End Sub

    ''' <summary>
    ''' Achar a encarnação em OUTRA pasta é evidência POSITIVA, e é o único
    ''' caminho para ausência que não depende de cobertura.
    ''' </summary>
    <TestMethod>
    Public Sub Encontrado_em_outra_pasta_confirma_ausencia_sem_depender_de_cobertura()
        Dim s = Suspeito()
        Dim r = PresencePolicy.AplicarVerificacao(s, s.Version, 2,
                    ProbeResult.EncontradoEmOutraPasta, FolderCoverage.Desconhecida, U)
        Assert.AreEqual(PresenceState.AusenteDaPasta, r.Presence)
    End Sub

    <TestMethod>
    Public Sub Inconclusivo_deixa_suspeito()
        Dim s = Suspeito()
        Dim r = PresencePolicy.AplicarVerificacao(s, s.Version, 2,
                    ProbeResult.Inconclusivo, FolderCoverage.Completa, U)
        Assert.AreEqual(PresenceState.Suspeito, r.Presence)
    End Sub

    ' ==================================================================
    ' Fencing (adversários 27, 28, 29)
    ' ==================================================================

    ''' <summary>
    ''' Adversário 27 — "o mais importante que ainda não estava explícito".
    '''
    ''' Uma verificação antiga chegando DEPOIS de uma geração nova ter
    ''' marcado o item presente sobrescreveria o novo com o velho. Some com
    ''' o que acabou de ser visto.
    ''' </summary>
    <TestMethod>
    Public Sub Verificacao_antiga_nao_sobrescreve_presenca_nova()
        Dim s = Suspeito()
        Dim versaoPedida = s.Version
        Dim geracaoPedida = s.GenerationKey

        ' Enquanto a verificacao corria, uma geracao nova viu o item.
        Dim agoraPresente = PresencePolicy.AplicarGeracao(s, Obs(True, 10))
        Assert.AreEqual(PresenceState.Presente, agoraPresente.Presence)

        ' A verificacao velha chega agora, dizendo "nao encontrei".
        Dim r = PresencePolicy.AplicarVerificacao(agoraPresente, versaoPedida, geracaoPedida,
                    ProbeResult.NaoEncontrado, FolderCoverage.Completa, U)

        Assert.AreEqual(PresenceState.Presente, r.Presence,
            "a verificacao velha nao pode apagar o que a geracao nova acabou de ver")
        Assert.AreEqual(agoraPresente.Version, r.Version, "nem sequer aplica transicao")
    End Sub

    <TestMethod>
    Public Sub Verificacao_com_cobertura_de_outro_universo_nao_autoriza()
        Dim s = Suspeito()
        Dim r = PresencePolicy.AplicarVerificacao(s, s.Version, 2,
                    ProbeResult.NaoEncontrado, FolderCoverage.Completa, OUTRO_U)
        Assert.AreEqual(PresenceState.Suspeito, r.Presence,
            "cobertura de outro universo nao vale — I7")
    End Sub

    <TestMethod>
    Public Sub Verificacao_reentregue_nao_aplica_duas_vezes()
        Dim s = Suspeito()
        Dim v = s.Version
        Dim primeira = PresencePolicy.AplicarVerificacao(s, v, 2,
                            ProbeResult.EncontradoNaMesmaPasta, FolderCoverage.Completa, U)
        Assert.AreEqual(PresenceState.Presente, primeira.Presence)

        ' A MESMA verificacao chega de novo, com a versao velha.
        Dim segunda = PresencePolicy.AplicarVerificacao(primeira, v, 2,
                            ProbeResult.EncontradoNaMesmaPasta, FolderCoverage.Completa, U)
        Assert.AreEqual(primeira.Version, segunda.Version, "nao aplica duas vezes")
    End Sub

    <TestMethod>
    Public Sub Geracao_mais_velha_nao_retrocede_estado()
        Dim s = PresencePolicy.AplicarGeracao(AssociationState.Nova(), Obs(True, 10))
        Dim velha = PresencePolicy.AplicarGeracao(s, Obs(False, 3))
        Assert.AreEqual(PresenceState.Presente, velha.Presence,
            "geracao 3 chegando depois da 10 nao vale")
    End Sub

    ' ==================================================================
    ' Observabilidade (adversários 41–46)
    ' ==================================================================

    <TestMethod>
    Public Sub Ausencia_numa_varredura_sozinha_deixa_observabilidade_desconhecida()
        Dim s = PresencePolicy.AplicarGeracao(Presente(), Obs(False, 2, FolderCoverage.Completa))
        Assert.AreEqual(Observability.Desconhecida, s.Observability,
            "nao ter visto nao prova que esta fora do universo")
    End Sub

    <TestMethod>
    Public Sub Com_cutoff_explicito_vira_nao_observavel()
        Dim s = PresencePolicy.AplicarGeracao(Presente(),
                    Obs(False, 2, FolderCoverage.Completa, U, cutoff:=True))
        Assert.AreEqual(Observability.NaoObservavelNoUniverso, s.Observability)
    End Sub

    <TestMethod>
    Public Sub Voltar_a_ser_observado_restaura_observavel()
        Dim fora = PresencePolicy.AplicarGeracao(Presente(),
                       Obs(False, 2, FolderCoverage.Completa, U, cutoff:=True))
        Dim volta = PresencePolicy.AplicarGeracao(fora, Obs(True, 3))
        Assert.AreEqual(Observability.Observavel, volta.Observability)
    End Sub

    ''' <summary>
    ''' Adversário 46: as duas dimensões são independentes. Mexer numa não
    ''' pode mexer na outra — se pudessem, seriam a mesma coisa e o
    ''' bloqueador do `fora_da_janela` voltaria.
    ''' </summary>
    <TestMethod>
    Public Sub Presenca_e_observabilidade_sao_independentes()
        ' Presente e observavel.
        Dim a = Presente()
        Assert.AreEqual(PresenceState.Presente, a.Presence)
        Assert.AreEqual(Observability.Observavel, a.Observability)

        ' Suspeito com observabilidade desconhecida.
        Dim b = PresencePolicy.AplicarGeracao(a, Obs(False, 2))
        Assert.AreEqual(PresenceState.Suspeito, b.Presence)
        Assert.AreEqual(Observability.Desconhecida, b.Observability)

        ' Suspeito com observabilidade "fora do universo": a mesma presenca,
        ' outra observabilidade.
        Dim c = PresencePolicy.AplicarGeracao(a, Obs(False, 2, FolderCoverage.Completa, U, cutoff:=True))
        Assert.AreEqual(PresenceState.Suspeito, c.Presence)
        Assert.AreEqual(Observability.NaoObservavelNoUniverso, c.Observability)
    End Sub

End Class

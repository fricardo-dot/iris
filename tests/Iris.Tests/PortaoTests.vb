Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Assist
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O portão da §29 — e as propriedades que ele tem de ter.</b>
'''
''' ------------------------------------------------------------------
''' <b>EXAUSTIVIDADE, EM DOIS NÍVEIS</b>
'''
''' Todo <c>LabelReadingKind</c> nega quando não está listado na autorização.
''' E os que descrevem o <b>fracasso de ler</b> — <c>Denied</c>,
''' <c>Unreadable</c>, <c>Unknown</c> e companhia — negam <b>mesmo listados</b>.
'''
''' A primeira versão deste arquivo tinha um teste exigindo o contrário:
''' <c>TODO_desfecho_LISTADO_passa</c> fixava que até <c>Unknown</c> e
''' <c>Denied</c> permitissem quando listados. Ou seja, a suíte estava verde
''' <b>garantindo</b> que desserialização incompleta virasse configuração
''' legítima.
'''
''' ------------------------------------------------------------------
''' <b>MONOTONICIDADE DE SEGURANÇA</b>
'''
''' Acrescentar incerteza, proteção ou mensagem não autorizada <b>nunca</b>
''' transforma um "não" em "sim".
'''
''' ------------------------------------------------------------------
''' <b>E O CONTROLE POSITIVO</b>
'''
''' Sem ele, um portão que negasse tudo sempre — inclusive por defeito —
''' passaria em cada teste deste arquivo. A produção nega tudo <b>por
''' decisão</b>, e a diferença entre decisão e defeito é o controle positivo
''' passar.
''' </summary>
<TestClass>
Public Class PortaoTests

    Private Const RotuloOk As String = "baea3331-0000-4000-8000-000000000001"
    Private Const RotuloEstranho As String = "0d35eb3e-0000-4000-8000-000000000002"
    Private Const Endereco As String = "https://exemplo.invalido/v1"

    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 25, 12, 0, 0, TimeSpan.Zero)

    ' ---- fixtures ------------------------------------------------------

    Private Shared Function Pasta() As FolderKey
        Return New FolderKey("store-1", "pasta-1")
    End Function

    Private Shared Function OutraPasta() As FolderKey
        Return New FolderKey("store-1", "pasta-2")
    End Function

    Private Shared Function Destino(Optional endpoint As String = Endereco,
                                    Optional modelo As String = "modelo-de-teste") _
                                    As AssistDestination
        Return New AssistDestination("provedor-de-teste", endpoint, modelo)
    End Function

    ''' <summary>Uma autorização completa, coerente, vigente e generosa.</summary>
    Private Shared Function Autorizacao(
            Optional leituras As IEnumerable(Of LabelReadingKind) = Nothing,
            Optional rotulos As IEnumerable(Of String) = Nothing,
            Optional bits As IEnumerable(Of Integer) = Nothing,
            Optional ignorarHistorico As Boolean = False) As ActivationRecord
        Return New ActivationRecord(
            id:="ativacao-de-teste", versao:=1,
            autoridade:="teste — FASE3 §28.3",
            quando:=Agora.AddDays(-1),
            provedor:="provedor-de-teste",
            endpoint:=Endereco,
            modelo:="modelo-de-teste",
            regiao:="local",
            retencaoAceita:="sem retenção",
            operacoes:={AssistOperation.Resumir},
            pastas:={Pasta()},
            rotulos:=If(rotulos, {RotuloOk}),
            leituras:=If(leituras, {LabelReadingKind.Absent, LabelReadingKind.Present}),
            contentBits:=If(bits, {0}),
            ate:=Agora.AddDays(30),
            ignorarHistorico:=ignorarHistorico)
    End Function

    Private Shared Function Chave(Optional entryId As String = "E-1") As ItemKey
        Return New ItemKey(entryId, "store-1")
    End Function

    Private Shared Function Evidencia(Optional entryId As String = "E-1",
                                      Optional changeKey As String = "AABB") _
                                      As LabelVersionEvidence
        Return New LabelVersionEvidence(entryId, Agora, changeKey)
    End Function

    ''' <summary>
    ''' Uma leitura <b>coerente</b> do desfecho pedido: <c>Present</c> ganha um
    ''' registro ativo, <c>HistoricalOnly</c> ganha um desligado, e os outros
    ''' ficam sem registro. É assim que o parser as produz.
    ''' </summary>
    Private Shared Function Leitura(kind As LabelReadingKind,
                                    Optional registros As IEnumerable(Of LabelRecord) = Nothing,
                                    Optional prova As LabelVersionEvidence = Nothing,
                                    Optional entryId As String = "E-1") As LabelReading
        ' 'versao' eclipsaria Versao(), 'pasta' eclipsaria Pasta(): VB e
        ' case-insensitive, e o CLAUDE.md ja lista sete ocorrencias disso.
        Dim rs = registros
        If rs Is Nothing Then
            Select Case kind
                Case LabelReadingKind.Present : rs = {Registro(RotuloOk, True, 0)}
                Case LabelReadingKind.HistoricalOnly : rs = {Registro(RotuloOk, False, 0)}
                Case Else : rs = Array.Empty(Of LabelRecord)()
            End Select
        End If

        Return New LabelReading(Chave(entryId), kind, LabelReadStage.Parse,
                                version:=If(prova, Evidencia(entryId)),
                                registros:=rs.ToList())
    End Function

    Private Shared Function Registro(id As String, ativo As Boolean, bits As Integer?,
                                     Optional ilegivel As Boolean = False) As LabelRecord
        Return New LabelRecord(id, ativo, bits, ilegivel, {"Enabled", "ContentBits"})
    End Function

    Private Shared Function Voo(Optional operacao As AssistOperation = AssistOperation.Resumir,
                                Optional onde As FolderKey = Nothing,
                                Optional aonde As AssistDestination = Nothing) _
                                As PreflightRequest
        ' 'destino' eclipsaria Destino(). Decima ocorrencia disso no projeto.
        Return New PreflightRequest(operacao, If(onde, Pasta()), If(aonde, Destino()))
    End Function

    Private Shared Function Pedido(ParamArray m As MessageClassification()) As DisclosureRequest
        Return New DisclosureRequest(Voo(), m)
    End Function

    Private Shared Function Mensagem(l As LabelReading,
                                     Optional onde As FolderKey = Nothing,
                                     Optional anexo As Boolean = False) As MessageClassification
        Return New MessageClassification(l.Item, If(onde, Pasta()), l, anexo)
    End Function

    Private Shared Function Decidir(a As ActivationRecord,
                                    p As DisclosureRequest) As DisclosureDecision
        Return New DisclosurePolicy(a).Decidir(p, Agora)
    End Function

    ' ==================================================================
    ' O estado da produção

    <TestMethod>
    Public Sub A_producao_nega_TUDO_por_falta_de_autorizacao()
        Dim d = DisclosurePolicy.DaProducao().Decidir(
            Pedido(Mensagem(Leitura(LabelReadingKind.Absent))), Agora)

        Assert.IsFalse(d.Permitido)
        Assert.AreEqual(DisclosureReason.SemAtivacao, d.Motivo)
        StringAssert.Contains(d.Explicacao, "não está habilitada",
            "o motivo tem de aparecer em portugues, nao so como enum")
    End Sub

    ''' <summary>
    ''' <b>Sem autorização, a classificação nem é INVOCADA.</b>
    '''
    ''' Este é o teste que a primeira versão não tinha. Ela provava que o
    ''' motivo devolvido era o da ativação — o que é precedência de motivo, e
    ''' não ausência de leitura. O pedido já chegava com as classificações
    ''' dentro, ou seja, alguém tinha ido ao COM antes.
    '''
    ''' Aqui a classificação é um <c>Func</c>, e o espião explode se for
    ''' chamado. Só assim "não se lê item sem autorização" vira propriedade
    ''' testada em vez de comentário.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_autorizacao_a_classificacao_nem_e_INVOCADA()
        Dim chamou = False
        Dim portao As New DisclosureGate(DisclosurePolicy.DaProducao(), Function() Agora)

        Dim d = portao.Avaliar(Voo(),
            Function()
                chamou = True
                Throw New InvalidOperationException(
                    "foi ao COM classificar sem autorizacao para tocar em item nenhum")
            End Function)

        Assert.IsFalse(chamou, "ninguem pode classificar antes de o preflight passar")
        Assert.AreEqual(DisclosureReason.SemAtivacao, d.Motivo)
    End Sub

    ''' <summary>O contraponto: com autorização, a classificação É invocada.</summary>
    <TestMethod>
    Public Sub Com_autorizacao_a_classificacao_E_invocada()
        Dim chamou = False
        Dim portao As New DisclosureGate(New DisclosurePolicy(Autorizacao()), Function() Agora)

        Dim d = portao.Avaliar(Voo(),
            Function()
                chamou = True
                Return New List(Of MessageClassification) From {
                    Mensagem(Leitura(LabelReadingKind.Absent))}
            End Function)

        Assert.IsTrue(chamou, "o preflight passou; classificar e o passo seguinte")
        Assert.IsTrue(d.Permitido, $"negou por {d.Motivo}")
    End Sub

    ''' <summary>
    ''' <b>A autorização que vence DURANTE a classificação nega.</b>
    '''
    ''' Classificar custa ~17 ms por item, e uma thread grande com o Outlook
    ''' ocupado leva bem mais — dá tempo de uma autorização expirar no meio.
    '''
    ''' A primeira versão passava o <b>mesmo instante</b> para as duas etapas.
    ''' A segunda conferência existia e conferia contra um relógio congelado,
    ''' então esta situação passava. Conferir de novo contra o mesmo tempo não
    ''' é conferir de novo.
    ''' </summary>
    <TestMethod>
    Public Sub Autorizacao_que_vence_DURANTE_a_classificacao_NEGA()
        Dim vence = Agora.AddMinutes(1)
        Dim a As New ActivationRecord("id", 1, "quem", Agora.AddDays(-1), "provedor-de-teste",
                                      Endereco, "modelo-de-teste", "local", "sem retenção",
                                      {AssistOperation.Resumir}, {Pasta()}, {RotuloOk},
                                      {LabelReadingKind.Absent}, {0}, ate:=vence)

        Dim leituras = 0
        Dim portao As New DisclosureGate(New DisclosurePolicy(a),
                                         Function()
                                             ' Primeira leitura: vigente.
                                             ' Segunda: ja passou do prazo.
                                             leituras += 1
                                             Return If(leituras = 1, Agora, vence.AddMinutes(1))
                                         End Function)

        Dim d = portao.Avaliar(Voo(),
            Function() CType({Mensagem(Leitura(LabelReadingKind.Absent))},
                             IReadOnlyList(Of MessageClassification)))

        Assert.AreEqual(2, leituras, "o relogio TEM de ser lido duas vezes")
        Assert.AreEqual(DisclosureReason.AtivacaoForaDeVigencia, d.Motivo,
            "a autorizacao venceu enquanto o COM classificava")
    End Sub

    ''' <summary>
    ''' <b>Registro sem <c>Enabled</c> declarado ao lado de um ativo NEGA.</b>
    '''
    ''' Um registro com <c>Enabled</c> ausente não está ligado nem desligado:
    ''' é um registro sobre o qual não se sabe. Aceitá-lo ao lado de um bem
    ''' formado deixaria passar "um ativo permitido MAIS um indeterminado"
    ''' como se fosse só o primeiro.
    ''' </summary>
    <TestMethod>
    Public Sub Registro_sem_Enabled_declarado_NEGA()
        Dim indeterminado As New LabelRecord(RotuloEstranho, Nothing, 0, False, {"SetDate"})
        Dim l = Leitura(LabelReadingKind.Present,
                        {Registro(RotuloOk, True, 0), indeterminado})

        Assert.AreEqual(DisclosureReason.ClassificacaoIncoerente,
                        Decidir(Autorizacao(), Pedido(Mensagem(l))).Motivo)
    End Sub

    ''' <summary>
    ''' <b>Histórico nega mesmo quando o GUID dele está na allowlist.</b>
    '''
    ''' A versão anterior só negava quando o GUID também estava de fora, o que
    ''' fazia a allowlist de rótulo <b>ativo</b> virar allowlist de histórico
    ''' sem ninguém ter decidido isso. "Este rótulo pode ser usado hoje" e
    ''' "ter tido este rótulo um dia não importa" são duas políticas
    ''' diferentes.
    ''' </summary>
    <TestMethod>
    Public Sub Historico_de_rotulo_PERMITIDO_tambem_nega_sem_a_declaracao()
        Dim l = Leitura(LabelReadingKind.HistoricalOnly, {Registro(RotuloOk, False, 0)})

        Assert.AreEqual(DisclosureReason.HistoricoNaoDeclarado,
                        Decidir(Autorizacao(leituras:={LabelReadingKind.HistoricalOnly}),
                                Pedido(Mensagem(l))).Motivo)
    End Sub

    ''' <summary>
    ''' Provedor e modelo comparam <b>com</b> distinção de caixa.
    '''
    ''' São identificadores de máquina, e nada prova que todo provedor os
    ''' trate como iguais. Canonizar é trabalho do adaptador, que sabe as
    ''' regras dele — não do portão.
    ''' </summary>
    <TestMethod>
    Public Sub Modelo_com_caixa_diferente_NEGA()
        Dim p As New DisclosureRequest(Voo(aonde:=Destino(modelo:="Modelo-De-Teste")),
                                       {Mensagem(Leitura(LabelReadingKind.Absent))})

        Assert.AreEqual(DisclosureReason.ProvedorNaoAutorizado, Decidir(Autorizacao(), p).Motivo)
    End Sub

    ' ==================================================================
    ' Controle positivo

    <TestMethod>
    Public Sub Autorizada_com_TUDO_no_lugar_PERMITE()
        Dim d = Decidir(Autorizacao(), Pedido(Mensagem(Leitura(LabelReadingKind.Absent))))
        Assert.IsTrue(d.Permitido, $"deveria permitir, negou por {d.Motivo}")
    End Sub

    <TestMethod>
    Public Sub Rotulo_permitido_com_ContentBits_aceito_PERMITE()
        Dim d = Decidir(Autorizacao(), Pedido(Mensagem(Leitura(LabelReadingKind.Present))))
        Assert.IsTrue(d.Permitido, $"negou por {d.Motivo}")
    End Sub

    ' ==================================================================
    ' Exaustividade

    ''' <summary>Todo desfecho não listado nega.</summary>
    <TestMethod>
    Public Sub TODO_desfecho_nao_listado_NEGA()
        Dim a = Autorizacao(leituras:=Array.Empty(Of LabelReadingKind)())

        For Each k As LabelReadingKind In [Enum].GetValues(GetType(LabelReadingKind))
            Assert.IsFalse(Decidir(a, Pedido(Mensagem(Leitura(k)))).Permitido,
                           $"{k} passou sem estar listado")
        Next
    End Sub

    ''' <summary>
    ''' <b>E o não elegível nega MESMO LISTADO.</b>
    '''
    ''' <c>Denied</c> quer dizer que a leitura foi negada; <c>Unreadable</c>,
    ''' que não deu para ler; <c>Unknown</c> é o zero do enum. Nenhum descreve
    ''' o item — descrevem o fracasso de descrevê-lo —, e nenhuma cerimônia
    ''' transforma ausência de informação em prova.
    '''
    ''' Note o desfecho: a autorização que os lista é <b>inválida por
    ''' inteiro</b>, e não "válida ignorando a entrada ruim". Ignorar em
    ''' silêncio deixaria o erro vivo e mudo.
    ''' </summary>
    <TestMethod>
    Public Sub Desfecho_NAO_ELEGIVEL_nega_mesmo_LISTADO()
        For Each k As LabelReadingKind In [Enum].GetValues(GetType(LabelReadingKind))
            If LabelPolicy.Elegivel(k) Then Continue For

            Dim d = Decidir(Autorizacao(leituras:={k}), Pedido(Mensagem(Leitura(k))))

            Assert.IsFalse(d.Permitido, $"{k} passou depois de ser listado")
            Assert.AreEqual(DisclosureReason.AtivacaoInvalida, d.Motivo,
                $"{k}: listar o que nao e prova invalida a autorizacao INTEIRA")
        Next
    End Sub

    ''' <summary>
    ''' E os elegíveis passam quando listados — o contraponto, sem o qual os
    ''' dois testes acima passariam num portão que negasse sempre.
    ''' </summary>
    <TestMethod>
    Public Sub Os_ELEGIVEIS_passam_quando_listados()
        Dim quantos = 0
        For Each k As LabelReadingKind In [Enum].GetValues(GetType(LabelReadingKind))
            If Not LabelPolicy.Elegivel(k) Then Continue For
            quantos += 1

            ' HistoricalOnly so passa com a politica declarando o que fazer
            ' com classificacao antiga — e essa declaracao e o assunto do
            ' Historico_de_rotulo_PERMITIDO_tambem_nega_sem_a_declaracao.
            Dim d = Decidir(Autorizacao(leituras:={k}, ignorarHistorico:=True),
                            Pedido(Mensagem(Leitura(k))))
            Assert.IsTrue(d.Permitido, $"{k} e elegivel, estava listado, e negou por {d.Motivo}")
        Next
        Assert.AreEqual(4, quantos, "Present, Absent, Blank e HistoricalOnly")
    End Sub

    ''' <summary>
    ''' Uma leitura <b>não elegível</b> nega mesmo com a autorização válida —
    ''' a defesa em profundidade.
    '''
    ''' Sem isto, a única barreira seria a validação da ativação, e um portão
    ''' construído por outro caminho aceitaria <c>Denied</c> caladamente.
    ''' </summary>
    <TestMethod>
    Public Sub Leitura_nao_elegivel_nega_com_autorizacao_VALIDA()
        Dim d = Decidir(Autorizacao(), Pedido(Mensagem(Leitura(LabelReadingKind.Denied))))

        Assert.AreEqual(DisclosureReason.LeituraEstruturalmenteInsegura, d.Motivo)
    End Sub

    <TestMethod>
    Public Sub Operacao_NAO_DECLARADA_nega()
        Dim p As New DisclosureRequest(Voo(AssistOperation.Nenhuma),
                                       {Mensagem(Leitura(LabelReadingKind.Absent))})
        Assert.AreEqual(DisclosureReason.OperacaoNaoAutorizada, Decidir(Autorizacao(), p).Motivo)
    End Sub

    ''' <summary>E listar a operação nula invalida a autorização.</summary>
    <TestMethod>
    Public Sub Autorizacao_que_lista_operacao_NULA_e_invalida()
        Dim a As New ActivationRecord("id", 1, "quem", Agora.AddDays(-1), "provedor-de-teste",
                                      Endereco, "modelo-de-teste", "local", "sem retenção",
                                      {AssistOperation.Nenhuma, AssistOperation.Resumir},
                                      {Pasta()}, {RotuloOk}, {LabelReadingKind.Absent}, {0}, ate:=Agora.AddDays(30))

        Assert.AreEqual(DisclosureReason.AtivacaoInvalida,
                        Decidir(a, Pedido(Mensagem(Leitura(LabelReadingKind.Absent)))).Motivo)
    End Sub

    ' ==================================================================
    ' O destino

    ''' <summary>
    ''' <b>O endpoint do PEDIDO tem de ser o autorizado.</b>
    '''
    ''' Sem isto a decisão autoriza "o provedor certo" e o transmissor manda
    ''' para outro lugar — a autorização teria dito sim a um destino que
    ''' ninguém conferiu.
    ''' </summary>
    <TestMethod>
    Public Sub Endpoint_diferente_do_autorizado_NEGA()
        Dim p As New DisclosureRequest(Voo(aonde:=Destino("https://outro.invalido/v1")),
                                       {Mensagem(Leitura(LabelReadingKind.Absent))})

        Assert.AreEqual(DisclosureReason.EndpointNaoAutorizado, Decidir(Autorizacao(), p).Motivo)
    End Sub

    ''' <summary>Nem caminho diferente no mesmo host passa.</summary>
    <TestMethod>
    Public Sub Endpoint_com_caminho_diferente_NEGA()
        Dim p As New DisclosureRequest(Voo(aonde:=Destino("https://exemplo.invalido/v2")),
                                       {Mensagem(Leitura(LabelReadingKind.Absent))})

        Assert.AreEqual(DisclosureReason.EndpointNaoAutorizado, Decidir(Autorizacao(), p).Motivo)
    End Sub

    <TestMethod>
    Public Sub Endpoint_que_nao_e_HTTPS_NEGA()
        Const inseguro = "http://exemplo.invalido/v1"
        Dim a As New ActivationRecord("id", 1, "quem", Agora.AddDays(-1), "provedor-de-teste",
                                      inseguro, "modelo-de-teste", "local", "sem retenção",
                                      {AssistOperation.Resumir}, {Pasta()}, {RotuloOk},
                                      {LabelReadingKind.Absent}, {0}, ate:=Agora.AddDays(30))
        Dim p As New DisclosureRequest(Voo(aonde:=Destino(inseguro)),
                                       {Mensagem(Leitura(LabelReadingKind.Absent))})

        Assert.AreEqual(DisclosureReason.EndpointInseguro, Decidir(a, p).Motivo)
    End Sub

    <TestMethod>
    Public Sub Modelo_diferente_do_autorizado_NEGA()
        Dim p As New DisclosureRequest(Voo(aonde:=Destino(modelo:="outro-modelo")),
                                       {Mensagem(Leitura(LabelReadingKind.Absent))})

        Assert.AreEqual(DisclosureReason.ProvedorNaoAutorizado, Decidir(Autorizacao(), p).Motivo)
    End Sub

    ' ==================================================================
    ' Identidade e coerência

    ''' <summary>
    ''' <b>A classificação tem de ser DA mensagem.</b>
    '''
    ''' DTO público não garante nada, e uma leitura de outro item anexada a
    ''' esta passaria por qualquer conferência que olhasse só o conteúdo da
    ''' leitura.
    ''' </summary>
    <TestMethod>
    Public Sub Classificacao_de_OUTRO_item_NEGA()
        Dim leituraDeOutro = Leitura(LabelReadingKind.Absent, entryId:="E-2")
        Dim m As New MessageClassification(Chave("E-1"), Pasta(), leituraDeOutro, temAnexo:=False)

        Assert.AreEqual(DisclosureReason.IdentidadeNaoBate,
                        Decidir(Autorizacao(), Pedido(m)).Motivo)
    End Sub

    ''' <summary>
    ''' <b>Um <c>Absent</c> forjado carregando rótulo restritivo NEGA.</b>
    '''
    ''' É a combinação mais perigosa que um DTO montado à mão produz: entraria
    ''' pela porta de "sem rótulo" levando um rótulo junto. O parser nunca a
    ''' produz; um portão que olhasse só o <c>Kind</c> a aceitaria.
    ''' </summary>
    <TestMethod>
    Public Sub Absent_COM_registro_ativo_NEGA_por_incoerencia()
        Dim forjada = Leitura(LabelReadingKind.Absent,
                              {Registro(RotuloEstranho, True, 7)})

        Assert.AreEqual(DisclosureReason.ClassificacaoIncoerente,
                        Decidir(Autorizacao(), Pedido(Mensagem(forjada))).Motivo)
    End Sub

    ''' <summary><c>Present</c> sem registro ativo nenhum também é incoerente.</summary>
    <TestMethod>
    Public Sub Present_SEM_registro_ativo_NEGA_por_incoerencia()
        Dim forjada = Leitura(LabelReadingKind.Present, Array.Empty(Of LabelRecord)())

        Assert.AreEqual(DisclosureReason.ClassificacaoIncoerente,
                        Decidir(Autorizacao(), Pedido(Mensagem(forjada))).Motivo)
    End Sub

    ''' <summary>Evidência de versão de OUTRO item não serve.</summary>
    <TestMethod>
    Public Sub Evidencia_de_versao_de_outro_item_NEGA()
        Dim l = Leitura(LabelReadingKind.Absent, Nothing, Evidencia("E-9"))

        Assert.AreEqual(DisclosureReason.SemEvidenciaDeVersao,
                        Decidir(Autorizacao(), Pedido(Mensagem(l))).Motivo)
    End Sub

    ''' <summary>E sem <c>PR_CHANGE_KEY</c> também não — veio em 20 de 20 no 3.0.</summary>
    <TestMethod>
    Public Sub Sem_ChangeKey_NEGA()
        Dim l = Leitura(LabelReadingKind.Absent, Nothing, Evidencia(changeKey:=""))

        Assert.AreEqual(DisclosureReason.SemEvidenciaDeVersao,
                        Decidir(Autorizacao(), Pedido(Mensagem(l))).Motivo)
    End Sub

    ' ==================================================================
    ' As provas, uma a uma

    <TestMethod>
    Public Sub Pasta_fora_da_lista_NEGA()
        Dim p As New DisclosureRequest(Voo(onde:=OutraPasta()),
                                       {Mensagem(Leitura(LabelReadingKind.Absent), OutraPasta())})

        Assert.AreEqual(DisclosureReason.PastaNaoAutorizada, Decidir(Autorizacao(), p).Motivo)
    End Sub

    <TestMethod>
    Public Sub Mensagem_de_outra_pasta_NEGA()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Absent)),
                               Mensagem(Leitura(LabelReadingKind.Absent, entryId:="E-2"),
                                        OutraPasta())))

        Assert.AreEqual(DisclosureReason.MensagemDeOutraPasta, d.Motivo)
    End Sub

    <TestMethod>
    Public Sub Anexo_NEGA_por_inteiro()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Absent), anexo:=True)))

        Assert.AreEqual(DisclosureReason.AnexoForaDeEscopo, d.Motivo)
    End Sub

    <TestMethod>
    Public Sub Rotulo_ativo_fora_da_lista_NEGA()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Present,
                                                {Registro(RotuloEstranho, True, 0)}))))

        Assert.AreEqual(DisclosureReason.RotuloNaoPermitido, d.Motivo)
    End Sub

    ''' <summary>
    ''' <c>ContentBits</c> <b>ausente</b> nega — não prova ausência de proteção.
    '''
    ''' O 3.0 mediu que o campo existe. Não mediu que seja autêntico, que esteja
    ''' atual, nem que cubra toda forma de IRM. "Não vi proteção" e "comprovei
    ''' que não há proteção" são coisas diferentes.
    ''' </summary>
    <TestMethod>
    Public Sub ContentBits_AUSENTE_nega()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Present,
                                                {Registro(RotuloOk, True, Nothing)}))))

        Assert.AreEqual(DisclosureReason.ContentBitsDesconhecido, d.Motivo)
    End Sub

    <TestMethod>
    Public Sub ContentBits_ILEGIVEL_nega()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Present,
                                                {Registro(RotuloOk, True, Nothing, True)}))))

        Assert.AreEqual(DisclosureReason.ContentBitsDesconhecido, d.Motivo)
    End Sub

    <TestMethod>
    Public Sub ContentBits_fora_da_lista_NEGA()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Present,
                                                {Registro(RotuloOk, True, 7)}))))

        Assert.AreEqual(DisclosureReason.ContentBitsNaoAceito, d.Motivo)
    End Sub

    ' ==================================================================
    ' Histórico

    ''' <summary>
    ''' <b>Registro desligado de rótulo não autorizado NEGA</b> — a menos que a
    ''' política declare ignorar histórico.
    '''
    ''' A primeira versão ignorava registro desligado em silêncio, com o
    ''' argumento de que senão uma mensagem que um dia teve rótulo seria
    ''' recusada para sempre. Isso é argumento de usabilidade, não prova de
    ''' segurança: <c>Enabled=False</c> mostra que o registro se diz inativo, e
    ''' não prova que o conteúdo deixou de ser sensível nem que a empresa
    ''' aceita desconsiderar rebaixamento.
    '''
    ''' O caso que a versão anterior deixava passar mudo é este: rótulo ativo
    ''' permitido <b>e</b> rótulo restritivo desligado no mesmo item.
    ''' </summary>
    <TestMethod>
    Public Sub Historico_de_rotulo_NAO_autorizado_NEGA()
        Dim comHistorico = Leitura(LabelReadingKind.Present,
                                   {Registro(RotuloOk, True, 0),
                                    Registro(RotuloEstranho, False, 0)})

        Assert.AreEqual(DisclosureReason.HistoricoNaoDeclarado,
                        Decidir(Autorizacao(), Pedido(Mensagem(comHistorico))).Motivo)
    End Sub

    ''' <summary>E passa quando a política declara, explicitamente, ignorá-lo.</summary>
    <TestMethod>
    Public Sub Historico_passa_quando_a_politica_DECLARA_ignorar()
        Dim comHistorico = Leitura(LabelReadingKind.Present,
                                   {Registro(RotuloOk, True, 0),
                                    Registro(RotuloEstranho, False, 0)})

        Dim d = Decidir(Autorizacao(ignorarHistorico:=True), Pedido(Mensagem(comHistorico)))

        Assert.IsTrue(d.Permitido, $"negou por {d.Motivo}")
    End Sub

    ' ==================================================================
    ' Ativação inválida

    <TestMethod>
    Public Sub Autorizacao_INCOMPLETA_nega()
        Dim a As New ActivationRecord("id", 1, "", Agora.AddDays(-1), "provedor-de-teste",
                                      Endereco, "modelo-de-teste", "local", "sem retenção",
                                      {AssistOperation.Resumir}, {Pasta()}, {RotuloOk},
                                      {LabelReadingKind.Absent}, {0}, ate:=Agora.AddDays(30))

        Assert.AreEqual(DisclosureReason.AtivacaoIncompleta,
                        Decidir(a, Pedido(Mensagem(Leitura(LabelReadingKind.Absent)))).Motivo)
    End Sub

    ''' <summary>Sem região também está incompleta — a cerimônia da §28.3 a exige.</summary>
    <TestMethod>
    Public Sub Autorizacao_SEM_REGIAO_esta_incompleta()
        Dim a As New ActivationRecord("id", 1, "quem", Agora.AddDays(-1), "provedor-de-teste",
                                      Endereco, "modelo-de-teste", "", "sem retenção",
                                      {AssistOperation.Resumir}, {Pasta()}, {RotuloOk},
                                      {LabelReadingKind.Absent}, {0}, ate:=Agora.AddDays(30))

        Assert.AreEqual(DisclosureReason.AtivacaoIncompleta,
                        Decidir(a, Pedido(Mensagem(Leitura(LabelReadingKind.Absent)))).Motivo)
    End Sub

    <TestMethod>
    Public Sub Autorizacao_VENCIDA_nega()
        Dim a As New ActivationRecord("id", 1, "quem", Agora.AddDays(-10), "provedor-de-teste",
                                      Endereco, "modelo-de-teste", "local", "sem retenção",
                                      {AssistOperation.Resumir}, {Pasta()}, {RotuloOk},
                                      {LabelReadingKind.Absent}, {0},
                                      ate:=Agora.AddDays(-1))

        Assert.AreEqual(DisclosureReason.AtivacaoForaDeVigencia,
                        Decidir(a, Pedido(Mensagem(Leitura(LabelReadingKind.Absent)))).Motivo)
    End Sub

    ''' <summary>
    ''' Prazo que termina antes de começar é <b>inválido</b>, não apenas "não
    ''' vigente por acaso". A diferença aparece quando alguém conserta o
    ''' relógio.
    ''' </summary>
    <TestMethod>
    Public Sub Prazo_invertido_INVALIDA_a_autorizacao()
        Dim a As New ActivationRecord("id", 1, "quem", Agora.AddDays(-1), "provedor-de-teste",
                                      Endereco, "modelo-de-teste", "local", "sem retenção",
                                      {AssistOperation.Resumir}, {Pasta()}, {RotuloOk},
                                      {LabelReadingKind.Absent}, {0},
                                      ate:=Agora.AddDays(-5))

        Assert.AreEqual(DisclosureReason.AtivacaoInvalida,
                        Decidir(a, Pedido(Mensagem(Leitura(LabelReadingKind.Absent)))).Motivo)
    End Sub

    ''' <summary>
    ''' Texto que não é GUID na lista de rótulos invalida.
    '''
    ''' Se passasse, casaria com nada — negando por confusão de formato em vez
    ''' de por política, que é um jeito ruim de acertar.
    ''' </summary>
    <TestMethod>
    Public Sub Rotulo_que_nao_e_GUID_INVALIDA_a_autorizacao()
        Dim a As New ActivationRecord("id", 1, "quem", Agora.AddDays(-1), "provedor-de-teste",
                                      Endereco, "modelo-de-teste", "local", "sem retenção",
                                      {AssistOperation.Resumir}, {Pasta()}, {"nao-sou-guid"},
                                      {LabelReadingKind.Absent}, {0}, ate:=Agora.AddDays(30))

        Assert.AreEqual(DisclosureReason.AtivacaoInvalida,
                        Decidir(a, Pedido(Mensagem(Leitura(LabelReadingKind.Absent)))).Motivo)
    End Sub

    ''' <summary>GUID em outra forma — chaves e maiúsculas — casa mesmo assim.</summary>
    <TestMethod>
    Public Sub GUID_em_outra_forma_casa()
        Dim d = Decidir(Autorizacao(rotulos:={"{BAEA3331-0000-4000-8000-000000000001}"}),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Present))))

        Assert.IsTrue(d.Permitido, $"negou por {d.Motivo}")
    End Sub

    <TestMethod>
    Public Sub Pedido_VAZIO_nega()
        Assert.AreEqual(DisclosureReason.PedidoVazio, Decidir(Autorizacao(), Pedido()).Motivo)
    End Sub

    ' ==================================================================
    ' Monotonicidade de segurança

    ''' <summary>
    ''' <b>Acrescentar mensagem nunca transforma "não" em "sim".</b>
    '''
    ''' A forma que o portão precisa não ter: decidir por maioria, por média,
    ''' ou pela primeira mensagem. Uma thread de trinta em que vinte e nove
    ''' passam e uma não é uma thread que <b>não passa</b> — e é o caso comum
    ''' desta caixa, onde rótulo é raro por mensagem e por isso deixa de ser
    ''' raro por thread.
    ''' </summary>
    <TestMethod>
    Public Sub Uma_mensagem_ruim_entre_muitas_boas_NEGA_a_thread()
        Dim boas = Enumerable.Range(1, 29).
                   Select(Function(i) Mensagem(Leitura(LabelReadingKind.Absent,
                                                       entryId:=$"E-{i}"))).ToList()
        Dim ruim = Mensagem(Leitura(LabelReadingKind.Present,
                                    {Registro(RotuloEstranho, True, 0)}, entryId:="E-99"))

        Assert.IsTrue(Decidir(Autorizacao(), Pedido(boas.ToArray())).Permitido,
                      "controle: as 29 sozinhas passam")

        Dim d = Decidir(Autorizacao(), Pedido(boas.Concat({ruim}).ToArray()))

        Assert.IsFalse(d.Permitido, "29 boas e uma ruim NAO e um resumo parcial: e um nao")
        Assert.AreEqual(DisclosureReason.RotuloNaoPermitido, d.Motivo)
    End Sub

    ''' <summary>
    ''' O veredito carrega <b>todas</b> as violações, não só a primeira — senão
    ''' o usuário conserta um problema para descobrir o próximo.
    ''' </summary>
    <TestMethod>
    Public Sub O_veredito_carrega_TODAS_as_violacoes()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Absent, entryId:="E-1"),
                                        anexo:=True),
                               Mensagem(Leitura(LabelReadingKind.Present,
                                                {Registro(RotuloEstranho, True, 0)},
                                                entryId:="E-2"))))

        Assert.IsFalse(d.Permitido)
        Assert.AreEqual(2, d.Total, "duas mensagens com problemas diferentes")
        CollectionAssert.AreEquivalent(
            {DisclosureReason.AnexoForaDeEscopo, DisclosureReason.RotuloNaoPermitido},
            d.Violacoes.Select(Function(x) x.Motivo).ToArray())
    End Sub

    ''' <summary>Encolher a lista de autorizações nunca amplia o que passa.</summary>
    <TestMethod>
    Public Sub Tirar_permissao_nunca_amplia_o_que_passa()
        Dim m = Mensagem(Leitura(LabelReadingKind.Present))

        Assert.IsTrue(Decidir(Autorizacao(leituras:={LabelReadingKind.Present}),
                              Pedido(m)).Permitido, "controle")

        Assert.IsFalse(Decidir(Autorizacao(leituras:={LabelReadingKind.Present},
                                           rotulos:=Array.Empty(Of String)()),
                               Pedido(m)).Permitido, "sem o rotulo listado")
        Assert.IsFalse(Decidir(Autorizacao(leituras:={LabelReadingKind.Present},
                                           bits:=Array.Empty(Of Integer)()),
                               Pedido(m)).Permitido, "sem o ContentBits listado")
        Assert.IsFalse(Decidir(Autorizacao(leituras:=Array.Empty(Of LabelReadingKind)()),
                               Pedido(m)).Permitido, "sem o desfecho listado")
    End Sub

    ''' <summary>Nenhum <c>ContentBits</c> além do listado passa.</summary>
    <TestMethod>
    Public Sub Nenhum_ContentBits_alem_do_listado_passa()
        Dim a = Autorizacao(bits:={0})

        For b = 0 To 8
            Dim d = Decidir(a, Pedido(Mensagem(Leitura(LabelReadingKind.Present,
                                                       {Registro(RotuloOk, True, b)}))))
            If b = 0 Then
                Assert.IsTrue(d.Permitido, "o listado tem de passar")
            Else
                Assert.IsFalse(d.Permitido, $"ContentBits={b} nao esta listado e passou")
            End If
        Next
    End Sub

End Class

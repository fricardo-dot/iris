Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Assist
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O portão da §29 — e as duas propriedades que ele tem de ter.</b>
'''
''' ------------------------------------------------------------------
''' <b>EXAUSTIVIDADE</b>
'''
''' Todo membro de <c>LabelReadingKind</c> nega, a menos que a autorização o
''' tenha listado <b>pelo nome</b>. Vale para membro que ainda não existe: o
''' teste percorre <c>Enum.GetValues</c>, então acrescentar um estado e
''' esquecer do portão quebra a suíte em vez de abrir uma porta.
'''
''' ------------------------------------------------------------------
''' <b>MONOTONICIDADE DE SEGURANÇA</b>
'''
''' Acrescentar incerteza, proteção ou mensagem não autorizada <b>nunca</b>
''' transforma um "não" em "sim". Parece óbvio, e é exatamente o que se
''' perde quando alguém troca uma conjunção por um atalho.
'''
''' ------------------------------------------------------------------
''' <b>E O CONTROLE POSITIVO</b>
'''
''' Sem ele, um portão que negasse tudo sempre — inclusive por defeito —
''' passaria em cada teste deste arquivo. A produção nega tudo <b>por
''' decisão</b>, e a diferença entre decisão e defeito é o
''' <c>Autorizada_com_TUDO_no_lugar_PERMITE</c> passar.
''' </summary>
<TestClass>
Public Class PortaoTests

    Private Const RotuloOk As String = "baea3331-0000-4000-8000-000000000001"
    Private Const RotuloEstranho As String = "0d35eb3e-0000-4000-8000-000000000002"

    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 25, 12, 0, 0, TimeSpan.Zero)

    Private Shared Function Pasta() As FolderKey
        Return New FolderKey("store-1", "pasta-1")
    End Function

    Private Shared Function OutraPasta() As FolderKey
        Return New FolderKey("store-1", "pasta-2")
    End Function

    ''' <summary>Uma autorização completa, vigente e generosa.</summary>
    Private Shared Function Autorizacao(
            Optional leituras As IEnumerable(Of LabelReadingKind) = Nothing,
            Optional rotulos As IEnumerable(Of String) = Nothing,
            Optional bits As IEnumerable(Of Integer) = Nothing) As ActivationRecord
        Return New ActivationRecord(
            autoridade:="teste — FASE3 §28.3",
            quando:=Agora.AddDays(-1),
            provedor:="provedor-de-teste",
            endpoint:="https://exemplo.invalido/v1",
            modelo:="modelo-de-teste",
            regiao:="local",
            retencaoAceita:="sem retenção",
            operacoes:={AssistOperation.Resumir},
            pastas:={Pasta()},
            rotulos:=If(rotulos, {RotuloOk}),
            leituras:=If(leituras, {LabelReadingKind.Absent, LabelReadingKind.Present}),
            contentBits:=If(bits, {0}))
    End Function

    Private Shared Function Versao() As LabelVersionEvidence
        Return New LabelVersionEvidence("E-1", Agora, "AABB")
    End Function

    Private Shared Function Leitura(kind As LabelReadingKind,
                                    Optional registros As IEnumerable(Of LabelRecord) = Nothing,
                                    Optional evidencia As LabelVersionEvidence = Nothing) As LabelReading
        ' 'versao' eclipsaria Versao() — VB e case-insensitive, e este projeto
        ' ja tem sete ocorrencias disso no CLAUDE.md.
        Return New LabelReading(New ItemKey("E-1", "store-1"), kind,
                                LabelReadStage.Parse,
                                version:=If(evidencia, Versao()),
                                registros:=If(registros, Array.Empty(Of LabelRecord)()).ToList())
    End Function

    Private Shared Function Registro(id As String, ativo As Boolean,
                                     bits As Integer?,
                                     Optional ilegivel As Boolean = False) As LabelRecord
        Return New LabelRecord(id, ativo, bits, ilegivel, {"Enabled", "ContentBits"})
    End Function

    Private Shared Function Pedido(ParamArray m As MessageClassification()) As DisclosureRequest
        Return New DisclosureRequest(AssistOperation.Resumir, Pasta(),
                                     "provedor-de-teste", "modelo-de-teste", m)
    End Function

    Private Shared Function Mensagem(l As LabelReading,
                                     Optional onde As FolderKey = Nothing,
                                     Optional anexo As Boolean = False) As MessageClassification
        ' 'pasta' eclipsaria Pasta(). Mesmo motivo do 'evidencia' acima.
        Return New MessageClassification(l.Item, If(onde, Pasta()), l, anexo)
    End Function

    Private Shared Function Decidir(a As ActivationRecord,
                                    p As DisclosureRequest) As DisclosureDecision
        Return New DisclosurePolicy(a).Decidir(p, Agora)
    End Function

    ' ==================================================================
    ' O estado da produção

    ''' <summary>
    ''' Em produção o portão nega <b>tudo</b>, e o motivo é
    ''' <c>SemAtivacao</c> — não um erro genérico.
    ''' </summary>
    <TestMethod>
    Public Sub A_producao_nega_TUDO_por_falta_de_autorizacao()
        Dim d = DisclosurePolicy.DaProducao().Decidir(
            Pedido(Mensagem(Leitura(LabelReadingKind.Absent))), Agora)

        Assert.IsFalse(d.Permitido)
        Assert.AreEqual(DisclosureReason.SemAtivacao, d.Motivo)
        StringAssert.Contains(d.Explicacao, "não está habilitada",
            "o motivo tem de aparecer em portugues para o usuario, nao so como enum")
    End Sub

    ''' <summary>
    ''' E nega <b>antes</b> de olhar rótulo nenhum.
    '''
    ''' Importa por dinheiro de COM: o 3.0 mediu ~17 ms por item para
    ''' classificar. Classificar uma thread inteira para depois descobrir que
    ''' está tudo desligado seria pagar meio segundo de fila da STA para nada.
    ''' Aqui isso é cobrado passando uma leitura que <b>seria</b> recusada por
    ''' outro motivo, e exigindo que o motivo devolvido seja o da ativação.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_autorizacao_nem_chega_a_olhar_o_rotulo()
        Dim ruim = Leitura(LabelReadingKind.Denied, {Registro(RotuloEstranho, True, 7)})

        Dim d = DisclosurePolicy.DaProducao().Decidir(Pedido(Mensagem(ruim)), Agora)

        Assert.AreEqual(DisclosureReason.SemAtivacao, d.Motivo,
            "a ausencia de autorizacao vence tudo, e vence PRIMEIRO")
    End Sub

    ' ==================================================================
    ' Controle positivo

    ''' <summary>
    ''' <b>Com tudo no lugar, PERMITE.</b>
    '''
    ''' Sem este teste, um portão que negasse por defeito passaria em todos os
    ''' outros deste arquivo — e "nega tudo" deixaria de ser uma decisão para
    ''' virar um bug que ninguém veria.
    ''' </summary>
    <TestMethod>
    Public Sub Autorizada_com_TUDO_no_lugar_PERMITE()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Absent))))

        Assert.IsTrue(d.Permitido, $"deveria permitir, negou por {d.Motivo}")
    End Sub

    ''' <summary>Rótulo ativo permitido, com <c>ContentBits</c> aceito, passa.</summary>
    <TestMethod>
    Public Sub Rotulo_permitido_com_ContentBits_aceito_PERMITE()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Present,
                                                {Registro(RotuloOk, True, 0)}))))

        Assert.IsTrue(d.Permitido, $"negou por {d.Motivo}")
    End Sub

    ' ==================================================================
    ' Exaustividade

    ''' <summary>
    ''' <b>Todo</b> <c>LabelReadingKind</c> nega quando não está listado — e
    ''' isso vale para membro que ainda não existe.
    '''
    ''' O teste percorre <c>Enum.GetValues</c> de propósito: acrescentar um
    ''' estado e esquecer de reler o portão quebra a suíte, em vez de abrir uma
    ''' porta que ninguém sabe que abriu.
    ''' </summary>
    <TestMethod>
    Public Sub TODO_desfecho_nao_listado_NEGA()
        ' Autorizacao que nao lista desfecho NENHUM.
        Dim a = Autorizacao(leituras:=Array.Empty(Of LabelReadingKind)())

        For Each k As LabelReadingKind In [Enum].GetValues(GetType(LabelReadingKind))
            Dim d = Decidir(a, Pedido(Mensagem(Leitura(k))))
            Assert.IsFalse(d.Permitido, $"{k} passou sem estar listado")
            Assert.AreEqual(DisclosureReason.LeituraNaoAceita, d.Motivo, $"{k}")
        Next
    End Sub

    ''' <summary>
    ''' E o contraponto: listado, passa. Sem isto o teste de cima passaria num
    ''' portão que ignorasse a lista e negasse sempre.
    ''' </summary>
    <TestMethod>
    Public Sub TODO_desfecho_LISTADO_passa_a_prova_da_leitura()
        For Each k As LabelReadingKind In [Enum].GetValues(GetType(LabelReadingKind))
            Dim d = Decidir(Autorizacao(leituras:={k}), Pedido(Mensagem(Leitura(k))))
            Assert.IsTrue(d.Permitido, $"{k} estava listado e negou por {d.Motivo}")
        Next
    End Sub

    ''' <summary>Operação não declarada — o zero do enum — nunca é autorizada.</summary>
    <TestMethod>
    Public Sub Operacao_NAO_DECLARADA_nega()
        Dim p As New DisclosureRequest(AssistOperation.Nenhuma, Pasta(),
                                       "provedor-de-teste", "modelo-de-teste",
                                       {Mensagem(Leitura(LabelReadingKind.Absent))})

        Assert.AreEqual(DisclosureReason.OperacaoNaoAutorizada, Decidir(Autorizacao(), p).Motivo)
    End Sub

    ' ==================================================================
    ' As provas, uma a uma

    <TestMethod>
    Public Sub Pasta_fora_da_lista_NEGA()
        Dim p As New DisclosureRequest(AssistOperation.Resumir, OutraPasta(),
                                       "provedor-de-teste", "modelo-de-teste",
                                       {Mensagem(Leitura(LabelReadingKind.Absent), OutraPasta())})

        Assert.AreEqual(DisclosureReason.PastaNaoAutorizada, Decidir(Autorizacao(), p).Motivo)
    End Sub

    ''' <summary>
    ''' Mensagem de <b>outra</b> pasta dentro de um pedido sobre pasta
    ''' autorizada. É o caminho por onde conteúdo de pasta proibida entraria
    ''' de carona numa thread.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_de_outra_pasta_NEGA()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Absent)),
                               Mensagem(Leitura(LabelReadingKind.Absent), OutraPasta())))

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
    ''' <c>ContentBits</c> <b>ausente</b> nega — não prova ausência de
    ''' proteção.
    '''
    ''' O 3.0 mediu que o campo existe. Não mediu que ele seja autêntico, que
    ''' esteja atual, nem que cubra toda forma de IRM. "Não vi proteção" e
    ''' "comprovei que não há proteção" são coisas diferentes, e só a segunda
    ''' autorizaria.
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
                                                {Registro(RotuloOk, True, Nothing, ilegivel:=True)}))))

        Assert.AreEqual(DisclosureReason.ContentBitsDesconhecido, d.Motivo)
    End Sub

    <TestMethod>
    Public Sub ContentBits_fora_da_lista_NEGA()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Present,
                                                {Registro(RotuloOk, True, 7)}))))

        Assert.AreEqual(DisclosureReason.ContentBitsNaoAceito, d.Motivo)
    End Sub

    ''' <summary>
    ''' Registro <b>desligado</b> não é conferido: ele não vale, então o rótulo
    ''' dele não precisa estar na lista.
    '''
    ''' É o contraponto que impede o portão de negar por histórico — uma
    ''' mensagem que já teve rótulo e não tem mais seria recusada para sempre,
    ''' o que não é o que a autorização diz.
    ''' </summary>
    <TestMethod>
    Public Sub Registro_DESLIGADO_nao_precisa_estar_na_lista()
        Dim d = Decidir(Autorizacao(leituras:={LabelReadingKind.HistoricalOnly}),
                        Pedido(Mensagem(Leitura(LabelReadingKind.HistoricalOnly,
                                                {Registro(RotuloEstranho, False, Nothing)}))))

        Assert.IsTrue(d.Permitido, $"negou por {d.Motivo}")
    End Sub

    <TestMethod>
    Public Sub Sem_evidencia_de_versao_NEGA()
        Dim d = Decidir(Autorizacao(),
                        Pedido(Mensagem(Leitura(LabelReadingKind.Absent, Nothing,
                                                New LabelVersionEvidence("", Nothing, Nothing)))))

        Assert.AreEqual(DisclosureReason.SemEvidenciaDeVersao, d.Motivo)
    End Sub

    <TestMethod>
    Public Sub Pedido_VAZIO_nega()
        Assert.AreEqual(DisclosureReason.PedidoVazio, Decidir(Autorizacao(), Pedido()).Motivo)
    End Sub

    <TestMethod>
    Public Sub Endpoint_que_nao_e_HTTPS_NEGA()
        Dim a As New ActivationRecord("quem", Agora.AddDays(-1), "provedor-de-teste",
                                      "http://exemplo.invalido/v1", "modelo-de-teste",
                                      "local", "sem retenção",
                                      {AssistOperation.Resumir}, {Pasta()}, {RotuloOk},
                                      {LabelReadingKind.Absent}, {0})

        Assert.AreEqual(DisclosureReason.EndpointInseguro,
                        Decidir(a, Pedido(Mensagem(Leitura(LabelReadingKind.Absent)))).Motivo)
    End Sub

    <TestMethod>
    Public Sub Autorizacao_VENCIDA_nega()
        Dim a As New ActivationRecord("quem", Agora.AddDays(-10), "provedor-de-teste",
                                      "https://exemplo.invalido/v1", "modelo-de-teste",
                                      "local", "sem retenção",
                                      {AssistOperation.Resumir}, {Pasta()}, {RotuloOk},
                                      {LabelReadingKind.Absent}, {0},
                                      ate:=Agora.AddDays(-1))

        Assert.AreEqual(DisclosureReason.AtivacaoForaDeVigencia,
                        Decidir(a, Pedido(Mensagem(Leitura(LabelReadingKind.Absent)))).Motivo)
    End Sub

    <TestMethod>
    Public Sub Modelo_diferente_do_autorizado_NEGA()
        Dim p As New DisclosureRequest(AssistOperation.Resumir, Pasta(),
                                       "provedor-de-teste", "outro-modelo",
                                       {Mensagem(Leitura(LabelReadingKind.Absent))})

        Assert.AreEqual(DisclosureReason.ProvedorNaoAutorizado, Decidir(Autorizacao(), p).Motivo)
    End Sub

    ''' <summary>Autorização pela metade não autoriza nada.</summary>
    <TestMethod>
    Public Sub Autorizacao_INCOMPLETA_nega()
        Dim a As New ActivationRecord("", Agora.AddDays(-1), "p", "https://x.invalido", "m",
                                      "local", "sem retenção",
                                      {AssistOperation.Resumir}, {Pasta()}, {RotuloOk},
                                      {LabelReadingKind.Absent}, {0})

        Assert.AreEqual(DisclosureReason.AtivacaoIncompleta,
                        Decidir(a, Pedido(Mensagem(Leitura(LabelReadingKind.Absent)))).Motivo)
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
                   Select(Function(i) Mensagem(Leitura(LabelReadingKind.Absent))).ToList()
        Dim ruim = Mensagem(Leitura(LabelReadingKind.Present,
                                    {Registro(RotuloEstranho, True, 0)}))

        Assert.IsTrue(Decidir(Autorizacao(), Pedido(boas.ToArray())).Permitido,
                      "controle: as 29 sozinhas passam")

        Dim d = Decidir(Autorizacao(), Pedido(boas.Concat({ruim}).ToArray()))

        Assert.IsFalse(d.Permitido, "29 boas e uma ruim NAO e um resumo parcial: e um nao")
        Assert.AreEqual(DisclosureReason.RotuloNaoPermitido, d.Motivo)
    End Sub

    ''' <summary>
    ''' <b>Encolher a lista de autorizações nunca amplia o que passa.</b>
    '''
    ''' Percorre cada desfecho: com ele listado, o pedido pode passar; tirando
    ''' <b>qualquer</b> dimensão da autorização, o que passava não pode voltar
    ''' a passar.
    ''' </summary>
    <TestMethod>
    Public Sub Tirar_permissao_nunca_amplia_o_que_passa()
        Dim m = Mensagem(Leitura(LabelReadingKind.Present, {Registro(RotuloOk, True, 0)}))

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

    ''' <summary>
    ''' <b>Trocar um registro conhecido por um desconhecido nunca libera.</b>
    '''
    ''' Percorre todo <c>ContentBits</c> plausível: só o listado passa, e a
    ''' ausência do campo não é tratada como zero.
    ''' </summary>
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

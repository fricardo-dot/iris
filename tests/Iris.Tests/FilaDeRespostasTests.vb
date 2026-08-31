Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A FILA DE RESPOSTAS — Fase 3, o coração.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTE ARQUIVO PRENDE</b>
'''
''' A fila responde uma pergunta só: <i>quem falou por último em cada conversa,
''' e há quantos dias?</i> Tudo o mais é consequência. E como ela vai ser lida
''' de manhã para decidir o dia, o perigo não é ela ficar vazia — é ela ficar
''' <b>cheia de mentira</b>, com a confiança de quem mediu.
'''
''' Por isso a maior parte daqui é sobre <i>o que ela se recusa a afirmar</i>:
'''
''' <list type="number">
''' <item>Sem os enviados varridos ela <b>não monta</b> — não devolve lista
''' parcial.</item>
''' <item>Empate no instante vira "não sei", e não desempate.</item>
''' <item>Direção desconhecida não vira linha.</item>
''' <item>O que fica de fora é <b>contado</b>, nunca descartado em
''' silêncio.</item>
''' </list>
'''
''' ------------------------------------------------------------------
''' <b>O CONTROLE NEGATIVO</b>
'''
''' <see cref="Sem_os_enviados_varridos_a_fila_RECUSA"/>. Sem ele, uma fila
''' montada só com a Caixa de Entrada mostraria como pendente <b>tudo o que o
''' dono já respondeu</b> — e passaria em todos os outros testes daqui, porque
''' todos os outros dão os enviados por varridos.
''' </summary>
<TestClass>
Public Class FilaDeRespostasTests

    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 31, 12, 0, 0, TimeSpan.Zero)

    Private Shared Function Eu() As MinhasIdentidades
        Return New MinhasIdentidades({"ricardo@empresa.com"})
    End Function

    ''' <summary>Uma mensagem na conversa <paramref name="conversa"/>.</summary>
    Private Shared Function Msg(conversa As String, deQuem As String,
                               diasAtras As Double,
                               Optional assunto As String = "assunto") As MensagemNaFila
        Return New MensagemNaFila(New ItemKey($"E-{conversa}-{diasAtras}", "store-1"),
                                  conversa, assunto, deQuem, deQuem,
                                  Agora.AddDays(-diasAtras))
    End Function

    Private Shared Function Montar(mensagens As IEnumerable(Of MensagemNaFila),
                                   Optional viuOsEnviados As Boolean = True,
                                   Optional dispensadas As String() = Nothing) As ResultadoDaFila
        Return FilaDeRespostas.Montar(mensagens, Eu(), Agora, viuOsEnviados, dispensadas)
    End Function

    ' ==================================================================
    ' AS DUAS FILAS

    <TestMethod>
    Public Sub A_ultima_e_do_outro_entao_pode_ser_a_minha_vez()
        Dim r = Montar({Msg("c1", "ricardo@empresa.com", 20),
                        Msg("c1", "caroline@outra.com", 5)})

        Assert.AreEqual(1, r.Linhas.Count)
        Assert.AreEqual(EstadoDaConversa.PossivelResposta, r.Linhas(0).Estado)
        Assert.AreEqual(5, r.Linhas(0).Dias)
        Assert.AreEqual(1, r.Minhas().Count)
        Assert.AreEqual(0, r.Deles().Count)
    End Sub

    <TestMethod>
    Public Sub A_ultima_e_minha_entao_a_espera_e_do_outro()
        Dim r = Montar({Msg("c1", "caroline@outra.com", 20),
                        Msg("c1", "ricardo@empresa.com", 3)})

        Assert.AreEqual(EstadoDaConversa.Aguardando, r.Linhas(0).Estado)
        Assert.AreEqual(3, r.Linhas(0).Dias)
        Assert.AreEqual(1, r.Deles().Count)
        Assert.AreEqual(0, r.Minhas().Count)
    End Sub

    ''' <summary>
    ''' <b>É a mesma máquina com o sinal trocado</b>, e este teste é onde isso
    ''' fica escrito: uma consulta, duas filas, e a única diferença é quem
    ''' escreveu por último.
    ''' </summary>
    <TestMethod>
    Public Sub Uma_consulta_produz_as_DUAS_filas()
        Dim r = Montar({Msg("eu-devo", "caroline@outra.com", 9),
                        Msg("me-devem", "ricardo@empresa.com", 4)})

        Assert.AreEqual(1, r.Minhas().Count)
        Assert.AreEqual(1, r.Deles().Count)
        Assert.AreEqual("eu-devo", r.Minhas()(0).Conversa)
        Assert.AreEqual("me-devem", r.Deles()(0).Conversa)
    End Sub

    ' ==================================================================
    ' O QUE ELA SE RECUSA A AFIRMAR

    ''' <summary>
    ''' <b>O controle negativo do arquivo.</b>
    '''
    ''' Uma fila montada sem ver os enviados mostraria como pendente tudo o que
    ''' o dono já respondeu — e com a cara de quem mediu. Recusar é a única
    ''' saída honesta, e "recusou" tem de ser distinguível de "não há nada".
    ''' </summary>
    <TestMethod>
    Public Sub Sem_os_enviados_varridos_a_fila_RECUSA()
        Dim r = Montar({Msg("c1", "caroline@outra.com", 20)}, viuOsEnviados:=False)

        Assert.IsFalse(r.Respondeu,
            "montou a fila sem ter visto os enviados: cada conversa ja " &
            "respondida vira uma cobranca falsa")
        Assert.AreEqual(0, r.Linhas.Count, "nao pode devolver lista parcial")

        ' E "recusou" NAO e "nao ha nada". Uma fila vazia de verdade responde.
        Dim vazia = Montar(Array.Empty(Of MensagemNaFila)())
        Assert.IsTrue(vazia.Respondeu, "fila vazia e uma resposta, e recusa nao e")
        Assert.AreEqual(0, vazia.Linhas.Count)
    End Sub

    ''' <summary>
    ''' <b>Empate é "não sei", e não desempate.</b>
    '''
    ''' Duas mensagens no mesmo instante com direções diferentes acontecem:
    ''' cópia de sistema, relógio de servidor, importação em lote. Escolher
    ''' sempre a minha esconderia pendência; escolher sempre a do outro criaria
    ''' pendência falsa. A terceira saída é não afirmar.
    ''' </summary>
    <TestMethod>
    Public Sub Empate_no_instante_nao_vira_linha()
        Dim r = Montar({Msg("c1", "ricardo@empresa.com", 5),
                        Msg("c1", "caroline@outra.com", 5)})

        Assert.AreEqual(0, r.Linhas.Count,
            "escolheu um lado num empate: inventou a resposta")
        Assert.AreEqual(1, r.Fora.SemDirecao, "e nao contou o que descartou")
    End Sub

    ''' <summary>
    ''' Remetente que não dá para identificar não vira linha. Uma linha que não
    ''' sabe de quem é a vez não diz nada e ocupa a fila.
    ''' </summary>
    <TestMethod>
    Public Sub Remetente_desconhecido_nao_vira_linha()
        Dim r = Montar({Msg("c1", "", 5)})

        Assert.AreEqual(0, r.Linhas.Count)
        Assert.AreEqual(1, r.Fora.SemDirecao)
    End Sub

    ''' <summary>
    ''' Conversa desconhecida não é conversa própria: juntar as vazias faria de
    ''' todas as mensagens ilegíveis uma conversa só, com dez pessoas dentro.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_conversa_e_sem_data_ficam_de_fora_E_SAO_CONTADAS()
        Dim semConversa = New MensagemNaFila(New ItemKey("E-1", "s"), "", "a", "x", "x@y.com",
                                             Agora.AddDays(-2))
        Dim semData = New MensagemNaFila(New ItemKey("E-2", "s"), "c9", "a", "x", "x@y.com",
                                         Nothing)

        Dim r = Montar({semConversa, semData})

        Assert.AreEqual(0, r.Linhas.Count)
        Assert.AreEqual(1, r.Fora.SemConversa)
        Assert.AreEqual(1, r.Fora.SemData)
        Assert.AreEqual(2, r.Fora.Total,
            "descartar em silencio faria a fila parecer completa quando " &
            "metade da caixa nao coube nela")
    End Sub

    ' ==================================================================
    ' O QUE O DONO DISPENSA

    ''' <summary>
    ''' "Não exige resposta" tira a conversa da fila <b>na hora</b>, e sem IA
    ''' nenhuma. É o que resolve a maior parte da dívida falsa antes de existir
    ''' classificador.
    ''' </summary>
    <TestMethod>
    Public Sub Conversa_dispensada_some_da_fila_e_e_contada()
        Dim r = Montar({Msg("c1", "caroline@outra.com", 20),
                        Msg("c2", "outro@outra.com", 10)},
                       dispensadas:={"c1"})

        Assert.AreEqual(1, r.Linhas.Count)
        Assert.AreEqual("c2", r.Linhas(0).Conversa)
        Assert.AreEqual(1, r.Fora.Dispensadas,
            "some da vista, nao do conhecimento")
    End Sub

    ' ==================================================================
    ' A ORDEM E AS FAIXAS

    ''' <summary>
    ''' A mais antiga primeiro — é a ordem que a tela existe para mostrar: o que
    ''' espera há mais tempo é o que se perde de vista.
    ''' </summary>
    <TestMethod>
    Public Sub A_mais_antiga_vem_primeiro()
        Dim r = Montar({Msg("nova", "caroline@outra.com", 1),
                        Msg("velha", "caroline@outra.com", 20),
                        Msg("media", "caroline@outra.com", 8)})

        CollectionAssert.AreEqual({"velha", "media", "nova"},
                                  r.Linhas.Select(Function(l) l.Conversa).ToArray())
    End Sub

    <TestMethod>
    Public Sub As_faixas_saem_dos_dias()
        Dim r = Montar({Msg("c15", "caroline@outra.com", 15),
                        Msg("c9", "caroline@outra.com", 9),
                        Msg("c4", "caroline@outra.com", 4),
                        Msg("c1", "caroline@outra.com", 1)})

        Dim porConversa = r.Linhas.ToDictionary(Function(l) l.Conversa)
        Assert.AreEqual(FaixaDeEspera.Critico, porConversa("c15").Faixa)
        Assert.AreEqual(FaixaDeEspera.Atrasado, porConversa("c9").Faixa)
        Assert.AreEqual(FaixaDeEspera.Atencao, porConversa("c4").Faixa)
        Assert.AreEqual(FaixaDeEspera.Normal, porConversa("c1").Faixa)
    End Sub

    ''' <summary>
    ''' Os limites, um por um. Faixa é corte sobre número, e corte errado por um
    ''' dia numa fila que ordena por dias muda a leitura de manhã.
    ''' </summary>
    <TestMethod>
    Public Sub Os_limites_das_faixas()
        Assert.AreEqual(FaixaDeEspera.Atrasado, Faixa(14), "14 ainda nao e critico")
        Assert.AreEqual(FaixaDeEspera.Critico, Faixa(15), "15 e critico")
        Assert.AreEqual(FaixaDeEspera.Atencao, Faixa(6), "6 ainda nao e atrasado")
        Assert.AreEqual(FaixaDeEspera.Atrasado, Faixa(7), "7 e atrasado")
        Assert.AreEqual(FaixaDeEspera.Normal, Faixa(2), "2 ainda e normal")
        Assert.AreEqual(FaixaDeEspera.Atencao, Faixa(3), "3 e atencao")
    End Sub

    Private Shared Function Faixa(dias As Integer) As FaixaDeEspera
        Return Montar({Msg("c", "caroline@outra.com", dias)}).Linhas(0).Faixa
    End Function

    ''' <summary>
    ''' Mensagem com data no futuro — relógio de servidor adiantado — não vira
    ''' dias negativos. Negativo ordenaria antes de tudo e diria "esperando
    ''' há -3 dias", que não quer dizer nada.
    ''' </summary>
    <TestMethod>
    Public Sub Data_no_futuro_nao_vira_dia_negativo()
        Dim r = Montar({Msg("c1", "caroline@outra.com", -3)})

        Assert.AreEqual(0, r.Linhas(0).Dias)
    End Sub

    ''' <summary>
    ''' A linha aponta para a <b>última</b> mensagem da conversa — é ela que a
    ''' tela abre. Apontar para a primeira levaria o dono ao começo de uma
    ''' conversa cujo fim é o que interessa.
    ''' </summary>
    <TestMethod>
    Public Sub A_linha_aponta_para_a_ULTIMA_mensagem()
        Dim antiga = Msg("c1", "caroline@outra.com", 20, "o comeco")
        Dim recente = Msg("c1", "caroline@outra.com", 2, "o fim")

        Dim r = Montar({antiga, recente})

        Assert.AreEqual("o fim", r.Linhas(0).Assunto)
        Assert.AreEqual(recente.Chave.EntryId, r.Linhas(0).Chave.EntryId)
    End Sub

End Class

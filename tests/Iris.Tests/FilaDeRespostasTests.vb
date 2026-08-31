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
''' <item>Sem os enviados varridos ela <b>não monta</b>.</item>
''' <item>Com tudo inclassificável ela <b>não diz "não há nada"</b> — são
''' estados diferentes, e a tela vai falar diferente.</item>
''' <item>Empate no instante vira "não sei", e não desempate.</item>
''' <item>O que fica de fora é contado <b>com a unidade no nome</b>.</item>
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

    ''' <summary>
    ''' Uma segunda-feira, meio-dia, em Brasília. O fuso importa: a contagem de
    ''' dias é de calendário, e calendário depende de onde se está.
    ''' </summary>
    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 31, 12, 0, 0, TimeSpan.FromHours(-3))

    ''' <summary>
    ''' Fuso fixo, sem horário de verão, para o teste não depender da máquina.
    ''' <c>TimeZoneInfo.Local</c> aqui faria o resultado mudar conforme quem roda.
    ''' </summary>
    Private Shared ReadOnly Fuso As TimeZoneInfo =
        TimeZoneInfo.CreateCustomTimeZone("iris-teste", TimeSpan.FromHours(-3),
                                          "Teste", "Teste")

    Private Shared Function Eu() As MinhasIdentidades
        Return New MinhasIdentidades({"ricardo@empresa.com"})
    End Function

    ''' <summary>Uma mensagem na conversa <paramref name="conversa"/>.</summary>
    Private Shared Function Msg(conversa As String, deQuem As String,
                               diasAtras As Double,
                               Optional assunto As String = "assunto",
                               Optional id As String = Nothing) As MensagemNaFila
        Return New MensagemNaFila(New ItemKey(If(id, $"E-{conversa}-{diasAtras}"), "store-1"),
                                  conversa, assunto, deQuem, deQuem,
                                  Agora.AddDays(-diasAtras))
    End Function

    Private Shared Function Montar(mensagens As IEnumerable(Of MensagemNaFila),
                                   Optional viuOsEnviados As Boolean = True,
                                   Optional dispensadas As String() = Nothing,
                                   Optional eu As MinhasIdentidades = Nothing) As ResultadoDaFila
        Return FilaDeRespostas.Montar(mensagens, If(eu, FilaDeRespostasTests.Eu()),
                                      Agora, Fuso, viuOsEnviados, dispensadas)
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
    '''
    ''' As conversas aqui têm <b>histórico</b>, e cada uma muda de lado ao longo
    ''' dele — a primeira versão deste teste dava uma mensagem para cada, o que
    ''' provava a classificação e não a agregação.
    ''' </summary>
    <TestMethod>
    Public Sub Uma_consulta_produz_as_DUAS_filas()
        Dim r = Montar({
            Msg("eu-devo", "caroline@outra.com", 30),
            Msg("eu-devo", "ricardo@empresa.com", 20),
            Msg("eu-devo", "caroline@outra.com", 9),
            Msg("me-devem", "ricardo@empresa.com", 25),
            Msg("me-devem", "outro@outra.com", 12),
            Msg("me-devem", "ricardo@empresa.com", 4)})

        Assert.AreEqual(2, r.Linhas.Count, "duas conversas, duas linhas")
        Assert.AreEqual(MotivoDaFila.Respondida, r.Motivo)

        Assert.AreEqual(1, r.Minhas().Count)
        Assert.AreEqual(1, r.Deles().Count)
        Assert.AreEqual("eu-devo", r.Minhas()(0).Conversa)
        Assert.AreEqual("me-devem", r.Deles()(0).Conversa)

        ' AS DUAS FILAS SAO DISJUNTAS E COBREM TUDO. Sem isto, uma linha podia
        ' cair nas duas -- ou em nenhuma -- e o teste nao veria.
        Assert.AreEqual(r.Linhas.Count, r.Minhas().Count + r.Deles().Count)

        ' E o lado vem da ULTIMA mensagem, e nao da primeira: as duas conversas
        ' comecam do lado oposto ao que terminam.
        Assert.AreEqual(9, r.Minhas()(0).Dias)
        Assert.AreEqual(4, r.Deles()(0).Dias)
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

        Assert.AreEqual(MotivoDaFila.SemOsEnviados, r.Motivo,
            "montou a fila sem ter visto os enviados: cada conversa ja " &
            "respondida vira uma cobranca falsa")
        Assert.IsFalse(r.Respondeu)
        Assert.AreEqual(0, r.Linhas.Count, "nao pode devolver lista parcial")

        ' E "recusou" NAO e "nao ha nada". Uma fila vazia de verdade responde.
        Dim vazia = Montar(Array.Empty(Of MensagemNaFila)())
        Assert.AreEqual(MotivoDaFila.Respondida, vazia.Motivo,
            "fila vazia e uma resposta, e recusa nao e")
        Assert.IsTrue(vazia.Respondeu)
    End Sub

    ''' <summary>
    ''' <b>Caixa cheia e nada classificável não é "não há nada para hoje".</b>
    '''
    ''' Com as identidades do dono incompletas — ou ausentes — toda conversa cai
    ''' em direção desconhecida, a lista fica vazia e a tela diria que o dia
    ''' está limpo. É a afirmação mais errada que esta fila consegue fazer, e
    ''' contar as descartadas num canto não protegia a afirmação principal.
    '''
    ''' Achado por revisão externa em 31/08/2026.
    ''' </summary>
    <TestMethod>
    Public Sub Nada_classificavel_NAO_e_fila_vazia()
        Dim semIdentidade = New MinhasIdentidades({})
        Dim r = Montar({Msg("c1", "caroline@outra.com", 20),
                        Msg("c2", "outro@outra.com", 9),
                        Msg("c3", "mais@outra.com", 2)},
                       eu:=semIdentidade)

        Assert.AreEqual(0, r.Linhas.Count)
        Assert.AreEqual(MotivoDaFila.NadaClassificavel, r.Motivo,
            "tres conversas viraram lista vazia com cara de dia limpo")
        Assert.IsFalse(r.Respondeu)
        Assert.AreEqual(3, r.Fora.ConversasSemDirecao)
        Assert.AreEqual(3, r.ConversasVistas,
            "sem o denominador, '3 sem direcao' nao diz se sobraram " &
            "trezentas ou nenhuma")
    End Sub

    ''' <summary>
    ''' O contraponto do de cima: <b>algumas</b> inclassificáveis com outras
    ''' classificadas continua sendo uma fila respondida. Sem este, a distinção
    ''' viraria "qualquer descarte estraga tudo".
    ''' </summary>
    <TestMethod>
    Public Sub Algumas_sem_direcao_ainda_e_fila_respondida()
        Dim r = Montar({Msg("boa", "caroline@outra.com", 9),
                        Msg("ruim", "", 3)})

        Assert.AreEqual(1, r.Linhas.Count)
        Assert.AreEqual(MotivoDaFila.Respondida, r.Motivo)
        Assert.AreEqual(1, r.Fora.ConversasSemDirecao)
        Assert.AreEqual(2, r.ConversasVistas)
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
        Assert.AreEqual(1, r.Fora.ConversasSemDirecao, "e nao contou o que descartou")
    End Sub

    ''' <summary>
    ''' <b>Empate de mesma direção escolhe pela chave, e não pela ordem de
    ''' entrada.</b>
    '''
    ''' A direção é segura, mas a mensagem que a tela abre não pode depender da
    ''' ordem em que o leitor devolveu as linhas: o plano da consulta SQLite
    ''' mudaria a linha mostrada sem nada no código mudar.
    ''' </summary>
    <TestMethod>
    Public Sub Empate_de_mesma_direcao_escolhe_a_MESMA_mensagem_sempre()
        Dim a = Msg("c1", "caroline@outra.com", 5, "assunto A", id:="E-aaa")
        Dim b = Msg("c1", "caroline@outra.com", 5, "assunto B", id:="E-bbb")

        Dim numaOrdem = Montar({a, b}).Linhas(0)
        Dim naOutra = Montar({b, a}).Linhas(0)

        Assert.AreEqual("E-aaa", numaOrdem.Chave.EntryId)
        Assert.AreEqual(numaOrdem.Chave.EntryId, naOutra.Chave.EntryId,
            "a linha mostrada mudou com a ordem de entrada")
        Assert.AreEqual(numaOrdem.Assunto, naOutra.Assunto)
    End Sub

    ''' <summary>
    ''' <b>Mesmo instante em fusos diferentes é o mesmo instante.</b>
    '''
    ''' A regra de empate depende disso e não é óbvia: <c>DateTimeOffset</c>
    ''' compara instantes, não a representação. Sem este teste, alguém que
    ''' trocasse a comparação por data local criaria empates falsos — ou
    ''' perderia empates verdadeiros — sem nada mais falhar.
    ''' </summary>
    <TestMethod>
    Public Sub O_mesmo_instante_em_fusos_diferentes_EMPATA()
        Dim instante = New DateTimeOffset(2026, 8, 26, 10, 0, 0, TimeSpan.FromHours(-3))
        Dim oMesmo = instante.ToOffset(TimeSpan.Zero)
        Assert.AreNotEqual(instante.Offset, oMesmo.Offset, "o preparo do teste esta errado")

        Dim r = Montar({
            New MensagemNaFila(New ItemKey("E-1", "s"), "c1", "a", "eu",
                               "ricardo@empresa.com", instante),
            New MensagemNaFila(New ItemKey("E-2", "s"), "c1", "a", "ela",
                               "caroline@outra.com", oMesmo)})

        Assert.AreEqual(0, r.Linhas.Count,
            "o mesmo instante escrito em dois fusos deixou de empatar")
        Assert.AreEqual(1, r.Fora.ConversasSemDirecao)
    End Sub

    <TestMethod>
    Public Sub Remetente_desconhecido_nao_vira_linha()
        Dim r = Montar({Msg("c1", "", 5), Msg("c2", "caroline@outra.com", 5)})

        Assert.AreEqual(1, r.Linhas.Count)
        Assert.AreEqual(1, r.Fora.ConversasSemDirecao)
    End Sub

    ''' <summary>
    ''' Conversa desconhecida não é conversa própria: juntar as vazias faria de
    ''' todas as mensagens ilegíveis uma conversa só, com dez pessoas dentro. E
    ''' <b>chave de espaços é chave vazia</b> — agruparia do mesmo jeito.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_conversa_e_sem_data_ficam_de_fora_E_SAO_CONTADAS()
        Dim semConversa = New MensagemNaFila(New ItemKey("E-1", "s"), "", "a", "x", "x@y.com",
                                             Agora.AddDays(-2))
        Dim soEspacos = New MensagemNaFila(New ItemKey("E-3", "s"), "   ", "a", "x", "x@y.com",
                                           Agora.AddDays(-2))
        Dim semData = New MensagemNaFila(New ItemKey("E-2", "s"), "c9", "a", "x", "x@y.com",
                                         Nothing)

        Dim r = Montar({semConversa, soEspacos, semData})

        Assert.AreEqual(0, r.Linhas.Count)
        Assert.AreEqual(2, r.Fora.MensagensSemConversa,
            "chave de espacos foi tratada como conversa de verdade")
        Assert.AreEqual(1, r.Fora.MensagensSemData)
    End Sub

    ''' <summary>Lista nula não estoura — é o que um leitor com falha devolve.</summary>
    <TestMethod>
    Public Sub Lista_nula_devolve_fila_vazia()
        Dim r = Montar(Nothing)

        Assert.AreEqual(0, r.Linhas.Count)
        Assert.AreEqual(MotivoDaFila.Respondida, r.Motivo)
    End Sub

    ' ==================================================================
    ' O QUE O DONO DISPENSA

    ''' <summary>
    ''' "Não exige resposta" tira a conversa da fila <b>na hora</b>, e sem IA
    ''' nenhuma. É o que resolve a maior parte da dívida falsa antes de existir
    ''' classificador.
    '''
    ''' E a contagem é de <b>conversas</b>, e não de mensagens: uma conversa de
    ''' cinco mensagens dispensada é uma dispensa, não cinco. A tela vai dizer
    ''' "3 conversas dispensadas", e o número tem de casar com o que o dono fez.
    ''' </summary>
    <TestMethod>
    Public Sub Conversa_dispensada_some_da_fila_e_conta_UMA_VEZ()
        Dim r = Montar({Msg("c1", "caroline@outra.com", 20),
                        Msg("c1", "ricardo@empresa.com", 18),
                        Msg("c1", "caroline@outra.com", 15),
                        Msg("c2", "outro@outra.com", 10)},
                       dispensadas:={"c1"})

        Assert.AreEqual(1, r.Linhas.Count)
        Assert.AreEqual("c2", r.Linhas(0).Conversa)
        Assert.AreEqual(1, r.Fora.ConversasDispensadas,
            "tres mensagens de UMA conversa dispensada contaram tres dispensas")
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

    ''' <summary>
    ''' <b>Dentro do mesmo dia, a ordem sai do instante</b> — e não do assunto.
    '''
    ''' Ordenar por <c>Dias</c> empatava 8 dias e 23 horas com 8 dias e 1 hora, e
    ''' o desempate caía na ordem alfabética: a fila prometia "a mais antiga
    ''' primeiro" e entregava outra coisa dentro do dia. Achado por revisão
    ''' externa.
    ''' </summary>
    <TestMethod>
    Public Sub Dentro_do_mesmo_dia_a_ordem_sai_do_INSTANTE()
        Dim cedo = Msg("zzz-mais-velha", "caroline@outra.com", 8.4)
        Dim tarde = Msg("aaa-mais-nova", "caroline@outra.com", 8.1)

        Dim r = Montar({tarde, cedo})

        Assert.AreEqual(r.Linhas(0).Dias, r.Linhas(1).Dias, "o preparo exige o mesmo dia")
        Assert.AreEqual("zzz-mais-velha", r.Linhas(0).Conversa,
            "dentro do dia a ordem caiu no alfabeto, e nao no relogio")
    End Sub

    ''' <summary>
    ''' Mesma hora e mesmo assunto: a ordem ainda tem de ser estável entre duas
    ''' aberturas. Lista que se reordena sozinha ensina a não confiar nela.
    ''' </summary>
    <TestMethod>
    Public Sub A_ordem_e_ESTAVEL_com_hora_e_assunto_iguais()
        Dim a = Msg("aaa", "caroline@outra.com", 5, "igual")
        Dim b = Msg("bbb", "caroline@outra.com", 5, "igual")

        CollectionAssert.AreEqual({"aaa", "bbb"},
            Montar({a, b}).Linhas.Select(Function(l) l.Conversa).ToArray())
        CollectionAssert.AreEqual({"aaa", "bbb"},
            Montar({b, a}).Linhas.Select(Function(l) l.Conversa).ToArray())
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

    ' ==================================================================
    ' DIA DE CALENDARIO

    ''' <summary>
    ''' <b>Sexta à noite até segunda de manhã são três dias, e não dois.</b>
    '''
    ''' São 56 horas — dois blocos de 24 e sobra. Para quem está esperando, e
    ''' para quem lê a fila de manhã, passaram três datas. A fila conta como a
    ''' pessoa conta.
    '''
    ''' Achado por revisão externa em 31/08/2026: a versão anterior usava
    ''' <c>Floor(TotalDays)</c> e dizia dois.
    ''' </summary>
    <TestMethod>
    Public Sub De_sexta_a_noite_para_segunda_de_manha_sao_TRES_dias()
        Dim sexta = New DateTimeOffset(2026, 8, 28, 23, 30, 0, TimeSpan.FromHours(-3))
        Dim segunda = New DateTimeOffset(2026, 8, 31, 8, 0, 0, TimeSpan.FromHours(-3))

        Assert.IsTrue((segunda - sexta).TotalHours < 72, "o preparo do teste esta errado")
        Assert.AreEqual(3, FilaDeRespostas.DiasDeCalendario(sexta, segunda, Fuso))
    End Sub

    ''' <summary>
    ''' O contraponto: hoje é zero dia, por mais cedo que a mensagem tenha
    ''' chegado. Sem ele, "dia de calendário" poderia virar "sempre pelo menos
    ''' um".
    ''' </summary>
    <TestMethod>
    Public Sub Hoje_e_zero_dia()
        Dim cedo = New DateTimeOffset(2026, 8, 31, 0, 5, 0, TimeSpan.FromHours(-3))
        Dim agora = New DateTimeOffset(2026, 8, 31, 23, 55, 0, TimeSpan.FromHours(-3))

        Assert.AreEqual(0, FilaDeRespostas.DiasDeCalendario(cedo, agora, Fuso))
    End Sub

    ''' <summary>
    ''' O fuso muda a resposta, e é por isso que ele é parâmetro. A mesma
    ''' mensagem, vista de dois lugares, atravessou um número diferente de
    ''' datas.
    ''' </summary>
    <TestMethod>
    Public Sub O_FUSO_muda_a_contagem()
        ' 23h30 em Brasilia e 02h30 do dia seguinte em UTC.
        Dim quando = New DateTimeOffset(2026, 8, 30, 23, 30, 0, TimeSpan.FromHours(-3))
        Dim agora = New DateTimeOffset(2026, 8, 31, 12, 0, 0, TimeSpan.FromHours(-3))

        Assert.AreEqual(1, FilaDeRespostas.DiasDeCalendario(quando, agora, Fuso))
        Assert.AreEqual(0, FilaDeRespostas.DiasDeCalendario(quando, agora, TimeZoneInfo.Utc),
            "em UTC as duas caem no mesmo dia 31")
    End Sub

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

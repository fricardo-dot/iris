Imports System.Collections.Generic
Imports System.Linq
Imports Iris.App.ViewModels
Imports Iris.Assist
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>OS CONSERTOS DA SEGUNDA RODADA DE REVISÕES.</b>
'''
''' Cinco passadas novas do revisor externo acharam onze graves e nove médios —
''' e três deles estavam <b>no código que a primeira rodada tinha acabado de
''' escrever</b>. Este arquivo prende os que dá para prender sem Outlook e sem
''' banco.
'''
''' ------------------------------------------------------------------
''' <b>O QUE ELE GUARDA</b>
'''
''' <list type="bullet">
''' <item>a reserva do rascunho não pode ser usurpada por quem não a
''' apresenta;</item>
''' <item>uma falha de leitura dos rótulos não fica guardada para sempre;</item>
''' <item>a fila não apaga um retrato bom quando a leitura seguinte falha;</item>
''' <item>as identidades que faltam deixam de errar em silêncio.</item>
''' </list>
''' </summary>
<TestClass>
Public Class ConsertosDaSegundaRodadaTests

    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 9, 1, 12, 0, 0, TimeSpan.Zero)

    Private Shared Function Chave(id As String) As ItemKey
        Return New ItemKey(id, "store-1")
    End Function

    ' ==================================================================
    ' A RESERVA NÃO É USURPÁVEL

    ''' <summary>
    ''' <b>Quem não apresenta reserva não passa na frente de quem tem uma.</b>
    '''
    ''' O ramo sem identidade soltava a reserva alheia antes de gravar: ele
    ''' escrevia por cima, liberava a vaga para uma terceira rodada, e a rodada
    ''' em voo descobria no fim que a reserva dela tinha sumido. A identidade só
    ''' protegia quem a apresentava.
    ''' </summary>
    <TestMethod>
    Public Sub Guardar_SEM_reserva_nao_atropela_quem_esta_em_voo()
        Dim sessao As New RascunhosDaSessao()
        Dim daRodada = sessao.Reservar(Chave("a"))
        Assert.IsTrue(daRodada.HasValue, "controle: a reserva tinha de ser dada")

        ' Um chamador sem reserva tenta gravar no meio do voo.
        Assert.IsFalse(sessao.Guardar(Chave("a"), "CK-1", "texto de fora"))

        ' E a rodada que reservou continua conseguindo gravar.
        Assert.IsTrue(sessao.Guardar(Chave("a"), "CK-1", "o texto certo", daRodada))
        Assert.AreEqual("o texto certo", sessao.Pegar(Chave("a"), "CK-1").Texto)
    End Sub

    ''' <summary>
    ''' E o contraponto: sem ninguém em voo, gravar sem reserva continua valendo.
    ''' Sem este, a guarda acima podia virar "só grava quem reservou", que
    ''' quebraria todo caminho manual.
    ''' </summary>
    <TestMethod>
    Public Sub Guardar_sem_reserva_vale_quando_nao_ha_voo()
        Dim sessao As New RascunhosDaSessao()

        Assert.IsTrue(sessao.Guardar(Chave("a"), "CK-1", "um texto"))
    End Sub

    ''' <summary>
    ''' <b>Soltar cita a reserva.</b> Uma rodada velha que solta sem citar
    ''' liberaria a vaga de uma rodada nova, e aí uma terceira pediria a mesma
    ''' redação — pagando duas vezes pelo mesmo texto.
    ''' </summary>
    <TestMethod>
    Public Sub Soltar_de_uma_reserva_VELHA_nao_libera_a_nova()
        Dim sessao As New RascunhosDaSessao()
        Dim velha = sessao.Reservar(Chave("a"))

        sessao.Esquecer()
        Dim nova = sessao.Reservar(Chave("a"))
        Assert.IsTrue(nova.HasValue)

        ' A rodada velha termina e solta citando a reserva dela.
        sessao.Soltar(Chave("a"), velha)

        ' A vaga da nova continua ocupada.
        Assert.IsNull(sessao.Reservar(Chave("a")),
            "a reserva da rodada nova foi liberada por uma rodada velha")
    End Sub

    ' ==================================================================
    ' A FALHA DE LEITURA NÃO FICA GUARDADA

    ''' <summary>
    ''' <b>Uma leitura que estoura não é carimbada.</b>
    '''
    ''' A versão anterior carimbava a falha para não repeti-la a cada linha
    ''' desenhada, e o preço era pior: uma falha transitória — o banco ocupado
    ''' por um lote de classificação — congelava "nenhum rótulo" até uma
    ''' publicação futura, que pode não vir nunca.
    ''' </summary>
    <TestMethod>
    Public Sub Leitura_que_falha_e_tentada_de_NOVO()
        Dim tentativas = 0
        Dim naMao As New RotulosNaMao(
            Function() As Iris.Integration.LeituraDeRotulos
                tentativas += 1
                If tentativas = 1 Then Throw New InvalidOperationException("banco ocupado")
                Return Nothing
            End Function,
            Function() 7)

        Assert.AreEqual(0, naMao.Atual().Rotulos.Count)
        naMao.Atual()

        Assert.AreEqual(2, tentativas,
            "a falha ficou guardada e a segunda leitura nem tentou")
    End Sub

    ''' <summary>
    ''' <b>Gravar rótulo não move o carimbo do acervo</b> — o carimbo conta
    ''' recargas, e gravar rótulo não republica pasta nenhuma. Por isso existe
    ''' <c>Esquecer</c>, e quem grava tem de chamá-lo.
    ''' </summary>
    <TestMethod>
    Public Sub Esquecer_faz_a_leitura_seguinte_ir_ao_banco()
        Dim leituras = 0
        Dim naMao As New RotulosNaMao(
            Function()
                leituras += 1
                Return New Iris.Integration.LeituraDeRotulos(
                    New Dictionary(Of ItemKey, String)(),
                    New Dictionary(Of ItemKey, IReadOnlyList(Of String))(), 0, 1)
            End Function,
            Function() 7)

        naMao.Atual()
        naMao.Atual()
        Assert.AreEqual(1, leituras, "controle: o carimbo parado guarda a leitura")

        naMao.Esquecer()
        naMao.Atual()
        Assert.AreEqual(2, leituras, "Esquecer não fez a leitura seguinte ir ao banco")
    End Sub

    ' ==================================================================
    ' A FILA NÃO APAGA O QUE ESTÁ BOM

    Private Shared Function Msg(conversa As String, deQuem As String,
                                diasAtras As Integer) As MensagemNaFila
        Return New MensagemNaFila(New ItemKey($"E-{conversa}", "s"), conversa,
                                  "assunto", deQuem, deQuem, Agora.AddDays(-diasAtras))
    End Function

    Private Shared Function UmaFila() As ResultadoDaFila
        Return FilaDeRespostas.Montar(
            {Msg("c1", "alguem@fora.com", 5)},
            New MinhasIdentidades({"ricardo@empresa.com"}),
            Agora, TimeZoneInfo.Utc,
            New Dictionary(Of String, DateTimeOffset) From {{"s", Agora}}, Nothing)
    End Function

    ''' <summary>
    ''' <b>Uma leitura que falha não troca dados válidos por uma tela vazia.</b>
    '''
    ''' A fila é a tela que se abre de manhã; perder a referência que já estava
    ''' ali por causa de uma falha transitória é pior do que mostrá-la com um
    ''' aviso de que é de antes — e a frase passa a dizer exatamente isso.
    ''' </summary>
    <TestMethod>
    Public Async Function Falha_na_releitura_NAO_apaga_o_retrato_bom() As Task
        Dim quebrar = False
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) As ResultadoDaFila
                If quebrar Then Throw New InvalidOperationException("o cache caiu")
                Return UmaFila()
            End Function,
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc)

        Await vm.Atualizar()
        Dim antes = vm.Minhas.Count
        Assert.IsTrue(antes > 0, "controle: o cenário tinha de produzir linhas")

        quebrar = True
        Await vm.Atualizar()

        Assert.AreEqual(antes, vm.Minhas.Count, "a falha apagou o retrato bom")
        StringAssert.Contains(vm.Frase, "leitura anterior")
    End Function

    ''' <summary>
    ''' E com a tela vazia a frase é a outra: não há retrato a preservar, e
    ''' dizer "o que está na tela é de antes" sobre uma lista vazia seria uma
    ''' frase sem referente.
    ''' </summary>
    <TestMethod>
    Public Async Function Falha_com_a_tela_vazia_diz_que_a_fila_nao_vale() As Task
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) As ResultadoDaFila
                Throw New InvalidOperationException("o cache caiu")
            End Function,
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc)

        Await vm.Atualizar()

        StringAssert.Contains(vm.Frase, "não vale")
        Assert.IsFalse(vm.Frase.Contains("leitura anterior"))
    End Function

    ' ==================================================================
    ' AS IDENTIDADES QUE FALTAM

    ''' <summary>
    ''' <b>Uma identidade que falta deixa de errar em silêncio.</b>
    '''
    ''' Com o conjunto parcialmente preenchido, uma mensagem que o dono enviou
    ''' por um alias não cadastrado é classificada como <c>DoOutro</c> com toda
    ''' a confiança — a conversa entra na fila como possível resposta dele, e
    ''' pode disparar um rascunho para algo que ele já respondeu.
    '''
    ''' A ressalva mostra os endereços, e não só a contagem: sem eles o dono não
    ''' sabe o que escrever no arquivo.
    ''' </summary>
    <TestMethod>
    Public Async Function A_ressalva_diz_QUAIS_enderecos_faltam() As Task
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) UmaFila(),
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc,
            quemFalta:=Function(eu) CType({"r.silva@empresa.com"}, IReadOnlyList(Of String)))

        Await vm.Atualizar()

        StringAssert.Contains(vm.Ressalva, "r.silva@empresa.com")
        StringAssert.Contains(vm.Ressalva, "identidades.txt")
    End Function

    ''' <summary>
    ''' <b>O diagnóstico não pode derrubar a fila.</b> Se a leitura das
    ''' identidades estourar, a fila continua valendo — ela é o produto; a
    ''' ressalva é o comentário.
    ''' </summary>
    <TestMethod>
    Public Async Function Diagnostico_que_estoura_nao_derruba_a_fila() As Task
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) UmaFila(),
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc,
            quemFalta:=Function(eu) As IReadOnlyList(Of String)
                           Throw New InvalidOperationException("estourou")
                       End Function)

        Await vm.Atualizar()

        Assert.IsTrue(vm.Minhas.Count > 0, "a fila sumiu por causa do diagnóstico")
    End Function

    <TestMethod>
    Public Async Function Sem_endereco_faltando_nao_ha_ressalva_de_identidade() As Task
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) UmaFila(),
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc,
            quemFalta:=Function(eu) CType(Array.Empty(Of String)(), IReadOnlyList(Of String)))

        Await vm.Atualizar()

        Assert.IsFalse(vm.Ressalva.Contains("identidades.txt"))
    End Function

    ' ==================================================================
    ' A DISPENSA VALE POR CAIXA

    ''' <summary>
    ''' <b>Dispensar numa caixa não apaga a conversa homônima da outra.</b>
    '''
    ''' O agrupamento já separava as caixas — o mesmo <c>ConversationID</c> pode
    ''' existir em duas — e a dispensa não: ela era gravada só pelo id. O dono
    ''' não teria como notar, porque o que some some.
    ''' </summary>
    <TestMethod>
    Public Sub Dispensar_numa_caixa_nao_apaga_a_da_OUTRA()
        Dim daPessoal As New MensagemNaFila(New ItemKey("E-1", "pessoal"), "c1",
                                            "assunto", "alguem", "alguem@fora.com",
                                            Agora.AddDays(-5))
        Dim daPartilhada As New MensagemNaFila(New ItemKey("E-2", "partilhada"), "c1",
                                               "assunto", "alguem", "alguem@fora.com",
                                               Agora.AddDays(-5))

        Dim cobertura As New Dictionary(Of String, DateTimeOffset) From {
            {"pessoal", Agora}, {"partilhada", Agora}}

        ' Dispensada SO na partilhada.
        Dim r = FilaDeRespostas.Montar(
            {daPessoal, daPartilhada},
            New MinhasIdentidades({"ricardo@empresa.com"}),
            Agora, TimeZoneInfo.Utc, cobertura,
            {"partilhada" & ControlChars.NullChar & "c1"})

        Assert.AreEqual(1, r.Linhas.Count, "a dispensa de uma caixa levou a outra")
        Assert.AreEqual("pessoal", r.Linhas.Single().Chave.StoreId)
    End Sub

    ''' <summary>
    ''' <b>A linha antiga continua valendo para todas as caixas.</b> É o que o
    ''' dono já escreveu no arquivo, e reinterpretá-la como "só na caixa tal"
    ''' faria conversas dispensadas voltarem sem ninguém pedir.
    ''' </summary>
    <TestMethod>
    Public Sub Dispensa_ANTIGA_so_com_o_id_continua_valendo_em_todas()
        Dim daPessoal As New MensagemNaFila(New ItemKey("E-1", "pessoal"), "c1",
                                            "assunto", "alguem", "alguem@fora.com",
                                            Agora.AddDays(-5))

        Dim r = FilaDeRespostas.Montar(
            {daPessoal},
            New MinhasIdentidades({"ricardo@empresa.com"}),
            Agora, TimeZoneInfo.Utc,
            New Dictionary(Of String, DateTimeOffset) From {{"pessoal", Agora}},
            {"c1"})

        Assert.AreEqual(0, r.Linhas.Count,
            "uma dispensa escrita antes da mudança deixou de valer")
    End Sub

End Class

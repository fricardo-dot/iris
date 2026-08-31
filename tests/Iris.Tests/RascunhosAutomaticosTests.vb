Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Integration
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>OS RASCUNHOS AUTOMÁTICOS — Fase 8.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTE ARQUIVO PRENDE</b>
'''
''' Duas coisas, e as duas são sobre não fazer nada nas costas do dono:
'''
''' <list type="number">
''' <item>só quem está esperando ganha rascunho, e há um teto — sem ele, mandar
''' varrer uma pasta grande dispararia centenas de redações de uma vez, e a
''' conta chegaria depois;</item>
''' <item>um rascunho escrito para uma versão anterior da mensagem <b>não é
''' entregue</b> — ele responde a um texto que não está mais lá, e o dono não
''' tem como perceber isso lendo o rascunho.</item>
''' </list>
'''
''' ------------------------------------------------------------------
''' <b>O CONTROLE NEGATIVO</b>
'''
''' <see cref="Rascunho_de_versao_ANTERIOR_nao_e_entregue"/>. Sem ele, um
''' armazenamento que ignorasse a versão passaria em todo o resto daqui — e
''' seria exatamente o que põe na tela uma resposta para um e-mail que mudou.
''' </summary>
<TestClass>
Public Class RascunhosAutomaticosTests

    Private Shared Function Mensagem(id As String,
                                     Optional dia As Integer = 10) As MensagemNaFila
        Return New MensagemNaFila(
            New ItemKey(id, "store-1"), "conversa-" & id, "assunto " & id,
            "Alguém", "alguem@exemplo.com",
            New DateTimeOffset(2026, 8, dia, 12, 0, 0, TimeSpan.Zero))
    End Function

    Private Shared Function Chave(id As String) As ItemKey
        Return New ItemKey(id, "store-1")
    End Function

    Private Shared Function Rotulos(ParamArray pares As String()) _
                                    As IReadOnlyDictionary(Of ItemKey, String)
        Dim mapa As New Dictionary(Of ItemKey, String)()
        For i = 0 To pares.Length - 1 Step 2
            mapa(Chave(pares(i))) = pares(i + 1)
        Next
        Return mapa
    End Function

    Private Shared Function Escolher(mensagens As MensagemNaFila(),
                                     rotulos As IReadOnlyDictionary(Of ItemKey, String),
                                     Optional jaTem As ItemKey() = Nothing,
                                     Optional dispensadas As ItemKey() = Nothing,
                                     Optional teto As Integer = RascunhosAutomaticos.PorRodada) _
                                     As IReadOnlyList(Of MensagemNaFila)
        Return RascunhosAutomaticos.Escolher(
            mensagens, rotulos,
            If(jaTem, Array.Empty(Of ItemKey)()),
            If(dispensadas, Array.Empty(Of ItemKey)()), teto)
    End Function

    ' ==================================================================
    ' QUEM MERECE

    ''' <summary>
    ''' Só quem está esperando. Redigir para uma newsletter é queimar dinheiro;
    ''' redigir para o que ele já respondeu é pior, porque produz um texto que
    ''' <i>parece</i> pendência.
    ''' </summary>
    <TestMethod>
    Public Sub So_o_que_espera_resposta_ganha_rascunho()
        Dim escolhidas = Escolher(
            {Mensagem("a"), Mensagem("b"), Mensagem("c")},
            Rotulos("a", "precisa_de_mim", "b", "newsletter", "c", "aguardando"))

        Assert.AreEqual(1, escolhidas.Count)
        Assert.AreEqual("a", escolhidas.Single().Chave.EntryId)
    End Sub

    <TestMethod>
    Public Sub Mensagem_sem_rotulo_nao_ganha_rascunho()
        Assert.AreEqual(0, Escolher({Mensagem("a")}, Rotulos()).Count)
    End Sub

    ''' <summary>
    ''' <b>Mais velha primeiro.</b> O rascunho existe para desatolar o que está
    ''' parado; começar pelo recente atenderia primeiro quem menos esperou, e o
    ''' teto faria o resto nunca chegar.
    ''' </summary>
    <TestMethod>
    Public Sub A_mais_velha_e_redigida_primeiro()
        Dim escolhidas = Escolher(
            {Mensagem("nova", dia:=25), Mensagem("velha", dia:=2)},
            Rotulos("nova", "precisa_de_mim", "velha", "precisa_de_mim"))

        Assert.AreEqual("velha", escolhidas.First().Chave.EntryId)
    End Sub

    ''' <summary>
    ''' <b>O teto.</b> Sem ele, mandar varrer uma pasta grande dispararia
    ''' centenas de redações de uma vez, e a conta chegaria depois.
    ''' </summary>
    <TestMethod>
    Public Sub O_teto_da_rodada_e_respeitado()
        Dim muitas = Enumerable.Range(1, 50).
                     Select(Function(i) Mensagem("m" & i, dia:=1 + i Mod 27)).ToArray()
        Dim rotulos As New Dictionary(Of ItemKey, String)()
        For Each m In muitas
            rotulos(m.Chave) = "precisa_de_mim"
        Next

        Assert.AreEqual(RascunhosAutomaticos.PorRodada,
                        RascunhosAutomaticos.Escolher(muitas, rotulos,
                                                      Array.Empty(Of ItemKey)(),
                                                      Array.Empty(Of ItemKey)()).Count)
    End Sub

    ''' <summary>
    ''' Um teto pedido maior que o da classe não vale. O parâmetro serve para
    ''' pedir <b>menos</b>; deixá-lo aumentar faria a proteção depender de quem
    ''' chama, que é o mesmo que não tê-la.
    ''' </summary>
    <TestMethod>
    Public Sub Teto_pedido_MAIOR_nao_passa_do_teto_da_classe()
        Dim muitas = Enumerable.Range(1, 50).Select(Function(i) Mensagem("m" & i)).ToArray()
        Dim rotulos As New Dictionary(Of ItemKey, String)()
        For Each m In muitas
            rotulos(m.Chave) = "precisa_de_mim"
        Next

        Assert.AreEqual(RascunhosAutomaticos.PorRodada,
                        Escolher(muitas, rotulos, teto:=500).Count)
    End Sub

    <TestMethod>
    Public Sub O_que_ja_tem_rascunho_nao_e_refeito()
        Dim escolhidas = Escolher({Mensagem("a")}, Rotulos("a", "precisa_de_mim"),
                                  jaTem:={Chave("a")})

        Assert.AreEqual(0, escolhidas.Count)
    End Sub

    ''' <summary>
    ''' A recusa dele vale mais que o rótulo. Uma mensagem que ele dispensou não
    ''' volta porque a classificação continua dizendo que espera resposta.
    ''' </summary>
    <TestMethod>
    Public Sub O_que_ele_dispensou_nao_volta()
        Dim escolhidas = Escolher({Mensagem("a")}, Rotulos("a", "precisa_de_mim"),
                                  dispensadas:={Chave("a")})

        Assert.AreEqual(0, escolhidas.Count)
    End Sub

    ' ==================================================================
    ' A SESSÃO

    <TestMethod>
    Public Sub O_rascunho_guardado_volta()
        Dim sessao As New RascunhosDaSessao()
        sessao.Guardar(Chave("a"), "CK-1", "Prezado, segue a resposta.")

        Assert.AreEqual("Prezado, segue a resposta.", sessao.Pegar(Chave("a"), "CK-1").Texto)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo.</b> Um rascunho escrito para a versão anterior
    ''' responde a um texto que não está mais lá — e o dono não tem como
    ''' perceber isso lendo o rascunho.
    ''' </summary>
    <TestMethod>
    Public Sub Rascunho_de_versao_ANTERIOR_nao_e_entregue()
        Dim sessao As New RascunhosDaSessao()
        sessao.Guardar(Chave("a"), "CK-1", "resposta à versão velha")

        Assert.IsNull(sessao.Pegar(Chave("a"), "CK-2"))
    End Sub

    ''' <summary>
    ''' Mas ele <b>não é apagado</b>: quem não entrega ainda sabe que existiu, e
    ''' o dono que se lembra de tê-lo visto merece a frase "aquele rascunho era
    ''' de uma versão anterior" em vez do silêncio.
    '''
    ''' E, por isso mesmo, a rodada não o refaz sozinha — refazer a cada
    ''' mudança de versão gastaria dinheiro sem ele ter pedido nada.
    ''' </summary>
    <TestMethod>
    Public Sub Rascunho_de_versao_anterior_nao_e_APAGADO_nem_refeito()
        Dim sessao As New RascunhosDaSessao()
        sessao.Guardar(Chave("a"), "CK-1", "resposta à versão velha")

        Assert.IsTrue(sessao.Tem(Chave("a")))
        Assert.AreEqual(0, Escolher({Mensagem("a")}, Rotulos("a", "precisa_de_mim"),
                                    jaTem:=sessao.Feitos().ToArray()).Count)
    End Sub

    ''' <summary>
    ''' Dispensar apaga o texto. Deixá-lo lá depois de ele recusar seria manter
    ''' na tela justamente o que ele mandou tirar.
    ''' </summary>
    <TestMethod>
    Public Sub Dispensar_APAGA_o_rascunho_que_havia()
        Dim sessao As New RascunhosDaSessao()
        sessao.Guardar(Chave("a"), "CK-1", "um texto")

        sessao.Dispensar(Chave("a"))

        Assert.IsFalse(sessao.Tem(Chave("a")))
        Assert.IsNull(sessao.Pegar(Chave("a"), "CK-1"))
        CollectionAssert.Contains(sessao.Dispensadas().ToArray(), Chave("a"))
    End Sub

    <TestMethod>
    Public Sub Esquecer_leva_tudo()
        Dim sessao As New RascunhosDaSessao()
        sessao.Guardar(Chave("a"), "CK-1", "um texto")
        sessao.Dispensar(Chave("b"))

        sessao.Esquecer()

        Assert.AreEqual(0, sessao.Feitos().Count)
        Assert.AreEqual(0, sessao.Dispensadas().Count)
    End Sub

    <TestMethod>
    Public Sub Texto_em_branco_nao_vira_rascunho()
        Dim sessao As New RascunhosDaSessao()

        Assert.IsFalse(sessao.Guardar(Chave("a"), "CK-1", "   "))
        Assert.IsFalse(sessao.Tem(Chave("a")))
    End Sub

    ''' <summary>
    ''' <b>Sem versão não se guarda.</b>
    '''
    ''' Duas ausências passavam por igualdade: guardado sem <c>PR_CHANGE_KEY</c>
    ''' e pedido sem <c>PR_CHANGE_KEY</c>, o rascunho voltava — e continuaria
    ''' voltando depois de a mensagem mudar, porque ninguém tinha como comparar
    ''' nada. Ausência não prova que nada mudou; prova que ninguém sabe. Achado
    ''' por revisão externa em 31/08/2026.
    ''' </summary>
    <TestMethod>
    Public Sub Rascunho_SEM_versao_nao_e_guardado()
        Dim sessao As New RascunhosDaSessao()

        Assert.IsFalse(sessao.Guardar(Chave("a"), Nothing, "um texto"))
        Assert.IsFalse(sessao.Tem(Chave("a")))
        Assert.IsNull(sessao.Pegar(Chave("a"), Nothing))
    End Sub


    ' ==================================================================
    ' A RODADA

    ''' <summary>
    ''' <b>Um pedido por mensagem, e não um lote.</b>
    '''
    ''' A classificação vai em lotes porque o resultado dela é uma palavra de um
    ''' enum. Aqui o resultado é texto escrito em nome do dono: dois corpos
    ''' hostis dividindo o mesmo contexto significaria que o conteúdo de um
    ''' e-mail pode influenciar a resposta que ele vai mandar para outra pessoa,
    ''' e essa é a única contaminação que não dá para consertar com uma
    ''' superfície fechada — porque a superfície aqui é prosa livre.
    ''' </summary>
    <TestMethod>
    Public Sub Cada_mensagem_tem_o_SEU_pedido()
        Dim sessao As New RascunhosDaSessao()
        Dim pedidos As New List(Of ItemKey)()

        Dim r = New RascunhosDeUmaRodada(sessao).Passar(
            {Mensagem("a"), Mensagem("b")},
            Rotulos("a", "precisa_de_mim", "b", "precisa_de_mim"),
            Function(m)
                pedidos.Add(m.Chave)
                Return New RedacaoFeita("resposta para " & m.Chave.EntryId, "CK-1")
            End Function)

        Assert.AreEqual(2, pedidos.Count)
        Assert.AreEqual(2, r.Escritos)
        Assert.AreEqual("resposta para a", sessao.Pegar(Chave("a"), "CK-1").Texto)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo da rodada.</b> A mensagem que falha pode ser
    ''' justamente a hostil, e deixá-la parar a rodada daria a ela o poder de
    ''' impedir os rascunhos de todas as outras.
    ''' </summary>
    <TestMethod>
    Public Sub Falha_numa_mensagem_NAO_derruba_a_rodada()
        Dim sessao As New RascunhosDaSessao()

        Dim r = New RascunhosDeUmaRodada(sessao).Passar(
            {Mensagem("a", dia:=1), Mensagem("b", dia:=2)},
            Rotulos("a", "precisa_de_mim", "b", "precisa_de_mim"),
            Function(m)
                If m.Chave.EntryId = "a" Then Throw New InvalidOperationException("estourou")
                Return New RedacaoFeita("resposta", "CK-1")
            End Function)

        Assert.AreEqual(1, r.Escritos)
        Assert.AreEqual(1, r.Falharam)
        Assert.IsTrue(sessao.Tem(Chave("b")))
        Assert.IsFalse(sessao.Tem(Chave("a")))
    End Sub

    ''' <summary>
    ''' Interromper para <b>entre</b> mensagens, e o que já foi escrito fica:
    ''' foi pago. Jogar fora o texto pronto porque o dono fechou o painel seria
    ''' cobrar duas vezes pela mesma coisa.
    ''' </summary>
    <TestMethod>
    Public Sub Interromper_guarda_o_que_ja_foi_escrito()
        Dim sessao As New RascunhosDaSessao()
        Dim parada As New Threading.CancellationTokenSource()

        Dim r = New RascunhosDeUmaRodada(sessao).Passar(
            {Mensagem("a", dia:=1), Mensagem("b", dia:=2)},
            Rotulos("a", "precisa_de_mim", "b", "precisa_de_mim"),
            Function(m)
                parada.Cancel()
                Return New RedacaoFeita("resposta para " & m.Chave.EntryId, "CK-1")
            End Function,
            parar:=parada.Token)

        Assert.IsTrue(r.Interrompida)
        Assert.AreEqual(1, r.Escritos)
        Assert.IsTrue(sessao.Tem(Chave("a")), "jogou fora o que já tinha sido pago")
    End Sub

    ''' <summary>
    ''' Rodar de novo não refaz o que já tem rascunho — e a rodada lê isso da
    ''' própria sessão, sem quem chama ter de lembrar.
    ''' </summary>
    <TestMethod>
    Public Sub A_segunda_rodada_nao_refaz_o_que_ja_existe()
        Dim sessao As New RascunhosDaSessao()
        Dim rodada As New RascunhosDeUmaRodada(sessao)

        Dim redator As RascunhosDeUmaRodada.Redigir =
            Function(m) New RedacaoFeita("resposta", "CK-1")

        rodada.Passar({Mensagem("a")}, Rotulos("a", "precisa_de_mim"), redator)
        Dim r = rodada.Passar({Mensagem("a")}, Rotulos("a", "precisa_de_mim"), redator)

        Assert.AreEqual(0, r.Escolhidas)
    End Sub

    ''' <summary>
    ''' Texto em branco vale como "não deu", e não como rascunho vazio. Um
    ''' rascunho vazio na tela diz "eu tentei e a resposta é nada", que é uma
    ''' afirmação — e faria a rodada seguinte pular a mensagem para sempre.
    ''' </summary>
    <TestMethod>
    Public Sub Texto_vazio_conta_como_falha_e_a_mensagem_volta()
        Dim sessao As New RascunhosDaSessao()
        Dim rodada As New RascunhosDeUmaRodada(sessao)

        Dim r = rodada.Passar({Mensagem("a")}, Rotulos("a", "precisa_de_mim"),
                              Function(m) New RedacaoFeita("", "CK-1"))

        Assert.AreEqual(1, r.Falharam)
        Assert.IsFalse(sessao.Tem(Chave("a")))

        ' E ela volta na rodada seguinte.
        Dim outra = rodada.Passar({Mensagem("a")}, Rotulos("a", "precisa_de_mim"),
                                  Function(m) New RedacaoFeita("agora foi", "CK-1"))
        Assert.AreEqual(1, outra.Escritos)
    End Sub


    ' ==================================================================
    ' AS CORRIDAS

    ''' <summary>
    ''' <b>Dispensar durante a redação vence a redação.</b>
    '''
    ''' Entre pedir e guardar passam segundos, e neles o dono pode dispensar. A
    ''' versão anterior guardava assim mesmo — a trava protegia os dicionários e
    ''' não a decisão —, e o texto que ele mandou tirar voltava. O trabalho pago
    ''' se perde, e é o lado certo de perder.
    ''' </summary>
    <TestMethod>
    Public Sub Dispensar_DURANTE_a_redacao_vence_a_redacao()
        Dim sessao As New RascunhosDaSessao()

        Dim r = New RascunhosDeUmaRodada(sessao).Passar(
            {Mensagem("a")}, Rotulos("a", "precisa_de_mim"),
            Function(m)
                ' Enquanto a redacao esta em voo, o dono dispensa.
                sessao.Dispensar(m.Chave)
                Return New RedacaoFeita("um texto que ele nao quer", "CK-1")
            End Function)

        Assert.IsFalse(sessao.Tem(Chave("a")),
            "o rascunho voltou depois de ele mandar tirar")
        Assert.AreEqual(1, r.Falharam)
        Assert.AreEqual(0, r.Escritos)
    End Sub

    ''' <summary>
    ''' <b>Esquecer durante a redação não deixa o texto ressuscitar</b> numa
    ''' sessão que já foi descartada. É a geração da reserva que o impede.
    ''' </summary>
    <TestMethod>
    Public Sub Esquecer_DURANTE_a_redacao_nao_deixa_ressuscitar()
        Dim sessao As New RascunhosDaSessao()

        Dim r = New RascunhosDeUmaRodada(sessao).Passar(
            {Mensagem("a")}, Rotulos("a", "precisa_de_mim"),
            Function(m)
                sessao.Esquecer()
                Return New RedacaoFeita("um texto de outra sessão", "CK-1")
            End Function)

        Assert.IsFalse(sessao.Tem(Chave("a")))
        Assert.AreEqual(1, r.Falharam)
    End Sub

    ''' <summary>
    ''' <b>A mesma mensagem duas vezes na entrada gasta um pedido só.</b> A lista
    ''' vem do acervo e não promete unicidade; sem isto, a duplicata gastava dois
    ''' pedidos para produzir o mesmo texto, e o segundo sobrescrevia o primeiro.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_repetida_na_entrada_gasta_um_pedido_so()
        Dim sessao As New RascunhosDaSessao()
        Dim pedidos = 0

        Dim r = New RascunhosDeUmaRodada(sessao).Passar(
            {Mensagem("a"), Mensagem("a")}, Rotulos("a", "precisa_de_mim"),
            Function(m)
                pedidos += 1
                Return New RedacaoFeita("resposta", "CK-1")
            End Function)

        Assert.AreEqual(1, pedidos)
        Assert.AreEqual(1, r.Escritos)
    End Sub

    ''' <summary>
    ''' Chave vazia não é mensagem: um <c>ItemKey</c> sem <c>EntryId</c> não
    ''' identifica nada, e duas dessas colidem entre si — a rodada gastaria uma
    ''' redação e a guardaria por cima da anterior.
    ''' </summary>
    <TestMethod>
    Public Sub Chave_VAZIA_nao_ganha_rascunho()
        Dim vazia As New MensagemNaFila(New ItemKey("", ""), "c", "assunto",
                                        "quem", "quem@exemplo.com", Nothing)
        Dim rotulos As New Dictionary(Of ItemKey, String) From {
            {New ItemKey("", ""), "precisa_de_mim"}}

        Assert.AreEqual(0, RascunhosAutomaticos.Escolher(
            {vazia}, rotulos, Array.Empty(Of ItemKey)(),
            Array.Empty(Of ItemKey)()).Count)
    End Sub

    ''' <summary>
    ''' <b>Cancelamento de dentro do provedor é falha daquela mensagem</b>, e não
    ''' interrupção do dono.
    '''
    ''' Antes, qualquer <c>OperationCanceledException</c> parava a rodada — e uma
    ''' mensagem hostil capaz de provocá-la derrubava os rascunhos de todas as
    ''' outras, que é exatamente o que o resto deste arquivo existe para impedir.
    ''' </summary>
    <TestMethod>
    Public Sub Cancelamento_de_DENTRO_do_provedor_nao_para_a_rodada()
        Dim sessao As New RascunhosDaSessao()

        Dim r = New RascunhosDeUmaRodada(sessao).Passar(
            {Mensagem("a", dia:=1), Mensagem("b", dia:=2)},
            Rotulos("a", "precisa_de_mim", "b", "precisa_de_mim"),
            Function(m)
                If m.Chave.EntryId = "a" Then
                    Throw New OperationCanceledException("timeout do provedor")
                End If
                Return New RedacaoFeita("resposta", "CK-1")
            End Function)

        Assert.IsFalse(r.Interrompida, "o timeout do provedor virou interrupção do dono")
        Assert.AreEqual(1, r.Escritos)
        Assert.AreEqual(1, r.Falharam)
    End Sub

    ''' <summary>
    ''' As contas fecham: <c>Tentadas = Escritos + Falharam</c>, e
    ''' <c>Escolhidas</c> conta as que <i>mereciam</i>. Numa rodada interrompida
    ''' as duas divergem, e é essa divergência que diz quantas não foram
    ''' tocadas.
    ''' </summary>
    <TestMethod>
    Public Sub Escolhidas_e_Tentadas_sao_coisas_diferentes()
        Dim sessao As New RascunhosDaSessao()
        Dim parada As New Threading.CancellationTokenSource()

        Dim r = New RascunhosDeUmaRodada(sessao).Passar(
            {Mensagem("a", dia:=1), Mensagem("b", dia:=2), Mensagem("c", dia:=3)},
            Rotulos("a", "precisa_de_mim", "b", "precisa_de_mim",
                    "c", "precisa_de_mim"),
            Function(m)
                parada.Cancel()
                Return New RedacaoFeita("resposta", "CK-1")
            End Function,
            parar:=parada.Token)

        Assert.AreEqual(3, r.Escolhidas)
        Assert.AreEqual(1, r.Tentadas)
        Assert.AreEqual(r.Escritos + r.Falharam, r.Tentadas)
    End Sub

End Class

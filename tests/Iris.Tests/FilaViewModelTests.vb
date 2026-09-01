Imports System.Collections.Generic
Imports System.Linq
Imports Iris.App.ViewModels
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>AS TRÊS FRASES DA FILA — e por que a do meio existe.</b>
'''
''' ------------------------------------------------------------------
''' <b>A FRASE É A PARTE QUE MAIS IMPORTA DESTA TELA</b>
'''
''' Uma lista vazia sem frase deixa o dono concluir "não tenho nada", e esse é o
''' único desfecho que a fila não pode produzir por engano. As três recusas não
''' se parecem:
'''
''' <list type="bullet">
''' <item><b>Sem os enviados</b> — não dá para montar.</item>
''' <item><b>Nada classificável</b> — vi conversas e não sei de quem é a vez;
''' quase sempre falta um endereço em <c>identidades.txt</c>.</item>
''' <item><b>Vazia</b> — olhei e não há nada esperando.</item>
''' </list>
'''
''' A do meio existe por causa de um defeito real: sem ela, uma caixa cheia com
''' as identidades incompletas produzia a tela dizendo que o dia estava limpo.
'''
''' ------------------------------------------------------------------
''' <b>ESTES TESTES OLHAM A FRASE, E NÃO O ENUM</b>
'''
''' Cobrar o <c>Motivo</c> provaria que o ViewModel copiou um campo. O que
''' importa é o dono conseguir <b>distinguir</b> os três casos lendo a tela — e
''' é por isso que cada teste exige que a frase diga a coisa daquele caso, e não
''' a dos outros.
''' </summary>
<TestClass>
Public Class FilaViewModelTests

    Private Shared ReadOnly Agora As New DateTimeOffset(2026, 8, 31, 12, 0, 0, TimeSpan.Zero)

    Private Shared Function Montar(mensagens As IEnumerable(Of MensagemNaFila),
                                   viuOsEnviados As Boolean,
                                   Optional eu As MinhasIdentidades = Nothing) As ResultadoDaFila
        Dim cobertura As IReadOnlyDictionary(Of String, DateTimeOffset) =
            If(viuOsEnviados,
               New Dictionary(Of String, DateTimeOffset) From {{"s", Agora}},
               New Dictionary(Of String, DateTimeOffset)())
        Return FilaDeRespostas.Montar(mensagens,
                                      If(eu, New MinhasIdentidades({"ricardo@empresa.com"})),
                                      Agora, TimeZoneInfo.Utc, cobertura, Nothing)
    End Function

    ''' <summary>
    ''' Um leitor que estoura. Em VB o <c>Throw</c> não é expressão, então ele
    ''' precisa de nome próprio — e ter nome deixa o teste dizer o que encena.
    ''' </summary>
    Private Shared Function Estourar() As ResultadoDaFila
        Throw New InvalidOperationException("o cache caiu")
    End Function

    Private Shared Function Msg(conversa As String, deQuem As String,
                                diasAtras As Integer) As MensagemNaFila
        Return New MensagemNaFila(New ItemKey($"E-{conversa}", "s"), conversa,
                                  "assunto", deQuem, deQuem, Agora.AddDays(-diasAtras))
    End Function

    ' ==================================================================
    ' AS TRES FRASES

    <TestMethod>
    Public Sub Sem_os_enviados_a_frase_diz_que_NAO_DA_para_montar()
        Dim frase = FilaViewModel.FraseDe(
            Montar({Msg("c1", "caroline@outra.com", 20)}, viuOsEnviados:=False))

        StringAssert.Contains(frase, "Itens Enviados",
            "a frase tem de dizer o que fazer, e nao so que falhou")
        Assert.IsFalse(frase.Contains("Nada esperando"),
            "recusa disfarcada de fila vazia e o defeito que esta tela nao " &
            "pode ter")
    End Sub

    ''' <summary>
    ''' <b>A frase do meio.</b> Caixa cheia, identidades incompletas, lista
    ''' vazia — e a tela tem de dizer que <b>não sabe</b>, apontando para o
    ''' arquivo que o dono consegue corrigir.
    ''' </summary>
    <TestMethod>
    Public Sub Nada_classificavel_a_frase_manda_olhar_as_identidades()
        Dim frase = FilaViewModel.FraseDe(
            Montar({Msg("c1", "caroline@outra.com", 20),
                    Msg("c2", "outro@outra.com", 9)},
                   viuOsEnviados:=True,
                   eu:=New MinhasIdentidades({})))

        StringAssert.Contains(frase, "identidades.txt",
            "a frase precisa dizer ONDE consertar; sem isso o dono so sabe " &
            "que algo esta errado")
        StringAssert.Contains(frase, "2", "e quantas conversas ele nao viu")
        Assert.IsFalse(frase.Contains("Nada esperando"),
            "'nao sei de quem e a vez' virou 'nao ha nada esperando'")
    End Sub

    <TestMethod>
    Public Sub Fila_vazia_de_verdade_diz_que_olhou()
        Dim frase = FilaViewModel.FraseDe(
            Montar(Array.Empty(Of MensagemNaFila)(), viuOsEnviados:=True))

        StringAssert.Contains(frase, "Nada esperando")
        StringAssert.Contains(frase, "Olhei", "'nada' sem 'olhei' nao distingue " &
            "fila vazia de fila que nao rodou")
    End Sub

    ''' <summary>
    ''' As três frases são <b>diferentes entre si</b>. É o controle: sem ele,
    ''' um ViewModel que devolvesse a mesma frase sempre passaria nos três
    ''' testes acima, desde que ela contivesse todas as palavras.
    ''' </summary>
    <TestMethod>
    Public Sub As_tres_frases_sao_DIFERENTES()
        Dim recusa = FilaViewModel.FraseDe(Montar({Msg("c1", "x@y.com", 5)}, False))
        Dim naoSei = FilaViewModel.FraseDe(
            Montar({Msg("c1", "x@y.com", 5)}, True, New MinhasIdentidades({})))
        Dim vazia = FilaViewModel.FraseDe(Montar(Array.Empty(Of MensagemNaFila)(), True))

        Dim todas = {recusa, naoSei, vazia}
        Assert.AreEqual(3, todas.Distinct().Count(),
            "duas das tres situacoes produzem a mesma frase, e o dono nao " &
            "consegue distingui-las")
    End Sub

    <TestMethod>
    Public Sub Com_linhas_a_frase_traz_o_numerador_e_o_denominador()
        Dim frase = FilaViewModel.FraseDe(
            Montar({Msg("c1", "caroline@outra.com", 20),
                    Msg("c2", "ricardo@empresa.com", 5)}, True))

        StringAssert.Contains(frase, "2 de 2")
    End Sub

    ' ==================================================================
    ' A RESSALVA

    ''' <summary>
    ''' A ressalva traz a <b>unidade</b> junto de cada número: mensagens e
    ''' conversas são coisas diferentes, e um total somando as duas não teria
    ''' significado.
    ''' </summary>
    <TestMethod>
    Public Sub A_ressalva_diz_a_unidade_de_cada_numero()
        Dim semConversa = New MensagemNaFila(New ItemKey("E-1", "s"), "", "a", "x",
                                             "x@y.com", Agora.AddDays(-2))
        Dim r = Montar({Msg("boa", "caroline@outra.com", 5),
                        Msg("ruim", "", 3),
                        semConversa}, True)

        Dim ressalva = FilaViewModel.RessalvaDe(r)

        StringAssert.Contains(ressalva, "conversa(s) sem saber de quem é a vez")
        StringAssert.Contains(ressalva, "mensagem(ns) sem conversa legível")
    End Sub

    ''' <summary>
    ''' Nada de fora, nada de ressalva. Ressalva que aparece sempre vira ruído, e
    ''' ruído ensina a não ler — inclusive quando ela tem algo a dizer.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_nada_de_fora_NAO_ha_ressalva()
        Dim r = Montar({Msg("c1", "caroline@outra.com", 5)}, True)

        Assert.AreEqual("", FilaViewModel.RessalvaDe(r))
    End Sub

    ''' <summary>
    ''' Recusa não tem ressalva: não houve triagem, então não há o que ficar de
    ''' fora. Uma ressalva ali sugeriria que a fila rodou.
    ''' </summary>
    <TestMethod>
    Public Sub Recusa_nao_tem_ressalva()
        Assert.AreEqual("", FilaViewModel.RessalvaDe(
            Montar({Msg("c1", "caroline@outra.com", 5)}, False)))
    End Sub

    ''' <summary>
    ''' <b>Clicar em Abrir e nada acontecer ensina que o botão não funciona.</b>
    '''
    ''' A fila conhece a chave e não conhece a lista: ela veio do cache, e a
    ''' lista mostra a pasta selecionada. Quando a mensagem não está lá, a tela
    ''' <b>diz</b> — em vez de fingir que abriu.
    ''' </summary>
    <TestMethod>
    Public Async Function Quando_nao_da_para_abrir_a_tela_DIZ() As Task
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) _
                Montar(Array.Empty(Of MensagemNaFila)(), True),
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc)

        Await vm.Atualizar()
        Dim antes = vm.Frase

        vm.NaoDeuParaAbrir()

        Assert.AreNotEqual(antes, vm.Frase, "clicou em Abrir e a tela nao mudou")
        StringAssert.Contains(vm.Frase, "pasta",
            "a frase precisa dizer o que fazer, e nao so que falhou")
    End Function

    ''' <summary>
    ''' <b>Falha do leitor não derruba a janela.</b>
    '''
    ''' <c>Atualizar</c> roda dentro de um comando do WPF, no dispatcher: uma
    ''' exceção do cache subiria sem ninguém para pegá-la, e o programa fecharia
    ''' porque uma lista não carregou. A fila some e diz por quê — que é o que
    ''' ela já faz nas outras recusas.
    '''
    ''' Achado por revisão externa em 31/08/2026.
    ''' </summary>
    <TestMethod>
    Public Async Function Excecao_do_leitor_NAO_derruba_a_tela() As Task
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) Estourar(),
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc)

        Await vm.Atualizar()

        Assert.IsFalse(vm.Respondeu)
        Assert.AreEqual(0, vm.Minhas.Count)
        StringAssert.Contains(vm.Frase, "não vale",
            "a tela ficou sem dizer que a fila nao vale")
    End Function

    ''' <summary>Leitor que devolve <c>Nothing</c> também não derruba.</summary>
    <TestMethod>
    Public Async Function Resultado_nulo_NAO_derruba_a_tela() As Task
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) _
                CType(Nothing, ResultadoDaFila),
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc)

        Await vm.Atualizar()

        Assert.IsFalse(vm.Respondeu)
        Assert.AreNotEqual("", vm.Frase)
    End Function

    ''' <summary>
    ''' <b>Atualizar preenche as duas coleções, e a segunda chamada substitui a
    ''' primeira.</b> Sem isto, uma fila que só acrescentasse duplicaria as
    ''' linhas a cada clique em Atualizar.
    ''' </summary>
    <TestMethod>
    Public Async Function Atualizar_SUBSTITUI_as_listas() As Task
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) _
                Montar({Msg("c1", "caroline@outra.com", 9),
                        Msg("c2", "ricardo@empresa.com", 4)}, True),
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc)

        Await vm.Atualizar()
        Await vm.Atualizar()

        Assert.AreEqual(1, vm.Minhas.Count, "a segunda leitura duplicou as linhas")
        Assert.AreEqual(1, vm.Deles.Count)
        Assert.IsTrue(vm.Respondeu)
    End Function

    ''' <summary>
    ''' <b>Dispensa que não grava não tira a linha da tela.</b>
    '''
    ''' Sumir com a linha depois de uma gravação que falhou deixaria o dono
    ''' achando que resolveu, e a conversa voltaria na abertura seguinte sem
    ''' explicação. Aqui o arquivo é impossível de criar.
    ''' </summary>
    <TestMethod>
    Public Async Function Dispensa_que_falha_MANTEM_a_linha() As Task
        Dim atrapalho = IO.Path.Combine(IO.Path.GetTempPath(),
                                        "iris-fila-vm-" & Guid.NewGuid().ToString("N"))
        Try
            IO.File.WriteAllText(atrapalho, "sou um arquivo, e nao uma pasta")

            Dim vm As New FilaViewModel(
                Function(eu, agora, fuso, dispensadas, ignorados) _
                    Montar({Msg("c1", "caroline@outra.com", 9)}, True),
                New Iris.Integration.DispensasDaFila(atrapalho),
                Nothing, Function() Agora, TimeZoneInfo.Utc)

            Await vm.Atualizar()
            Assert.AreEqual(1, vm.Minhas.Count, "o preparo do teste esta errado")

            vm.Minhas(0).DispensarCommand.Execute(Nothing)

            Assert.AreEqual(1, vm.Minhas.Count,
                "a linha sumiu apesar de a dispensa nao ter sido gravada")
            StringAssert.Contains(vm.Frase, "continua na fila")
        Finally
            If IO.File.Exists(atrapalho) Then IO.File.Delete(atrapalho)
        End Try
    End Function


    ' ==================================================================
    ' A ORDEM POR PRIORIDADE (Fase 9)

    ''' <summary>
    ''' Duas conversas com a última palavra do outro lado: uma de vinte dias e
    ''' uma de dois.
    ''' </summary>
    Private Shared Function Duas() As ResultadoDaFila
        Return Montar({Msg("velha", "cliente@fora.com", 20),
                       Msg("nova", "outro@fora.com", 2)}, viuOsEnviados:=True)
    End Function

    Private Shared Function ComPrioridade(Optional rotulo As Func(Of ItemKey, String) = Nothing) _
                                          As FilaViewModel
        Return New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) Duas(),
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc,
            rotulo:=rotulo)
    End Function

    ''' <summary>
    ''' <b>Por padrão, a ordem é a da idade.</b> Ela é a única que não depende de
    ''' opinião nenhuma, e a nota é feita de pesos que ninguém mediu.
    ''' </summary>
    <TestMethod>
    Public Sub A_ordem_por_prioridade_vem_DESLIGADA()
        Assert.IsFalse(ComPrioridade().PorPrioridade)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo da fase.</b> Ligar a prioridade REORDENA e não
    ''' esconde: as mesmas linhas, com os dias intactos.
    '''
    ''' Uma ordenação que também filtrasse esconderia justamente o caso em que a
    ''' nota errou — e o dono não teria como descobrir que ela errou.
    ''' </summary>
    ''' <summary>Chave e dias de cada linha, para comparar CONJUNTOS iguais.</summary>
    Private Shared Function Retrato(vm As FilaViewModel) As List(Of String)
        Return vm.Minhas.Concat(vm.Deles).
               Select(Function(l) l.Conversa & "|" & l.Dias).
               OrderBy(Function(x) x, StringComparer.Ordinal).ToList()
    End Function

    <TestMethod>
    Public Async Function Ligar_a_prioridade_NAO_esconde_nem_apaga_os_dias() As Task
        Dim vm = ComPrioridade()
        Await vm.Atualizar()
        Dim antes = Retrato(vm)

        Await vm.Atualizar()
        vm.PorPrioridade = True

        Assert.IsTrue(antes.Count > 0, "controle: o cenário tinha de produzir linhas")
        ' AS MESMAS LINHAS, e nao "a mesma quantidade de dias": duas conversas
        ' com a mesma idade podiam ser trocadas uma pela outra sem ninguem ver.
        CollectionAssert.AreEqual(antes, Retrato(vm),
            "a ordenação escondeu, trocou ou mexeu nos dias de alguma linha")
        Assert.IsTrue(vm.Minhas.Concat(vm.Deles).All(Function(l) l.Espera.Length > 0),
                      "a coluna de espera sumiu de alguma linha")
    End Function

    ''' <summary>
    ''' A nota vem acompanhada da conta. Um número sozinho na tela é um palpite
    ''' com cara de conta, e o dono que discordar não terá do que discordar.
    ''' </summary>
    <TestMethod>
    Public Async Function Toda_linha_carrega_a_nota_E_a_conta() As Task
        Dim vm = ComPrioridade()
        Await vm.Atualizar()

        Assert.IsTrue(vm.Minhas.Count > 0)
        For Each l In vm.Minhas
            StringAssert.Contains(l.PorQue, "esperando")
            StringAssert.Contains(l.PorQue, "total:")
        Next
    End Function

    ''' <summary>
    ''' <b>A explicação bate com a ordem.</b> A mesma função que ordena é a que
    ''' explica — duas contas separadas divergiriam, e a divergência apareceria
    ''' como uma tela cuja explicação não corresponde à própria ordem.
    ''' </summary>
    <TestMethod>
    Public Async Function A_ordem_segue_a_nota_que_a_tela_mostra() As Task
        Dim vm = ComPrioridade()
        Await vm.Atualizar()
        vm.PorPrioridade = True

        Dim pontos = vm.Minhas.Select(Function(l) l.Pontos).ToList()
        CollectionAssert.AreEqual(pontos.OrderByDescending(Function(p) p).ToList(), pontos,
            "a ordem na tela não é a das notas que ela mostra")
    End Function

    ''' <summary>
    ''' O rótulo pesa: quem espera resposta sobe na frente de quem esperou mais
    ''' tempo sem esperar nada. Vinte dias de FYI contra dois dias de "espera
    ''' você" — e a segunda ganha, porque a diferença entre esperar e não esperar
    ''' é maior que a diferença entre dois dias e vinte.
    ''' </summary>
    <TestMethod>
    Public Async Function Quem_espera_resposta_sobe_na_frente_da_mais_velha() As Task
        Dim vm = ComPrioridade(
            Function(k) If(k.EntryId = "E-nova", "precisa_de_mim", "fyi"))

        Await vm.Atualizar()
        vm.PorPrioridade = True

        Assert.AreEqual(2, vm.Minhas.First().Dias,
            "a de dois dias que espera resposta não passou na frente da de vinte")
    End Function


    ''' <summary>
    ''' <b>Ligar a prioridade não relê o acervo.</b>
    '''
    ''' Chamar <c>Atualizar</c> ali montava a fila de novo, e entre a primeira
    ''' carga e o clique pode ter chegado mensagem, mudado o relógio ou entrado
    ''' uma dispensa: linhas apareciam, sumiam ou mudavam de idade. O botão dizia
    ''' "reordena" e trocava o conteúdo.
    '''
    ''' O leitor aqui devolve <b>coisas diferentes</b> a cada chamada, de
    ''' propósito: é o que faz o defeito aparecer.
    ''' </summary>
    <TestMethod>
    Public Async Function Ligar_a_prioridade_nao_vai_buscar_o_acervo_DE_NOVO() As Task
        Dim leituras = 0
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados)
                leituras += 1
                If leituras = 1 Then Return Duas()
                Return Montar({Msg("terceira", "mais@fora.com", 5)}, viuOsEnviados:=True)
            End Function,
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc)

        Await vm.Atualizar()
        Dim antes = Retrato(vm)

        vm.PorPrioridade = True

        Assert.AreEqual(1, leituras, "o botão de ordenar foi ao acervo de novo")
        CollectionAssert.AreEqual(antes, Retrato(vm))
    End Function

    ''' <summary>
    ''' <b>A nota que ordena é a mesma que a tela mostra — a MESMA avaliação.</b>
    '''
    ''' Ela era calculada duas vezes: uma na chave da ordenação e outra no
    ''' construtor da linha. É a mesma função em código, e não a mesma avaliação:
    ''' as fontes do rótulo e das regras são delegates, e nada promete que
    ''' devolvem o mesmo duas vezes.
    '''
    ''' O rótulo aqui alterna a cada chamada. Com duas avaliações, a linha é
    ''' ordenada com uma nota e mostra outra.
    ''' </summary>
    <TestMethod>
    Public Async Function A_nota_que_ORDENA_e_a_que_a_tela_MOSTRA() As Task
        ' SO A SEGUNDA CHAMADA diz "precisa_de_mim", e a escolha e do numero.
        '
        ' Alternar por paridade nao serviria: com duas avaliacoes a paridade de
        ' cada linha se repete e as contas coincidem por acidente. E a PRIMEIRA
        ' chamada tambem nao serve: ela cai na linha mais velha, que a ordem por
        ' idade ja poria na frente -- entao a ordem sabotada e a correta
        ' coincidem. A segunda cai na linha mais nova, e ai as duas ordens
        ' DIVERGEM, que e a unica situacao em que este teste tem o que provar.
        Dim chamadas = 0
        Dim vm = ComPrioridade(
            Function(k)
                chamadas += 1
                Return If(chamadas = 2, "precisa_de_mim", "fyi")
            End Function)

        Await vm.Atualizar()
        vm.PorPrioridade = True
        Dim linhas = vm.Minhas.Concat(vm.Deles).ToList()

        ' 1. Numero e explicacao vem do mesmo objeto.
        ' PONTOS MENOS OS DIAS, e nao os pontos: a parcela da espera sozinha
        ' chega a 20 numa linha de 20 dias, e ai a comparacao acusaria uma linha
        ' correta. O que se quer isolar e a parcela do rotulo.
        For Each l In linhas
            Dim tem = l.PorQue.Contains("alguém espera uma resposta sua")
            Assert.AreEqual(tem, (l.Pontos - l.Dias) >= PrioridadeDaFila.PorEsperarResposta,
                "a nota da ordem e a da explicação vieram de avaliações diferentes")
        Next

        ' 2. E A ORDEM SEGUE ESSE MESMO OBJETO.
        '
        ' A asserção 1 sozinha nao provava nada sobre a ORDEM: Pontos e PorQue
        ' sao dois derivados do mesmo valor, entao eles concordam mesmo que a
        ' chave da ordenacao tenha vindo de uma terceira avaliacao. Com o
        ' provedor mutavel, uma segunda avaliacao na chave produz uma lista que
        ' NAO esta em ordem decrescente dos pontos mostrados. Achado por revisao
        ' externa em 31/08/2026.
        Dim mostrados = linhas.Select(Function(l) l.Pontos).ToList()
        CollectionAssert.AreEqual(
            mostrados.OrderByDescending(Function(p) p).ToList(), mostrados,
            "a ordem veio de uma avaliação diferente da que a tela mostra")
    End Function

    ''' <summary>
    ''' <b>Empate completo não troca de lugar entre duas leituras.</b>
    '''
    ''' Nota, dias e assunto iguais deixavam a ordem por conta de como o acervo
    ''' enumerou — e ela troca. O último critério é a conversa, que é única.
    ''' </summary>
    <TestMethod>
    Public Async Function Empate_COMPLETO_nao_troca_de_lugar() As Task
        Dim invertido = False
        Dim iguais = {Msg("aaa", "um@fora.com", 5), Msg("bbb", "dois@fora.com", 5)}

        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados)
                Dim quais = If(invertido, iguais.Reverse().ToArray(), iguais)
                Return Montar(quais, viuOsEnviados:=True)
            End Function,
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc)

        Await vm.Atualizar()
        vm.PorPrioridade = True
        Dim antes = vm.Minhas.Concat(vm.Deles).Select(Function(l) l.Conversa).ToList()

        invertido = True
        Await vm.Atualizar()
        Dim depois = vm.Minhas.Concat(vm.Deles).Select(Function(l) l.Conversa).ToList()

        CollectionAssert.AreEqual(antes, depois,
            "duas linhas empatadas trocaram de lugar porque o acervo as enumerou " &
            "em outra ordem")
    End Function

    ''' <summary>
    ''' <b>As parcelas escritas na tela somam o total escrito na tela.</b>
    '''
    ''' "Conferível" quer dizer conferível com papel e caneta, sobre o que está
    ''' ali. Um arredondamento mais curto no total do que nas parcelas faria os
    ''' números não fecharem — hoje não aparece, porque os pesos são inteiros;
    ''' com o primeiro peso fracionário, apareceria.
    ''' </summary>
    <TestMethod>
    Public Async Function Os_numeros_ESCRITOS_fecham_a_conta() As Task
        Dim vm = ComPrioridade(Function(k) "precisa_de_mim")
        Await vm.Atualizar()
        vm.PorPrioridade = True

        For Each l In vm.Minhas.Concat(vm.Deles)
            Dim linhas = l.PorQue.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            Dim soma = linhas.Where(Function(x) Not x.StartsWith("total:")).
                       Sum(Function(x) Double.Parse(x.Substring(x.LastIndexOf(": ") + 2),
                                                    Globalization.CultureInfo.InvariantCulture))
            Dim total = Double.Parse(linhas.Last().Substring("total: ".Length),
                                     Globalization.CultureInfo.InvariantCulture)

            Assert.AreEqual(total, soma, 0.0001,
                "as parcelas escritas não somam o total escrito")
        Next
    End Function

End Class

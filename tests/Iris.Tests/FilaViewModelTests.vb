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
        Return FilaDeRespostas.Montar(mensagens,
                                      If(eu, New MinhasIdentidades({"ricardo@empresa.com"})),
                                      Agora, TimeZoneInfo.Utc, viuOsEnviados, Nothing)
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
    Public Sub Quando_nao_da_para_abrir_a_tela_DIZ()
        Dim vm As New FilaViewModel(
            Function(eu, agora, fuso, dispensadas, ignorados) _
                Montar(Array.Empty(Of MensagemNaFila)(), True),
            Nothing, Nothing, Function() Agora, TimeZoneInfo.Utc)

        vm.Atualizar()
        Dim antes = vm.Frase

        vm.NaoDeuParaAbrir()

        Assert.AreNotEqual(antes, vm.Frase, "clicou em Abrir e a tela nao mudou")
        StringAssert.Contains(vm.Frase, "pasta",
            "a frase precisa dizer o que fazer, e nao so que falhou")
    End Sub

End Class

Imports System.Collections.Generic
Imports System.Linq
Imports Iris.App.ViewModels
Imports Iris.Integration
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A BUSCA NA TELA.</b>
'''
''' ------------------------------------------------------------------
''' <c>BuscaNoAcervoTests</c> prova a busca. Este arquivo prova a <b>tela</b>,
''' e o que ele persegue é diferente: que a ressalva não desapareça quando a
''' busca acha, que "procurou e não achou" seja distinguível de "ainda não
''' procurou", e que falha de banco não vire "não achei".
'''
''' A última é a que importa mais. Confundir "não consegui olhar" com "olhei e
''' não tem" é a §23 na forma mais fácil de cometer.
''' </summary>
<TestClass>
Public Class BuscaViewModelTests

    ''' <summary>
    ''' Um resultado montado à mão, sem banco.
    '''
    ''' O <c>ResultadoDaBusca</c> tem construtor <c>Friend</c>, e a suíte está
    ''' no mesmo assembly de amigos — então dá para montar o caso sem abrir
    ''' SQLite. É o que mantém esta classe fora do <c>DoNotParallelize</c>.
    ''' </summary>
    Private Shared Function Resultado(quantos As Integer) As ResultadoDaBusca
        Dim achados = Enumerable.Range(1, quantos).
                      Select(Function(i) New AchadoDaBusca(1, "Caixa de Entrada",
                                                           Item($"assunto {i}", "Quem"), GrauDoAchado.Exato)).
                      ToList()
        Return New ResultadoDaBusca(New TermoDeBusca("assunto"), achados,
                                    {Consultada()}, Array.Empty(Of PastaConsultada)())
    End Function

    Private Shared Function Item(assunto As String, remetente As String) As ManifestItem
        Return New ManifestItem("E-1", assunto, remetente,
                                New DateTimeOffset(2026, 8, 20, 9, 0, 0, TimeSpan.Zero).ToString("o"),
                                False, Iris.Sync.PresenceState.Presente)
    End Function

    Private Shared Function Consultada() As PastaConsultada
        Return New PastaConsultada(1, "Caixa de Entrada", 1,
                                   Iris.Sync.FolderCoverage.Parcial, "2026-08-28T00:00:00Z",
                                   "Acervo parcial.", 5)
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>CONTROLE POSITIVO: a tela mostra o que a busca achou.</b>
    ''' </summary>
    <TestMethod>
    Public Sub Controle_os_achados_chegam_a_tela()
        Dim vm As New BuscaViewModel(Function(t) Resultado(3))
        vm.Termo = "assunto"
        vm.ProcurarCommand.Execute(Nothing)

        Assert.AreEqual(3, vm.Achados.Count)
        Assert.AreEqual("assunto 1", vm.Achados(0).Assunto)
        Assert.AreEqual("Caixa de Entrada", vm.Achados(0).Pasta)
        Assert.IsFalse(vm.SemAchados)
    End Sub

    ''' <summary>
    ''' <b>A ressalva aparece TAMBÉM quando a busca acha.</b>
    '''
    ''' Uma ressalva que só aparece no resultado vazio ensina o usuário a
    ''' lê-la como "não achei" — e ela não é isso. Ela diz onde se procurou,
    ''' com que alcance, e que o corpo da mensagem não é procurável. Isso vale
    ''' igual com dez achados na tela.
    ''' </summary>
    <TestMethod>
    Public Sub A_ressalva_nao_some_quando_a_busca_ACHA()
        Dim vm As New BuscaViewModel(Function(t) Resultado(3))
        vm.Termo = "assunto"
        vm.ProcurarCommand.Execute(Nothing)

        Assert.IsTrue(vm.Achados.Count > 0, "controle: achou")
        Assert.IsTrue(vm.TemRessalva, "a ressalva sumiu justamente onde ela nao pode sumir")
        StringAssert.Contains(vm.Ressalva, "corpo da mensagem")
    End Sub

    ''' <summary>
    ''' <b>"Ainda não procurei" e "procurei e não achei" são estados diferentes.</b>
    '''
    ''' Sem essa distinção, a tela abriria dizendo "nada encontrado" antes de
    ''' alguém ter digitado qualquer coisa — e o usuário concluiria alguma
    ''' coisa do silêncio. É o mesmo erro que a lista de mensagens já teve,
    ''' quando "selecione uma pasta" e "esta pasta está vazia" eram a mesma
    ''' frase.
    ''' </summary>
    <TestMethod>
    Public Sub Antes_de_procurar_NAO_e_o_mesmo_que_nao_achou()
        Dim vm As New BuscaViewModel(Function(t) Resultado(0))

        Assert.IsFalse(vm.Procurou, "ninguem procurou ainda")
        Assert.IsFalse(vm.SemAchados, "nao pode dizer 'nada encontrado' antes de procurar")
        Assert.IsFalse(vm.TemRessalva)

        vm.Termo = "coisa"
        vm.ProcurarCommand.Execute(Nothing)

        Assert.IsTrue(vm.Procurou)
        Assert.IsTrue(vm.SemAchados, "agora sim: procurou e nao achou")
    End Sub

    ''' <summary>
    ''' <b>Falha ao procurar NÃO vira "não achei".</b>
    '''
    ''' Banco travado, arquivo sumindo, janela fechando com a caixa em foco.
    ''' Nenhum desses diz nada sobre o que existe na caixa, e tratá-los como
    ''' resultado vazio seria afirmar ausência a partir de uma falha — a §23
    ''' na forma mais fácil de cometer.
    ''' </summary>
    <TestMethod>
    Public Sub Falha_ao_procurar_nao_afirma_ausencia()
        Dim vm As New BuscaViewModel(Function(t)
                                         Throw New InvalidOperationException("banco travado")
                                     End Function)
        vm.Termo = "coisa"
        vm.ProcurarCommand.Execute(Nothing)

        Assert.AreEqual(0, vm.Achados.Count)
        StringAssert.Contains(vm.Ressalva, "Não consegui procurar")
        StringAssert.Contains(vm.Ressalva, "não diz nada sobre o que existe")
        Assert.IsFalse(vm.Ressalva.Contains("Nada no acervo observado"),
            "falha nao pode usar a frase de 'procurei e nao achei'")
    End Sub

    ''' <summary>
    ''' <b>Limpar volta ao estado de antes de procurar.</b>
    '''
    ''' Inclusive o <c>Procurou</c>: limpar que deixasse "nada encontrado" na
    ''' tela seria limpar pela metade, e a metade que fica é justamente a
    ''' afirmação.
    ''' </summary>
    <TestMethod>
    Public Sub Limpar_volta_ao_estado_de_antes()
        Dim vm As New BuscaViewModel(Function(t) Resultado(2))
        vm.Termo = "assunto"
        vm.ProcurarCommand.Execute(Nothing)
        Assert.IsTrue(vm.Procurou, "controle: procurou")

        vm.LimparCommand.Execute(Nothing)

        Assert.AreEqual("", vm.Termo)
        Assert.AreEqual(0, vm.Achados.Count)
        Assert.IsFalse(vm.Procurou)
        Assert.IsFalse(vm.SemAchados)
        Assert.IsFalse(vm.TemRessalva)
    End Sub

    ''' <summary>
    ''' <b>Limpar só é oferecido depois de haver o que limpar.</b>
    ''' </summary>
    <TestMethod>
    Public Sub Limpar_so_habilita_depois_de_procurar()
        Dim vm As New BuscaViewModel(Function(t) Resultado(1))
        Assert.IsFalse(vm.LimparCommand.CanExecute(Nothing))

        vm.Termo = "assunto"
        vm.ProcurarCommand.Execute(Nothing)
        Assert.IsTrue(vm.LimparCommand.CanExecute(Nothing))
    End Sub

    ''' <summary>
    ''' <b>Mensagem suspeita chega à tela COM o aviso.</b>
    '''
    ''' É o que impede o resultado de parecer o estado corrente da caixa: uma
    ''' linha que o acervo tem e o Outlook já não confirma pode ter sido
    ''' apagada.
    ''' </summary>
    <TestMethod>
    Public Sub Mensagem_suspeita_chega_com_aviso()
        Dim suspeita As New ManifestItem("E-9", "assunto sumido", "Quem",
                                         New DateTimeOffset(2026, 8, 20, 9, 0, 0, TimeSpan.Zero).ToString("o"),
                                         False, Iris.Sync.PresenceState.Suspeito)
        Dim r As New ResultadoDaBusca(New TermoDeBusca("sumido"),
                                      {New AchadoDaBusca(1, "Caixa de Entrada", suspeita, GrauDoAchado.Exato)},
                                      {Consultada()}, Array.Empty(Of PastaConsultada)())

        Dim vm As New BuscaViewModel(Function(t) r)
        vm.Termo = "sumido"
        vm.ProcurarCommand.Execute(Nothing)

        Assert.IsTrue(vm.Achados(0).TemAviso)
        StringAssert.Contains(vm.Achados(0).Aviso, "apagada")
    End Sub

    ''' <summary>
    ''' <b>Assunto vazio não vira linha em branco.</b>
    '''
    ''' Uma linha sem texto nenhum na lista de achados é indistinguível de um
    ''' defeito de renderização.
    ''' </summary>
    <TestMethod>
    Public Sub Assunto_vazio_vira_texto_explicito()
        Dim r As New ResultadoDaBusca(New TermoDeBusca("quem"),
                                      {New AchadoDaBusca(1, "Caixa de Entrada", Item("", "Quem"), GrauDoAchado.Exato)},
                                      {Consultada()}, Array.Empty(Of PastaConsultada)())
        Dim vm As New BuscaViewModel(Function(t) r)
        vm.Termo = "quem"
        vm.ProcurarCommand.Execute(Nothing)

        Assert.AreEqual("(sem assunto)", vm.Achados(0).Assunto)
    End Sub

End Class

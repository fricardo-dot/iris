Imports System.Collections.Generic
Imports System.Linq
Imports Iris.App.ViewModels
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A CAIXA DIVIDIDA NA TELA — Fase 7.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTE ARQUIVO PRENDE</b>
'''
''' A legenda. Uma caixa dividida em gavetas bonitas dá a impressão de que
''' <i>tudo</i> foi olhado; se a varredura classificou quarenta de novecentas,
''' isso é falso, e o dono só descobriria abrindo a última gaveta.
'''
''' ------------------------------------------------------------------
''' <b>OS CONTROLES NEGATIVOS</b>
'''
''' <see cref="Esconder_as_vazias_NAO_esconde_a_conta_das_nao_classificadas"/> —
''' sem ele, o botão "esconder gavetas vazias" viraria um jeito de a tela
''' mentir.
'''
''' <see cref="Leitura_que_falha_nao_vira_caixa_LIMPA"/> — sem ele, um banco
''' fora do ar apareceria como "nada aqui", que é a única conclusão que esta
''' tela não pode produzir por engano.
''' </summary>
<TestClass>
Public Class CaixasViewModelTests

    Private Shared Function Mensagem(id As String) As MensagemNaFila
        Return New MensagemNaFila(
            New ItemKey(id, "store-1"), "conversa-" & id, "assunto " & id,
            "Alguém", "alguem@exemplo.com",
            New DateTimeOffset(2026, 8, 20, 12, 0, 0, TimeSpan.Zero))
    End Function

    Private Shared Function Montar(mensagens As IReadOnlyList(Of MensagemNaFila),
                                   rotulos As IReadOnlyDictionary(Of ItemKey, String),
                                   Optional regras As IReadOnlyList(Of String) = Nothing) _
                                   As CaixasViewModel
        Return New CaixasViewModel(
            Function() mensagens,
            Function() rotulos,
            Function() New Dictionary(Of ItemKey, IReadOnlyList(Of String))(),
            Function() If(regras, CType(Array.Empty(Of String)(), IReadOnlyList(Of String))))
    End Function

    Private Shared Function Rotulos(ParamArray pares As String()) _
                                    As IReadOnlyDictionary(Of ItemKey, String)
        Dim mapa As New Dictionary(Of ItemKey, String)()
        For i = 0 To pares.Length - 1 Step 2
            mapa(New ItemKey(pares(i), "store-1")) = pares(i + 1)
        Next
        Return mapa
    End Function

    ' ==================================================================
    ' A LEGENDA

    ''' <summary>
    ''' <b>A frase dá o número, sempre.</b> Duas de três classificadas não é
    ''' "classificado": é dois terços, e o dono precisa saber disso antes de
    ''' concluir qualquer coisa da tela.
    ''' </summary>
    <TestMethod>
    Public Sub A_legenda_conta_as_que_faltam()
        Dim vm = Montar({Mensagem("a"), Mensagem("b"), Mensagem("c")},
                        Rotulos("a", "fyi", "b", "promocao"))
        vm.Atualizar()

        StringAssert.Contains(vm.Cobertura, "1 de 3")
    End Sub

    <TestMethod>
    Public Sub Tudo_classificado_diz_que_esta_tudo_classificado()
        Dim vm = Montar({Mensagem("a")}, Rotulos("a", "fyi"))
        vm.Atualizar()

        StringAssert.Contains(vm.Cobertura, "1 de 1")
    End Sub

    ''' <summary>
    ''' Nada classificado tem frase própria. "0 de 900 classificadas" e "nenhuma
    ''' das 900 foi classificada" dizem o mesmo, e a segunda é a que alguém lê
    ''' de manhã sem interpretar.
    ''' </summary>
    <TestMethod>
    Public Sub Nada_classificado_tem_frase_propria()
        Dim vm = Montar({Mensagem("a"), Mensagem("b")}, Rotulos())
        vm.Atualizar()

        StringAssert.Contains(vm.Cobertura, "Nenhuma das 2")
    End Sub

    ''' <summary>
    ''' <b>Caixa sem mensagem nenhuma não é "tudo classificado".</b> É caixa não
    ''' varrida, ou vazia — e dizer "0 de 0 classificadas" seria uma afirmação
    ''' de completude sobre coisa nenhuma.
    ''' </summary>
    <TestMethod>
    Public Sub Caixa_vazia_nao_diz_que_esta_tudo_classificado()
        Dim vm = Montar(Array.Empty(Of MensagemNaFila)(), Rotulos())
        vm.Atualizar()

        Assert.IsFalse(vm.Cobertura.Contains("classificadas."))
        StringAssert.Contains(vm.Cobertura, "Varra")
    End Sub

    ' ==================================================================
    ' AS GAVETAS

    <TestMethod>
    Public Sub As_gavetas_vazias_aparecem_por_padrao()
        Dim vm = Montar({Mensagem("a")}, Rotulos("a", "fyi"))
        vm.Atualizar()

        Assert.IsTrue(vm.Gavetas.Any(Function(g) g.Vazia))
        Assert.AreEqual(7, vm.Gavetas.Count)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo.</b> Esconder as gavetas vazias é escolha do dono
    ''' — mas a legenda continua contando as não classificadas, senão esconder
    ''' viraria um jeito de a tela mentir.
    ''' </summary>
    <TestMethod>
    Public Sub Esconder_as_vazias_NAO_esconde_a_conta_das_nao_classificadas()
        Dim vm = Montar({Mensagem("a"), Mensagem("b")}, Rotulos("a", "fyi"))
        vm.MostrarVazias = False

        Assert.AreEqual(2, vm.Gavetas.Count)
        StringAssert.Contains(vm.Cobertura, "1 de 2")
    End Sub

    <TestMethod>
    Public Sub A_gaveta_do_dono_se_identifica_na_tela()
        Dim vm = Montar({Mensagem("a")}, Rotulos("a", "fyi"), {"clientes reclamando"})
        vm.Atualizar()

        Assert.IsTrue(vm.Gavetas.First(Function(g) g.Nome = "clientes reclamando").DoDono)
    End Sub

    <TestMethod>
    Public Sub A_mensagem_chega_na_gaveta_com_assunto_e_quem()
        Dim vm = Montar({Mensagem("a")}, Rotulos("a", "fyi"))
        vm.Atualizar()

        Dim gaveta = vm.Gavetas.First(Function(g) g.Nome = "Só para saber")
        Assert.AreEqual("assunto a", gaveta.Mensagens.Single().Assunto)
        Assert.AreEqual("Alguém", gaveta.Mensagens.Single().Quem)
    End Sub

    ' ==================================================================
    ' QUANDO DÁ ERRADO

    ''' <summary>
    ''' <b>O outro controle negativo.</b> Um banco fora do ar não pode virar
    ''' "nada aqui" — essa é a única conclusão que esta tela não pode produzir
    ''' por engano, e sem gaveta nenhuma o dono não tem o que confundir com
    ''' resultado.
    ''' </summary>
    <TestMethod>
    Public Sub Leitura_que_falha_nao_vira_caixa_LIMPA()
        ' THROW NAO E EXPRESSAO em lambda de uma linha -- armadilha do
        ' CLAUDE.md. Precisa da lambda de varias linhas.
        Dim vm As New CaixasViewModel(
            Function() As IReadOnlyList(Of MensagemNaFila)
                Throw New InvalidOperationException("banco fora do ar")
            End Function,
            Function() Nothing, Function() Nothing, Function() Nothing)

        vm.Atualizar()

        Assert.AreEqual(0, vm.Gavetas.Count)
        StringAssert.Contains(vm.Cobertura, "Não deu para ler")
    End Sub

    <TestMethod>
    Public Sub Abrir_sem_quem_abra_nao_estoura()
        Dim vm = Montar({Mensagem("a")}, Rotulos("a", "fyi"))
        vm.Atualizar()

        Dim linha = vm.Gavetas.First(Function(g) g.Nome = "Só para saber").Mensagens.Single()
        Assert.IsFalse(linha.AbrirCommand.CanExecute(Nothing))
    End Sub

    <TestMethod>
    Public Sub Abrir_leva_a_chave_da_mensagem_certa()
        Dim aberta As ItemKey = Nothing
        Dim vm As New CaixasViewModel(
            Function() CType({Mensagem("a")}, IReadOnlyList(Of MensagemNaFila)),
            Function() Rotulos("a", "fyi"),
            Function() New Dictionary(Of ItemKey, IReadOnlyList(Of String))(),
            Function() CType(Array.Empty(Of String)(), IReadOnlyList(Of String)),
            Sub(k) aberta = k)

        vm.Atualizar()
        vm.Gavetas.First(Function(g) g.Nome = "Só para saber").
           Mensagens.Single().AbrirCommand.Execute(Nothing)

        Assert.AreEqual("a", aberta.EntryId)
    End Sub

End Class

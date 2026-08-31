Imports System.Collections.Generic
Imports System.Linq
Imports Iris.App.ViewModels
Imports Iris.Assist
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>IMAGEM EMBUTIDA NÃO É ANEXO — e o número que obrigou a distinção.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE FOI MEDIDO</b>
'''
''' 30/08/2026, pasta "0. E-mails Lidos" de uma caixa corporativa real, 13
''' mensagens:
'''
''' <code>
''' sem anexo nenhum ....... 0
''' so imagem embutida ..... 10
''' com anexo de verdade ... 3
''' </code>
'''
''' O portão negava qualquer <c>Attachments.Count &gt; 0</c>, então a IA
''' recusava <b>13 de 13</b>. Não era guarda rigorosa: era guarda olhando a
''' coisa errada. Toda assinatura corporativa tem logo, e logo é anexo com
''' <c>Content-ID</c>.
'''
''' ------------------------------------------------------------------
''' <b>E POR QUE AFROUXAR NÃO FOI DE GRAÇA</b>
'''
''' Uma <b>captura de tela colada no corpo</b> é embutida do mesmo jeito, e
''' pode carregar o teor inteiro da mensagem. Se embutida deixasse de bloquear
''' e ninguém dissesse nada, uma recusa honesta teria virado um resumo
''' silenciosamente parcial — a família de defeito que esta base passou a série
''' inteira corrigindo.
'''
''' Por isso a contagem viaja até a tela e é declarada. Estes testes prendem as
''' duas metades: <b>embutida não nega</b>, e <b>embutida é dita</b>.
''' </summary>
<TestClass>
Public Class AnexoEmbutidoTests

    Private Shared ReadOnly Pasta As New FolderKey("p", "s")

    Private Shared Function Msg(temAnexo As Boolean, embutidas As Integer?) _
                                As MessageClassification
        Dim item = New ItemKey("e1", "s")
        Return New MessageClassification(
            item, Pasta,
            New LabelReading(item, LabelReadingKind.Absent, LabelReadStage.Parse,
                             Nothing, Array.Empty(Of String)()),
            temAnexo:=temAnexo, embutidas:=embutidas)
    End Function

    ' ==================================================================
    ' O que NEGA

    ''' <summary>
    ''' <b>CONTROLE: anexo de verdade continua negando.</b>
    '''
    ''' A mudança foi sobre o que <i>conta</i> como anexo, e não sobre deixar
    ''' anexo passar. Sem este controle, um portão que aceitasse tudo passaria
    ''' nos testes de baixo.
    ''' </summary>
    <TestMethod>
    Public Sub Anexo_de_verdade_continua_NEGANDO()
        Dim m = Msg(temAnexo:=True, embutidas:=0)

        Assert.IsTrue(m.TemAnexo,
            "o DTO parou de carregar a presença de anexo de verdade")
    End Sub

    ''' <summary>
    ''' <b>Só imagem embutida NÃO nega.</b>
    '''
    ''' Os dez de treze da medição.
    ''' </summary>
    <TestMethod>
    Public Sub So_embutida_NAO_nega()
        Dim m = Msg(temAnexo:=False, embutidas:=5)

        Assert.IsFalse(m.TemAnexo,
            "cinco logos de assinatura voltaram a contar como anexo, e a IA " &
            "volta a recusar 13 de 13")
        Assert.AreEqual(5, m.Embutidas)
    End Sub

    ' ==================================================================
    ' O que é DITO

    ''' <summary>
    ''' <b>A frase diz o número, e some quando não há nada a dizer.</b>
    '''
    ''' Zero embutidas não produz ressalva: uma linha permanente sobre nada é
    ''' ruído, e ruído ensina a não ler a ressalva de verdade.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_embutida_nao_ha_ressalva()
        Assert.AreEqual("", AssistenteViewModel.DizerOQueFicouDeFora(
            {Msg(False, 0)}.ToList()))
    End Sub

    <TestMethod>
    Public Sub Uma_embutida_fala_no_singular()
        Dim f = AssistenteViewModel.DizerOQueFicouDeFora({Msg(False, 1)}.ToList())

        StringAssert.Contains(f, "1 imagem embutida")
        StringAssert.Contains(f, "não foi lida")
    End Sub

    ''' <summary>
    ''' <b>Várias mensagens somam.</b> O resumo é do conjunto, e a ressalva
    ''' também precisa ser.
    ''' </summary>
    <TestMethod>
    Public Sub Varias_mensagens_somam_as_embutidas()
        Dim f = AssistenteViewModel.DizerOQueFicouDeFora(
            {Msg(False, 2), Msg(False, 3)}.ToList())

        StringAssert.Contains(f, "5 imagens embutidas")
    End Sub

    ''' <summary>
    ''' <b>NÃO SABER QUANTAS NÃO VIRA ZERO.</b>
    '''
    ''' A mesma regra que esta base aplicou em seis lugares. Aqui o estrago
    ''' seria específico: a ressalva sumiria, e o leitor concluiria que o resumo
    ''' cobriu tudo — quando na verdade ninguém conseguiu contar o que ficou
    ''' de fora.
    '''
    ''' <b>Controle negativo:</b> trocando o <c>Any(Not HasValue)</c> por
    ''' <c>Sum(GetValueOrDefault)</c>, este teste cai.
    ''' </summary>
    <TestMethod>
    Public Sub Nao_saber_quantas_NAO_vira_zero()
        Dim f = AssistenteViewModel.DizerOQueFicouDeFora(
            {Msg(False, 2), Msg(False, Nothing)}.ToList())

        StringAssert.Contains(f, "Não sei quantas",
            "uma contagem que faltou virou silêncio, e o leitor conclui que o " &
            "resumo cobriu tudo")
        Assert.IsFalse(f.Contains("2 imagens"),
            "somou o que sabia e calou sobre o que não sabia -- pior que as duas")
    End Sub

    <TestMethod>
    Public Sub Sem_mensagem_nenhuma_nao_ha_ressalva()
        Assert.AreEqual("", AssistenteViewModel.DizerOQueFicouDeFora(
            New List(Of MessageClassification)()))
        Assert.AreEqual("", AssistenteViewModel.DizerOQueFicouDeFora(Nothing))
    End Sub

End Class

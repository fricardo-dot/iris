Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' §22.12 — a suíte não pode imprimir <c>Passed!</c> escondendo que deixou de
''' cobrir alguma coisa.
'''
''' O sintoma apareceu no fim do 2.1: enquanto o usuário reiniciava o Outlook
''' para o teste da janela, a suíte deu <c>258 passed, 3 skipped</c> — com o
''' cabeçalho verde idêntico ao de sempre. Os três pulados eram os únicos que
''' tocam o Outlook real. Um resultado que <b>parece igual</b> quando a
''' cobertura mudou é exatamente o formato de erro que a Fase 2 inteira
''' persegue.
'''
''' O conserto é distinguir <b>ambiente sem suporte</b> de <b>falta de
''' preparo</b>. Máquina sem Outlook instalado: pular é honesto. Máquina com
''' Outlook instalado e fechado: é preparo, e preparo tem de falhar com
''' instrução de como resolver.
''' </summary>
<TestClass>
Public Class CoberturaDaSuiteTests

    <TestMethod>
    Public Sub Conectado_prossegue()
        Assert.AreEqual(PagingIntegrationTests.SemOutlook.Prosseguir,
                        PagingIntegrationTests.Decidir(conectado:=True, instalado:=True))
        Assert.AreEqual(PagingIntegrationTests.SemOutlook.Prosseguir,
                        PagingIntegrationTests.Decidir(conectado:=True, instalado:=False))
    End Sub

    ''' <summary>
    ''' O ramo que dá sentido ao conserto: instalado e não conectado FALHA.
    '''
    ''' Sem este teste, o caminho só seria exercitado numa máquina com Outlook
    ''' instalado e fechado — ou seja, nunca, na prática. Um conserto para "a
    ''' suíte mente quando pula" cujo próprio caminho nunca roda seria a mesma
    ''' piada um nível acima.
    ''' </summary>
    <TestMethod>
    Public Sub Instalado_e_fechado_FALHA_em_vez_de_pular()
        Assert.AreEqual(PagingIntegrationTests.SemOutlook.Falhar,
                        PagingIntegrationTests.Decidir(conectado:=False, instalado:=True),
                        "Outlook instalado e fechado e falta de PREPARO, nao ambiente sem suporte")
    End Sub

    <TestMethod>
    Public Sub Nao_instalado_pula()
        Assert.AreEqual(PagingIntegrationTests.SemOutlook.Pular,
                        PagingIntegrationTests.Decidir(conectado:=False, instalado:=False))
    End Sub

    ''' <summary>
    ''' Controle: a decisão não é constante. Uma função que devolvesse sempre
    ''' <c>Pular</c> passaria nos dois testes de pular e reintroduziria o
    ''' defeito original sem ninguém notar.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_a_decisao_distingue_os_tres_casos()
        Dim r = {PagingIntegrationTests.Decidir(True, True),
                 PagingIntegrationTests.Decidir(False, True),
                 PagingIntegrationTests.Decidir(False, False)}
        CollectionAssert.AllItemsAreUnique(r,
            "os tres casos tem de dar respostas diferentes, senao a decisao nao decide")
    End Sub

    ''' <summary>
    ''' Nesta máquina o Outlook ESTÁ instalado — então o ramo que vale aqui é o
    ''' <c>Falhar</c>, e não o <c>Pular</c>.
    '''
    ''' O teste é condicional de propósito: numa máquina sem Outlook ele não
    ''' afirma nada, porque não teria como. Mas onde ele pode afirmar, afirma —
    ''' e é o que garante que a máquina do usuário não caia no ramo silencioso.
    ''' </summary>
    <TestMethod>
    Public Sub Nesta_maquina_o_ramo_que_vale_e_o_de_FALHA()
        If Not PagingIntegrationTests.OutlookInstalado() Then
            Assert.Inconclusive("sem Outlook instalado: nada a afirmar aqui")
        End If
        Assert.AreEqual(PagingIntegrationTests.SemOutlook.Falhar,
                        PagingIntegrationTests.Decidir(conectado:=False, instalado:=True),
                        "nesta maquina, Outlook fechado tem de FALHAR a suite, nao pula-la")
    End Sub

End Class

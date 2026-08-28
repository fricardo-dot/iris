Imports Iris.Core
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>QUAIS PASTAS A ÁRVORE MOSTRA.</b>
'''
''' ------------------------------------------------------------------
''' <b>ESTA CLASSE NUNCA TINHA SIDO TESTADA, E ESCONDEU UMA FASE INTEIRA</b>
'''
''' A <c>FolderVisibilityPolicy</c> é uma classe de quatro linhas úteis, e
''' passou da Fase 1 à Fase 6 sem um único teste. Em 28/08/2026 ela escondeu
''' a Fase 6 inteira: o calendário era lido, os testes contra o Outlook real
''' passavam, e a pasta <b>não aparecia na árvore</b> — porque a política
''' ainda dizia "só correio".
'''
''' Ninguém percebeu porque os testes do calendário achavam a pasta pelo
''' broker, contornando a árvore. Prova de leitura não é prova de alcance.
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTES TESTES GUARDAM</b>
'''
''' A regra é "não ofereça porta que não abre". Então a lista de tipos
''' visíveis tem de crescer <b>junto com a tela</b>, e não junto com a
''' leitura — e cada tipo de fora precisa estar de fora por decisão, não por
''' esquecimento.
''' </summary>
<TestClass>
Public Class VisibilidadeDePastasTests

    Private Shared Function Pasta(tipo As FolderContentKind,
                                  Optional oculta As Boolean = False) As FolderInfo
        Return New FolderInfo With {
            .Key = New FolderKey("e-1", "s-1"),
            .Name = tipo.ToString(),
            .ContentKind = tipo,
            .IsHidden = oculta}
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>Correio aparece.</b> Controle positivo: uma política que esconde
    ''' tudo passaria em todos os testes de "não mostra o que não abre".
    ''' </summary>
    <TestMethod>
    Public Sub Correio_aparece()
        Assert.IsTrue(New FolderVisibilityPolicy().IsVisible(Pasta(FolderContentKind.Mail)))
    End Sub

    ''' <summary>
    ''' <b>Calendário aparece — e é o teste que não existia.</b>
    '''
    ''' A Fase 6 entregou a leitura e a faixa da agenda. Sem esta linha, a
    ''' política podia voltar a "só correio" e a agenda ficaria inalcançável de
    ''' novo, com toda a suíte verde.
    ''' </summary>
    <TestMethod>
    Public Sub Calendario_aparece()
        Assert.IsTrue(New FolderVisibilityPolicy().IsVisible(Pasta(FolderContentKind.Calendar)),
            "a pasta de calendario voltou a ficar invisivel, e a agenda com ela")
    End Sub

    ''' <summary>
    ''' <b>Contatos, Tarefas, Observações e Diário NÃO aparecem.</b>
    '''
    ''' O contraponto, e ele é o que dá sentido ao teste acima. Sem ele, a
    ''' política poderia passar a mostrar tudo e os dois primeiros testes
    ''' continuariam verdes — oferecendo quatro portas que não abrem.
    '''
    ''' Cada um destes sai da lista no dia em que houver tela, e não no dia em
    ''' que houver leitura.
    ''' </summary>
    <TestMethod>
    Public Sub O_que_nao_tem_tela_NAO_aparece()
        Dim p As New FolderVisibilityPolicy()
        For Each tipo In {FolderContentKind.Contacts, FolderContentKind.Tasks,
                          FolderContentKind.Notes, FolderContentKind.Journal,
                          FolderContentKind.Unknown}
            Assert.IsFalse(p.IsVisible(Pasta(tipo)),
                $"{tipo} apareceu na arvore, e nao ha tela para ela: e uma porta que nao abre")
        Next
    End Sub

    ''' <summary>
    ''' <b>Pasta oculta do Outlook continua oculta</b>, mesmo sendo de correio.
    '''
    ''' <c>PR_ATTR_HIDDEN</c> marca coisas como "Conversation Action Settings".
    ''' O próprio Outlook não as mostra.
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_oculta_continua_oculta()
        Dim p As New FolderVisibilityPolicy()
        Assert.IsFalse(p.IsVisible(Pasta(FolderContentKind.Mail, oculta:=True)))

        ' E o diagnóstico continua podendo vê-las.
        p.IncludeHidden = True
        Assert.IsTrue(p.IsVisible(Pasta(FolderContentKind.Mail, oculta:=True)))
    End Sub

    ''' <summary>
    ''' <b>Desligar o filtro mostra tudo.</b>
    '''
    ''' É a outra metade do controle: prova que os testes acima medem o
    ''' <i>filtro</i>, e não uma política que rejeita por outro motivo.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_o_filtro_tudo_aparece()
        Dim p As New FolderVisibilityPolicy() With {.SoOQueAbre = False}
        For Each tipo In {FolderContentKind.Contacts, FolderContentKind.Tasks,
                          FolderContentKind.Journal}
            Assert.IsTrue(p.IsVisible(Pasta(tipo)),
                $"com o filtro desligado, {tipo} tinha de aparecer")
        Next
    End Sub

    ''' <summary>Pasta nula não aparece, e não explode.</summary>
    <TestMethod>
    Public Sub Pasta_nula_nao_aparece()
        Assert.IsFalse(New FolderVisibilityPolicy().IsVisible(Nothing))
    End Sub

End Class

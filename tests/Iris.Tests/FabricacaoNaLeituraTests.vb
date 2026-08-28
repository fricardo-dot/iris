Imports System.Linq
Imports System.Threading
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>QUANDO A LEITURA FABRICA, ELA CONTA.</b>
'''
''' ------------------------------------------------------------------
''' <b>O DEFEITO, E POR QUE ELE NÃO VIROU MIGRAÇÃO</b>
'''
''' A revisão externa de 28/08/2026 apontou que o caminho de paginação
''' transforma <b>ausência em fato</b>: célula nula vira <c>0</c> para
''' tamanho, <c>False</c> para não-lida e para anexo, e texto vazio. É a
''' mesma família do <c>message_class</c>, que era constante fabricada — e
''' este é pior, porque <c>False</c> em "tem anexo" é uma afirmação que o
''' usuário lê como fato.
'''
''' <b>Antes de corrigir, medi.</b> <c>tools/medir-nulos-da-table.ps1</c>
''' contou <b>zero</b> nulos nas oito colunas, em 1.109 linhas da Caixa de
''' Entrada real. O defeito existe no contrato e não se manifesta nesta
''' caixa.
'''
''' Tornar os campos anuláveis até a tela seria migração de esquema mais uma
''' forma de mostrar "não sei" sem parecer "não" — desproporcional para zero
''' ocorrências. <b>Mas o silêncio não precisava continuar.</b> O contador
''' custa nada e faz a fabricação aparecer no dia em que acontecer, no mesmo
''' lugar onde o descarte já aparece.
'''
''' É a regra que a varredura já segue: <i>recusa declarada é mais forte que
''' recusa silenciosa</i>.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ESTES TESTES NÃO PRECISAM DO OUTLOOK</b>
'''
''' As conversões são lógica pura sobre <c>Object</c>. Exercitá-las por um
''' Outlook aberto exigiria uma caixa com célula nula dentro — que é
''' justamente o que a medição mostrou não existir aqui. Um contador que só
''' pode ser exercitado por um estado que ninguém consegue produzir é um
''' contador sem prova.
''' </summary>
<TestClass>
Public Class FabricacaoNaLeituraTests

    ' ==================================================================

    ''' <summary>
    ''' <b>CONTROLE POSITIVO: valor presente não conta como fabricação.</b>
    '''
    ''' Vem primeiro. Um contador que incrementa sempre marcaria toda leitura
    ''' como suspeita, e o número vira ruído que se aprende a ignorar — que é
    ''' o mesmo destino de um aviso que aparece o tempo todo.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_valor_presente_NAO_conta()
        Dim f As New MessagePaging.TableRowSource()

        Assert.AreEqual("assunto", f.ComoTexto("assunto"))
        Assert.AreEqual(1234, f.ComoInteiro(1234))
        Assert.IsTrue(f.ComoBooleano(True))
        Assert.IsFalse(f.ComoBooleano(False))

        Assert.AreEqual(0, f.Fabricadas,
            "leitura limpa marcou fabricacao: o contador vira ruido")
    End Sub

    ''' <summary>
    ''' <b>Texto ausente vira vazio, e conta.</b>
    '''
    ''' Vazio e ausente são coisas diferentes: um assunto em branco é uma
    ''' mensagem sem assunto; um <c>Nothing</c> é o provedor não tendo
    ''' respondido. Colapsar os dois é o que o contador denuncia.
    '''
    ''' Note que <b>string vazia de verdade não conta</b> — ela é um valor.
    ''' </summary>
    <TestMethod>
    Public Sub Texto_ausente_conta_e_texto_vazio_NAO()
        Dim f As New MessagePaging.TableRowSource()

        Assert.AreEqual("", f.ComoTexto(Nothing))
        Assert.AreEqual(1, f.Fabricadas, "texto ausente tinha de contar")

        Assert.AreEqual("", f.ComoTexto(""))
        Assert.AreEqual(1, f.Fabricadas,
            "string vazia e um VALOR, e nao ausencia: nao pode contar")
    End Sub

    ''' <summary>
    ''' <b>Número ausente ou ilegível vira <c>0</c>, e conta.</b>
    '''
    ''' Os dois casos, porque são falhas diferentes com o mesmo resultado:
    ''' a coluna não veio, ou veio com algo que não converte.
    ''' </summary>
    <TestMethod>
    Public Sub Numero_ausente_e_ilegivel_contam()
        Dim f As New MessagePaging.TableRowSource()

        Assert.AreEqual(0, f.ComoInteiro(Nothing))
        Assert.AreEqual(1, f.Fabricadas)

        Assert.AreEqual(0, f.ComoInteiro("isto nao e numero"))
        Assert.AreEqual(2, f.Fabricadas, "conversao que falha tambem fabrica")
    End Sub

    ''' <summary>
    ''' <b>Booleano ausente vira <c>False</c>, e conta — e este é o pior.</b>
    '''
    ''' <c>False</c> em "não lida" e em "tem anexo" são afirmações que o
    ''' usuário lê como fato. O <c>MailSummary</c> já não tem
    ''' <c>IsProtected</c> justamente para não afirmar "não é protegida" sem
    ''' ter medido; aqui a mesma afirmação escapava por dentro da conversão.
    ''' </summary>
    <TestMethod>
    Public Sub Booleano_ausente_conta()
        Dim f As New MessagePaging.TableRowSource()

        Assert.IsFalse(f.ComoBooleano(Nothing))
        Assert.AreEqual(1, f.Fabricadas)

        Assert.IsFalse(f.ComoBooleano("nem isto"))
        Assert.AreEqual(2, f.Fabricadas)
    End Sub

    ''' <summary>
    ''' <b>A chave ausente também conta.</b>
    '''
    ''' Ela tem tratamento próprio mais acima — chave vazia derruba a pasta
    ''' inteira para o caminho lento —, mas a conversão continua sendo o lugar
    ''' onde a ausência acontece, e o número tem de sair de lá.
    ''' </summary>
    <TestMethod>
    Public Sub Chave_ausente_conta_e_bytes_NAO()
        Dim f As New MessagePaging.TableRowSource()

        Assert.AreEqual("", f.ComoEntryId(Nothing))
        Assert.AreEqual(1, f.Fabricadas)

        ' O caminho normal: PT_BINARY vira hex MAIUSCULO, e nao conta.
        Assert.AreEqual("0A1B", f.ComoEntryId(New Byte() {&HA, &H1B}))
        Assert.AreEqual(1, f.Fabricadas, "bytes legiveis nao sao fabricacao")
    End Sub

    ''' <summary>
    ''' <b>O contador soma as fabricações de uma página.</b>
    '''
    ''' Ele acumula dentro de uma página. Quem recomeça é <c>Zerar</c>, e o
    ''' teste abaixo cobra isso — este aqui só prova a soma.
    ''' </summary>
    <TestMethod>
    Public Sub O_contador_soma_dentro_da_pagina()
        Dim f As New MessagePaging.TableRowSource()

        f.ComoTexto(Nothing)
        f.ComoInteiro(Nothing)
        f.ComoBooleano(Nothing)
        f.ComoBooleano(Nothing)

        Assert.AreEqual(4, f.Fabricadas)
    End Sub

    ''' <summary>
    ''' <b>O contador é de PÁGINA, e a página pode custar vários lotes.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ZERAR NO LUGAR ERRADO SUBCONTAVA</b>
    '''
    ''' A primeira versão zerava no início de <c>Ler</c>, e parecia certo. O
    ''' <c>CursorPaging</c> chama <c>Ler</c> <b>várias vezes</b> por página, para
    ''' drenar o grupo do último instante — então o DTO recebia só o último lote,
    ''' e uma fabricação no primeiro sumia.
    '''
    ''' Quem zera agora é <c>Zerar</c>, chamado uma vez por página por quem monta
    ''' a página. Este teste cobra as duas metades: que acumula sem <c>Zerar</c>,
    ''' e que <c>Zerar</c> recomeça.
    '''
    ''' <b>O que ele NÃO cobre, e é preciso dizer:</b> o caminho inteiro
    ''' <c>ReadPage → LerPorTabela → MessagePage.FabricatedCells</c>. Isso pede
    ''' uma <c>Table</c> de verdade com célula nula dentro, e a medição de 28/08
    ''' mostrou que esta caixa não tem nenhuma. O que se prova aqui é o contrato
    ''' do contador; a ligação até o DTO é sustentada por leitura.
    ''' </summary>
    <TestMethod>
    Public Sub O_contador_acumula_entre_lotes_e_Zerar_recomeca()
        Dim f As New MessagePaging.TableRowSource()

        ' Primeiro lote.
        f.ComoTexto(Nothing)
        f.ComoInteiro(Nothing)
        Assert.AreEqual(2, f.Fabricadas)

        ' Segundo lote da MESMA pagina: acumula, e nao recomeca.
        f.ComoBooleano(Nothing)
        Assert.AreEqual(3, f.Fabricadas,
            "o contador recomecou entre lotes: a pagina perderia as fabricacoes " &
            "de todos os lotes menos o ultimo")

        ' Pagina nova.
        f.Zerar()
        Assert.AreEqual(0, f.Fabricadas, "Zerar tinha de recomecar a contagem")

        f.ComoTexto(Nothing)
        Assert.AreEqual(1, f.Fabricadas)
    End Sub

End Class

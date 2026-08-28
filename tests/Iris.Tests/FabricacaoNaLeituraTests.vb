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
    ''' <b>O acumulador soma enquanto uma linha é convertida.</b>
    '''
    ''' O comentário aqui ainda mandava o leitor a um <c>Zerar()</c> que foi
    ''' removido junto com o contador de página — a revisão externa pegou. Quem
    ''' recomeça hoje é o <c>ConverterLinhas</c>, uma vez por linha, e é o teste
    ''' <c>Cada_linha_leva_o_SEU_numero</c> que cobra isso na produção.
    ''' Este aqui só prova que a soma acontece.
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
    ''' <b>A PÁGINA SOMA SÓ AS LINHAS QUE ENTRARAM NELA.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ESTE CONTADOR ERROU NAS DUAS DIREÇÕES, EM DIAS SEGUIDOS</b>
    '''
    ''' Ele começou na <i>fonte</i>, zerado a cada <c>Ler</c> — e <b>subcontava</b>,
    ''' porque o <c>CursorPaging</c> chama <c>Ler</c> várias vezes por página e o
    ''' DTO recebia só o último lote.
    '''
    ''' Eu movi o reset para uma vez por página, e aí ele <b>sobrecontava</b>: o
    ''' <c>CursorPaging</c> lê um lote inteiro e para na primeira linha de outro
    ''' instante, então as linhas de <i>read-ahead</i> — que não entram nesta
    ''' página — já tinham sido convertidas e contadas. Na página seguinte elas
    ''' seriam contadas de novo.
    '''
    ''' Cada conserto de um lado abria o outro, porque o número morava no lugar
    ''' errado: entre <i>quem converte</i> e <i>quem escolhe o que entra</i>.
    '''
    ''' <b>Agora ele mora na LINHA.</b> Quem entra na página leva o seu número
    ''' junto, e a página soma o que recebeu. Não há o que errar — e é por isso
    ''' que este teste é sobre a soma, e não sobre o reset.
    ''' </summary>
    <TestMethod>
    Public Sub A_pagina_soma_so_as_linhas_que_recebeu()
        Dim entraram = {
            New MessagePaging.TableRow With {.EntryId = "E-1", .Fabricadas = 2},
            New MessagePaging.TableRow With {.EntryId = "E-2", .Fabricadas = 0},
            New MessagePaging.TableRow With {.EntryId = "E-3", .Fabricadas = 1}
        }

        Assert.AreEqual(3, MessagePaging.Fabricadas(entraram),
            "a pagina tem de somar as fabricacoes das linhas que recebeu")

        ' A linha de read-ahead NAO esta na lista, entao nao entra na conta --
        ' e sera contada na pagina dela, uma vez so.
        Assert.AreEqual(2, MessagePaging.Fabricadas({entraram(0)}),
            "somou linha que nao foi passada")
    End Sub

    ''' <summary>
    ''' <b>Lista vazia e lista nula somam zero, e não explodem.</b>
    '''
    ''' Uma página legitimamente vazia é comum — fim da pasta —, e um contador
    ''' que estourasse ali derrubaria a listagem por causa da instrumentação.
    ''' </summary>
    <TestMethod>
    Public Sub Pagina_vazia_soma_zero()
        Assert.AreEqual(0, MessagePaging.Fabricadas(Array.Empty(Of MessagePaging.TableRow)()))
        Assert.AreEqual(0, MessagePaging.Fabricadas(Nothing))
    End Sub

    ''' <summary>
    ''' <b>CADA LINHA LEVA O SEU NÚMERO — e é a produção que faz isso.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O TESTE ANTERIOR PASSARIA COM A CORREÇÃO DESFEITA</b>
    '''
    ''' Ele chamava os conversores em sequência e fazia <c>f.Fabricadas = 0</c>
    ''' entre as duas "linhas" — <b>com a própria mão</b>. Ou seja: provava que
    ''' zerar zera. Apagar o <c>Fabricadas = 0</c> da produção o deixaria verde,
    ''' e foi exatamente essa a crítica da revisão externa.
    '''
    ''' Por isso o laço saiu de dentro do <c>Ler</c>, que precisa de uma
    ''' <c>Table</c> do Outlook, e virou <c>ConverterLinhas</c>, que recebe o
    ''' bloco cru. Agora quem zera e quem colhe é o código de produção, e o
    ''' teste só entrega os dados e confere.
    '''
    ''' <b>Controle negativo, nas duas direções:</b> sem o <c>Fabricadas = 0</c>,
    ''' a segunda linha herda as quatro da primeira e a asserção do zero cai;
    ''' colhendo antes do inicializador em vez de depois, a primeira linha vem
    ''' com zero e a outra asserção cai.
    ''' </summary>
    <TestMethod>
    Public Sub Cada_linha_leva_o_SEU_numero()
        Dim f As New MessagePaging.TableRowSource()

        ' Duas linhas, oito colunas -- a mesma forma que o GetArray devolve.
        Dim bruto(1, 7) As Object

        ' LINHA 0: quatro buracos -- EntryID, assunto, tamanho e anexo.
        bruto(0, 0) = Nothing
        bruto(0, 1) = Nothing
        bruto(0, 2) = "Caroline Abreu"
        bruto(0, 3) = New DateTime(2026, 8, 25, 9, 0, 0)
        bruto(0, 4) = Nothing
        bruto(0, 5) = True
        bruto(0, 6) = "IPM.Note"
        bruto(0, 7) = Nothing

        ' LINHA 1: inteira. Ela vem DEPOIS de propósito -- e o zero dela e o
        ' que prova o reset.
        bruto(1, 0) = New Byte() {&HA, &H1B}
        bruto(1, 1) = "Aditivo contratual"
        bruto(1, 2) = "Marcos Vinicius"
        bruto(1, 3) = New DateTime(2026, 8, 26, 10, 0, 0)
        bruto(1, 4) = 4096
        bruto(1, 5) = False
        bruto(1, 6) = "IPM.Note"
        bruto(1, 7) = True

        Dim convertidas = f.ConverterLinhas(bruto)

        Assert.AreEqual(2, convertidas.Count)
        Assert.AreEqual(4, convertidas(0).Fabricadas,
            "a linha com quatro buracos nao levou o seu numero")
        Assert.AreEqual(0, convertidas(1).Fabricadas,
            "a linha inteira herdou a fabricacao da anterior -- o reset sumiu")

        ' E as celulas boas continuam chegando certas, para o teste nao passar
        ' com um conversor que so conta e nao converte.
        Assert.AreEqual("0A1B", convertidas(1).EntryId)
        Assert.AreEqual("Aditivo contratual", convertidas(1).Subject)
        Assert.AreEqual(4096, convertidas(1).SizeBytes)
        Assert.IsTrue(convertidas(1).HasAttachments)
    End Sub

    ''' <summary>
    ''' <b>O CAMINHO LEGADO CONTA — inclusive a ausência sem exceção.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DUAS INSTRUMENTAÇÕES DISCORDANDO É PIOR QUE UMA SÓ</b>
    '''
    ''' A listagem tem dois caminhos: <c>Table</c>, rápido, e a iteração por
    ''' <c>MailItem</c>, para o store que recusa a coluna de EntryID longo. Até
    ''' 28/08/2026 só o rápido contava, então o zero mostrado numa pasta lida
    ''' pelo legado era um zero <b>fabricado</b>.
    '''
    ''' Instrumentei o legado, e a revisão externa achou o buraco que sobrou:
    ''' eu contei só o <c>Catch</c>. O getter pode devolver <c>Nothing</c>
    ''' <b>sem lançar nada</b> — e o <c>ComoTexto</c> do caminho rápido já
    ''' contava esse caso. O número passava a depender de qual caminho a pasta
    ''' tomou, que é a pior propriedade possível para um número que sobe para a
    ''' tela como aviso.
    '''
    ''' <b>Fora deste teste:</b> <c>ContarAnexos</c> também conta agora, e não
    ''' tem teste — ele precisa de um <c>MailItem</c> real que falhe ao abrir
    ''' <c>Attachments</c>. Está dito com esse nome no código.
    ''' </summary>
    <TestMethod>
    Public Sub O_legado_conta_a_excecao_E_o_Nothing_calado()
        Dim n = 0

        ' Controle: valor bom nao e fabricacao.
        Assert.AreEqual("Aditivo", MessagePaging.TextoDoItem(Function() "Aditivo", n))
        Assert.AreEqual(0, n, "valor legivel foi contado como fabricacao")

        ' O CASO QUE ESCAPOU: Nothing sem excecao.
        Assert.AreEqual("", MessagePaging.TextoDoItem(Function() CType(Nothing, String), n))
        Assert.AreEqual(1, n, "Nothing sem excecao virou vazio EM SILENCIO")

        ' E a excecao, que era a unica contada antes.
        Assert.AreEqual("", MessagePaging.TextoDoItem(AddressOf TextoQueExplode, n))
        Assert.AreEqual(2, n)

        Assert.AreEqual(0, MessagePaging.NumeroDoItem(AddressOf NumeroQueExplode, n))
        Assert.AreEqual(3, n)

        Assert.IsFalse(MessagePaging.BooleanoDoItem(AddressOf BooleanoQueExplode, n))
        Assert.AreEqual(4, n)

        ' Controles dos outros dois, para o teste nao passar com auxiliares
        ' que contam sempre.
        Assert.AreEqual(4096, MessagePaging.NumeroDoItem(Function() 4096, n))
        Assert.IsTrue(MessagePaging.BooleanoDoItem(Function() True, n))
        Assert.AreEqual(4, n, "valor legivel foi contado como fabricacao")
    End Sub

    ' Propriedade COM ilegivel, que e o que os auxiliares do legado existem
    ' para sobreviver: item corrompido, offline ou baixado pela metade.
    Private Shared Function TextoQueExplode() As String
        Throw New InvalidOperationException("propriedade COM ilegivel")
    End Function

    Private Shared Function NumeroQueExplode() As Integer
        Throw New InvalidOperationException("propriedade COM ilegivel")
    End Function

    Private Shared Function BooleanoQueExplode() As Boolean
        Throw New InvalidOperationException("propriedade COM ilegivel")
    End Function

End Class

Imports System.Collections.Generic
Imports System.IO
Imports Iris.Assist
Imports Iris.App.ViewModels
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O AVISO QUE NÃO DESAPARECIA SOZINHO — e por que encolher exigiu um
''' reconhecimento.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ELE DIZ</b>
'''
''' Que um envio à IA ficou sem desfecho conhecido: <i>pode ter saído
''' conteúdo, e não dá para saber</i>. É uma divulgação possível, não um erro
''' de operação — e por isso o texto era um parágrafo, e ficava.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ELE FICAVA PARA SEMPRE</b>
'''
''' Não havia como o dono dizer "eu vi". Isso estava anotado no ESCOPO como
''' dívida desde antes de existir esta tela. E um parágrafo permanente é um
''' parágrafo que se aprende a não ler — o aviso ia perdendo força justamente
''' por insistir.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE UM ÍCONE SOZINHO SERIA PIOR</b>
'''
''' Trocar o texto por um ícone, sem reconhecimento, transformaria uma
''' divulgação <b>não vista</b> em uma marca discreta. O reconhecimento é o que
''' torna o ícone honesto: ele diz "você já leu isto", e essa frase precisa ser
''' verdadeira.
'''
''' ------------------------------------------------------------------
''' <b>E O NÚMERO, EM VEZ DE UM "JÁ VI"</b>
'''
''' É a decisão que estes testes prendem. Guardar um booleano faria o clique de
''' hoje calar a ambiguidade de amanhã — a forma mais silenciosa possível de
''' perder uma divulgação. Guardando o número, envio ambíguo novo traz o
''' parágrafo de volta inteiro.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class ReconhecimentoDeAmbiguasTests

    Private _guardado As String

    ''' <summary>
    ''' O reconhecimento mora num caminho fixo, sob <c>LocalApplicationData</c>.
    ''' A suíte não pode deixar lixo na máquina de quem roda, nem ler o estado
    ''' real do dono — então guarda o que houver e devolve no fim.
    ''' </summary>
    <TestInitialize>
    Public Sub Antes()
        Dim c = AssistenteViewModel.CaminhoDoReconhecimento()
        _guardado = If(File.Exists(c), File.ReadAllText(c), Nothing)
        If File.Exists(c) Then File.Delete(c)
    End Sub

    <TestCleanup>
    Public Sub Depois()
        Dim c = AssistenteViewModel.CaminhoDoReconhecimento()
        Try
            If _guardado Is Nothing Then
                If File.Exists(c) Then File.Delete(c)
            Else
                File.WriteAllText(c, _guardado)
            End If
        Catch
        End Try
    End Sub

    Private Shared Function Com(ambiguas As Integer, reconhecidas As Integer) As ReconciliationResult
        Return ReconciliationResult.Rodar(New DiarioQueDevolve(ambiguas),
                                          DateTimeOffset.Now).ComReconhecimento(reconhecidas)
    End Function

    ''' <summary>
    ''' So o Reconciliar importa aqui; o resto da porta existe para o
    ''' caminho de envio, que estes testes nao exercitam.
    ''' </summary>
    Private NotInheritable Class DiarioQueDevolve
        Implements IDisclosureJournal

        Private ReadOnly _n As Integer
        Public Sub New(n As Integer)
            _n = n
        End Sub

        Public Function Reconciliar(q As DateTimeOffset) As Integer _
            Implements IDisclosureJournal.Reconciliar
            Return _n
        End Function

        Public Function Intencao(c As DisclosureCapability, q As DateTimeOffset) As Boolean _
            Implements IDisclosureJournal.Intencao
            Return True
        End Function

        Public Function Iniciando(r As Guid, q As DateTimeOffset) As Boolean _
            Implements IDisclosureJournal.Iniciando
            Return True
        End Function

        Public Function Concluir(r As Guid, q As DateTimeOffset,
                                 codigoHttp As Integer?) As Boolean _
            Implements IDisclosureJournal.Concluir
            Return True
        End Function

        Public Function Falhar(r As Guid, q As DateTimeOffset, n As DisclosureNote,
                               a As Boolean, codigoHttp As Integer?) As Boolean _
                               Implements IDisclosureJournal.Falhar
            Return True
        End Function

        Public Function NaoEnviou(r As Guid, q As DateTimeOffset, n As DisclosureNote,
                                  Optional motivo As DisclosureReason =
                                      DisclosureReason.NaoDecidido) As Boolean _
                                  Implements IDisclosureJournal.NaoEnviou
            Return True
        End Function

        Public Function Ler(n As Integer) As IReadOnlyList(Of DisclosureEntry) _
            Implements IDisclosureJournal.Ler
            Return Array.Empty(Of DisclosureEntry)()
        End Function
    End Class

    ' ==================================================================

    ''' <summary>
    ''' <b>CONTROLE: sem reconhecimento, o parágrafo inteiro aparece.</b>
    ''' </summary>
    <TestMethod>
    Public Sub Controle_sem_reconhecimento_o_paragrafo_aparece()
        Dim r = Com(ambiguas:=1, reconhecidas:=0)

        Assert.IsTrue(r.TemNovidade)
        StringAssert.Contains(r.Aviso, "sem desfecho conhecido")
        StringAssert.Contains(r.Aviso, "Pode ter saído conteúdo",
            "o aviso parou de dizer o que está em jogo")
        Assert.AreEqual("", r.Marca, "mostrou a marca curta E o parágrafo")
    End Sub

    ''' <summary>
    ''' <b>Reconhecido vira MARCA, e a marca não some.</b>
    '''
    ''' A ambiguidade continua sendo verdade depois de reconhecida — o que mudou
    ''' foi o tamanho com que ela é dita. Uma marca que sumisse seria o ícone
    ''' desonesto que este desenho existe para evitar.
    ''' </summary>
    <TestMethod>
    Public Sub Reconhecido_vira_marca_e_a_marca_NAO_some()
        Dim r = Com(ambiguas:=3, reconhecidas:=3)

        Assert.IsFalse(r.TemNovidade)
        Assert.AreEqual("", r.Aviso, "o parágrafo ficou depois de reconhecido")
        Assert.IsTrue(r.TemMarca, "a ambiguidade sumiu da tela ao ser reconhecida")
        StringAssert.Contains(r.Marca, "3")
        StringAssert.Contains(r.Marca, "sem desfecho conhecido",
            "a marca não diz mais do que se trata")
    End Sub

    ''' <summary>
    ''' <b>AMBIGUIDADE NOVA TRAZ O PARÁGRAFO DE VOLTA.</b>
    '''
    ''' O teste que justifica guardar o número em vez de um "já vi". Com um
    ''' booleano, o clique de ontem calaria a divulgação de hoje.
    '''
    ''' <b>Controle negativo:</b> trocando <c>Ambiguas &gt; Reconhecidas</c> por
    ''' <c>Reconhecidas = 0</c>, este teste cai.
    ''' </summary>
    <TestMethod>
    Public Sub Ambiguidade_NOVA_traz_o_paragrafo_de_volta()
        ' Tres reconhecidas, e agora sao cinco: duas sao novidade.
        Dim r = Com(ambiguas:=5, reconhecidas:=3)

        Assert.IsTrue(r.TemNovidade,
            "duas divulgações novas ficaram caladas pelo reconhecimento das antigas")
        StringAssert.Contains(r.Aviso, "sem desfecho conhecido")
        Assert.AreEqual("", r.Marca, "mostrou a marca curta havendo novidade")
    End Sub

    ''' <summary>
    ''' <b>Zero ambíguas não mostra nada.</b> Nem parágrafo, nem marca — não há
    ''' o que ressalvar, e uma marca permanente sobre nada é ruído que ensina a
    ''' ignorar a marca de verdade.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_ambiguas_nao_ha_marca_nem_aviso()
        Dim r = Com(ambiguas:=0, reconhecidas:=0)

        Assert.AreEqual("", r.Aviso)
        Assert.AreEqual("", r.Marca)
        Assert.IsFalse(r.TemMarca)
    End Sub

    ''' <summary>
    ''' <b>A reconciliação que NÃO terminou vence o reconhecimento.</b>
    '''
    ''' "Não consegui conferir o registro" não é uma ambiguidade que se
    ''' reconhece: é não saber. Reconhecer o que não se sabe seria a versão
    ''' assinada de ignorar.
    '''
    ''' <b>E este teste NÃO alcança a guarda que parece alcançar.</b> O
    ''' controle negativo mostrou: tirando o <c>Not Terminou</c> do
    ''' <c>Marca</c>, ele continua passando, porque <c>NaoRodou</c> devolve
    ''' zero ambíguas e a condição do zero já pega. Ele prende o
    ''' <i>comportamento</i> — não ter conferido não vira marca — e não a
    ''' linha que o produz.
    ''' </summary>
    <TestMethod>
    Public Sub Nao_ter_conferido_nao_vira_marca()
        Dim r = ReconciliationResult.NaoRodou().ComReconhecimento(99)

        StringAssert.Contains(r.Aviso, "Não foi possível conferir")
        Assert.AreEqual("", r.Marca,
            "encolheu para uma marca um estado que é 'não sei', e não 'já vi'")
    End Sub

    ''' <summary>
    ''' <b>Não conseguir LER o reconhecimento vale ZERO.</b>
    '''
    ''' Falha fechada, como as outras deste projeto: não conseguir ler que o
    ''' dono viu não é o dono ter visto. O custo de errar para este lado é um
    ''' aviso a mais; para o outro lado, é uma divulgação a menos.
    ''' </summary>
    <TestMethod>
    Public Sub Reconhecimento_ilegivel_vale_zero()
        Dim c = AssistenteViewModel.CaminhoDoReconhecimento()
        Directory.CreateDirectory(Path.GetDirectoryName(c))
        File.WriteAllText(c, "isto não é um número")

        ' A LEITURA DIRETO, e nao pelo Rodar: quem le o arquivo e o
        ' AssistenteViewModel na construcao, e o Rodar so pergunta ao banco.
        ' A primeira versao deste teste chamava o Rodar e conferia
        ' Reconhecidas -- ele passaria com a leitura QUEBRADA, porque aquele
        ' caminho nunca toca no arquivo.
        Assert.AreEqual(0, AssistenteViewModel.LerReconhecidas(),
            "conteúdo ilegível virou 'já reconheceu', e a divulgação encolheria")

        Dim r = ReconciliationResult.Rodar(New DiarioQueDevolve(2), DateTimeOffset.Now).
                ComReconhecimento(AssistenteViewModel.LerReconhecidas())
        Assert.IsTrue(r.TemNovidade)
    End Sub

End Class

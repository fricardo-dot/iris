Imports System.Collections.Generic
Imports Iris.Model

Namespace Global.Iris.Assist

    ''' <summary>
    ''' Em que ponto do envio uma divulgação está — ou parou.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A DISTINÇÃO QUE O ENUM INTEIRO EXISTE PARA FAZER</b>
    '''
    ''' Entre <see cref="Intencionada"/> e <see cref="EmVoo"/> mora a diferença
    ''' entre "nada saiu" e "não dá para saber". Se o processo morre no
    ''' primeiro, o conteúdo não foi para lugar nenhum; se morre no segundo,
    ''' <b>talvez tenha ido</b>, e ninguém nunca vai saber.
    '''
    ''' É a mesma disciplina do <c>ErrorKind.Ambiguous</c> que o CLAUDE.md
    ''' impõe às mutações, aplicada ao egress.
    ''' </summary>
    Public Enum DisclosureStage
        ''' <summary>Zero: registro incompleto. Nunca significa "não saiu".</summary>
        Desconhecido = 0

        ''' <summary>
        ''' A intenção foi gravada e a transmissão <b>não começou</b>. Morrer
        ''' aqui é seguro: nada saiu.
        ''' </summary>
        Intencionada

        ''' <summary>
        ''' A transmissão <b>começou</b>. Morrer aqui é o caso ambíguo — os
        ''' bytes podem ter chegado ao provedor.
        ''' </summary>
        EmVoo

        ''' <summary>Terminou, e o provedor respondeu.</summary>
        Concluida

        ''' <summary>
        ''' Falhou de um jeito que se sabe que <b>não</b> chegou — recusa antes
        ''' de transmitir, portão negando, capability recusada.
        ''' </summary>
        NaoEnviada

        ''' <summary>
        ''' <b>Pode ter chegado, e não dá para saber.</b> Timeout, cancelamento
        ''' depois de começar, conexão caindo, ou o processo morrendo em voo.
        ''' Nunca vira "não enviou" depois.
        ''' </summary>
        Ambigua
    End Enum

    ''' <summary>
    ''' Uma linha do diário. <b>Nunca carrega conteúdo</b> — nem trecho, nem
    ''' assunto, nem nome de rótulo.
    '''
    ''' O R11 do ESCOPO é explícito: <i>log do que foi enviado à IA registrando
    ''' metadados, hash, modelo e tamanho, não o conteúdo — um log com o texto
    ''' cria mais uma cópia sensível</i>.
    ''' </summary>
    Public NotInheritable Class DisclosureEntry

        Public ReadOnly Property RequestId As Guid
        Public ReadOnly Property CapabilityId As Guid
        Public ReadOnly Property Estagio As DisclosureStage
        Public ReadOnly Property AtivacaoId As String
        Public ReadOnly Property AtivacaoVersao As Integer
        Public ReadOnly Property Operacao As AssistOperation
        Public ReadOnly Property Provedor As String
        Public ReadOnly Property Endpoint As String
        Public ReadOnly Property Modelo As String
        ''' <summary>SHA-256 dos bytes. <c>Nothing</c> quando nada foi autorizado.</summary>
        Public ReadOnly Property Hash As String
        Public ReadOnly Property Bytes As Integer
        Public ReadOnly Property Mensagens As Integer
        Public ReadOnly Property Quando As DateTimeOffset
        ''' <summary>
        ''' O motivo, em forma de código — nunca texto do provedor.
        '''
        ''' Corpo de resposta de erro pode <b>ecoar o conteúdo enviado</b>, e um
        ''' diário que o guardasse viraria a cópia que ele existe para não criar.
        ''' </summary>
        Public ReadOnly Property Motivo As String

        Public Sub New(requestId As Guid, capabilityId As Guid, estagio As DisclosureStage,
                       ativacaoId As String, ativacaoVersao As Integer,
                       operacao As AssistOperation, provedor As String, endpoint As String,
                       modelo As String, hash As String, bytes As Integer,
                       mensagens As Integer, quando As DateTimeOffset, motivo As String)
            Me.RequestId = requestId
            Me.CapabilityId = capabilityId
            Me.Estagio = estagio
            Me.AtivacaoId = If(ativacaoId, "")
            Me.AtivacaoVersao = ativacaoVersao
            Me.Operacao = operacao
            Me.Provedor = If(provedor, "")
            Me.Endpoint = If(endpoint, "")
            Me.Modelo = If(modelo, "")
            Me.Hash = hash
            Me.Bytes = bytes
            Me.Mensagens = mensagens
            Me.Quando = quando
            Me.Motivo = If(motivo, "")
        End Sub

    End Class

    ' ==================================================================

    ''' <summary>
    ''' <b>O diário do egress, e o protocolo de crash dele.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE REGISTRAR DEPOIS NÃO SERVE</b>
    '''
    ''' Um diário escrito no fim registra os envios que terminaram — e perde
    ''' justamente os que importam. Se o processo morre durante a transmissão,
    ''' não há linha nenhuma, e o registro passa a afirmar, por omissão, que
    ''' nada saiu. Foi exatamente isso que aconteceu.
    '''
    ''' Então são <b>cinco passos</b>, nesta ordem:
    '''
    '''   1. <see cref="Intencao"/> — durável, <b>antes</b> de qualquer tentativa;
    '''   2. o hash dos bytes exatos vai junto, na intenção;
    '''   3. <see cref="Iniciando"/> — a transmissão começou;
    '''   4. <see cref="Concluir"/> ou <see cref="Falhar"/>;
    '''   5. <see cref="Reconciliar"/> na abertura seguinte: o que ficou em voo
    '''      vira <see cref="DisclosureStage.Ambigua"/>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DECISÃO NEGADA NÃO REGISTRA HASH</b>
    '''
    ''' Registrar o hash de algo que não saiu confundiria as duas coisas na
    ''' leitura de depois — e "houve um hash aqui" é o que alguém vai procurar
    ''' quando a pergunta for se o conteúdo vazou.
    ''' </summary>
    Public Interface IDisclosureJournal

        ''' <summary>
        ''' Grava a intenção, com o hash dos bytes. <b>Antes</b> de transmitir,
        ''' e de forma durável — se não durar, não serve.
        ''' </summary>
        Sub Intencao(c As DisclosureCapability, mensagens As Integer,
                     quando As DateTimeOffset)

        ''' <summary>A transmissão começou. Daqui em diante, morrer é ambíguo.</summary>
        Sub Iniciando(requestId As Guid, quando As DateTimeOffset)

        Sub Concluir(requestId As Guid, quando As DateTimeOffset)

        ''' <summary>
        ''' Terminou sem sucesso.
        ''' </summary>
        ''' <param name="podeTerChegado">
        ''' <c>True</c> quando não dá para saber — timeout, cancelamento depois
        ''' de começar, conexão caindo. Vira <see cref="DisclosureStage.Ambigua"/>,
        ''' e <b>nunca</b> volta a ser "não enviou".
        ''' </param>
        Sub Falhar(requestId As Guid, quando As DateTimeOffset, motivo As String,
                   podeTerChegado As Boolean)

        ''' <summary>
        ''' Registra uma divulgação que <b>não aconteceu</b> — o portão negou, a
        ''' capability foi recusada, o conteúdo não passou. Sem hash.
        ''' </summary>
        Sub NaoEnviou(requestId As Guid, quando As DateTimeOffset, motivo As String)

        ''' <summary>
        ''' Na abertura: o que ficou <see cref="DisclosureStage.EmVoo"/> de uma
        ''' execução anterior vira <see cref="DisclosureStage.Ambigua"/>, e o
        ''' que ficou <see cref="DisclosureStage.Intencionada"/> vira
        ''' <see cref="DisclosureStage.NaoEnviada"/>.
        '''
        ''' Devolve quantas viraram ambíguas — número que a UI mostra, porque
        ''' "pode ter saído conteúdo e ninguém sabe" não é detalhe de log.
        ''' </summary>
        Function Reconciliar(quando As DateTimeOffset) As Integer

        Function Ler(quantas As Integer) As IReadOnlyList(Of DisclosureEntry)

    End Interface

End Namespace

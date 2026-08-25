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
    ''' Por que uma divulgação parou onde parou — <b>em código, nunca em texto</b>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE NÃO É UMA STRING</b>
    '''
    ''' O campo era <c>String</c>, e "o diário nunca guarda conteúdo" virava
    ''' convenção: qualquer adaptador podia passar a mensagem de uma exceção ou
    ''' o corpo de erro do provedor — e corpo de erro <b>ecoa o que foi
    ''' enviado</b>. Um diário que aceitasse texto arbitrário seria a cópia
    ''' sensível que ele existe para não criar.
    '''
    ''' Enum fechado não tem esse problema, e a tradução para português mora na
    ''' apresentação, que é onde ela serve.
    ''' </summary>
    Public Enum DisclosureNote
        Nenhuma = 0
        ''' <summary>O portão negou. O motivo específico vai em campo próprio.</summary>
        PortaoNegou
        ''' <summary>O cofre recusou a capability.</summary>
        CapabilityRecusada
        ''' <summary>O envelope não pôde ser montado.</summary>
        EnvelopeRecusado
        ''' <summary>O conteúdo não pôde ser preparado.</summary>
        ConteudoRecusado
        ''' <summary>
        ''' O provedor não está em condição de transmitir — sem credencial,
        ''' endereço inseguro, nenhum provedor configurado.
        '''
        ''' Distinta de <see cref="CapabilityRecusada"/>: a capability foi
        ''' consumida com sucesso, e o que faltou foi do outro lado. Registrar as
        ''' duas com a mesma nota faria "o cofre recusou" e "não havia credencial"
        ''' virarem a mesma linha no diário.
        ''' </summary>
        ProvedorIndisponivel
        ''' <summary>O tempo acabou. <b>Não</b> quer dizer que não chegou.</summary>
        Timeout
        ''' <summary>O usuário mandou parar.</summary>
        Cancelado
        ''' <summary>A conexão caiu.</summary>
        ConexaoCaiu
        ''' <summary>O provedor respondeu com erro.</summary>
        ProvedorRecusou
        ''' <summary>O processo terminou em voo.</summary>
        ProcessoMorreuEmVoo
        ''' <summary>O processo terminou antes de transmitir.</summary>
        ProcessoMorreuAntesDeTransmitir
    End Enum

    ''' <summary>
    ''' <b>Que combinações de nota e motivo fazem sentido.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ENUM DO .NET NÃO É FECHADO</b>
    '''
    ''' Trocar <c>String</c> por enum tirou o texto arbitrário, e não fechou a
    ''' porta: <c>CType(999, DisclosureNote)</c> compila e roda. E mesmo entre
    ''' valores definidos há combinações que não descrevem nada —
    ''' <c>PortaoNegou</c> sem dizer o que o portão negou, ou <c>Timeout</c>
    ''' acompanhado de um motivo de portão que não teve nada a ver.
    '''
    ''' Um diário com registro incoerente é pior que um diário sem o registro:
    ''' ele parece resposta.
    ''' </summary>
    Public Module DisclosureNotes

        ''' <summary>O valor está definido no enum? <c>CType(999, …)</c> não está.</summary>
        Public Function Definida(n As DisclosureNote) As Boolean
            Return [Enum].IsDefined(GetType(DisclosureNote), n)
        End Function

        Public Function Definido(r As DisclosureReason) As Boolean
            Return [Enum].IsDefined(GetType(DisclosureReason), r)
        End Function

        ''' <summary>
        ''' A dupla é coerente? <c>PortaoNegou</c> <b>exige</b> um motivo; toda
        ''' outra nota exige <c>NaoDecidido</c>.
        ''' </summary>
        Public Function Coerente(n As DisclosureNote, r As DisclosureReason) As Boolean
            If Not Definida(n) OrElse Not Definido(r) Then Return False
            If n = DisclosureNote.PortaoNegou Then Return r <> DisclosureReason.NaoDecidido
            Return r = DisclosureReason.NaoDecidido
        End Function

        ''' <summary>
        ''' Notas que descrevem um envio que <b>já tinha começado</b> ou que
        ''' terminou mal — as únicas que <c>Falhar</c> aceita.
        ''' </summary>
        Public Function DeTransporte(n As DisclosureNote) As Boolean
            Select Case n
                Case DisclosureNote.Timeout, DisclosureNote.Cancelado,
                     DisclosureNote.ConexaoCaiu, DisclosureNote.ProvedorRecusou,
                     DisclosureNote.ProcessoMorreuEmVoo
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' Notas de coisa que impediu o envio <b>antes</b> dele — as únicas que
        ''' <c>NaoEnviou</c> aceita.
        ''' </summary>
        Public Function AnteriorAoEnvio(n As DisclosureNote) As Boolean
            Select Case n
                Case DisclosureNote.PortaoNegou, DisclosureNote.CapabilityRecusada,
                     DisclosureNote.EnvelopeRecusado, DisclosureNote.ConteudoRecusado,
                     DisclosureNote.ProvedorIndisponivel,
                     DisclosureNote.ProcessoMorreuAntesDeTransmitir
                    Return True
                Case Else
                    Return False
            End Select
        End Function

    End Module

    ''' <summary>
    ''' Uma linha do diário. <b>Nunca carrega conteúdo</b> — nem trecho, nem
    ''' assunto, nem nome de rótulo, nem texto vindo do provedor.
    '''
    ''' O R11 do ESCOPO é explícito: <i>log do que foi enviado à IA registrando
    ''' metadados, hash, modelo e tamanho, não o conteúdo — um log com o texto
    ''' cria mais uma cópia sensível</i>.
    ''' </summary>
    Public NotInheritable Class DisclosureEntry

        ''' <summary>Ordem de inserção. Desempate estável — o <c>Guid</c> não é.</summary>
        Public ReadOnly Property Sequencia As Long
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

        ''' <summary>
        ''' Quando a intenção foi gravada. <b>Imutável</b>, e é por ela que a
        ''' ordem histórica se guia.
        '''
        ''' Havia um <c>at</c> único, sobrescrito a cada passo. Depois de uma
        ''' reconciliação, uma intenção abandonada há meses aparecia como
        ''' atividade recente — e a evidência de <i>quando</i> cada passo
        ''' aconteceu simplesmente sumia.
        ''' </summary>
        Public ReadOnly Property Intencionada As DateTimeOffset
        ''' <summary>Quando entrou em voo, se entrou.</summary>
        Public ReadOnly Property Iniciada As DateTimeOffset?
        ''' <summary>Quando terminou — por conclusão, falha ou reconciliação.</summary>
        Public ReadOnly Property Terminada As DateTimeOffset?

        Public ReadOnly Property Nota As DisclosureNote
        ''' <summary>
        ''' Quando a nota é <see cref="DisclosureNote.PortaoNegou"/>, qual foi o
        ''' motivo. Também enum fechado.
        ''' </summary>
        Public ReadOnly Property MotivoDoPortao As DisclosureReason

        Public Sub New(sequencia As Long, requestId As Guid, capabilityId As Guid,
                       estagio As DisclosureStage, ativacaoId As String,
                       ativacaoVersao As Integer, operacao As AssistOperation,
                       provedor As String, endpoint As String, modelo As String,
                       hash As String, bytes As Integer, mensagens As Integer,
                       intencionada As DateTimeOffset, iniciada As DateTimeOffset?,
                       terminada As DateTimeOffset?, nota As DisclosureNote,
                       motivoDoPortao As DisclosureReason)
            Me.Sequencia = sequencia
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
            Me.Intencionada = intencionada
            Me.Iniciada = iniciada
            Me.Terminada = terminada
            Me.Nota = nota
            Me.MotivoDoPortao = motivoDoPortao
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
    ''' nada saiu.
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
    ''' <b>TODA TRANSIÇÃO DIZ SE PEGOU</b>
    '''
    ''' Os passos devolvem <c>Boolean</c>, e não é decoração. Um
    ''' <c>Iniciando</c> que não persistisse — pedido inexistente, estado
    ''' errado, corrida — passava em silêncio, e quem chamou seguia para o HTTP
    ''' assim mesmo. Resultado: <b>egress sem registro de voo</b>, que é
    ''' exatamente o buraco que o diário existe para não ter.
    '''
    ''' <b>Quem transmite só toca na rede depois de <c>Iniciando</c> devolver
    ''' <c>True</c>.</b>
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
        ''' <returns><c>False</c> se já existe registro para este pedido.</returns>
        ''' <remarks>
        ''' A contagem de mensagens vem da <b>própria capability</b>, e não como
        ''' parâmetro. Enquanto vinha de fora, o diário podia registrar uma
        ''' quantidade diferente da autorizada — e o número de mensagens é
        ''' justamente o que alguém confere quando a pergunta for quanto saiu.
        ''' </remarks>
        Function Intencao(c As DisclosureCapability, quando As DateTimeOffset) As Boolean

        ''' <summary>
        ''' A transmissão vai começar. Daqui em diante, morrer é ambíguo.
        ''' </summary>
        ''' <returns>
        ''' <c>False</c> quando a transição <b>não</b> aconteceu. Quem chama
        ''' <b>não pode</b> tocar na rede nesse caso.
        ''' </returns>
        Function Iniciando(requestId As Guid, quando As DateTimeOffset) As Boolean

        Function Concluir(requestId As Guid, quando As DateTimeOffset) As Boolean

        ''' <summary>
        ''' Terminou sem sucesso.
        ''' </summary>
        ''' <param name="podeTerChegado">
        ''' <c>True</c> quando não dá para saber — timeout, cancelamento depois
        ''' de começar, conexão caindo. Vira <see cref="DisclosureStage.Ambigua"/>,
        ''' e <b>nunca</b> volta a ser "não enviou".
        ''' </param>
        ''' <remarks>
        ''' Só aceita nota de <b>transporte</b>
        ''' (<see cref="DisclosureNotes.DeTransporte"/>). "O portão negou" não é
        ''' um jeito de a transmissão falhar: ela nem teria começado.
        ''' </remarks>
        Function Falhar(requestId As Guid, quando As DateTimeOffset,
                        nota As DisclosureNote, podeTerChegado As Boolean) As Boolean

        ''' <summary>
        ''' Registra uma divulgação que <b>não aconteceu</b> — o portão negou, a
        ''' capability foi recusada, o conteúdo não passou. Sem hash novo.
        '''
        ''' Só aceita nota <b>anterior ao envio</b>
        ''' (<see cref="DisclosureNotes.AnteriorAoEnvio"/>), e a dupla nota/motivo
        ''' tem de ser coerente: <c>PortaoNegou</c> exige dizer o que o portão
        ''' negou.
        ''' </summary>
        Function NaoEnviou(requestId As Guid, quando As DateTimeOffset,
                           nota As DisclosureNote,
                           Optional motivoDoPortao As DisclosureReason =
                               DisclosureReason.NaoDecidido) As Boolean

        ''' <summary>
        ''' Na abertura: o que ficou <see cref="DisclosureStage.EmVoo"/> de uma
        ''' execução anterior vira <see cref="DisclosureStage.Ambigua"/>, e o
        ''' que ficou <see cref="DisclosureStage.Intencionada"/> vira
        ''' <see cref="DisclosureStage.NaoEnviada"/>.
        '''
        ''' <b>Não é trabalho da UI.</b> É recuperação de segurança, e roda na
        ''' composição, antes de o assistente ficar apto a transmitir — se ela
        ''' falhar ou não terminar, o egress fica fechado. A UI só mostra o
        ''' número.
        '''
        ''' Devolve quantas viraram ambíguas, porque "pode ter saído conteúdo e
        ''' ninguém sabe" não é detalhe de log.
        ''' </summary>
        Function Reconciliar(quando As DateTimeOffset) As Integer

        Function Ler(quantas As Integer) As IReadOnlyList(Of DisclosureEntry)

    End Interface

End Namespace

Imports System.Threading

Namespace Global.Iris.Assist

    ''' <summary>Como uma tentativa de falar com o modelo terminou.</summary>
    Public Enum ProviderStatus
        ''' <summary>Zero: não decidido. Nunca significa sucesso.</summary>
        Desconhecido = 0

        ''' <summary>O provedor respondeu.</summary>
        Respondeu
        ''' <summary>O provedor respondeu <b>com erro</b>.</summary>
        Recusou
        ''' <summary>O tempo acabou. <b>Não</b> quer dizer que não chegou.</summary>
        Timeout
        ''' <summary>O usuário mandou parar depois de a chamada começar.</summary>
        Cancelado
        ''' <summary>A conexão caiu.</summary>
        ConexaoCaiu
        ''' <summary>
        ''' Nem começou — configuração recusada, endereço não-HTTPS, sem
        ''' credencial. Nenhum byte saiu, e isso <b>se sabe</b>.
        ''' </summary>
        NaoComecou

        ''' <summary>
        ''' A resposta passou do teto.
        '''
        ''' Estado próprio, e não sucesso: devolver o pedaço que coube
        ''' apresentaria uma resposta <b>parcial</b> como se fosse completa — e um
        ''' resumo cortado no meio parece um resumo.
        ''' </summary>
        RespostaGrandeDemais

        ''' <summary>
        ''' Respondeu, e <b>não deu para ler</b> a resposta.
        '''
        ''' Estado próprio, e não <see cref="Respondeu"/> com texto vazio: as
        ''' duas coisas são diferentes para quem lê a tela depois. "O provedor
        ''' não tinha o que dizer" e "o provedor disse algo que o Iris não
        ''' entendeu" levam a investigações diferentes.
        '''
        ''' E não é <see cref="NaoComecou"/>: o conteúdo <b>saiu</b>.
        ''' </summary>
        RespostaIlegivel
    End Enum

    ''' <summary>
    ''' O que o provedor devolveu.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O TEXTO É DADO, NUNCA COMANDO</b>
    '''
    ''' <see cref="Texto"/> é a resposta do modelo, e ela vem de um lugar que
    ''' leu o e-mail — que por sua vez veio de fora. Ela não escolhe endpoint,
    ''' não pede nova chamada, não aciona COM, não cria destinatário, não envia,
    ''' não abre URL e não vira Markdown ativo em lugar nenhum.
    '''
    ''' Campos estruturais no envelope reduzem ambiguidade e <b>não</b> impedem
    ''' o modelo obedecer ao e-mail. A barreira é esta: a saída é passiva.
    ''' </summary>
    Public NotInheritable Class ProviderOutcome

        Public ReadOnly Property Status As ProviderStatus
        ''' <summary>Texto passivo. Vazio quando não houve resposta.</summary>
        Public ReadOnly Property Texto As String

        ''' <summary>
        ''' <b>Código HTTP, quando houve resposta.</b> Nunca o corpo do erro.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>ELE É FILTRADO AQUI, NA ENTRADA</b>
        '''
        ''' Este é o ponto onde a palavra do provedor entra no Iris, e é aqui
        ''' que ela é conferida — uma vez só, e não em cada camada que a
        ''' carrega depois.
        '''
        ''' Duas coisas não passam, por motivos <b>diferentes</b>:
        '''
        ''' <list type="bullet">
        ''' <item>número fora de <c>100..599</c>. Um servidor hostil escolhe o
        ''' que responde, então isto <b>não</b> é sinal de defeito de
        ''' adaptador — é só um número que não descreve resposta nenhuma;</item>
        ''' <item>código num estado que <b>não chegou a receber resposta</b> —
        ''' <c>ConexaoCaiu</c> com 418, por exemplo. Isso não vem de servidor
        ''' nenhum: é o adaptador se contradizendo. E chegaria à tela como "o
        ''' provedor respondeu HTTP 418" logo abaixo de uma frase dizendo que
        ''' ele não respondeu.</item>
        ''' </list>
        '''
        ''' Nos dois casos o campo fica <c>Nothing</c>, e <c>Nothing</c> tem um
        ''' sentido só: <b>não há status por que responder</b>.
        ''' </summary>
        Public ReadOnly Property Codigo As Integer?

        ''' <summary>
        ''' <b>O custo que o provedor INFORMOU</b>, em dólares. Não é fato
        ''' conferido: é a palavra dele.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>"NÚMERO NÃO ECOA CONTEÚDO" NÃO É O MOTIVO — ERA A DESCULPA</b>
        '''
        ''' A primeira versão deste comentário dizia que custo e tokens entram
        ''' "pelo mesmo motivo que o código HTTP", porque número não ecoa nada.
        ''' <b>Está errado</b>, e a revisão pegou: um servidor hostil escolhe
        ''' esses números exatamente como escolheria texto. Confiabilidade não
        ''' vem do tipo.
        '''
        ''' O motivo verdadeiro é outro, e é mais modesto: são dados
        ''' <b>opcionais e passivos</b>, e o pior caso é um número mentiroso na
        ''' ficha. Não há execução, não há injeção, e o tamanho é limitado por
        ''' <c>Decimal</c> e <c>Int32</c>. O usuário fica com uma ideia errada
        ''' do gasto — e é por isso que a tela diz <b>informado</b>, e não
        ''' "custou".
        '''
        ''' <b>Daqui não sai orçamento, bloqueio nem auditoria.</b> Para isso a
        ''' fonte é o painel do provedor, que é de quem cobra. O limite de gasto
        ''' da chave existe justamente porque este número não serve de freio.
        '''
        ''' O que continua de fora é o <c>provider</c> devolvido, que é texto:
        ''' quem diz o agente e o modelo na tela é a <b>ativação</b>, que o
        ''' usuário assinou.
        ''' </summary>
        Public ReadOnly Property Custo As Decimal?

        ''' <summary>Tokens somados da chamada, quando o provedor conta.</summary>
        Public ReadOnly Property Tokens As Integer?

        Public Sub New(status As ProviderStatus, texto As String,
                       Optional codigo As Integer? = Nothing,
                       Optional custo As Decimal? = Nothing,
                       Optional tokens As Integer? = Nothing)
            Me.Status = status
            Me.Texto = If(texto, "")
            Me.Codigo = Confiavel(status, codigo)
            ' Negativo nao descreve custo nem contagem de nada. Vira Nothing
            ' pelo mesmo motivo que codigo fora da faixa vira: um numero
            ' estranho num campo de diagnostico nao pode virar afirmacao.
            Me.Custo = If(custo.HasValue AndAlso custo.Value >= 0D, custo, Nothing)
            Me.Tokens = If(tokens.HasValue AndAlso tokens.Value >= 0, tokens, Nothing)
        End Sub

        ''' <summary>
        ''' Estados em que <b>houve resposta</b> — os únicos que podem trazer
        ''' status. <c>Select Case</c> sem <c>Case Else</c> permissivo: estado
        ''' novo no enum entra aqui recusando, e não aceitando por omissão.
        ''' </summary>
        Public Shared Function PodeTerCodigo(s As ProviderStatus) As Boolean
            Select Case s
                Case ProviderStatus.Respondeu, ProviderStatus.Recusou,
                     ProviderStatus.RespostaGrandeDemais,
                     ProviderStatus.RespostaIlegivel
                    Return True
                Case Else
                    ' Desconhecido, Timeout, Cancelado, ConexaoCaiu,
                    ' NaoComecou: nenhum deles leu uma resposta.
                    Return False
            End Select
        End Function

        ''' <summary>O código que se pode afirmar, ou <c>Nothing</c>.</summary>
        Public Shared Function Confiavel(s As ProviderStatus,
                                         codigo As Integer?) As Integer?
            If Not codigo.HasValue Then Return Nothing
            If Not PodeTerCodigo(s) Then Return Nothing
            If codigo.Value < 100 OrElse codigo.Value > 599 Then Return Nothing
            Return codigo
        End Function

        ''' <summary>
        ''' Os bytes podem ter chegado ao provedor?
        '''
        ''' Só <see cref="ProviderStatus.NaoComecou"/> responde não com
        ''' segurança. Todo o resto — inclusive <see cref="ProviderStatus.Recusou"/>,
        ''' porque para recusar ele leu — pode ter recebido o conteúdo.
        ''' </summary>
        Public ReadOnly Property PodeTerChegado As Boolean
            Get
                Return Status <> ProviderStatus.NaoComecou
            End Get
        End Property

    End Class

    ''' <summary>
    ''' <b>A porta externa: chamar o modelo.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ISTO NÃO É O SERVIÇO DE APLICAÇÃO</b>
    '''
    ''' Aplicar política, montar contexto, registrar intenção e publicar
    ''' resultado é <see cref="AssistTransmitter"/>. Aqui é só o transporte —
    ''' e a separação existe para ninguém confundir "chamar o modelo" com
    ''' "executar a operação segura".
    '''
    ''' ------------------------------------------------------------------
    ''' <b>RECEBE BYTES, E SÓ BYTES</b>
    '''
    ''' Não recebe texto, nem mensagem, nem DTO que o provedor serialize. Se
    ''' recebesse, haveria uma segunda serialização — e o que foi autorizado
    ''' deixaria de ser o que sai, que é a garantia inteira do 3.2.
    ''' </summary>
    Public Interface IAssistantProvider

        ''' <summary>
        ''' Manda <b>exatamente estes bytes</b>. Sem retry: começou, acabou.
        '''
        ''' A regra "leitura tem retry, mutação não" do CLAUDE.md vale aqui —
        ''' egress é mutação do mundo, e repetir depois de começar pode mandar
        ''' o mesmo conteúdo duas vezes.
        ''' </summary>
        ''' <summary>
        ''' <b>Traduz o envelope no corpo que este provedor aceita.</b> Local,
        ''' puro, e <b>sem rede</b>.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE ISTO EXISTE, E POR QUE ACONTECE ANTES DA CAPABILITY</b>
        '''
        ''' O envelope é o formato do Iris; nenhum provedor real aceita ele. A
        ''' primeira versão deste desenho deixava o adaptador reescrever os bytes
        ''' na hora de enviar — e aí <b>o que ia no fio não era o que a
        ''' capability tinha coberto</b>. A autorização falava de um artefato e a
        ''' rede transportava outro.
        '''
        ''' Agora a tradução acontece <b>antes</b> da emissão, e a capability
        ''' cobre os dois: o envelope, pela proveniência, e o corpo, porque é ele
        ''' que sai. <see cref="Enviar"/> volta a valer ao pé da letra: manda
        ''' exatamente estes bytes.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O QUE UMA IMPLEMENTAÇÃO PODE E NÃO PODE FAZER</b>
        '''
        ''' Pode <b>embrulhar</b> o envelope e acrescentar o que vem da
        ''' <b>autorização</b> — modelo, restrição de roteamento. Não pode
        ''' acrescentar conteúdo que a autorização não cobre, não pode ler nada
        ''' de fora, e não pode tocar na rede.
        '''
        ''' <c>Nothing</c> quer dizer que não deu para preparar, e nada sai.
        ''' </summary>
        Function Preparar(envelope As Byte()) As Byte()

        Function Enviar(bytes As Byte(), ct As CancellationToken) As ProviderOutcome

        ''' <summary>Para onde este provedor manda. Fixo, e nunca vindo do prompt.</summary>
        ReadOnly Property Destino As AssistDestination

        ''' <summary>
        ''' <b>Dá para transmitir?</b> Perguntado <b>antes</b> de o diário marcar
        ''' o voo.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE A PERGUNTA EXISTE</b>
        '''
        ''' O diário marca "em voo" antes de chamar o transporte, e de lá em
        ''' diante toda falha é <b>ambígua</b> — porque, do lado de cá, "a conexão
        ''' caiu" e "a conexão caiu depois de o servidor ler o corpo" não se
        ''' distinguem.
        '''
        ''' Mas há recusas que se sabem <b>antes</b> de qualquer byte: endereço
        ''' que não é HTTPS, credencial ausente, URL que não parseia, provedor
        ''' nenhum configurado. Marcá-las como ambíguas encheria de ruído
        ''' justamente a contagem que a UI mostra — e "pode ter saído conteúdo"
        ''' tem de significar alguma coisa.
        '''
        ''' Então quem sabe responde, e responde <b>sem tocar na rede</b>.
        ''' </summary>
        Function Pronto() As Boolean

    End Interface

    ''' <summary>
    ''' <b>O provedor que a produção tem: nenhum.</b>
    '''
    ''' Não é lacuna. É a §28.2: a política corporativa aplicável não é
    ''' inferível desta máquina e a escolha do provedor é do usuário. Enquanto
    ''' isso não acontecer, a única política defensável é não mandar nada — e
    ''' este objeto é essa política em forma de código.
    '''
    ''' Ele existe em vez de <c>Nothing</c> porque <c>Nothing</c> vira
    ''' <c>NullReferenceException</c> em algum caminho esquecido, e "explodiu"
    ''' e "recusou por decisão" não são a mesma coisa para quem lê depois.
    ''' </summary>
    Public NotInheritable Class AssistenteIndisponivel
        Implements IAssistantProvider

        Public ReadOnly Property Destino As AssistDestination _
                                 Implements IAssistantProvider.Destino
            Get
                Return New AssistDestination("", "", "")
            End Get
        End Property

        Public Function Pronto() As Boolean Implements IAssistantProvider.Pronto
            Return False
        End Function

        ''' <summary>Não prepara nada: não há para onde mandar.</summary>
        Public Function Preparar(envelope As Byte()) As Byte() _
                                 Implements IAssistantProvider.Preparar
            Return Nothing
        End Function

        Public Function Enviar(bytes As Byte(), ct As CancellationToken) As ProviderOutcome _
                               Implements IAssistantProvider.Enviar
            Return New ProviderOutcome(ProviderStatus.NaoComecou, "")
        End Function

    End Class

End Namespace

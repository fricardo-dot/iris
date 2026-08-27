Imports System.Threading
Imports Iris.Model

Namespace Global.Iris.Assist

    ''' <summary>Como a operação inteira terminou, do ponto de vista de quem pediu.</summary>
    Public Enum AssistOutcomeKind
        ''' <summary>Zero: não decidido. Nunca significa sucesso.</summary>
        Desconhecido = 0
        ''' <summary>Veio resposta.</summary>
        Respondeu
        ''' <summary>O portão negou. Nada saiu.</summary>
        Negado
        ''' <summary>O cofre recusou a capability. Nada saiu.</summary>
        Recusado
        ''' <summary>
        ''' O diário não pôde registrar a <b>intenção</b>, ou o voo. Nada saiu —
        ''' e isso se sabe, porque a transmissão não chegou a ser tentada.
        ''' </summary>
        SemDiario

        ''' <summary>
        ''' <b>Pode ter saído, e o diário não conseguiu fechar o registro.</b>
        '''
        ''' Existe porque <see cref="SemDiario"/> passou a mentir: ele dizia
        ''' "nada saiu", e era devolvido também quando <c>Concluir</c> ou
        ''' <c>Falhar</c> falhava <b>depois</b> do HTTP — quando conteúdo pode ter
        ''' saído.
        '''
        ''' A UI precisa dizer as duas coisas: que pode ter saído, e que o
        ''' registro do desfecho não foi gravado. "Erro de diário" sozinho
        ''' esconderia a primeira metade.
        ''' </summary>
        AmbiguoSemFechamentoDoDiario
        ''' <summary>Falhou, e <b>não</b> chegou a começar.</summary>
        NaoComecou
        ''' <summary>
        ''' Falhou depois de começar. <b>Pode ter chegado</b>, e ninguém vai
        ''' saber.
        ''' </summary>
        Ambiguo
    End Enum

    Public NotInheritable Class AssistOutcome
        Public ReadOnly Property Kind As AssistOutcomeKind
        ''' <summary>Texto passivo do modelo. Vazio quando não houve resposta.</summary>
        Public ReadOnly Property Texto As String
        Public ReadOnly Property RequestId As Guid
        Public ReadOnly Property Nota As DisclosureNote
        Public ReadOnly Property MotivoDoPortao As DisclosureReason

        ''' <summary>
        ''' <b>O código HTTP, quando o provedor chegou a responder.</b>
        '''
        ''' Vai para a tela junto com o aviso. Sem ele, "não dá para saber se o
        ''' conteúdo chegou" é tudo o que o usuário vê — e descobrir se o caso
        ''' era credencial ou roteamento exigiu, uma vez, três ferramentas de
        ''' linha de comando e o banco aberto na mão.
        ''' </summary>
        Public ReadOnly Property CodigoHttp As Integer?

        ''' <summary>O que a chamada custou, quando o provedor contou.</summary>
        Public ReadOnly Property Custo As Decimal?
        ''' <summary>Tokens somados, quando o provedor contou.</summary>
        Public ReadOnly Property Tokens As Integer?

        Friend Sub New(kind As AssistOutcomeKind, texto As String, requestId As Guid,
                       nota As DisclosureNote, motivoDoPortao As DisclosureReason,
                       Optional codigoHttp As Integer? = Nothing,
                       Optional custo As Decimal? = Nothing,
                       Optional tokens As Integer? = Nothing)
            Me.Kind = kind
            Me.Texto = If(texto, "")
            Me.RequestId = requestId
            Me.Nota = nota
            Me.MotivoDoPortao = motivoDoPortao
            ' SEM normalizar de novo: o que chega aqui vem de
            ' ProviderOutcome.Codigo, ja filtrado na entrada. Repetir o filtro
            ' em cada camada faz parecer que cada uma tem uma regra propria --
            ' e no dia em que uma delas mudar, elas divergem em silencio.
            Me.CodigoHttp = codigoHttp
            Me.Custo = custo
            Me.Tokens = tokens
        End Sub
    End Class

    ' ==================================================================

    ''' <summary>
    ''' <b>O serviço de aplicação: executar a operação segura.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A ORDEM É A GARANTIA</b>
    '''
    '''   1. o portão decide, e emite o grant;
    '''   2. o cofre emite a capability sobre <b>aqueles</b> bytes;
    '''   3. o diário grava a <b>intenção</b>, durável, com o hash;
    '''   4. o cofre <b>consome</b> a capability — conferindo bytes, itens,
    '''      versões, destino e operação;
    '''   5. o diário registra o <b>início do voo</b>, e só se ele
    '''      <b>confirmar</b> é que se toca na rede;
    '''   6. o provedor manda;
    '''   7. o diário conclui, ou registra a falha — ambígua quando for.
    '''
    ''' Cada passo que falha para tudo, e o diário fica dizendo <b>onde</b>
    ''' parou. Nenhum deles é opcional, e nenhum pode ser reordenado sem abrir
    ''' um buraco que já foi aberto antes.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O CUSTO DECLARADO DE MARCAR O VOO ANTES</b>
    '''
    ''' Registrar "em voo" antes de chamar o transporte faz uma falha que
    ''' comprovadamente não enviou nada — conexão recusada, por exemplo — ser
    ''' contada como <b>ambígua</b>. Isso infla o número de ambíguos.
    '''
    ''' A alternativa seria marcar depois do primeiro byte, e para isso seria
    ''' preciso confiar no transporte para dizer quando o byte saiu. Quem erra
    ''' nessa direção esconde egress; quem erra nesta conta a mais. Está
    ''' escolhido, e o preço fica escrito.
    '''
    ''' O que dá para tirar desse preço sem trapacear é o que
    ''' <see cref="IAssistantProvider.Pronto"/> tira: recusas que se sabem
    ''' <b>antes</b> de qualquer byte — endereço não-HTTPS, credencial ausente,
    ''' provedor nenhum — são perguntadas antes de o voo ser marcado, e por isso
    ''' não entram na contagem de ambíguos.
    ''' </summary>
    Public NotInheritable Class AssistTransmitter

        Private ReadOnly _politica As DisclosurePolicy
        Private ReadOnly _cofre As CapabilityLedger
        Private ReadOnly _diario As IDisclosureJournal
        Private ReadOnly _provedor As IAssistantProvider
        Private ReadOnly _relogio As Func(Of DateTimeOffset)

        Public Sub New(politica As DisclosurePolicy, cofre As CapabilityLedger,
                       diario As IDisclosureJournal, provedor As IAssistantProvider,
                       relogio As Func(Of DateTimeOffset))
            _politica = politica
            _cofre = cofre
            _diario = diario
            _provedor = provedor
            _relogio = relogio
        End Sub

        ''' <summary>
        ''' Faz a operação inteira, ou para no primeiro passo que não passar.
        ''' </summary>
        Public Function Executar(pedido As PreflightRequest,
                                 classificar As Func(Of IReadOnlyList(Of MessageClassification)),
                                 montar As Func(Of EnvelopeResult),
                                 ct As CancellationToken) As AssistOutcome

            ' 1. o portao. O classificador so e INVOCADO se o preflight passar
            '    — quem garante isso e o DisclosureGate, e o motivo esta la.
            Dim portao As New DisclosureGate(_politica, _relogio)
            Dim decisao = portao.Avaliar(pedido, classificar)
            If Not decisao.Permitido Then
                Return Parar(AssistOutcomeKind.Negado, DisclosureNote.PortaoNegou, decisao.Motivo)
            End If

            Dim env = montar()
            If Not env.Ok Then
                Return Parar(AssistOutcomeKind.Recusado, DisclosureNote.EnvelopeRecusado)
            End If

            ' 2. O CORPO, traduzido para o formato do provedor — local, puro,
            '    sem rede, e ANTES da capability.
            '
            '    A ordem e o ponto. Traduzir depois de autorizar faria a
            '    capability cobrir o envelope e a rede transportar outra coisa:
            '    a autorizacao falaria de um artefato e o fio levaria outro.
            Dim corpo As Byte()
            Try
                Dim devolvido = _provedor.Preparar(env.Envelope.Bytes())
                ' COPIA, E A COPIA E NOSSA.
                '
                ' O arranjo devolvido foi criado pelo provedor, e nada impede
                ' que ele guarde a referencia e mexa depois — em Pronto(), que
                ' roda DEPOIS do consumo da capability, ou de outra thread. O
                ' hash seria de um conteudo e o fio levaria outro, que e
                ' exatamente o furo que a capability sobre o corpo existe para
                ' fechar.
                '
                ' A partir daqui so a copia e usada: hash, consumo e envio.
                corpo = If(devolvido Is Nothing, Nothing, CType(devolvido.Clone(), Byte()))
            Catch
                corpo = Nothing
            End Try
            If corpo Is Nothing OrElse corpo.Length = 0 Then
                Return Parar(AssistOutcomeKind.Recusado, DisclosureNote.CorpoNaoPreparado)
            End If

            ' 3. a capability, sobre AQUELES bytes — o envelope e o corpo.
            Dim agora = _relogio()
            Dim cap = _cofre.Emitir(decisao, env.Envelope, corpo, agora)
            If cap Is Nothing Then
                Return Parar(AssistOutcomeKind.Recusado, DisclosureNote.CapabilityRecusada)
            End If

            ' 3. a INTENCAO, duravel, antes de qualquer tentativa.
            '
            '    Sem registro nao se transmite. Um envio sem rastro e pior que
            '    um envio que nao acontece — e o diario pode falhar de duas
            '    formas: devolvendo False, ou LANCANDO (disco cheio, banco
            '    travado, I/O). Conferir so o Boolean deixaria a segunda
            '    escapar, e ai a excecao subiria de um ponto em que nada saiu.
            If Not Duravel(Function() _diario.Intencao(cap, agora)) Then
                Return SemDiario(cap.RequestId, DisclosureNote.Nenhuma)
            End If

            ' 4. o consumo — e ele confere bytes, itens, versoes, destino e
            '    operacao contra o que foi autorizado.
            Dim uso = _cofre.Consumir(cap, env.Envelope, corpo, _provedor.Destino,
                                      pedido.Operacao, _relogio())
            If Not uso.Autorizado Then
                ' Nada saiu — mas se o diario nao registrar isso, o registro
                ' fica "Intencionada" e a tela diria "recusado" sobre uma linha
                ' que nao conta a mesma historia.
                If Not Duravel(Function() _diario.NaoEnviou(
                        cap.RequestId, _relogio(), DisclosureNote.CapabilityRecusada)) Then
                    Return SemDiario(cap.RequestId, DisclosureNote.CapabilityRecusada)
                End If
                Return New AssistOutcome(AssistOutcomeKind.Recusado, "", cap.RequestId,
                                         DisclosureNote.CapabilityRecusada,
                                         DisclosureReason.NaoDecidido)
            End If

            ' 5a. o provedor esta pronto? A pergunta vem ANTES de marcar o
            '     voo, e sem tocar na rede: endereco que nao e HTTPS,
            '     credencial ausente e provedor nenhum sao recusas que se SABEM,
            '     e marca-las como ambiguas encheria de ruido a contagem que a
            '     UI mostra.
            '
            '     A chamada vai dentro de Try: um provedor que EXPLODE ao ser
            '     perguntado nao pode escapar da maquina de estados. E a nota e
            '     ProvedorIndisponivel, nao CapabilityRecusada - a capability
            '     foi consumida com sucesso no passo anterior, e credencial
            '     ausente nao e recusa do cofre.
            Dim pronto As Boolean
            Try
                pronto = _provedor.Pronto()
            Catch
                pronto = False
            End Try

            If Not pronto Then
                If Not Duravel(Function() _diario.NaoEnviou(
                        cap.RequestId, _relogio(), DisclosureNote.ProvedorIndisponivel)) Then
                    Return SemDiario(cap.RequestId, DisclosureNote.ProvedorIndisponivel)
                End If
                Return New AssistOutcome(AssistOutcomeKind.NaoComecou, "", cap.RequestId,
                                         DisclosureNote.ProvedorIndisponivel,
                                         DisclosureReason.NaoDecidido)
            End If

            ' 5b. o VOO, e so depois de ele confirmar e que se toca na rede.
            If Not Duravel(Function() _diario.Iniciando(cap.RequestId, _relogio())) Then
                Return SemDiario(cap.RequestId, DisclosureNote.Nenhuma)
            End If

            ' 6. a rede.
            '
            '    Excecao do provedor DEPOIS do voo nao pode escapar: escapando,
            '    a linha fica EmVoo, quem pediu recebe uma excecao em vez de um
            '    desfecho, e o texto dela — que pode carregar host, caminho, ou
            '    pedaco do que foi enviado — atravessa a fronteira.
            Dim r As ProviderOutcome
            Try
                r = _provedor.Enviar(corpo, ct)
            Catch
                r = New ProviderOutcome(ProviderStatus.ConexaoCaiu, "")
            End Try

            ' 7. o desfecho.
            '
            ' DEPOIS DO VOO, TODO INSUCESSO E AMBIGUO. Inclusive um provedor
            ' dizendo NaoComecou: ele prometeu estar pronto no passo 5a, e uma
            ' promessa quebrada nao vira autoridade sobre o que saiu. Pronto()
            ' e otimizacao ANTES do voo, nao palavra final depois.
            '
            ' Devolver NaoComecou aqui faria a UI dizer uma coisa e o diario
            ' registrar outra — e a UI e o diario nao podem discordar sobre se
            ' conteudo pode ter saido.
            If r.Status = ProviderStatus.Respondeu Then
                If Not Duravel(Function() _diario.Concluir(cap.RequestId, _relogio(),
                                                           r.Codigo)) Then
                    ' O HTTP respondeu e o diario nao fechou. Dizer "respondeu"
                    ' deixaria o registro em voo para sempre, e a reconciliacao
                    ' da proxima abertura marcaria ambiguo — a UI teria dito
                    ' sucesso sobre algo que o diario chama de incerto.
                    '
                    ' E o desfecho NAO e "erro de diario": conteudo saiu. Quem
                    ' le tem de ver as duas coisas.
                    ' O CODIGO VAI JUNTO TAMBEM AQUI.
                    '
                    ' Este ramo cai na MESMA frase da faixa que o ramo de
                    ' falha -- e sem o numero ele era o unico caminho
                    ' pos-resposta em que a tela perdia o diagnostico. Um
                    ' desfecho que a UI trata igual nao pode carregar menos
                    ' informacao que o irmao.
                    Return New AssistOutcome(AssistOutcomeKind.AmbiguoSemFechamentoDoDiario,
                                             "", cap.RequestId,
                                             DisclosureNote.Nenhuma, DisclosureReason.NaoDecidido,
                                             r.Codigo)
                End If
                Return New AssistOutcome(AssistOutcomeKind.Respondeu, r.Texto, cap.RequestId,
                                         DisclosureNote.Nenhuma, DisclosureReason.NaoDecidido,
                                         r.Codigo, r.Custo, r.Tokens)
            End If

            ' O CODIGO HTTP VAI PARA O DIARIO.
            '
            ' Sem ele, "ProvedorRecusou" nao distingue "a chave nao vale" de
            ' "nenhum provedor atende a esta politica de dados" -- e as duas
            ' levam a acoes opostas. Ja custou tres ferramentas de linha de
            ' comando para descobrir, por fora, o que esta linha devia ter
            ' contado. O provedor ja devolvia o numero; quem o jogava fora era
            ' este ponto aqui.
            Dim nota = NotaDe(r.Status)
            If Not Duravel(Function() _diario.Falhar(cap.RequestId, _relogio(), nota,
                                                     podeTerChegado:=True,
                                                     codigoHttp:=r.Codigo)) Then
                Return New AssistOutcome(AssistOutcomeKind.AmbiguoSemFechamentoDoDiario,
                                         "", cap.RequestId, nota, DisclosureReason.NaoDecidido,
                                         r.Codigo)
            End If

            Return New AssistOutcome(AssistOutcomeKind.Ambiguo, "", cap.RequestId,
                                     nota, DisclosureReason.NaoDecidido, r.Codigo)
        End Function

        ' ==============================================================

        ''' <summary>
        ''' Para antes de haver capability — então <b>não há o que registrar</b>.
        '''
        ''' O diário registra divulgações; um pedido que o portão negou antes de
        ''' qualquer envelope não é uma divulgação, e inventar uma linha para ele
        ''' encheria o diário de coisas que não aconteceram — justamente onde
        ''' alguém vai procurar o que aconteceu.
        ''' </summary>
        Private Shared Function Parar(kind As AssistOutcomeKind, nota As DisclosureNote,
                                      Optional motivo As DisclosureReason =
                                          DisclosureReason.NaoDecidido) As AssistOutcome
            Return New AssistOutcome(kind, "", Guid.Empty, nota, motivo)
        End Function

        ''' <summary>
        ''' Um passo do diário: <c>True</c> só se ele <b>disse</b> que gravou.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>LANÇAR TAMBÉM É NÃO GRAVAR</b>
        '''
        ''' Os passos devolvem <c>Boolean</c>, e conferir só isso deixava escapar
        ''' a outra metade: disco cheio, banco travado, I/O falhando — o SQLite
        ''' <b>lança</b>, e a exceção subiria de um ponto onde a máquina de
        ''' estados sabe exatamente o que dizer.
        '''
        ''' Aqui as duas viram a mesma resposta: não gravou.
        ''' </summary>
        Private Shared Function Duravel(passo As Func(Of Boolean)) As Boolean
            Try
                Return passo()
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' O diário não gravou, e a transmissão <b>não foi tentada</b>.
        '''
        ''' Distinto de <see cref="AssistOutcomeKind.AmbiguoSemFechamentoDoDiario"/>,
        ''' que é o mesmo problema depois de o conteúdo poder ter saído.
        ''' </summary>
        Private Shared Function SemDiario(requestId As Guid,
                                          nota As DisclosureNote) As AssistOutcome
            Return New AssistOutcome(AssistOutcomeKind.SemDiario, "", requestId,
                                     nota, DisclosureReason.NaoDecidido)
        End Function

        Private Shared Function NotaDe(s As ProviderStatus) As DisclosureNote
            Select Case s
                Case ProviderStatus.Timeout : Return DisclosureNote.Timeout
                Case ProviderStatus.Cancelado : Return DisclosureNote.Cancelado
                Case ProviderStatus.Recusou, ProviderStatus.RespostaGrandeDemais
                    Return DisclosureNote.ProvedorRecusou
                Case ProviderStatus.RespostaIlegivel
                    Return DisclosureNote.RespostaIlegivel
                Case Else : Return DisclosureNote.ConexaoCaiu
            End Select
        End Function

    End Class

End Namespace

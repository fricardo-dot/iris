Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Model

Namespace Global.Iris.Assist

    ''' <summary>
    ''' <b>A autorização de sair, presa a bytes concretos.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE UM VEREDITO NÃO BASTA</b>
    '''
    ''' O portão diz "este conteúdo pode sair". Entre esse "pode" e o envio
    ''' cabem várias coisas que ninguém confere: o veredito é reutilizado para
    ''' outro texto, o texto muda depois de aprovado, o destino é trocado, ou o
    ''' mesmo veredito serve para dois envios.
    '''
    ''' Uma capability fecha isso ao <b>não descrever conteúdo</b>: ela
    ''' descreve o <see cref="AssistEnvelope.Hash"/> de um buffer específico, o
    ''' comprimento dele, o destino, a operação, a ativação que a permitiu, e
    ''' um prazo curto. Bytes diferentes, capability diferente.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ELA NÃO CARREGA</b>
    '''
    ''' Texto. Nem o conteúdo nem parte dele — só o hash. Uma capability
    ''' persistida ou registrada em log não pode virar mais uma cópia do
    ''' e-mail.
    ''' </summary>
    Public NotInheritable Class DisclosureCapability

        Public ReadOnly Property Id As Guid
        Public ReadOnly Property AtivacaoId As String
        Public ReadOnly Property AtivacaoVersao As Integer
        Public ReadOnly Property Operacao As AssistOperation
        Public ReadOnly Property Destino As AssistDestination
        ''' <summary>SHA-256 dos bytes autorizados, em hex minúsculo.</summary>
        Public ReadOnly Property Hash As String
        Public ReadOnly Property Comprimento As Integer
        Public ReadOnly Property Itens As IReadOnlyList(Of ItemKey)
        Public ReadOnly Property Emitida As DateTimeOffset
        Public ReadOnly Property Expira As DateTimeOffset

        ''' <summary>
        ''' <c>Friend</c>, e é o ponto: só o <see cref="CapabilityLedger"/>
        ''' emite. Uma capability que qualquer um pudesse construir seria uma
        ''' afirmação de que houve autorização, sem que tivesse havido.
        ''' </summary>
        Friend Sub New(id As Guid, ativacaoId As String, ativacaoVersao As Integer,
                       operacao As AssistOperation, destino As AssistDestination,
                       hash As String, comprimento As Integer,
                       itens As IReadOnlyList(Of ItemKey),
                       emitida As DateTimeOffset, expira As DateTimeOffset)
            Me.Id = id
            Me.AtivacaoId = ativacaoId
            Me.AtivacaoVersao = ativacaoVersao
            Me.Operacao = operacao
            Me.Destino = destino
            Me.Hash = hash
            Me.Comprimento = comprimento
            Me.Itens = itens
            Me.Emitida = emitida
            Me.Expira = expira
        End Sub

    End Class

    ''' <summary>Por que o consumo foi recusado.</summary>
    Public Enum CapabilityRefusal
        Nenhuma = 0
        ''' <summary>Capability que este cofre não emitiu.</summary>
        Desconhecida
        ''' <summary>Já foi usada. Consumo é único.</summary>
        JaConsumida
        ''' <summary>O prazo passou.</summary>
        Expirada
        ''' <summary>Os bytes não são os que ela autorizou.</summary>
        BytesDiferentes
        ''' <summary>O destino do envio não é o que ela autorizou.</summary>
        DestinoDiferente
        ''' <summary>A operação não é a que ela autorizou.</summary>
        OperacaoDiferente
        ''' <summary>O envelope não confere consigo mesmo.</summary>
        EnvelopeCorrompido
    End Enum

    Public NotInheritable Class CapabilityUse
        Public ReadOnly Property Autorizado As Boolean
        Public ReadOnly Property Recusa As CapabilityRefusal

        Friend Sub New(autorizado As Boolean, recusa As CapabilityRefusal)
            Me.Autorizado = autorizado
            Me.Recusa = recusa
        End Sub
    End Class

    ' ==================================================================

    ''' <summary>
    ''' Emite capabilities e as consome <b>uma vez só</b>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE CONSUMO ÚNICO PRECISA DE ESTADO</b>
    '''
    ''' Uma capability imutável não sabe se já foi usada. Sem alguém guardando
    ''' isso, "consumo único" seria um comentário: o mesmo objeto autorizaria
    ''' dois envios, e o segundo mandaria o mesmo conteúdo de novo sem
    ''' aparecer em lugar nenhum.
    '''
    ''' O cofre é o que sabe. E ele guarda <b>id</b> — não bytes, não hash de
    ''' conteúdo que alguém possa cruzar depois.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A CONFERÊNCIA DO CONSUMO</b>
    '''
    ''' <see cref="Consumir"/> recebe o envelope <b>de novo</b> e confere que
    ''' o hash bate. Não é redundância com a emissão: entre uma e outra pode
    ''' ter passado uma classificação, uma troca de tela, ou um bug — e a
    ''' pergunta "os bytes que vão sair são os que foram autorizados?" só tem
    ''' resposta honesta no instante de sair.
    ''' </summary>
    Public NotInheritable Class CapabilityLedger

        ''' <summary>
        ''' Quanto tempo uma capability vale. Curto de propósito: ela existe
        ''' para atravessar a distância entre autorizar e enviar, não para ser
        ''' guardada.
        ''' </summary>
        Public Shared ReadOnly Validade As TimeSpan = TimeSpan.FromMinutes(2)

        Private ReadOnly _emitidas As New ConcurrentDictionary(Of Guid, DisclosureCapability)()
        Private ReadOnly _consumidas As New ConcurrentDictionary(Of Guid, Byte)()

        ''' <summary>
        ''' Emite. <b>Só depois de o portão permitir</b> — quem chama tem de ter
        ''' a decisão em mãos, e uma decisão negada não emite nada.
        ''' </summary>
        Public Function Emitir(decisao As DisclosureDecision, ativacao As ActivationRecord,
                               pedido As PreflightRequest, envelope As AssistEnvelope,
                               agora As DateTimeOffset) As DisclosureCapability

            If decisao Is Nothing OrElse Not decisao.Permitido Then Return Nothing
            If ativacao Is Nothing OrElse envelope Is Nothing OrElse pedido Is Nothing Then
                Return Nothing
            End If

            Dim c As New DisclosureCapability(
                Guid.NewGuid(), ativacao.Id, ativacao.Versao,
                pedido.Operacao, pedido.Destino,
                envelope.Hash, envelope.Comprimento, envelope.Itens,
                agora, agora + Validade)

            _emitidas(c.Id) = c
            Return c
        End Function

        ''' <summary>
        ''' Confere e <b>gasta</b>. Chamada imediatamente antes de transmitir.
        '''
        ''' A ordem das conferências importa: identidade primeiro, depois
        ''' consumo, depois prazo, e só então os bytes. Conferir os bytes de uma
        ''' capability que não é deste cofre seria trabalho sobre um objeto que
        ''' já não vale.
        ''' </summary>
        Public Function Consumir(c As DisclosureCapability, envelope As AssistEnvelope,
                                 destino As AssistDestination, operacao As AssistOperation,
                                 agora As DateTimeOffset) As CapabilityUse

            If c Is Nothing OrElse Not _emitidas.ContainsKey(c.Id) Then
                Return Recusar(CapabilityRefusal.Desconhecida)
            End If

            ' Marca o consumo ANTES de qualquer outra coisa poder falhar por
            ' motivo transitorio. Uma capability que fosse "devolvida" quando a
            ' conferencia falha viraria um oraculo: da para testar hash ate
            ' acertar.
            If Not _consumidas.TryAdd(c.Id, 0) Then
                Return Recusar(CapabilityRefusal.JaConsumida)
            End If

            If agora > c.Expira Then Return Recusar(CapabilityRefusal.Expirada)

            If envelope Is Nothing Then Return Recusar(CapabilityRefusal.BytesDiferentes)
            If Not envelope.Integro() Then Return Recusar(CapabilityRefusal.EnvelopeCorrompido)

            If Not String.Equals(envelope.Hash, c.Hash, StringComparison.Ordinal) OrElse
               envelope.Comprimento <> c.Comprimento Then
                Return Recusar(CapabilityRefusal.BytesDiferentes)
            End If

            If operacao <> c.Operacao Then Return Recusar(CapabilityRefusal.OperacaoDiferente)

            If destino Is Nothing OrElse c.Destino Is Nothing OrElse
               Not String.Equals(destino.Endpoint.Trim(), c.Destino.Endpoint.Trim(),
                                 StringComparison.Ordinal) OrElse
               Not String.Equals(destino.Provedor.Trim(), c.Destino.Provedor.Trim(),
                                 StringComparison.Ordinal) OrElse
               Not String.Equals(destino.Modelo.Trim(), c.Destino.Modelo.Trim(),
                                 StringComparison.Ordinal) Then
                Return Recusar(CapabilityRefusal.DestinoDiferente)
            End If

            Return New CapabilityUse(True, CapabilityRefusal.Nenhuma)
        End Function

        Private Shared Function Recusar(r As CapabilityRefusal) As CapabilityUse
            Return New CapabilityUse(False, r)
        End Function

    End Class

End Namespace

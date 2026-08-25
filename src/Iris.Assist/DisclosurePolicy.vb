Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Model

Namespace Global.Iris.Assist

    ''' <summary>Por que o portão negou. Um motivo, o primeiro que valeu.</summary>
    Public Enum DisclosureReason
        ''' <summary>Zero: não decidido. Nunca significa permitido.</summary>
        NaoDecidido = 0

        ''' <summary>Não há autorização nenhuma. É o estado da produção.</summary>
        SemAtivacao
        ''' <summary>Há autorização, e ela está incompleta.</summary>
        AtivacaoIncompleta
        ''' <summary>Há autorização, e ela venceu ou ainda não vale.</summary>
        AtivacaoForaDeVigencia
        ''' <summary>O endpoint autorizado não é HTTPS.</summary>
        EndpointInseguro
        ''' <summary>A operação pedida não está na autorização.</summary>
        OperacaoNaoAutorizada
        ''' <summary>O provedor ou o modelo pedido não é o autorizado.</summary>
        ProvedorNaoAutorizado
        ''' <summary>A pasta não está na lista explícita.</summary>
        PastaNaoAutorizada
        ''' <summary>Não há mensagem nenhuma no pedido.</summary>
        PedidoVazio
        ''' <summary>Alguma mensagem do pedido está fora da pasta autorizada.</summary>
        MensagemDeOutraPasta

        ''' <summary>O desfecho da leitura do rótulo não está na lista aceita.</summary>
        LeituraNaoAceita
        ''' <summary>Há rótulo ativo cujo GUID não está na lista permitida.</summary>
        RotuloNaoPermitido
        ''' <summary>O <c>ContentBits</c> do registro não está na lista aceita.</summary>
        ContentBitsNaoAceito
        ''' <summary>O <c>ContentBits</c> não veio, ou veio ilegível.</summary>
        ContentBitsDesconhecido
        ''' <summary>Falta evidência de versão do item.</summary>
        SemEvidenciaDeVersao
        ''' <summary>Anexo — fora do escopo da fase, por inteiro.</summary>
        AnexoForaDeEscopo
    End Enum

    ''' <summary>O veredito. <b>Nasce negado</b>, e é preciso prova para virar.</summary>
    Public NotInheritable Class DisclosureDecision

        Public ReadOnly Property Permitido As Boolean
        Public ReadOnly Property Motivo As DisclosureReason
        ''' <summary>O item que motivou a negativa, quando foi um item.</summary>
        Public ReadOnly Property Culpado As ItemKey
        ''' <summary>Explicação em português, para o usuário ler.</summary>
        Public ReadOnly Property Explicacao As String

        Private Sub New(permitido As Boolean, motivo As DisclosureReason,
                        culpado As ItemKey, explicacao As String)
            Me.Permitido = permitido
            Me.Motivo = motivo
            Me.Culpado = culpado
            Me.Explicacao = explicacao
        End Sub

        Friend Shared Function Negar(motivo As DisclosureReason, explicacao As String,
                                     Optional culpado As ItemKey = Nothing) As DisclosureDecision
            Return New DisclosureDecision(False, motivo, culpado, explicacao)
        End Function

        Friend Shared Function Permitir() As DisclosureDecision
            Return New DisclosureDecision(True, DisclosureReason.NaoDecidido, Nothing, "")
        End Function

    End Class

    ''' <summary>Uma mensagem do pedido, já classificada.</summary>
    Public NotInheritable Class MessageClassification
        Public ReadOnly Property Item As ItemKey
        Public ReadOnly Property Pasta As FolderKey
        Public ReadOnly Property Leitura As LabelReading
        ''' <summary>Tem anexo. Anexo está fora desta fase, e nega.</summary>
        Public ReadOnly Property TemAnexo As Boolean

        Public Sub New(item As ItemKey, pasta As FolderKey, leitura As LabelReading,
                       Optional temAnexo As Boolean = False)
            Me.Item = item
            Me.Pasta = pasta
            Me.Leitura = leitura
            Me.TemAnexo = temAnexo
        End Sub
    End Class

    ''' <summary>O que se quer divulgar.</summary>
    Public NotInheritable Class DisclosureRequest
        Public ReadOnly Property Operacao As AssistOperation
        Public ReadOnly Property Pasta As FolderKey
        Public ReadOnly Property Provedor As String
        Public ReadOnly Property Modelo As String
        Public ReadOnly Property Mensagens As IReadOnlyList(Of MessageClassification)

        Public Sub New(operacao As AssistOperation, pasta As FolderKey,
                       provedor As String, modelo As String,
                       mensagens As IEnumerable(Of MessageClassification))
            Me.Operacao = operacao
            Me.Pasta = pasta
            Me.Provedor = If(provedor, "")
            Me.Modelo = If(modelo, "")
            Me.Mensagens = If(mensagens Is Nothing,
                              CType(Array.Empty(Of MessageClassification)(),
                                    IReadOnlyList(Of MessageClassification)),
                              mensagens.ToList())
        End Sub
    End Class

    ' ==================================================================

    ''' <summary>
    ''' <b>O portão: conteúdo desta caixa pode sair desta máquina?</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>PERMISSÃO É CONJUNÇÃO FECHADA DE PROVAS POSITIVAS</b>
    '''
    ''' Nunca "não achei motivo suficiente para negar". A diferença não é
    ''' estilística: um portão escrito como lista de negativas libera todo
    ''' caso que ninguém pensou em proibir, e o caso que ninguém pensou é
    ''' exatamente o que vaza.
    '''
    ''' Então cada prova é uma pergunta cuja resposta tem de ser <b>sim</b>:
    ''' existe autorização; ela está completa; está vigente; o endpoint é
    ''' HTTPS; a operação está listada; o provedor e o modelo são os
    ''' autorizados; a pasta está listada; há mensagem; e cada mensagem passa
    ''' por todas as provas dela.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>UM MEMBRO QUE NÃO PASSA NEGA A THREAD INTEIRA</b>
    '''
    ''' Não a mensagem — a thread. Resumo parcial é fácil demais de confundir
    ''' com resumo completo, e o usuário não tem como saber que faltou pedaço.
    ''' É a regra da §29.1.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A ASSIMETRIA DA P16</b>
    '''
    ''' A medição do 3.0 mostrou que <c>MSIP_Labels</c> mora no namespace de
    ''' cabeçalhos de internet. Cabeçalho recebido pode ter origem fora do
    ''' mecanismo corporativo, então ler o valor com perfeição <b>não prova</b>
    ''' que ninguém apresenta uma classificação baixa falsa.
    '''
    ''' Por isso rótulo <b>nunca autoriza sozinho</b>. Ele entra como mais uma
    ''' condição a satisfazer, e o que autoriza é a autorização — que uma
    ''' pessoa emitiu, com a política corporativa na mão.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E O QUE A PRODUÇÃO TEM HOJE</b>
    '''
    ''' <see cref="ActivationRecord.DaProducao"/> é <c>Nothing</c>, então em
    ''' produção este portão nega <b>tudo</b>, sempre, com
    ''' <see cref="DisclosureReason.SemAtivacao"/>. Isso é a §28.2, não uma
    ''' pendência.
    ''' </summary>
    Public NotInheritable Class DisclosurePolicy

        Private ReadOnly _ativacao As ActivationRecord

        Public Sub New(ativacao As ActivationRecord)
            _ativacao = ativacao
        End Sub

        ''' <summary>O portão como a produção o tem: sem autorização nenhuma.</summary>
        Public Shared Function DaProducao() As DisclosurePolicy
            Return New DisclosurePolicy(ActivationRecord.DaProducao)
        End Function

        ' ==============================================================

        Public Function Decidir(pedido As DisclosureRequest,
                                agora As DateTimeOffset) As DisclosureDecision

            If pedido Is Nothing Then
                Return DisclosureDecision.Negar(DisclosureReason.PedidoVazio,
                                                "Não há nada a enviar.")
            End If

            ' ---- as provas sobre a AUTORIZAÇÃO -----------------------
            '
            ' Vêm primeiro, e por um motivo prático além do lógico: sem
            ' autorização não se gasta uma ida ao COM lendo rótulo de coisa
            ' nenhuma. Classificar para depois descobrir que estava tudo
            ' desligado seria pagar ~17 ms por item para nada.
            If _ativacao Is Nothing Then
                Return DisclosureDecision.Negar(DisclosureReason.SemAtivacao,
                    "A IA externa não está habilitada. Nada deste computador é enviado " &
                    "para fora enquanto você não autorizar, com a política da empresa " &
                    "e um provedor à sua escolha.")
            End If

            If Not _ativacao.Completo() Then
                Return DisclosureDecision.Negar(DisclosureReason.AtivacaoIncompleta,
                    "A autorização está incompleta — falta declarar quem autorizou, o " &
                    "provedor, o modelo, a retenção aceita, as operações ou as pastas.")
            End If

            If Not _ativacao.Vigente(agora) Then
                Return DisclosureDecision.Negar(DisclosureReason.AtivacaoForaDeVigencia,
                    "A autorização não está vigente nesta data.")
            End If

            If Not _ativacao.EndpointSeguro() Then
                Return DisclosureDecision.Negar(DisclosureReason.EndpointInseguro,
                    "O endereço autorizado não é HTTPS.")
            End If

            If Not _ativacao.Operacoes.Contains(pedido.Operacao) Then
                Return DisclosureDecision.Negar(DisclosureReason.OperacaoNaoAutorizada,
                    "Esta operação não está entre as autorizadas.")
            End If

            If Not Igual(pedido.Provedor, _ativacao.Provedor) OrElse
               Not Igual(pedido.Modelo, _ativacao.Modelo) Then
                Return DisclosureDecision.Negar(DisclosureReason.ProvedorNaoAutorizado,
                    "O provedor ou o modelo do pedido não é o autorizado.")
            End If

            If Not _ativacao.Pastas.Any(Function(f) MesmaPasta(f, pedido.Pasta)) Then
                Return DisclosureDecision.Negar(DisclosureReason.PastaNaoAutorizada,
                    "Esta pasta não está entre as autorizadas.")
            End If

            If pedido.Mensagens.Count = 0 Then
                Return DisclosureDecision.Negar(DisclosureReason.PedidoVazio,
                    "Não há nada a enviar.")
            End If

            ' ---- e as provas sobre CADA mensagem ----------------------
            For Each m In pedido.Mensagens
                Dim v = Conferir(m, pedido.Pasta)
                If Not v.Permitido Then Return v
            Next

            Return DisclosureDecision.Permitir()
        End Function

        ' ==============================================================

        ''' <summary>
        ''' As provas de uma mensagem. Falhar aqui nega o pedido <b>inteiro</b>.
        ''' </summary>
        Private Function Conferir(m As MessageClassification,
                                  pasta As FolderKey) As DisclosureDecision

            If m Is Nothing OrElse m.Leitura Is Nothing Then
                Return DisclosureDecision.Negar(DisclosureReason.LeituraNaoAceita,
                    "Uma das mensagens não foi classificada.")
            End If

            If Not MesmaPasta(m.Pasta, pasta) Then
                Return DisclosureDecision.Negar(DisclosureReason.MensagemDeOutraPasta,
                    "Uma das mensagens está em outra pasta que não a autorizada.",
                    m.Item)
            End If

            If m.TemAnexo Then
                Return DisclosureDecision.Negar(DisclosureReason.AnexoForaDeEscopo,
                    "Mensagem com anexo. Anexo não é tratado nesta fase, por inteiro.",
                    m.Item)
            End If

            ' O desfecho da leitura tem de estar LISTADO. Estado que ninguem
            ' listou nega - inclusive um que ainda nao existe.
            If Not _ativacao.Leituras.Contains(m.Leitura.Kind) Then
                Return DisclosureDecision.Negar(DisclosureReason.LeituraNaoAceita,
                    $"A classificação de uma das mensagens não é aceita ({m.Leitura.Kind}).",
                    m.Item)
            End If

            ' Evidencia de versao: sem ela nao da nem para perceber que o item
            ' mudou depois de classificado.
            If m.Leitura.Version Is Nothing OrElse
               m.Leitura.Version.EntryId.Length = 0 Then
                Return DisclosureDecision.Negar(DisclosureReason.SemEvidenciaDeVersao,
                    "Não foi possível saber qual versão da mensagem foi classificada.",
                    m.Item)
            End If

            For Each r In m.Leitura.Registros.Where(Function(x) x.Ativo)
                If Not _ativacao.Rotulos.Contains(r.Id) Then
                    Return DisclosureDecision.Negar(DisclosureReason.RotuloNaoPermitido,
                        "Uma das mensagens tem classificação de sensibilidade que não " &
                        "está entre as autorizadas.", m.Item)
                End If

                ' ContentBits ausente ou ilegivel NAO prova ausencia de
                ' protecao. O 3.0 mediu que o campo existe; nao mediu que ele
                ' seja autentico, atual, ou que cubra toda forma de IRM.
                If r.ContentBitsIlegivel OrElse Not r.ContentBits.HasValue Then
                    Return DisclosureDecision.Negar(DisclosureReason.ContentBitsDesconhecido,
                        "Não foi possível saber se uma das mensagens está protegida.",
                        m.Item)
                End If

                If Not _ativacao.ContentBits.Contains(r.ContentBits.Value) Then
                    Return DisclosureDecision.Negar(DisclosureReason.ContentBitsNaoAceito,
                        "Uma das mensagens tem proteção que não está entre as autorizadas.",
                        m.Item)
                End If
            Next

            Return DisclosureDecision.Permitir()
        End Function

        Private Shared Function Igual(a As String, b As String) As Boolean
            Return String.Equals(If(a, "").Trim(), If(b, "").Trim(),
                                 StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function MesmaPasta(a As FolderKey, b As FolderKey) As Boolean
            If a Is Nothing OrElse b Is Nothing Then Return False
            Return String.Equals(a.EntryId, b.EntryId, StringComparison.Ordinal) AndAlso
                   String.Equals(a.StoreId, b.StoreId, StringComparison.Ordinal)
        End Function

    End Class

End Namespace

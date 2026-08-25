Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Model

Namespace Global.Iris.Assist

    ''' <summary>Por que o portão negou.</summary>
    Public Enum DisclosureReason
        ''' <summary>Zero: não decidido. Nunca significa permitido.</summary>
        NaoDecidido = 0

        ''' <summary>Não há autorização nenhuma. É o estado da produção.</summary>
        SemAtivacao
        ''' <summary>Há autorização, e ela está incompleta.</summary>
        AtivacaoIncompleta
        ''' <summary>
        ''' A autorização declara coisa que não pode ser declarada — desfecho
        ''' não elegível, operação nula, GUID inválido, prazo invertido.
        ''' </summary>
        AtivacaoInvalida
        ''' <summary>Há autorização, e ela venceu ou ainda não vale.</summary>
        AtivacaoForaDeVigencia
        ''' <summary>O endpoint autorizado não é HTTPS.</summary>
        EndpointInseguro
        ''' <summary>O endpoint do pedido não é o endpoint autorizado.</summary>
        EndpointNaoAutorizado
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

        ''' <summary>O desfecho da leitura não está na lista aceita.</summary>
        LeituraNaoAceita
        ''' <summary>
        ''' O desfecho descreve o <b>fracasso de ler</b>, e nenhuma autorização
        ''' pode torná-lo prova. Nega mesmo listado.
        ''' </summary>
        LeituraEstruturalmenteInsegura
        ''' <summary>A leitura declara um desfecho que a forma dela desmente.</summary>
        ClassificacaoIncoerente
        ''' <summary>A classificação não é da mensagem que está no pedido.</summary>
        IdentidadeNaoBate
        ''' <summary>Há rótulo ativo cujo GUID não está na lista permitida.</summary>
        RotuloNaoPermitido
        ''' <summary>Há registro desligado, e a política não declarou ignorá-lo.</summary>
        HistoricoNaoDeclarado
        ''' <summary>O <c>ContentBits</c> do registro não está na lista aceita.</summary>
        ContentBitsNaoAceito
        ''' <summary>O <c>ContentBits</c> não veio, ou veio ilegível.</summary>
        ContentBitsDesconhecido
        ''' <summary>Falta evidência de versão suficiente do item.</summary>
        SemEvidenciaDeVersao
        ''' <summary>Anexo — fora do escopo da fase, por inteiro.</summary>
        AnexoForaDeEscopo
    End Enum

    ''' <summary>Uma violação, presa ao item que a causou.</summary>
    Public NotInheritable Class DisclosureViolation
        Public ReadOnly Property Motivo As DisclosureReason
        ''' <summary>O item culpado, quando foi um item.</summary>
        Public ReadOnly Property Item As ItemKey
        ''' <summary>
        ''' Explicação em português. <b>Nunca</b> carrega assunto, corpo ou
        ''' nome de rótulo — o nome é texto escolhido pela empresa e pode ele
        ''' próprio ser sensível.
        ''' </summary>
        Public ReadOnly Property Explicacao As String

        Public Sub New(motivo As DisclosureReason, item As ItemKey, explicacao As String)
            Me.Motivo = motivo
            Me.Item = item
            Me.Explicacao = If(explicacao, "")
        End Sub
    End Class

    ''' <summary>
    ''' O veredito. <b>Nasce negado</b>, e é preciso prova para virar.
    '''
    ''' Carrega <b>todas</b> as violações das mensagens, não só a primeira.
    ''' Numa thread de trinta com três problemas diferentes, devolver um motivo
    ''' só faz o usuário consertar um para descobrir o próximo — e o
    ''' <see cref="Motivo"/> continua existindo, determinístico, para quem só
    ''' quer uma linha.
    ''' </summary>
    Public NotInheritable Class DisclosureDecision

        ''' <summary>Quantas violações cabem no veredito. O resto é contado.</summary>
        Private Const Teto As Integer = 20

        Public ReadOnly Property Permitido As Boolean
        Public ReadOnly Property Violacoes As IReadOnlyList(Of DisclosureViolation)
        ''' <summary>Quantas violações houve ao todo, inclusive as não listadas.</summary>
        Public ReadOnly Property Total As Integer

        Private Sub New(permitido As Boolean, violacoes As IReadOnlyList(Of DisclosureViolation),
                        total As Integer)
            Me.Permitido = permitido
            Me.Violacoes = violacoes
            Me.Total = total
        End Sub

        ''' <summary>O motivo principal: o primeiro, e por isso determinístico.</summary>
        Public ReadOnly Property Motivo As DisclosureReason
            Get
                If Violacoes.Count = 0 Then Return DisclosureReason.NaoDecidido
                Return Violacoes(0).Motivo
            End Get
        End Property

        Public ReadOnly Property Explicacao As String
            Get
                If Violacoes.Count = 0 Then Return ""
                Return Violacoes(0).Explicacao
            End Get
        End Property

        Public ReadOnly Property Culpado As ItemKey
            Get
                If Violacoes.Count = 0 Then Return Nothing
                Return Violacoes(0).Item
            End Get
        End Property

        Friend Shared Function Negar(motivo As DisclosureReason, explicacao As String,
                                     Optional culpado As ItemKey = Nothing) As DisclosureDecision
            Return Negar({New DisclosureViolation(motivo, culpado, explicacao)})
        End Function

        Friend Shared Function Negar(v As IReadOnlyList(Of DisclosureViolation)) As DisclosureDecision
            Return New DisclosureDecision(False, v.Take(Teto).ToList(), v.Count)
        End Function

        Friend Shared Function Permitir() As DisclosureDecision
            Return New DisclosureDecision(True, Array.Empty(Of DisclosureViolation)(), 0)
        End Function

    End Class

    ' ==================================================================

    ''' <summary>
    ''' O destino: <b>onde</b> o conteúdo vai parar.
    '''
    ''' Provedor e modelo não bastam. A autorização registra um endpoint, e sem
    ''' o pedido declarar o endpoint que o transmissor vai usar, a decisão
    ''' autoriza "o provedor certo" e o transmissor manda para outro lugar.
    ''' </summary>
    Public NotInheritable Class AssistDestination
        Public ReadOnly Property Provedor As String
        Public ReadOnly Property Endpoint As String
        Public ReadOnly Property Modelo As String

        Public Sub New(provedor As String, endpoint As String, modelo As String)
            Me.Provedor = If(provedor, "")
            Me.Endpoint = If(endpoint, "")
            Me.Modelo = If(modelo, "")
        End Sub
    End Class

    ''' <summary>Tudo o que dá para decidir <b>antes</b> de classificar item nenhum.</summary>
    Public NotInheritable Class PreflightRequest
        Public ReadOnly Property Operacao As AssistOperation
        Public ReadOnly Property Pasta As FolderKey
        Public ReadOnly Property Destino As AssistDestination

        Public Sub New(operacao As AssistOperation, pasta As FolderKey,
                       destino As AssistDestination)
            Me.Operacao = operacao
            Me.Pasta = pasta
            Me.Destino = destino
        End Sub
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
        Public ReadOnly Property Preflight As PreflightRequest
        Public ReadOnly Property Mensagens As IReadOnlyList(Of MessageClassification)

        Public Sub New(preflight As PreflightRequest,
                       mensagens As IEnumerable(Of MessageClassification))
            Me.Preflight = preflight
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
    ''' estilística: um portão escrito como lista de negativas libera todo caso
    ''' que ninguém pensou em proibir, e o caso que ninguém pensou é exatamente
    ''' o que vaza.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DUAS ETAPAS, E A ORDEM É PARTE DA GARANTIA</b>
    '''
    ''' <see cref="Preflight"/> decide tudo o que não depende de item nenhum:
    ''' autorização, vigência, destino, operação, pasta. Só depois dele o
    ''' chamador vai ao COM classificar mensagem.
    '''
    ''' Não é economia: o 3.0 mediu ~17 ms por item, e classificar uma thread
    ''' inteira para descobrir que a IA está desligada seria pagar meio segundo
    ''' de fila da STA para nada — além de tocar itens sem autorização para
    ''' tocá-los. Quem garante a ordem na prática é o
    ''' <see cref="DisclosureGate"/>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>UM MEMBRO QUE NÃO PASSA NEGA A THREAD INTEIRA</b>
    '''
    ''' Não a mensagem — a thread. Resumo parcial é fácil demais de confundir
    ''' com resumo completo, e o usuário não tem como saber que faltou pedaço.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A ASSIMETRIA DA P16</b>
    '''
    ''' <c>MSIP_Labels</c> mora no namespace de cabeçalhos de internet.
    ''' Cabeçalho recebido pode ter origem fora do mecanismo corporativo, então
    ''' ler o valor com perfeição <b>não prova</b> que ninguém apresenta uma
    ''' classificação baixa falsa. Rótulo <b>nunca autoriza sozinho</b>: entra
    ''' como mais uma condição, e o que autoriza é a autorização.
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

        ''' <summary>
        ''' As provas que <b>não</b> dependem de item nenhum. Chamar isto antes
        ''' de classificar é o que impede ir ao COM sem autorização.
        ''' </summary>
        Public Function Preflight(pedido As PreflightRequest,
                                  agora As DateTimeOffset) As DisclosureDecision

            If _ativacao Is Nothing Then
                Return DisclosureDecision.Negar(DisclosureReason.SemAtivacao,
                    "A IA externa não está habilitada. Nada deste computador é enviado " &
                    "para fora enquanto você não autorizar, com a política da empresa " &
                    "e um provedor à sua escolha.")
            End If

            If Not _ativacao.Completo() Then
                Return DisclosureDecision.Negar(DisclosureReason.AtivacaoIncompleta,
                    "A autorização está incompleta — falta declarar quem autorizou, o " &
                    "provedor, o endpoint, o modelo, a região, a retenção aceita, as " &
                    "operações ou as pastas.")
            End If

            If Not _ativacao.Coerente() Then
                Return DisclosureDecision.Negar(DisclosureReason.AtivacaoInvalida,
                    "A autorização declara algo que não pode ser declarado — um estado " &
                    "de leitura que não é prova de nada, uma operação em branco, um " &
                    "identificador de classificação inválido, ou um prazo invertido.")
            End If

            If Not _ativacao.Vigente(agora) Then
                Return DisclosureDecision.Negar(DisclosureReason.AtivacaoForaDeVigencia,
                    "A autorização não está vigente nesta data.")
            End If

            If Not _ativacao.EndpointSeguro() Then
                Return DisclosureDecision.Negar(DisclosureReason.EndpointInseguro,
                    "O endereço autorizado não é HTTPS.")
            End If

            If pedido Is Nothing OrElse pedido.Destino Is Nothing Then
                Return DisclosureDecision.Negar(DisclosureReason.EndpointNaoAutorizado,
                    "O pedido não declarou para onde o conteúdo iria.")
            End If

            If Not _ativacao.Operacoes.Contains(pedido.Operacao) Then
                Return DisclosureDecision.Negar(DisclosureReason.OperacaoNaoAutorizada,
                    "Esta operação não está entre as autorizadas.")
            End If

            If Not Igual(pedido.Destino.Provedor, _ativacao.Provedor) OrElse
               Not Igual(pedido.Destino.Modelo, _ativacao.Modelo) Then
                Return DisclosureDecision.Negar(DisclosureReason.ProvedorNaoAutorizado,
                    "O provedor ou o modelo do pedido não é o autorizado.")
            End If

            ' O ENDPOINT do pedido, e nao so o do registro. Sem isto a decisao
            ' autoriza "o provedor certo" e o transmissor manda para outro
            ' lugar - a autorizacao teria dito sim a um destino que ninguem
            ' conferiu. Comparacao ORDINAL: diferenca de caminho ou de porta e
            ' diferenca de destino.
            If Not String.Equals(pedido.Destino.Endpoint.Trim(), _ativacao.Endpoint.Trim(),
                                 StringComparison.Ordinal) Then
                Return DisclosureDecision.Negar(DisclosureReason.EndpointNaoAutorizado,
                    "O endereço para onde o conteúdo iria não é o autorizado.")
            End If

            If Not _ativacao.Pastas.Any(Function(f) MesmaPasta(f, pedido.Pasta)) Then
                Return DisclosureDecision.Negar(DisclosureReason.PastaNaoAutorizada,
                    "Esta pasta não está entre as autorizadas.")
            End If

            Return DisclosureDecision.Permitir()
        End Function

        ''' <summary>O preflight, e depois cada mensagem.</summary>
        Public Function Decidir(pedido As DisclosureRequest,
                                agora As DateTimeOffset) As DisclosureDecision

            If pedido Is Nothing Then
                Return DisclosureDecision.Negar(DisclosureReason.PedidoVazio,
                                                "Não há nada a enviar.")
            End If

            Dim antes = Preflight(pedido.Preflight, agora)
            If Not antes.Permitido Then Return antes

            If pedido.Mensagens.Count = 0 Then
                Return DisclosureDecision.Negar(DisclosureReason.PedidoVazio,
                                                "Não há nada a enviar.")
            End If

            ' Todas as mensagens sao conferidas, e nao so ate a primeira falha:
            ' numa thread grande, devolver um motivo so faz o usuario consertar
            ' um problema para descobrir o proximo.
            Dim violacoes As New List(Of DisclosureViolation)()
            For Each m In pedido.Mensagens
                violacoes.AddRange(Conferir(m, pedido.Preflight.Pasta))
            Next

            If violacoes.Count > 0 Then Return DisclosureDecision.Negar(violacoes)
            Return DisclosureDecision.Permitir()
        End Function

        ' ==============================================================

        ''' <summary>
        ''' As provas de uma mensagem. Qualquer uma que falhe nega o pedido
        ''' <b>inteiro</b>.
        ''' </summary>
        Private Function Conferir(m As MessageClassification,
                                  pasta As FolderKey) As IReadOnlyList(Of DisclosureViolation)

            Dim v As New List(Of DisclosureViolation)()

            If m Is Nothing OrElse m.Leitura Is Nothing Then
                v.Add(New DisclosureViolation(DisclosureReason.LeituraNaoAceita, Nothing,
                    "Uma das mensagens não foi classificada."))
                Return v
            End If

            ' A classificacao e DESTA mensagem? DTO publico nao garante nada, e
            ' uma leitura de outro item anexada a esta passaria por qualquer
            ' conferencia que olhasse so o conteudo da leitura.
            If m.Item Is Nothing OrElse Not m.Item.Equals(m.Leitura.Item) Then
                v.Add(New DisclosureViolation(DisclosureReason.IdentidadeNaoBate, m.Item,
                    "A classificação de uma das mensagens não é dela."))
                Return v
            End If

            If Not MesmaPasta(m.Pasta, pasta) Then
                v.Add(New DisclosureViolation(DisclosureReason.MensagemDeOutraPasta, m.Item,
                    "Uma das mensagens está em outra pasta que não a autorizada."))
            End If

            If m.TemAnexo Then
                v.Add(New DisclosureViolation(DisclosureReason.AnexoForaDeEscopo, m.Item,
                    "Mensagem com anexo. Anexo não é tratado nesta fase, por inteiro."))
            End If

            ' Estruturalmente inseguro nega ANTES da lista, e nega mesmo
            ' listado: "nao consegui ler" nunca vira prova de nada.
            If Not LabelPolicy.Elegivel(m.Leitura.Kind) Then
                v.Add(New DisclosureViolation(DisclosureReason.LeituraEstruturalmenteInsegura,
                    m.Item, "Não foi possível classificar uma das mensagens com segurança."))
                Return v
            End If

            If Not _ativacao.Leituras.Contains(m.Leitura.Kind) Then
                v.Add(New DisclosureViolation(DisclosureReason.LeituraNaoAceita, m.Item,
                    "A classificação de uma das mensagens não está entre as aceitas."))
            End If

            ' A leitura diz uma coisa e a forma dela diz outra: um Present sem
            ' registro ativo, um Absent COM registro ativo. O parser nunca
            ' produz isso; um DTO montado a mao produz.
            If Not LabelPolicy.Coerente(m.Leitura) Then
                v.Add(New DisclosureViolation(DisclosureReason.ClassificacaoIncoerente, m.Item,
                    "A classificação de uma das mensagens não faz sentido internamente."))
                Return v
            End If

            v.AddRange(ConferirVersao(m))
            v.AddRange(ConferirRegistros(m))
            Return v
        End Function

        ''' <summary>
        ''' Sem evidência de versão não dá nem para perceber que o item mudou
        ''' depois de classificado. E a evidência tem de ser <b>daquele</b>
        ''' item: um <c>EntryId</c> qualquer preencheria o campo sem dizer nada.
        ''' </summary>
        Private Shared Function ConferirVersao(m As MessageClassification) _
                                               As IReadOnlyList(Of DisclosureViolation)
            Dim v As New List(Of DisclosureViolation)()
            Dim e = m.Leitura.Version

            If e Is Nothing OrElse e.EntryId.Length = 0 OrElse
               Not String.Equals(e.EntryId, m.Item.EntryId, StringComparison.Ordinal) OrElse
               String.IsNullOrEmpty(e.ChangeKey) Then
                v.Add(New DisclosureViolation(DisclosureReason.SemEvidenciaDeVersao, m.Item,
                    "Não foi possível saber qual versão de uma das mensagens foi classificada."))
            End If
            Return v
        End Function

        Private Function ConferirRegistros(m As MessageClassification) _
                                           As IReadOnlyList(Of DisclosureViolation)
            Dim v As New List(Of DisclosureViolation)()

            For Each r In m.Leitura.Registros
                If Not r.Ativo Then
                    ' Registro DESLIGADO. Ignora-lo e decisao de politica, e a
                    ' politica tem de declarar: Enabled=False mostra que o
                    ' registro se diz inativo, e nao prova que o conteudo
                    ' deixou de ser sensivel nem que a empresa aceita
                    ' desconsiderar rebaixamento.
                    If Not _ativacao.IgnorarHistorico AndAlso
                       Not _ativacao.Rotulos.Contains(r.Id) Then
                        v.Add(New DisclosureViolation(DisclosureReason.HistoricoNaoDeclarado,
                            m.Item, "Uma das mensagens teve uma classificação que não está " &
                                    "entre as autorizadas, e a autorização não declarou " &
                                    "ignorar classificações antigas."))
                    End If
                    Continue For
                End If

                If Not _ativacao.Rotulos.Contains(r.Id) Then
                    v.Add(New DisclosureViolation(DisclosureReason.RotuloNaoPermitido, m.Item,
                        "Uma das mensagens tem classificação de sensibilidade que não " &
                        "está entre as autorizadas."))
                End If

                ' ContentBits ausente ou ilegivel NAO prova ausencia de
                ' protecao. O 3.0 mediu que o campo existe; nao mediu que seja
                ' autentico, atual, ou que cubra toda forma de IRM.
                If r.ContentBitsIlegivel OrElse Not r.ContentBits.HasValue Then
                    v.Add(New DisclosureViolation(DisclosureReason.ContentBitsDesconhecido,
                        m.Item, "Não foi possível saber se uma das mensagens está protegida."))
                ElseIf Not _ativacao.ContentBits.Contains(r.ContentBits.Value) Then
                    v.Add(New DisclosureViolation(DisclosureReason.ContentBitsNaoAceito, m.Item,
                        "Uma das mensagens tem proteção que não está entre as autorizadas."))
                End If
            Next
            Return v
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

    ' ==================================================================

    ''' <summary>
    ''' <b>Quem garante que ninguém classifica antes de ser autorizado.</b>
    '''
    ''' O <see cref="DisclosurePolicy"/> sozinho não garante: ele recebe um
    ''' pedido que <b>já vem com as classificações dentro</b>, o que significa
    ''' que alguém foi ao COM antes de chamá-lo. Provar que o motivo devolvido
    ''' é o da ativação prova precedência de motivo, não ausência de leitura.
    '''
    ''' Aqui a classificação é um <c>Func</c> que só é <b>invocado</b> se o
    ''' preflight passar. É essa a fronteira, e há teste com um espião que
    ''' explode se for chamado.
    ''' </summary>
    Public NotInheritable Class DisclosureGate

        Private ReadOnly _politica As DisclosurePolicy

        Public Sub New(politica As DisclosurePolicy)
            _politica = politica
        End Sub

        Public Function Avaliar(pedido As PreflightRequest, agora As DateTimeOffset,
                                classificar As Func(Of IReadOnlyList(Of MessageClassification))) _
                                As DisclosureDecision

            Dim antes = _politica.Preflight(pedido, agora)
            If Not antes.Permitido Then Return antes

            Return _politica.Decidir(New DisclosureRequest(pedido, classificar()), agora)
        End Function

    End Class

End Namespace

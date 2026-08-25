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
        ''' <summary>
        ''' A versão de cada item, na mesma ordem — a versão que passou pelo
        ''' portão, não "o item como estiver".
        ''' </summary>
        Public ReadOnly Property Versoes As IReadOnlyList(Of String)
        Public ReadOnly Property Pasta As FolderKey
        ''' <summary>
        ''' Liga capability, diário e resposta. Sem ele, três registros da mesma
        ''' operação são três coisas que <i>parecem</i> a mesma.
        ''' </summary>
        Public ReadOnly Property RequestId As Guid
        Public ReadOnly Property Emitida As DateTimeOffset
        Public ReadOnly Property Expira As DateTimeOffset

        ''' <summary>
        ''' <c>Friend</c>, e é o ponto: só o <see cref="CapabilityLedger"/>
        ''' emite. Uma capability que qualquer um pudesse construir seria uma
        ''' afirmação de que houve autorização, sem que tivesse havido.
        ''' </summary>
        Friend Sub New(id As Guid, requestId As Guid, grant As DisclosureGrant,
                       hash As String, comprimento As Integer,
                       emitida As DateTimeOffset, expira As DateTimeOffset)
            Me.Id = id
            Me.RequestId = requestId
            AtivacaoId = grant.AtivacaoId
            AtivacaoVersao = grant.AtivacaoVersao
            Operacao = grant.Operacao
            Destino = grant.Destino
            Pasta = grant.Pasta
            Itens = grant.Itens
            Versoes = grant.Versoes
            Me.Hash = hash
            Me.Comprimento = comprimento
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
        ''' <summary>
        ''' Os itens do envelope não são os aprovados. Existe porque o
        ''' <c>EntryID</c> deliberadamente não entra nos bytes: dois envelopes
        ''' com o mesmo texto e itens diferentes têm o <b>mesmo hash</b>, então
        ''' o hash sozinho não prova proveniência.
        ''' </summary>
        ProveniencaDiferente
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
        ''' Emite, a partir do <see cref="DisclosureGrant"/> — e <b>só</b> dele.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE NÃO BASTA "A DECISÃO FOI PERMITIDA"</b>
        '''
        ''' A versão anterior pedia <c>decisao.Permitido</c> e depois aceitava,
        ''' em parâmetros separados, qualquer ativação, qualquer destino e
        ''' qualquer envelope. Um "sim" dado para a ativação A, destino A e
        ''' mensagens A emitia autorização para a ativação B, destino B e
        ''' conteúdo B. O veredito era verdadeiro; a emissão era sobre outra
        ''' coisa.
        '''
        ''' Agora tudo vem do grant, e o que o chamador traz de fora — o
        ''' envelope — é <b>conferido contra ele</b>.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>TRUNCADO E INCOMPLETO NÃO SÃO EMITÍVEIS</b>
        '''
        ''' A §29.1 diz que um membro não permitido nega a thread inteira. Pela
        ''' mesma razão, uma thread que não coube <b>não</b> vira uma thread
        ''' menor: o resumo sairia parecendo completo. Idem corpo pela metade.
        '''
        ''' O envelope continua sabendo dizer que truncou — é o que a UI usa
        ''' para explicar. Ele só não vira autorização.
        ''' </summary>
        Public Function Emitir(decisao As DisclosureDecision, envelope As AssistEnvelope,
                               agora As DateTimeOffset) As DisclosureCapability

            If decisao Is Nothing OrElse Not decisao.Permitido Then Return Nothing
            Dim g = decisao.Grant
            If g Is Nothing OrElse envelope Is Nothing Then Return Nothing

            ' Envelope que nao confere consigo mesmo nao vira autorizacao.
            If Not envelope.Integro() Then Return Nothing

            ' Meia thread e meio corpo nao sao autorizaveis.
            If envelope.Truncado OrElse envelope.CorpoIncompleto Then Return Nothing

            ' A operacao DENTRO dos bytes tem de ser a aprovada. Sem isto, um
            ' grant para Resumir emitia capability sobre um envelope montado
            ' como Redigir: a capability recebia a operacao do grant, e o hash
            ' era dos bytes de outra coisa.
            If g.Operacao <> envelope.Operacao Then Return Nothing

            ' E os itens tem de ser EXATAMENTE os aprovados, na ordem aprovada.
            ' desligado

            ' O prazo e o MENOR entre a validade curta da capability e o fim da
            ' propria ativacao. Uma capability que sobrevivesse a autorizacao
            ' que a gerou seria uma autorizacao a mais, emitida por ninguem.
            Dim expira = agora + Validade
            If g.AtivacaoAte.HasValue AndAlso g.AtivacaoAte.Value < expira Then
                expira = g.AtivacaoAte.Value
            End If

            Dim c As New DisclosureCapability(Guid.NewGuid(), Guid.NewGuid(), g,
                                              envelope.Hash, envelope.Comprimento,
                                              agora, expira)
            _emitidas(c.Id) = c
            Return c
        End Function

        ''' <summary>
        ''' Confere e <b>gasta</b>. Chamada imediatamente antes de transmitir.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>VALIDA PRIMEIRO, GASTA DEPOIS</b>
        '''
        ''' A versão anterior marcava o consumo antes de conferir, com o
        ''' argumento de que devolver a capability faria dela um oráculo — daria
        ''' para tentar envelope atrás de envelope até um bater.
        '''
        ''' O argumento não se sustenta: a capability <b>já expõe</b> o hash
        ''' esperado, e as recusas já são distintas. Não há segredo a adivinhar.
        ''' O que existia era o contrário — qualquer código com uma referência à
        ''' capability podia queimá-la passando destino errado.
        '''
        ''' A ordem certa é validar tudo e, por último, fazer o
        ''' <c>TryAdd</c> atômico: duas transmissões simultâneas continuam
        ''' impossíveis, e uma conferência local errada não destrói a
        ''' autorização.
        ''' </summary>
        Public Function Consumir(c As DisclosureCapability, envelope As AssistEnvelope,
                                 destino As AssistDestination, operacao As AssistOperation,
                                 agora As DateTimeOffset) As CapabilityUse

            ' A CANONICA, e nao o objeto apresentado. O ledger conferia os
            ' campos de 'c' - outro objeto com o mesmo Id, construivel dentro
            ' do assembly ou por desserializacao futura, apresentaria hash,
            ' destino e operacao diferentes dos que o ledger emitiu, e a
            ' conferencia bateria consigo mesma.
            Dim canonica As DisclosureCapability = Nothing
            If c Is Nothing OrElse Not _emitidas.TryGetValue(c.Id, canonica) Then
                Return Recusar(CapabilityRefusal.Desconhecida)
            End If
            If Not ReferenceEquals(c, canonica) Then
                Return Recusar(CapabilityRefusal.Desconhecida)
            End If
            c = canonica

            If _consumidas.ContainsKey(c.Id) Then Return Recusar(CapabilityRefusal.JaConsumida)

            If agora > c.Expira Then Return Recusar(CapabilityRefusal.Expirada)

            If envelope Is Nothing Then Return Recusar(CapabilityRefusal.BytesDiferentes)
            If Not envelope.Integro() Then Return Recusar(CapabilityRefusal.EnvelopeCorrompido)

            If Not String.Equals(envelope.Hash, c.Hash, StringComparison.Ordinal) OrElse
               envelope.Comprimento <> c.Comprimento Then
                Return Recusar(CapabilityRefusal.BytesDiferentes)
            End If

            ' O hash NAO cobre a proveniencia: o EntryID nao entra nos bytes,
            ' entao dois envelopes com o mesmo texto e itens diferentes tem o
            ' mesmo hash. Sem esta conferencia, conteudo aprovado para uma
            ' mensagem sairia registrado como vindo de outra.
            If Not MesmosItens(envelope.Itens, c.Itens) OrElse
               Not MesmasVersoes(envelope.Versoes, c.Versoes) Then
                Return Recusar(CapabilityRefusal.ProveniencaDiferente)
            End If

            ' A operacao pedida, a da capability e a que esta DENTRO dos bytes:
            ' as tres.
            If operacao <> c.Operacao OrElse envelope.Operacao <> c.Operacao Then
                Return Recusar(CapabilityRefusal.OperacaoDiferente)
            End If

            If destino Is Nothing OrElse c.Destino Is Nothing OrElse
               Not String.Equals(destino.Endpoint.Trim(), c.Destino.Endpoint.Trim(),
                                 StringComparison.Ordinal) OrElse
               Not String.Equals(destino.Provedor.Trim(), c.Destino.Provedor.Trim(),
                                 StringComparison.Ordinal) OrElse
               Not String.Equals(destino.Modelo.Trim(), c.Destino.Modelo.Trim(),
                                 StringComparison.Ordinal) Then
                Return Recusar(CapabilityRefusal.DestinoDiferente)
            End If

            ' Por ULTIMO, e atomico: quem vencer o TryAdd transmite.
            If Not _consumidas.TryAdd(c.Id, 0) Then
                Return Recusar(CapabilityRefusal.JaConsumida)
            End If

            Return New CapabilityUse(True, CapabilityRefusal.Nenhuma)
        End Function

        Private Shared Function MesmasVersoes(a As IReadOnlyList(Of String),
                                              b As IReadOnlyList(Of String)) As Boolean
            If a Is Nothing OrElse b Is Nothing OrElse a.Count <> b.Count Then Return False
            For i = 0 To a.Count - 1
                If Not String.Equals(a(i), b(i), StringComparison.Ordinal) Then Return False
            Next
            Return True
        End Function

        Private Shared Function MesmosItens(a As IReadOnlyList(Of ItemKey),
                                            b As IReadOnlyList(Of ItemKey)) As Boolean
            If a Is Nothing OrElse b Is Nothing OrElse a.Count <> b.Count Then Return False
            For i = 0 To a.Count - 1
                If Not a(i).Equals(b(i)) Then Return False
            Next
            Return True
        End Function

        Private Shared Function Recusar(r As CapabilityRefusal) As CapabilityUse
            Return New CapabilityUse(False, r)
        End Function

    End Class

End Namespace

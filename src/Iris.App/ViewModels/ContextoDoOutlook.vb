Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports Iris.Assist
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>O contexto de verdade: a mensagem aberta, lida do Outlook.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ELE EXISTE AGORA, SEM PROVEDOR ESCOLHIDO</b>
    '''
    ''' Ler a mensagem, classificar cada membro e montar o envelope são
    ''' requisitos <b>do Iris</b>, e independem de qual API vai receber os
    ''' bytes. Deixá-los para depois faria a frase "implementação e provas
    ''' locais concluídas" ser falsa: mesmo depois da cerimônia de ativação
    ''' ainda faltaria o caminho central até o Outlook.
    '''
    ''' O que continua pendente é só o <b>adaptador externo</b> — o formato que
    ''' o provedor aceita, a autenticação dele, a semântica de resposta.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A ORDEM, E POR QUE ELA É SÍNCRONA AQUI</b>
    '''
    ''' <see cref="Classificar"/> só é invocado pelo <c>DisclosureGate</c>
    ''' <b>depois</b> de o preflight passar — é o que impede ir ao COM sem
    ''' autorização para tocar em item nenhum.
    '''
    ''' As chamadas ao broker são assíncronas e aqui são esperadas em bloco. O
    ''' <c>AssistTransmitter</c> roda tudo isto fora da thread da UI, e o broker
    ''' tem STA própria — bloquear esta thread não trava a janela nem a fila do
    ''' COM. Fazer a porta assíncrona só para não esperar aqui espalharia
    ''' <c>Task</c> por uma cadeia que o portão precisa chamar em ordem.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ELE NÃO FAZ</b>
    '''
    ''' Não decide nada. Item que não dá para ler, classificar ou preparar
    ''' simplesmente <b>não entra</b> — e aí a contagem não bate com o que o
    ''' portão aprovou, o grant não cobre o envelope, e o cofre recusa. Quem
    ''' nega é o portão; aqui é só a coleta.
    ''' </summary>
    Friend NotInheritable Class ContextoDoOutlook
        Implements IAssistContext

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _destino As AssistDestination
        Private ReadOnly _selecao As Func(Of (Pasta As FolderKey, Itens As IReadOnlyList(Of ItemKey)))

        Friend Sub New(broker As IOutlookBroker, destino As AssistDestination,
                       selecao As Func(Of (Pasta As FolderKey, Itens As IReadOnlyList(Of ItemKey))))
            _broker = broker
            _destino = destino
            _selecao = selecao
        End Sub

        Public Function Pedido(operacao As AssistOperation) As PreflightRequest _
                               Implements IAssistContext.Pedido
            Return New PreflightRequest(operacao, _selecao().Pasta, _destino)
        End Function

        ''' <summary>
        ''' Lê o rótulo e a presença de anexo de cada item selecionado.
        '''
        ''' Item cuja leitura falhou entra com o desfecho que ela deu — e não
        ''' fica de fora: um item sumindo da lista faria a thread parecer menor
        ''' do que é, e o portão aprovaria uma thread que não é a do usuário.
        '''
        ''' <b>O anexo é lido, e não suposto.</b> Aqui havia <c>temAnexo:=False</c>
        ''' fixo — o portão nega mensagem com anexo, e o caminho de produção lhe
        ''' afirmava que não havia. Se a contagem falhar, para o item ou para a
        ''' chamada inteira, o valor é <b>tem</b>: o portão nega, e nega dizendo
        ''' por quê.
        ''' </summary>
        Public Function Classificar() As IReadOnlyList(Of MessageClassification) _
                                        Implements IAssistContext.Classificar
            Dim sel = _selecao()
            If sel.Itens.Count = 0 Then Return Array.Empty(Of MessageClassification)()

            Dim r = _broker.GetSensitivityLabelsAsync(sel.Itens, CancellationToken.None).
                    GetAwaiter().GetResult()
            If Not r.Succeeded Then Return Array.Empty(Of MessageClassification)()

            Dim anexo = Anexos(sel.Itens)

            Return r.Value.Select(
                Function(l) New MessageClassification(
                    l.Item, sel.Pasta, l,
                    temAnexo:=TemAnexo(anexo, l.Item.EntryId))).ToList()
        End Function

        ''' <summary>
        ''' Item sem resposta, ou com resposta <c>Nothing</c>, conta como
        ''' <b>tem anexo</b>. Fechado por falta de prova.
        ''' </summary>
        Private Shared Function TemAnexo(mapa As Dictionary(Of String, Boolean?),
                                         chave As String) As Boolean
            Dim tem As Boolean? = Nothing
            If Not mapa.TryGetValue(chave, tem) Then Return True
            If Not tem.HasValue Then Return True
            Return tem.Value
        End Function

        ''' <summary>
        ''' A presença de anexo por item, indexada pelo <c>EntryId</c>.
        '''
        ''' Chamada que falhou devolve dicionário vazio, e item ausente do
        ''' dicionário vira <b>tem anexo</b> lá em cima. Fechado por falta de
        ''' prova, como o resto da §29.
        ''' </summary>
        Private Function Anexos(itens As IReadOnlyList(Of ItemKey)) _
                                As Dictionary(Of String, Boolean?)
            Dim saida As New Dictionary(Of String, Boolean?)(StringComparer.Ordinal)
            Dim r = _broker.GetAttachmentPresenceAsync(itens, CancellationToken.None).
                    GetAwaiter().GetResult()
            If Not r.Succeeded Then Return saida
            For Each p In r.Value
                saida(p.Item.EntryId) = p.Tem
            Next
            Return saida
        End Function

        ''' <summary>
        ''' Captura cada mensagem numa leitura só, prepara pelo pipeline, e
        ''' monta.
        '''
        ''' Mensagem que o pipeline recusa — corpo pela metade, referência
        ''' embutida, HTML que não dá para interpretar — <b>não entra</b>. O
        ''' envelope sai com menos itens que o grant aprovou, o
        ''' <c>Cobre</c> falha, e o cofre não emite. É o desfecho certo: uma
        ''' thread com um membro faltando não é a thread.
        ''' </summary>
        Public Function Montar(operacao As AssistOperation, instrucao As String) _
                               As EnvelopeResult Implements IAssistContext.Montar
            Dim partes As New List(Of MessagePart)()

            For Each item In _selecao().Itens
                Dim s = _broker.GetMessageSnapshotAsync(item, CancellationToken.None).
                        GetAwaiter().GetResult()
                If Not s.Succeeded Then Continue For

                Dim pronto = ContentPipeline.Preparar(s.Value)
                If Not pronto.Ok Then Continue For

                partes.Add(pronto.Parte)
            Next

            Return New EnvelopeBuilder().Montar(operacao, instrucao, partes)
        End Function

    End Class

End Namespace

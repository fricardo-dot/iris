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
    ''' <b>POR QUE ELE NÃO SABE QUAL É O PROVEDOR</b>
    '''
    ''' Ler a mensagem, classificar cada membro e montar o envelope são
    ''' requisitos <b>do Iris</b>, e independem de qual API vai receber os bytes.
    ''' Ele recebe o destino pronto e não escolhe nada.
    '''
    ''' <b>O adaptador externo deixou de ser pendência</b> — há um de OpenRouter,
    ''' escolhido pela ativação em <c>MainViewModel.ProvedorPara</c>. Este
    ''' comentário dizia o contrário, e nascia falso no caminho que instancia a
    ''' classe. Achado por revisão externa em 01/09/2026.
    '''
    ''' <b>E a borda em lote também deixou de ser pendência</b>, em 01/09/2026:
    ''' este contexto lê as N mensagens da seleção numa visita só ao Outlook, e
    ''' é ele mesmo que serve a classificação de pasta. Não há um segundo
    ''' caminho de divulgação — <b>é este</b>, com outra seleção e com fichas.
    ''' Um segundo caminho seria um segundo lugar para o portão ser esquecido.
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

        ''' <summary>
        ''' <b>A ficha de cada item</b>, quando a seleção é um lote de
        ''' classificação. <c>Nothing</c> no caminho por mensagem, que não tem lote
        ''' nem precisa dizer <i>de quem</i> a resposta fala: ela fala da única.
        ''' </summary>
        Private ReadOnly _ficha As Func(Of ItemKey, String)

        ''' <summary>
        ''' O corpo mais longo que cabe numa parte <b>desta</b> seleção.
        '''
        ''' No caminho por mensagem o envelope inteiro é da mensagem, então o teto é
        ''' o do pipeline. Num lote de vinte, uma mensagem grande sozinha estoura o
        ''' envelope e derruba as outras dezenove — ver
        ''' <c>ContentPipeline.Preparar</c>, que tem o caso por extenso.
        ''' </summary>
        Private ReadOnly _tetoDoCorpo As Integer

        Friend Sub New(broker As IOutlookBroker, destino As AssistDestination,
                       selecao As Func(Of (Pasta As FolderKey, Itens As IReadOnlyList(Of ItemKey))),
                       Optional ficha As Func(Of ItemKey, String) = Nothing,
                       Optional tetoDoCorpo As Integer = ContentPipeline.MaxCorpo)
            _broker = broker
            _destino = destino
            _selecao = selecao
            _ficha = ficha
            _tetoDoCorpo = tetoDoCorpo
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
            Dim onde = PastasObservadas(sel.Itens)

            Return r.Value.Select(
                Function(l) New MessageClassification(
                    l.Item, PastaDe(onde, l.Item.EntryId), l,
                    temAnexo:=TemAnexo(anexo, l.Item.EntryId),
                    embutidas:=Embutidas(anexo, l.Item.EntryId))).ToList()
        End Function

        ''' <summary>
        ''' <b>A pasta que o Outlook diz, e não a que o chamador disse.</b>
        '''
        ''' Esta linha era <c>sel.Pasta</c> — a mesma pasta que vai no pedido. A
        ''' regra do portão <i>"mensagem de outra pasta nega"</i> comparava a
        ''' afirmação do chamador com ela mesma, e sempre concordava.
        '''
        ''' No caminho por mensagem isso quase não doía: a seleção era a pasta
        ''' aberta, e a lista viera dela. Na classificação em lote passou a doer,
        ''' porque as chaves vêm do cache — um retrato de quando a varredura rodou.
        ''' Uma mensagem movida para uma pasta confidencial depois disso sairia sob
        ''' a autorização da pasta antiga. Achado por revisão externa em 01/09/2026.
        '''
        ''' <b>Falha de leitura devolve pasta vazia</b>, que nunca é igual à do
        ''' pedido — então o portão nega, e nega dizendo por quê.
        ''' </summary>
        Private Function PastasObservadas(itens As IReadOnlyList(Of ItemKey)) _
                                          As Dictionary(Of String, FolderKey)
            Dim saida As New Dictionary(Of String, FolderKey)(StringComparer.Ordinal)
            Dim r = _broker.GetItemFoldersAsync(itens, CancellationToken.None).
                    GetAwaiter().GetResult()
            If Not r.Succeeded OrElse r.Value Is Nothing Then Return saida
            For Each p In r.Value
                If p Is Nothing OrElse p.Item Is Nothing Then Continue For
                saida(p.Item.EntryId) = p.Pasta
            Next
            Return saida
        End Function

        Private Shared Function PastaDe(mapa As Dictionary(Of String, FolderKey),
                                        chave As String) As FolderKey
            Dim achada As FolderKey = Nothing
            If mapa.TryGetValue(chave, achada) AndAlso achada IsNot Nothing Then
                Return achada
            End If
            Return New FolderKey("", "")
        End Function

        ''' <summary>
        ''' Item sem resposta, ou com resposta <c>Nothing</c>, conta como
        ''' <b>tem anexo</b>. Fechado por falta de prova.
        ''' </summary>
        Private Shared Function TemAnexo(mapa As Dictionary(Of String, AttachmentPresence),
                                         chave As String) As Boolean
            Dim p As AttachmentPresence = Nothing
            If Not mapa.TryGetValue(chave, p) OrElse p Is Nothing Then Return True
            If Not p.Tem.HasValue Then Return True
            Return p.Tem.Value
        End Function

        ''' <summary>
        ''' A presença de anexo por item, indexada pelo <c>EntryId</c>.
        '''
        ''' Chamada que falhou devolve dicionário vazio, e item ausente do
        ''' dicionário vira <b>tem anexo</b> lá em cima. Fechado por falta de
        ''' prova, como o resto da §29.
        ''' </summary>
        Private Function Anexos(itens As IReadOnlyList(Of ItemKey)) _
                                As Dictionary(Of String, AttachmentPresence)
            Dim saida As New Dictionary(Of String, AttachmentPresence)(StringComparer.Ordinal)
            Dim r = _broker.GetAttachmentPresenceAsync(itens, CancellationToken.None).
                    GetAwaiter().GetResult()
            If Not r.Succeeded Then Return saida
            For Each p In r.Value
                saida(p.Item.EntryId) = p
            Next
            Return saida
        End Function

        ''' <summary>
        ''' Quantas imagens embutidas. <c>Nothing</c> quando não houve resposta
        ''' para o item — e <c>Nothing</c> aqui não fecha nada, porque embutida
        ''' não bloqueia: ele só faz a tela dizer "não sei quantas" em vez de
        ''' "nenhuma".
        ''' </summary>
        Private Shared Function Embutidas(mapa As Dictionary(Of String, AttachmentPresence),
                                          chave As String) As Integer?
            Dim p As AttachmentPresence = Nothing
            If Not mapa.TryGetValue(chave, p) OrElse p Is Nothing Then Return Nothing
            Return p.Embutidas
        End Function

        ''' <summary>
        ''' Captura as mensagens da seleção numa visita só, prepara cada uma pelo
        ''' pipeline, e monta.
        '''
        ''' Mensagem que o pipeline recusa — corpo pela metade, referência
        ''' embutida, HTML que não dá para interpretar — <b>não entra</b>. O
        ''' envelope sai com menos itens que o grant aprovou, o
        ''' <c>Cobre</c> falha, e o cofre não emite. É o desfecho certo: uma
        ''' thread com um membro faltando não é a thread.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>UMA VISITA, E O ALINHAMENTO QUE ELA EXIGE</b>
        '''
        ''' Era uma ida ao broker por mensagem. Para uma thread de três, tanto faz;
        ''' para um lote de vinte, são vinte entradas na STA e vinte aquisições de
        ''' <c>Application</c> e <c>NameSpace</c> para produzir um envelope só.
        '''
        ''' A leitura em lote devolve <b>uma posição por item pedido</b>, e é por
        ''' isso que este laço anda por índice em vez de percorrer a saída: casar
        ''' ficha com mensagem pela ordem da lista devolvida só funciona enquanto
        ''' as duas listas tiverem o mesmo tamanho, e a graça do contrato é ele
        ''' valer quando <i>não</i> tiverem. Ficha trocada é a resposta do modelo
        ''' sendo aplicada à mensagem errada.
        '''
        ''' A conferência do tamanho é explícita e não confia no contrato: quem
        ''' implementa a interface pode errar, e este é o ponto do programa em que
        ''' errar significa rotular a mensagem errada.
        ''' </summary>
        Public Function Montar(operacao As AssistOperation, instrucao As String) _
                               As EnvelopeResult Implements IAssistContext.Montar
            Return New EnvelopeBuilder().Montar(operacao, instrucao, Partes())
        End Function

        ''' <summary>
        ''' <b>As partes da seleção — e é aqui que a leitura em lote acontece.</b>
        '''
        ''' Saiu de dentro do <see cref="Montar"/> porque a classificação de pasta
        ''' precisa exatamente disto e <b>não</b> do envelope: ela acrescenta o
        ''' controle do lote às partes antes de montar. Duas rotinas fazendo isto
        ''' divergiriam, e a divergência seria uma mensagem que atravessa o
        ''' pipeline por um caminho e é recusada pelo outro.
        ''' </summary>
        Friend Function Partes() As IReadOnlyList(Of MessagePart)
            ' "saida", e nao "partes": em VB um local nao pode ter o nome da
            ' funcao que o contem -- e este e o unico membro da familia do
            ' eclipse que o compilador acusa no lugar certo.
            Dim saida As New List(Of MessagePart)()
            Dim itens = _selecao().Itens
            If itens.Count = 0 Then
                Return saida
            End If

            Dim lidos = _broker.GetMessageSnapshotsAsync(itens, CancellationToken.None).
                        GetAwaiter().GetResult()
            If Not lidos.Succeeded OrElse lidos.Value Is Nothing Then
                Return saida
            End If

            ' TAMANHO DIFERENTE NAO E "LER O QUE DEU": e nao ler nada.
            ' Sem isto, uma implementacao que encolhesse a saida faria a ficha da
            ' mensagem 5 viajar com o corpo da mensagem 6.
            If lidos.Value.Count <> itens.Count Then
                Return saida
            End If

            Dim pedida = _selecao().Pasta

            For i = 0 To itens.Count - 1
                Dim retrato = lidos.Value(i)
                If retrato Is Nothing Then Continue For

                ' A PASTA E CONFERIDA DE NOVO, AQUI, presa a ESTE corpo.
                '
                ' O portao ja conferiu a pasta observada numa visita anterior, e
                ' entre aquela visita e esta a mensagem pode ter se movido. Conferir
                ' so la deixaria a janela aberta; conferir aqui a fecha, porque este
                ' e o corpo que vira bytes. Mesmo desenho do anexo.
                '
                ' Pasta vazia -- "nao deu para ler" -- tambem para: nao saber onde a
                ' mensagem esta nunca vira prova de que ela esta onde deveria.
                If retrato.Pasta Is Nothing OrElse Not retrato.Pasta.Equals(pedida) Then
                    Continue For
                End If

                Dim pronto = ContentPipeline.Preparar(retrato, FichaDe(itens(i)),
                                                      _tetoDoCorpo)
                If Not pronto.Ok Then Continue For

                saida.Add(pronto.Parte)
            Next

            Return saida
        End Function

        ''' <summary>
        ''' A ficha do item, ou vazio. <b>Exceção aqui vira vazio</b>, e não sobe:
        ''' um lote sem ficha é recusado adiante pela conferência do
        ''' <c>LoteDeClassificacao</c>, o que é bem melhor do que uma exceção
        ''' atravessando a montagem do envelope.
        ''' </summary>
        Private Function FichaDe(chave As ItemKey) As String
            If _ficha Is Nothing Then Return Nothing
            Try
                Return _ficha(chave)
            Catch
                Return Nothing
            End Try
        End Function

    End Class

End Namespace

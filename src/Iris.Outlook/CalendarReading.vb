Imports System.Collections.Generic
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' <b>Leitura do calendário — e a armadilha que ela existe para não cair.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A ORDEM IMPORTA, E ERRAR É SILENCIOSO</b>
    '''
    ''' Para que uma janela de datas inclua as ocorrências de séries, o OOM
    ''' exige <b>três coisas nesta ordem</b>:
    '''
    '''   1. <c>Items.Sort("[Start]")</c>
    '''   2. <c>Items.IncludeRecurrences = True</c>
    '''   3. <c>Items.Restrict(filtro)</c>
    '''
    ''' Fora dessa ordem o Outlook <b>não devolve erro</b>. Ele devolve uma
    ''' lista plausível — e o que ela tem de errado é pior do que eu supunha.
    '''
    ''' <b>MEDIDO em 28/08/2026</b>, contra a caixa real, invertendo para
    ''' <c>Restrict</c> antes de <c>IncludeRecurrences</c>: a leitura devolveu
    ''' <b>65 compromissos fora da janela pedida</b>. Ocorrências de
    ''' <i>REUNIÃO PERIÓDICA DE P&amp;D</i> de janeiro apareceram numa janela de
    ''' ±30 dias em torno de agosto.
    '''
    ''' Não é só perder ocorrência, então: a ordem errada faz a expansão
    ''' acontecer <b>ignorando o filtro</b>. Uma agenda com sete meses de
    ''' atraso, e nenhum erro em lugar nenhum.
    '''
    ''' É a mesma família dos defeitos que a Fase 0 catalogou: comportamento
    ''' que parece sucesso. Por isso a ordem está aqui em código, comentada, e
    ''' com o número que ela custa.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E POR QUE UMA JANELA FECHADA, SEMPRE</b>
    '''
    ''' Com <c>IncludeRecurrences</c> ligado, uma série sem data-fim é
    ''' <b>infinita</b>: iterar sem limite superior não termina. O filtro
    ''' fecha os dois lados, e o laço ainda tem um teto — porque um filtro
    ''' malformado degeneraria no mesmo laço infinito, e um travamento na fila
    ''' única da STA congela o aplicativo inteiro.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O FORMATO DA DATA NÃO É ESCOLHA</b>
    '''
    ''' O <c>Restrict</c> do OOM espera a data no formato da <b>localidade do
    ''' Outlook</b>, e não em ISO. Numa máquina em português, mandar
    ''' <c>2026-08-28</c> faz o filtro casar errado ou não casar. O formato
    ''' usado aqui — <c>MM/dd/yyyy hh:mm tt</c> com cultura invariante — é o
    ''' que a documentação do OOM prescreve e o único que não depende de qual
    ''' idioma o Office está usando.
    ''' </summary>
    Friend Module CalendarReading

        ''' <summary>
        ''' Teto de segurança do laço. Uma janela de sete dias com centenas de
        ''' ocorrências continua muito abaixo disto; alcançá-lo significa que o
        ''' filtro degenerou, e aí parar é melhor que travar a STA.
        ''' </summary>
        Private Const TetoDeItens As Integer = 5000

        Friend Function Ler(ns As OL.NameSpace, folder As FolderKey,
                            de As DateTimeOffset, ate As DateTimeOffset) _
                            As OperationResult(Of AppointmentWindow)

            If ate <= de Then
                Return OperationResult(Of AppointmentWindow).Fail(
                    ErrorKind.Unexpected, "janela invertida ou vazia")
            End If

            Dim pasta As OL.Folder = Nothing
            Try
                pasta = TryCast(ns.GetFolderFromID(folder.EntryId, folder.StoreId), OL.Folder)
            Catch ex As COMException
                Return OperationResult(Of AppointmentWindow).Fail(
                    ErrorKind.NotFound, "pasta de calendário não encontrada")
            End Try
            If pasta Is Nothing Then
                Return OperationResult(Of AppointmentWindow).Fail(
                    ErrorKind.NotFound, "pasta de calendário não encontrada")
            End If

            Dim itens As OL.Items = Nothing
            Dim janela As OL.Items = Nothing
            Try
                itens = pasta.Items

                ' A ORDEM. Ver o cabeçalho deste módulo: trocá-la não dá erro,
                ' dá uma lista errada com cara de certa.
                itens.Sort("[Start]", False)
                itens.IncludeRecurrences = True

                janela = itens.Restrict(CalendarFilter.Janela(de, ate))

                Dim resultado As New AppointmentWindow With {.De = de, .Ate = ate}
                Dim recusadas = 0
                Dim vistos = 0

                ' For Each sobre coleção COM é o que a §11 desaconselha, mas
                ' com IncludeRecurrences NÃO EXISTE Count confiável: a
                ' expansão é preguiçosa, e Count força a materialização
                ' inteira antes de qualquer leitura. GetFirst/GetNext é o
                ' idioma que o OOM oferece para isto.
                Dim atual As Object = janela.GetFirst()
                While atual IsNot Nothing AndAlso vistos < TetoDeItens
                    vistos += 1
                    Dim compromisso = TryCast(atual, OL.AppointmentItem)
                    If compromisso Is Nothing Then
                        ' Não é compromisso. Recusa DECLARADA, e não silêncio —
                        ' mesma disciplina do descarte da varredura, onde "28
                        ' de 30" viraria mistério sem o número.
                        recusadas += 1
                    Else
                        ' O contador e da JANELA, e cada item soma no dele.
                        ' Ver o que a mesma conta ja errou nas duas direcoes
                        ' em MessagePaging: aqui ela nasce por item, somando.
                        '
                        ' E SO SOMA SE O ITEM ENTROU. A primeira versao somava
                        ' antes de conferir o Nothing, entao um item que
                        ' fabricasse uma celula e depois fosse recusado
                        ' aparecia nas DUAS contas -- recusado e fabricado, que
                        ' o resumo apresenta como coisas diferentes. Fabricacao
                        ' de item que nao entrou nao descreve nada na tela.
                        Dim daqui = 0
                        Dim dto = Traduzir(compromisso, daqui)
                        If dto Is Nothing Then
                            recusadas += 1
                        Else
                            resultado.FabricatedCells += daqui
                            resultado.Items.Add(dto)
                            If dto.IsRecurring Then
                                resultado.FromRecurrence += 1
                            End If
                        End If
                    End If

                    Dim proximo As Object = Nothing
                    Try
                        proximo = janela.GetNext()
                    Catch ex As Exception
                        ' FALHA NO MEIO DA ENUMERACAO NAO E FIM DA COLECAO.
                        '
                        ' Ate 28/08/2026 este Catch devolvia Nothing, e o laco
                        ' terminava como se a janela tivesse acabado: a leitura
                        ' saia Ok, com uma lista plausivel e incompleta. Uma
                        ' agenda que perde o fim da semana sem avisar e pior
                        ' que uma agenda que falha.
                        resultado.Truncada = True
                        resultado.MotivoDoCorte =
                            "a leitura foi interrompida no meio (" & ex.GetType().Name & ")"
                        proximo = Nothing
                    End Try
                    ComHelpers.Release(atual)
                    atual = proximo
                End While

                ' TETO ALCANCADO: tambem e truncamento, e tambem era silencioso.
                If atual IsNot Nothing Then
                    resultado.Truncada = True
                    resultado.MotivoDoCorte =
                        $"a janela tem mais de {TetoDeItens} ocorrencias e a leitura parou no teto"
                End If
                ComHelpers.Release(atual)

                resultado.Skipped = recusadas
                Return OperationResult(Of AppointmentWindow).Ok(resultado)

            Catch ex As COMException
                ' O broker classifica o HRESULT no ReadAsync, que e quem sabe
                ' de retry e de message filter. Aqui a excecao vira uma recusa
                ' TRANSITORIA por padrao -- e o Busy do Outlook e o caso comum.
                Return OperationResult(Of AppointmentWindow).Fail(
                    ErrorKind.Busy, "falha ao ler o calendario: " & ex.GetType().Name)
            Finally
                ' Ordem inversa à aquisição, R7.
                ComHelpers.Release(janela)
                ComHelpers.Release(itens)
                ComHelpers.Release(pasta)
            End Try
        End Function

        Private Function Traduzir(a As OL.AppointmentItem, ByRef fabricadas As Integer) As AppointmentInfo
            Try
                Dim inicio = Safe(Function() a.Start)
                Dim fim = Safe(Function() a.End)

                Return New AppointmentInfo With {
                    .Key = New ItemKey(TextoDoCompromisso(Function() a.EntryID, fabricadas),
                                       StoreDe(a, fabricadas)),
                    .Subject = TextoDoCompromisso(Function() a.Subject, fabricadas),
                    .Location = TextoDoCompromisso(Function() a.Location, fabricadas),
                    .Start = New DateTimeOffset(inicio),
                    .End = New DateTimeOffset(fim),
                    .AllDayEvent = BooleanoDoCompromisso(Function() a.AllDayEvent, fabricadas),
                    .Organizer = TextoDoCompromisso(Function() a.Organizer, fabricadas),
                    .BusyStatus = DoBusy(a),
                    .ResponseStatus = DaResposta(a),
                    .IsRecurring = BooleanoDoCompromisso(Function() a.IsRecurring, fabricadas),
                    .RecipientCount = Contar(a, fabricadas)
                }
            Catch
                ' Item que não se deixa ler não derruba a janela inteira: ele
                ' vira recusa contada. Uma agenda que não abre porque um item
                ' está corrompido é pior que uma agenda com "1 recusado".
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' O <c>StoreID</c> do compromisso, pela PASTA.
        '''
        ''' <c>a.Parent</c> devolve <c>Object</c>, e <c>a.Parent.StoreID</c> seria
        ''' ligacao tardia — que o <c>Option Strict On</c> recusa, e com razao:
        ''' ela esconderia um RCW intermediario sem dono, que e a R7 outra vez.
        ''' Aqui o pai recebe nome, tipo, e e liberado.
        ''' </summary>
        Private Function StoreDe(a As OL.AppointmentItem, ByRef fabricadas As Integer) As String
            Dim pai As OL.Folder = Nothing
            Try
                pai = TryCast(a.Parent, OL.Folder)
                ' StoreID vazio nao e "sem store": e chave que nunca casa, e o
                ' sintoma aparece longe -- exatamente como o EntryID fabricado
                ' da paginacao.
                Dim id = If(pai?.StoreID, "")
                If id = "" Then fabricadas += 1
                Return id
            Catch
                fabricadas += 1
                Return ""
            Finally
                ComHelpers.Release(pai)
            End Try
        End Function

        Private Function Contar(a As OL.AppointmentItem, ByRef fabricadas As Integer) As Integer
            ' a.Recipients é objeto COM PRÓPRIO. `a.Recipients.Count` criaria
            ' um RCW intermediário sem dono — R7, já violado quatro vezes neste
            ' projeto, sempre em código que "só lia uma contagem".
            Dim r As OL.Recipients = Nothing
            Try
                r = a.Recipients
                Return r.Count
            Catch
                ' Zero participante e uma AFIRMACAO, e ela sai igual quando a
                ' leitura falha. Por isso conta.
                fabricadas += 1
                Return 0
            Finally
                ComHelpers.Release(r)
            End Try
        End Function

        Private Function DoBusy(a As OL.AppointmentItem) As AppointmentBusy
            Try
                Select Case a.BusyStatus
                    Case OL.OlBusyStatus.olFree : Return AppointmentBusy.Livre
                    Case OL.OlBusyStatus.olTentative : Return AppointmentBusy.Provisorio
                    Case OL.OlBusyStatus.olBusy : Return AppointmentBusy.Ocupado
                    Case OL.OlBusyStatus.olOutOfOffice : Return AppointmentBusy.ForaDoEscritorio
                    Case OL.OlBusyStatus.olWorkingElsewhere : Return AppointmentBusy.TrabalhandoEmOutroLugar
                    Case Else : Return AppointmentBusy.Desconhecido
                End Select
            Catch
                Return AppointmentBusy.Desconhecido
            End Try
        End Function

        Private Function DaResposta(a As OL.AppointmentItem) As AppointmentResponse
            Try
                Select Case a.ResponseStatus
                    Case OL.OlResponseStatus.olResponseNone : Return AppointmentResponse.NaoEhReuniao
                    Case OL.OlResponseStatus.olResponseOrganized : Return AppointmentResponse.Organizador
                    Case OL.OlResponseStatus.olResponseAccepted : Return AppointmentResponse.Aceito
                    Case OL.OlResponseStatus.olResponseTentative : Return AppointmentResponse.Provisorio
                    Case OL.OlResponseStatus.olResponseDeclined : Return AppointmentResponse.Recusado
                    Case OL.OlResponseStatus.olResponseNotResponded : Return AppointmentResponse.NaoRespondeu
                    Case Else : Return AppointmentResponse.Desconhecido
                End Select
            Catch
                Return AppointmentResponse.Desconhecido
            End Try
        End Function

        ''' <summary>
        ''' Texto ilegível <b>ou ausente</b> vira vazio — e conta nos dois casos.
        '''
        ''' <b>Sem teste do fio inteiro, e dito com esse nome:</b> os auxiliares
        ''' têm teste e o resumo da agenda tem teste, mas o trecho entre os dois
        ''' — <c>Traduzir</c> somando e o laço acumulando na janela — precisa de
        ''' um <c>AppointmentItem</c> real. É a mesma lacuna declarada do
        ''' <c>ContarAnexos</c> na paginação.
        '''
        ''' O sufixo <c>-DoCompromisso</c> existe porque estes passaram a ser
        ''' <c>Friend</c> para ter teste, e <c>Friend</c> em <c>Module</c> vale
        ''' para o assembly: <c>Texto</c> e <c>Booleano</c> colidiriam com os
        ''' homônimos de <c>MessageReading</c>, <c>DraftWriting</c> e
        ''' <c>MessagePaging</c>. É a armadilha que o CLAUDE.md lista doze vezes.
        ''' </summary>
        Friend Function TextoDoCompromisso(f As Func(Of String), ByRef fabricadas As Integer) As String
            Dim lido As String
            Try
                lido = f()
            Catch
                fabricadas += 1
                Return ""
            End Try

            If lido Is Nothing Then
                fabricadas += 1
                Return ""
            End If
            Return lido
        End Function

        ''' <summary>
        ''' Booleano ilegível vira <c>False</c> — e conta.
        '''
        ''' <c>AllDayEvent = False</c> e <c>IsRecurring = False</c> são
        ''' afirmações que a tela mostra como fato.
        ''' </summary>
        Friend Function BooleanoDoCompromisso(f As Func(Of Boolean), ByRef fabricadas As Integer) As Boolean
            Try
                Return f()
            Catch
                fabricadas += 1
                Return False
            End Try
        End Function

        Private Function Safe(f As Func(Of Date)) As Date
            Return f()
        End Function

    End Module

End Namespace

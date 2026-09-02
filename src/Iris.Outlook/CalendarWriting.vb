Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' <b>Escrita no calendário — e a invariante que a governa.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>SALVAR UM COMPROMISSO COM PARTICIPANTE É ENVIAR E-MAIL</b>
    '''
    ''' Esta é a única coisa que importa saber antes de ler o resto. Um
    ''' <c>AppointmentItem</c> com <c>MeetingStatus</c> diferente de
    ''' <c>olNonMeeting</c> é uma <b>reunião</b>: o Outlook manda convite ao
    ''' criar, manda atualização ao editar e manda cancelamento ao apagar. Não
    ''' existe "salvar sem avisar ninguém" — o envio é o comportamento normal
    ''' do objeto, e não um efeito colateral que dê para desligar.
    '''
    ''' O Iris tem uma regra que não admite exceção: <b>nada sai por e-mail
    ''' sem o usuário mandar</b>. Então este módulo:
    '''
    ''' <list type="bullet">
    ''' <item>nunca preenche <c>Recipients</c> ao criar — o compromisso nasce
    ''' e permanece um bloco na agenda de quem o criou;</item>
    ''' <item><b>recusa</b> editar ou apagar qualquer item que já seja reunião,
    ''' antes de tocar nele.</item>
    ''' </list>
    '''
    ''' A recusa é <see cref="ErrorKind.Denied"/> e não <c>Stale</c>: não é que
    ''' a chave envelheceu, é que a operação não é permitida neste item. Quem
    ''' quiser mexer numa reunião faz isso no Outlook, onde vê para quem o
    ''' convite vai.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E A IDENTIDADE MUDA NO SAVE</b>
    '''
    ''' Como em toda operação que grava neste projeto, o <c>EntryID</c>
    ''' <b>pode</b> mudar num <c>Save</c> — não é garantido que mude, e é por
    ''' isso que não dá para apostar em nenhuma das duas hipóteses. Toda
    ''' função daqui relê o item depois de gravar e devolve o compromisso
    ''' redescrito, e não só o desfecho. Esquecer isso já custou um
    ''' <c>NotFound</c> longe daqui, no <c>AddAttachment</c>.
    ''' </summary>
    Friend Module CalendarWriting

        ''' <summary>
        ''' Cria um compromisso <b>na pasta indicada</b>, e não no calendário
        ''' padrão.
        '''
        ''' <c>Items.Add</c> na pasta escolhida, em vez de
        ''' <c>Application.CreateItem</c> seguido de <c>Move</c>: criar no
        ''' calendário padrão e mover depois deixa, por um instante, um
        ''' compromisso na agenda de verdade de quem estiver olhando — e um
        ''' <c>Move</c> que falhe deixa ele lá.
        ''' </summary>
        ''' <param name="marcar">
        ''' Acionado <b>imediatamente antes</b> do primeiro efeito que fica no mundo.
        ''' É o que separa <i>"falhou e nada aconteceu"</i> de <i>"falhou e não se
        ''' sabe"</i> — ver <c>OutlookBroker.MutateAsync</c>, que tem o motivo por
        ''' extenso.
        ''' </param>
        Public Function Create(ns As OL.NameSpace, pasta As FolderKey,
                               rascunho As AppointmentDraft,
                               Optional marcar As Action = Nothing) _
                               As OperationResult(Of AppointmentInfo)
            ' A CERCA PASSA A SABER SE O EFEITO COMECOU.
            '
            ' Este Catch dizia mutationAttemptStarted:=True sempre. Uma falha em
            ' GetFolderFromID, em Items, na abertura do item ou numa atribuicao de
            ' propriedade -- tudo ANTES do Save -- era apresentada como possivel
            ' mutacao consumida, e a tela dizia que algo podia ter acontecido sobre
            ' o que nao aconteceu.
            '
            ' E o mesmo defeito que o despacho do broker tinha, um nivel abaixo: a
            ' captura local impedia a fase de verdade de chegar ao classificador.
            ' Achado por revisao externa em 02/09/2026.
            Dim comecou = False
            Dim aoComecar As Action =
                Sub()
                    comecou = True
                    If marcar IsNot Nothing Then marcar()
                End Sub

            If rascunho Is Nothing Then
                Return OperationResult(Of AppointmentInfo).Fail(ErrorKind.Unexpected, "rascunho nulo")
            End If

            Dim recusa = RecusarRascunho(rascunho)
            If recusa IsNot Nothing Then
                Return OperationResult(Of AppointmentInfo).Fail(ErrorKind.Denied, recusa)
            End If

            Dim destino As OL.Folder = Nothing
            Dim itens As OL.Items = Nothing
            Dim item As OL.AppointmentItem = Nothing
            Try
                destino = TryCast(ns.GetFolderFromID(pasta.EntryId, pasta.StoreId), OL.Folder)
                If destino Is Nothing Then
                    Return OperationResult(Of AppointmentInfo).Fail(ErrorKind.NotFound, "pasta")
                End If

                itens = destino.Items
                item = TryCast(itens.Add(OL.OlItemType.olAppointmentItem), OL.AppointmentItem)
                If item Is Nothing Then
                    Return OperationResult(Of AppointmentInfo).Fail(ErrorKind.Unexpected, "Items.Add")
                End If

                Aplicar(item, rascunho)

                ' NASCE E CONTINUA SEM PARTICIPANTE. Nao ha Recipients aqui, e
                ' e por isso que o Save nao manda convite nenhum.
                aoComecar()
                item.Save()

                Return DepoisDoSave(Descrever(item))
            Catch ex As COMException
                Return OperationResult(Of AppointmentInfo).Fail(
                    OutlookFailurePolicy.ClassifyFailure(
                        ex.HResult, isMutation:=True, mutationAttemptStarted:=comecou),
                    "falha ao criar compromisso: " & ex.GetType().Name)
            Finally
                ComHelpers.Release(item)
                ComHelpers.Release(itens)
                ComHelpers.Release(destino)
            End Try
        End Function

        ''' <summary>
        ''' Edita um compromisso que <b>não</b> é reunião.
        '''
        ''' A conferência acontece <b>antes</b> de qualquer atribuição: mudar o
        ''' assunto e só então descobrir que era reunião deixaria o item sujo,
        ''' e o <c>Save</c> seguinte — feito por qualquer outra coisa — mandaria
        ''' a atualização.
        ''' </summary>
        ''' <param name="marcar">
        ''' Acionado <b>imediatamente antes</b> do primeiro efeito que fica no mundo.
        ''' É o que separa <i>"falhou e nada aconteceu"</i> de <i>"falhou e não se
        ''' sabe"</i> — ver <c>OutlookBroker.MutateAsync</c>, que tem o motivo por
        ''' extenso.
        ''' </param>
        Public Function Update(ns As OL.NameSpace, chave As AppointmentKey,
                               rascunho As AppointmentDraft,
                               Optional marcar As Action = Nothing) _
                               As OperationResult(Of AppointmentInfo)
            ' A CERCA PASSA A SABER SE O EFEITO COMECOU.
            '
            ' Este Catch dizia mutationAttemptStarted:=True sempre. Uma falha em
            ' GetFolderFromID, em Items, na abertura do item ou numa atribuicao de
            ' propriedade -- tudo ANTES do Save -- era apresentada como possivel
            ' mutacao consumida, e a tela dizia que algo podia ter acontecido sobre
            ' o que nao aconteceu.
            '
            ' E o mesmo defeito que o despacho do broker tinha, um nivel abaixo: a
            ' captura local impedia a fase de verdade de chegar ao classificador.
            ' Achado por revisao externa em 02/09/2026.
            Dim comecou = False
            Dim aoComecar As Action =
                Sub()
                    comecou = True
                    If marcar IsNot Nothing Then marcar()
                End Sub

            If rascunho Is Nothing Then
                Return OperationResult(Of AppointmentInfo).Fail(ErrorKind.Unexpected, "rascunho nulo")
            End If

            Dim recusa = RecusarRascunho(rascunho)
            If recusa IsNot Nothing Then
                Return OperationResult(Of AppointmentInfo).Fail(ErrorKind.Denied, recusa)
            End If

            Dim item As OL.AppointmentItem = Nothing
            Try
                item = Abrir(ns, chave)
                If item Is Nothing Then
                    Return OperationResult(Of AppointmentInfo).Fail(ErrorKind.NotFound, "compromisso")
                End If

                If EhReuniao(item) Then
                    Return OperationResult(Of AppointmentInfo).Fail(ErrorKind.Denied, MotivoDaReuniao)
                End If

                Aplicar(item, rascunho)
                aoComecar()
                item.Save()

                Return DepoisDoSave(Descrever(item))
            Catch ex As COMException
                Return OperationResult(Of AppointmentInfo).Fail(
                    OutlookFailurePolicy.ClassifyFailure(
                        ex.HResult, isMutation:=True, mutationAttemptStarted:=comecou),
                    "falha ao editar compromisso: " & ex.GetType().Name)
            Finally
                ComHelpers.Release(item)
            End Try
        End Function

        ''' <summary>
        ''' Apaga um compromisso que <b>não</b> é reunião.
        '''
        ''' Apagar reunião manda cancelamento para todo mundo — é a forma mais
        ''' cara de violar a invariante, porque o estrago chega a terceiros.
        ''' </summary>
        ''' <param name="marcar">
        ''' Acionado <b>imediatamente antes</b> do primeiro efeito que fica no mundo.
        ''' É o que separa <i>"falhou e nada aconteceu"</i> de <i>"falhou e não se
        ''' sabe"</i> — ver <c>OutlookBroker.MutateAsync</c>, que tem o motivo por
        ''' extenso.
        ''' </param>
        Public Function Delete(ns As OL.NameSpace, chave As AppointmentKey,
                               Optional marcar As Action = Nothing) _
                               As OperationResult(Of Boolean)
            ' A CERCA PASSA A SABER SE O EFEITO COMECOU.
            '
            ' Este Catch dizia mutationAttemptStarted:=True sempre. Uma falha em
            ' GetFolderFromID, em Items, na abertura do item ou numa atribuicao de
            ' propriedade -- tudo ANTES do Save -- era apresentada como possivel
            ' mutacao consumida, e a tela dizia que algo podia ter acontecido sobre
            ' o que nao aconteceu.
            '
            ' E o mesmo defeito que o despacho do broker tinha, um nivel abaixo: a
            ' captura local impedia a fase de verdade de chegar ao classificador.
            ' Achado por revisao externa em 02/09/2026.
            Dim comecou = False
            Dim aoComecar As Action =
                Sub()
                    comecou = True
                    If marcar IsNot Nothing Then marcar()
                End Sub

            Dim item As OL.AppointmentItem = Nothing
            Try
                item = Abrir(ns, chave)
                If item Is Nothing Then
                    Return OperationResult(Of Boolean).Fail(ErrorKind.NotFound, "compromisso")
                End If

                If EhReuniao(item) Then
                    Return OperationResult(Of Boolean).Fail(ErrorKind.Denied, MotivoDaReuniao)
                End If

                aoComecar()
                item.Delete()
                Return OperationResult(Of Boolean).Ok(True)
            Catch ex As COMException
                Return OperationResult(Of Boolean).Fail(
                    OutlookFailurePolicy.ClassifyFailure(
                        ex.HResult, isMutation:=True, mutationAttemptStarted:=comecou),
                    "falha ao apagar compromisso: " & ex.GetType().Name)
            Finally
                ComHelpers.Release(item)
            End Try
        End Function

        ' ==============================================================

        Friend Const MotivoDaReuniao As String =
            "este compromisso é uma reunião com participantes, e salvá-lo " &
            "mandaria convite, atualização ou cancelamento por e-mail. O Iris " &
            "não envia. Use o Outlook, onde você vê para quem vai."

        ''' <summary>
        ''' <b>Este item manda e-mail se for salvo?</b>
        '''
        ''' <c>MeetingStatus</c> é a pergunta certa, e não "tem destinatário":
        ''' uma reunião cancelada ou recebida também dispara envio no
        ''' <c>Save</c>, e pode estar com a lista vazia no momento em que se
        ''' olha. Qualquer coisa que não seja <c>olNonMeeting</c> está fora.
        '''
        ''' Falha ao ler <b>fecha</b>: não saber se é reunião tem de valer como
        ''' "é". O contrário deixaria uma leitura ruim autorizar um envio.
        ''' </summary>
        Friend Function EhReuniao(item As OL.AppointmentItem) As Boolean
            Try
                Return item.MeetingStatus <> OL.OlMeetingStatus.olNonMeeting
            Catch
                Return True
            End Try
        End Function

        ''' <summary>
        ''' O que impede um rascunho de virar compromisso. <c>Nothing</c> quando
        ''' não há impedimento.
        '''
        ''' Existe separada para ter teste sem COM — é a única parte da guarda
        ''' que dá para provar sem um <c>AppointmentItem</c> real, e é a que
        ''' decide se o Iris chega a tocar no Outlook.
        ''' </summary>
        Friend Function RecusarRascunho(rascunho As AppointmentDraft) As String
            If rascunho Is Nothing Then Return "rascunho nulo"

            If String.IsNullOrWhiteSpace(rascunho.Subject) Then
                Return "um compromisso sem assunto vira um bloco anônimo na agenda"
            End If

            ' FIM ANTES DO INICIO. O OOM aceita e a agenda mostra coisa
            ' impossivel -- e o Restrict da leitura passa a nao achar o item,
            ' que e o sintoma aparecendo longe da causa.
            If rascunho.Ate < rascunho.De Then
                Return "o fim do compromisso é anterior ao início"
            End If

            Return Nothing
        End Function

        Private Function Abrir(ns As OL.NameSpace, chave As AppointmentKey) As OL.AppointmentItem
            If chave Is Nothing Then Return Nothing
            Try
                Return TryCast(ns.GetItemFromID(chave.Item.EntryId, chave.Item.StoreId),
                               OL.AppointmentItem)
            Catch ex As COMException
                Return Nothing
            End Try
        End Function

        Private Sub Aplicar(item As OL.AppointmentItem, rascunho As AppointmentDraft)
            item.Subject = rascunho.Subject
            item.Location = If(rascunho.Location, "")
            item.Body = If(rascunho.Body, "")
            item.AllDayEvent = rascunho.AllDayEvent

            ' HORA LOCAL SEM KIND, que e o que o OOM espera. Mandar UTC aqui
            ' desloca o compromisso pelo offset, e o sintoma e o item aparecer
            ' na hora errada -- o mesmo motivo pelo qual a leitura converte na
            ' direcao contraria.
            item.Start = rascunho.De.LocalDateTime
            item.End = rascunho.Ate.LocalDateTime
        End Sub

        ''' <summary>
        ''' Relê o item recém-gravado. <b>Depois</b> do <c>Save</c>, sempre:
        ''' o <c>EntryID</c> pode ter mudado.
        ''' </summary>
        ''' <summary>
        ''' <b>Depois do <c>Save</c>, chave vazia é ambiguidade — não sucesso.</b>
        '''
        ''' <c>Descrever</c> engole falha de leitura: <c>TextoDoItem</c> devolve vazio
        ''' e apenas incrementa um contador de campos fabricados, que ninguém olhava.
        ''' Um erro COM transitório ao reler o <c>EntryID</c> novo produzia um
        ''' compromisso <b>gravado</b> com uma chave que ninguém consegue reabrir —
        ''' devolvido como sucesso, e sem como o classificador do broker enxergar,
        ''' porque a exceção nunca subiu.
        '''
        ''' É a regra fundadora do projeto: <i>toda operação que salva devolve a
        ''' identidade nova</i>. Rascunho, contato e tarefa ganharam esta guarda em
        ''' 01/09/2026; o calendário ficou de fora porque ninguém foi procurar os
        ''' irmãos. Achado por revisão externa em 02/09/2026.
        ''' </summary>
        Private Function DepoisDoSave(info As AppointmentInfo) _
                         As OperationResult(Of AppointmentInfo)
            If info Is Nothing OrElse info.Key Is Nothing OrElse info.Key.IsEmpty Then
                Return OperationResult(Of AppointmentInfo).Fail(ErrorKind.Ambiguous,
                    "o compromisso foi gravado e a identidade nova nao pode ser lida")
            End If
            Return OperationResult(Of AppointmentInfo).Ok(info)
        End Function

        Private Function Descrever(item As OL.AppointmentItem) As AppointmentInfo
            Dim fabricadas = 0
            Return New AppointmentInfo With {
                .Key = New ItemKey(TextoDoItem(Function() item.EntryID, fabricadas),
                                   StoreDoItem(item, fabricadas)),
                .Subject = TextoDoItem(Function() item.Subject, fabricadas),
                .Location = TextoDoItem(Function() item.Location, fabricadas),
                .Start = New DateTimeOffset(DateTime.SpecifyKind(item.Start, DateTimeKind.Local)),
                .End = New DateTimeOffset(DateTime.SpecifyKind(item.End, DateTimeKind.Local)),
                .AllDayEvent = item.AllDayEvent,
                .Organizer = TextoDoItem(Function() item.Organizer, fabricadas),
                .BusyStatus = AppointmentBusy.Desconhecido,
                .ResponseStatus = AppointmentResponse.NaoEhReuniao,
                .IsRecurring = False,
                .RecipientCount = 0
            }
        End Function

        Private Function TextoDoItem(getter As Func(Of String), ByRef fabricadas As Integer) As String
            Dim lido As String
            Try
                lido = getter()
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
        ''' O <c>StoreID</c> pela PASTA. <c>item.Parent</c> devolve
        ''' <c>Object</c>, e encadear seria ligação tardia mais um RCW sem dono
        ''' — a R7, que este projeto já violou quatro vezes.
        ''' </summary>
        Private Function StoreDoItem(item As OL.AppointmentItem, ByRef fabricadas As Integer) As String
            Dim pai As OL.Folder = Nothing
            Try
                pai = TryCast(item.Parent, OL.Folder)
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

    End Module

End Namespace

Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' <b>Tarefas — leitura e escrita, e a invariante que as governa.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>TAREFA ATRIBUÍDA CONVERSA POR E-MAIL</b>
    '''
    ''' É a mesma armadilha do calendário, num objeto diferente.
    ''' <c>TaskItem.Assign()</c> seguido de <c>Send()</c> manda um pedido de
    ''' tarefa <b>por e-mail</b>; depois disso, cada mudança de status vira uma
    ''' atualização que vai e volta pela caixa. Salvar uma tarefa atribuída não
    ''' é uma escrita local.
    '''
    ''' A regra do Iris não admite exceção: <b>nada sai por e-mail sem o
    ''' usuário mandar</b>. Então:
    '''
    ''' <list type="bullet">
    ''' <item><see cref="TaskDraft"/> não tem responsável — não existe caminho
    ''' para o Iris atribuir;</item>
    ''' <item>concluir ou apagar <b>recusa</b> tarefa atribuída, antes de tocar
    ''' nela.</item>
    ''' </list>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E O "SEM PRAZO" DO OUTLOOK É UMA DATA</b>
    '''
    ''' <c>TaskItem.DueDate</c> nunca é nulo: quando não há prazo, ele vale
    ''' <c>4501-01-01</c>. Traduzir isso para um <c>DateTimeOffset</c> comum
    ''' transformaria "não tem prazo" em "vence em 4501" — ausência virando
    ''' fato, que é a família de defeito que esta base passou uma série inteira
    ''' de revisões corrigindo. Aqui ela vira <c>Nothing</c>, e a tela diz "sem
    ''' prazo".
    ''' </summary>
    Friend Module TaskWriting

        ''' <summary>O sentinela de "sem data" do Outlook.</summary>
        Friend ReadOnly SemData As New Date(4501, 1, 1)

        Friend Const MotivoDaAtribuicao As String =
            "esta tarefa está atribuída a alguém, e mexer nela manda " &
            "atualização por e-mail. O Iris não envia. Use o Outlook, onde " &
            "você vê com quem a tarefa está."

        ''' <summary>
        ''' Lê as tarefas de uma pasta. Leitura pura: nada aqui grava.
        '''
        ''' <c>teto</c> existe pelo mesmo motivo do teto da agenda: uma pasta
        ''' de tarefas de uma caixa antiga pode ter milhares, e ler tudo para
        ''' mostrar dez é gastar o Outlook à toa. Bater no teto é
        ''' <b>truncamento</b>, e truncamento é dito.
        ''' </summary>
        Public Function Ler(ns As OL.NameSpace, pasta As FolderKey,
                            teto As Integer) As OperationResult(Of TaskList)
            Dim destino As OL.Folder = Nothing
            Dim itens As OL.Items = Nothing
            Try
                destino = TryCast(ns.GetFolderFromID(pasta.EntryId, pasta.StoreId), OL.Folder)
                If destino Is Nothing Then
                    Return OperationResult(Of TaskList).Fail(ErrorKind.NotFound, "pasta")
                End If

                Dim saida As New TaskList()
                Dim recusadas = 0

                itens = destino.Items
                itens.Sort("[DueDate]")

                Dim atual As Object = itens.GetFirst()
                While atual IsNot Nothing
                    If saida.Items.Count >= teto Then
                        saida.Truncada = True
                        saida.MotivoDoCorte =
                            $"a pasta tem mais de {teto} tarefas e a leitura parou no teto"
                        ComHelpers.Release(atual)
                        Exit While
                    End If

                    Dim t = TryCast(atual, OL.TaskItem)
                    If t Is Nothing Then
                        ' Item que nao e tarefa numa pasta de tarefas: existe, e
                        ' recusa contada e melhor que silencio.
                        recusadas += 1
                    Else
                        Dim dto = Traduzir(t)
                        If dto Is Nothing Then recusadas += 1 Else saida.Items.Add(dto)
                    End If

                    Dim proximo As Object = Nothing
                    Try
                        proximo = itens.GetNext()
                    Catch ex As Exception
                        ' FALHA NO MEIO NAO E FIM DA COLECAO -- a mesma licao
                        ' que a leitura do calendario aprendeu em 28/08.
                        saida.Truncada = True
                        saida.MotivoDoCorte =
                            "a leitura foi interrompida no meio (" & ex.GetType().Name & ")"
                        proximo = Nothing
                    End Try
                    ComHelpers.Release(atual)
                    atual = proximo
                End While

                saida.Skipped = recusadas
                Return OperationResult(Of TaskList).Ok(saida)
            Catch ex As COMException
                Return OperationResult(Of TaskList).Fail(
                    ErrorKind.Busy, "falha ao ler tarefas: " & ex.GetType().Name)
            Finally
                ComHelpers.Release(itens)
                ComHelpers.Release(destino)
            End Try
        End Function

        ''' <summary>
        ''' Cria uma tarefa <b>na pasta indicada</b>.
        '''
        ''' <c>Items.Add</c> na pasta escolhida, e não criar na padrão e mover:
        ''' um <c>Move</c> que falhe deixaria a tarefa na lista de verdade.
        ''' </summary>
        Public Function Create(ns As OL.NameSpace, pasta As FolderKey,
                               rascunho As TaskDraft) As OperationResult(Of TaskInfo)
            Dim recusa = RecusarRascunho(rascunho)
            If recusa IsNot Nothing Then
                Return OperationResult(Of TaskInfo).Fail(ErrorKind.Denied, recusa)
            End If

            Dim destino As OL.Folder = Nothing
            Dim itens As OL.Items = Nothing
            Dim t As OL.TaskItem = Nothing
            Try
                destino = TryCast(ns.GetFolderFromID(pasta.EntryId, pasta.StoreId), OL.Folder)
                If destino Is Nothing Then
                    Return OperationResult(Of TaskInfo).Fail(ErrorKind.NotFound, "pasta")
                End If

                itens = destino.Items
                t = TryCast(itens.Add(OL.OlItemType.olTaskItem), OL.TaskItem)
                If t Is Nothing Then
                    Return OperationResult(Of TaskInfo).Fail(ErrorKind.Unexpected, "Items.Add")
                End If

                t.Subject = rascunho.Subject
                t.Body = If(rascunho.Body, "")
                If rascunho.Vence.HasValue Then
                    t.DueDate = rascunho.Vence.Value.LocalDateTime.Date
                Else
                    t.DueDate = SemData
                End If

                ' NASCE E CONTINUA SEM RESPONSAVEL. Nao ha Assign aqui, e e por
                ' isso que o Save nao manda pedido nenhum.
                t.Save()

                ' O SAVE ACONTECEU. Ver o comentario gemeo em Concluir.
                Return DepoisDoSave(t, "criada")
            Catch ex As COMException
                Return OperationResult(Of TaskInfo).Fail(
                    OutlookFailurePolicy.ClassifyFailure(
                        ex.HResult, isMutation:=True, mutationAttemptStarted:=True),
                    "falha ao criar tarefa: " & ex.GetType().Name)
            Finally
                ComHelpers.Release(t)
                ComHelpers.Release(itens)
                ComHelpers.Release(destino)
            End Try
        End Function

        ''' <summary>
        ''' Marca como concluída — e <b>recusa</b> tarefa atribuída, porque ali
        ''' concluir manda atualização a quem atribuiu.
        ''' </summary>
        Public Function Concluir(ns As OL.NameSpace, chave As TaskKey) As OperationResult(Of TaskInfo)
            Dim t As OL.TaskItem = Nothing
            Try
                t = Abrir(ns, chave)
                If t Is Nothing Then
                    Return OperationResult(Of TaskInfo).Fail(ErrorKind.NotFound, "tarefa")
                End If

                If EhAtribuida(t) Then
                    Return OperationResult(Of TaskInfo).Fail(ErrorKind.Denied, MotivoDaAtribuicao)
                End If

                t.Complete = True
                t.Save()

                ' O SAVE ACONTECEU. Se a releitura falhar daqui para a frente, a
                ' tarefa MUDOU e a identidade nova se perdeu -- e isso e
                ' ambiguidade, nao sucesso. Ok(Nothing) dava ao chamador um
                ' resultado formalmente bem-sucedido sem tarefa e sem EntryID.
                ' Achado por revisao externa em 01/09/2026.
                Return DepoisDoSave(t, "concluida")
            Catch ex As COMException
                Return OperationResult(Of TaskInfo).Fail(
                    OutlookFailurePolicy.ClassifyFailure(
                        ex.HResult, isMutation:=True, mutationAttemptStarted:=True),
                    "falha ao concluir tarefa: " & ex.GetType().Name)
            Finally
                ComHelpers.Release(t)
            End Try
        End Function

        ' ==============================================================

        ''' <summary>
        ''' <b>Mexer nesta tarefa conversa com alguém?</b>
        '''
        ''' <c>DelegationState</c> é a pergunta certa: ela cobre a tarefa que eu
        ''' atribuí a outro <i>e</i> a que outro me atribuiu, que são os dois
        ''' lados em que um <c>Save</c> vira e-mail.
        '''
        ''' <b>Falha ao ler fecha.</b> Não saber tem de valer como "sim" — o
        ''' contrário deixaria uma leitura ruim autorizar um envio.
        ''' </summary>
        Friend Function EhAtribuida(t As OL.TaskItem) As Boolean
            Try
                Return t.DelegationState <> OL.OlTaskDelegationState.olTaskNotDelegated
            Catch
                Return True
            End Try
        End Function

        ''' <summary>
        ''' O que impede um rascunho de virar tarefa. <c>Nothing</c> quando não
        ''' há impedimento. Separada para ter teste sem COM.
        ''' </summary>
        Friend Function RecusarRascunho(rascunho As TaskDraft) As String
            If rascunho Is Nothing Then Return "rascunho nulo"
            If String.IsNullOrWhiteSpace(rascunho.Subject) Then
                Return "uma tarefa sem assunto não diz o que é para fazer"
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' A data do Outlook virando <c>DateTimeOffset?</c>.
        '''
        ''' O sentinela <c>4501-01-01</c> vira <c>Nothing</c>: "sem prazo" é
        ''' ausência, e mantê-la como ausência é o que impede a tela de
        ''' inventar um vencimento.
        ''' </summary>
        Friend Function Vencimento(bruta As Date) As DateTimeOffset?
            If bruta.Date >= SemData Then Return Nothing
            Return New DateTimeOffset(DateTime.SpecifyKind(bruta, DateTimeKind.Local))
        End Function

        Private Function Abrir(ns As OL.NameSpace, chave As TaskKey) As OL.TaskItem
            If chave Is Nothing Then Return Nothing
            Try
                Return TryCast(ns.GetItemFromID(chave.Item.EntryId, chave.Item.StoreId),
                               OL.TaskItem)
            Catch ex As COMException
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Relê a tarefa. <b>Depois</b> do <c>Save</c>, sempre: o
        ''' <c>EntryID</c> pode ter mudado.
        ''' </summary>
        ''' <summary>
        ''' <b>Depois do <c>Save</c>, releitura que falha é ambiguidade.</b>
        '''
        ''' A mutação aconteceu; o que se perdeu foi a identidade nova. Devolver
        ''' sucesso com <c>Nothing</c> — ou com uma chave vazia, que <c>Traduzir</c>
        ''' também sabe produzir — dava ao chamador um resultado formalmente bom e
        ''' inutilizável, e sem como notar. É a regra do projeto: <i>toda operação
        ''' que salva devolve a identidade nova</i>.
        ''' </summary>
        Private Function DepoisDoSave(t As OL.TaskItem, oQue As String) As OperationResult(Of TaskInfo)
            Dim descrita = Traduzir(t)
            If descrita Is Nothing OrElse descrita.Key Is Nothing OrElse descrita.Key.IsEmpty Then
                Return OperationResult(Of TaskInfo).Fail(ErrorKind.Ambiguous,
                    $"a tarefa foi {oQue} e a identidade nova nao pode ser lida")
            End If
            Return OperationResult(Of TaskInfo).Ok(descrita)
        End Function

        Private Function Traduzir(t As OL.TaskItem) As TaskInfo
            Try
                Return New TaskInfo With {
                    .Key = New ItemKey(Texto(Function() t.EntryID), StoreDe(t)),
                    .Subject = Texto(Function() t.Subject),
                    .Vence = Vencimento(t.DueDate),
                    .Concluida = t.Complete,
                    .Atribuida = EhAtribuida(t)
                }
            Catch
                ' Item que nao se deixa ler nao derruba a lista: vira recusa
                ' contada, como no calendario.
                Return Nothing
            End Try
        End Function

        Private Function Texto(getter As Func(Of String)) As String
            Try
                Return If(getter(), "")
            Catch
                Return ""
            End Try
        End Function

        Private Function StoreDe(t As OL.TaskItem) As String
            Dim pai As OL.Folder = Nothing
            Try
                pai = TryCast(t.Parent, OL.Folder)
                Return If(pai?.StoreID, "")
            Catch
                Return ""
            Finally
                ComHelpers.Release(pai)
            End Try
        End Function

    End Module

End Namespace

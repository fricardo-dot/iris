Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' <b>Contatos — leitura e escrita, e as três coisas que os separam das
    ''' outras fases.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A PRIMEIRA: SALVAR CONTATO NÃO MANDA E-MAIL — MAS ENCAMINHAR
    ''' MANDA</b>
    '''
    ''' Ao contrário da reunião e da tarefa atribuída, um <c>ContactItem</c>
    ''' comum é escrita local: <c>Save()</c> não fala com ninguém. O único
    ''' caminho de envio deste objeto é <c>ForwardAsVcard()</c>, que devolve um
    ''' <c>MailItem</c> pronto para sair.
    '''
    ''' Então a regra aqui não é uma guarda condicional, é uma <b>ausência</b>:
    ''' não existe operação de encaminhar cartão, e há teste que varre este
    ''' fonte atrás de <c>Forward</c> e <c>Send</c>. É o mesmo teste que a Fase
    ''' 5 ganhou depois da revisão, pelo mesmo motivo — o desenho estar certo
    ''' hoje não impede alguém de acrescentar a chamada amanhã.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A SEGUNDA: PASTA VAZIA NÃO É CATÁLOGO VAZIO</b>
    '''
    ''' Numa conta corporativa os contatos vivem no <b>GAL</b>, e o GAL está
    ''' fora de escopo pela §8. Medido em 28/08/2026 nesta caixa: a pasta padrão
    ''' de Contatos tem <b>0 itens</b>, com a organização inteira endereçável.
    '''
    ''' Uma leitura que devolvesse lista vazia e mais nada faria a tela dizer
    ''' "nenhum contato", que é afirmar ausência a partir de não ter olhado —
    ''' sobre pessoas. Por isso <c>ForaDoAlcance</c> é preenchido <b>aqui</b>,
    ''' junto com o resultado, e não montado pela tela: ressalva que depende de
    ''' alguém lembrar de escrever é ressalva que some na próxima tela.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A TERCEIRA: CRIAR CONTATO DUPLICADO ESTRAGA EM SILÊNCIO</b>
    '''
    ''' Um compromisso duplicado aparece duas vezes na agenda e alguém apaga.
    ''' Um contato duplicado fica, e o catálogo de endereços passa a ter duas
    ''' fichas da mesma pessoa com dados diferentes — e quem completar
    ''' endereço no Outlook vai acertar uma das duas por sorte.
    '''
    ''' <see cref="Procurar"/> existe para a tela poder avisar antes. Ela não
    ''' bloqueia: encontrar homônimo é comum e legítimo. Ela informa.
    ''' </summary>
    Friend Module ContactWriting

        ''' <summary>
        ''' Lê os contatos de uma pasta. Leitura pura: nada aqui grava.
        '''
        ''' O teto tem o mesmo motivo do teto da agenda e do das tarefas — uma
        ''' pasta antiga pode ter milhares, e ler tudo para mostrar dez gasta o
        ''' Outlook à toa. Bater no teto é truncamento, e truncamento é dito.
        ''' </summary>
        Public Function Ler(ns As OL.NameSpace, pasta As FolderKey,
                            teto As Integer) As OperationResult(Of ContactList)
            Dim destino As OL.Folder = Nothing
            Dim itens As OL.Items = Nothing
            Try
                destino = TryCast(ns.GetFolderFromID(pasta.EntryId, pasta.StoreId), OL.Folder)
                If destino Is Nothing Then
                    Return OperationResult(Of ContactList).Fail(ErrorKind.NotFound, "pasta")
                End If

                Dim saida As New ContactList With {.ForaDoAlcance = RegrasDeContato.ForaDoAlcance}
                Dim recusados = 0

                itens = destino.Items
                itens.Sort("[FileAs]")

                Dim atual As Object = itens.GetFirst()
                While atual IsNot Nothing
                    If saida.Items.Count >= teto Then
                        saida.Truncada = True
                        saida.MotivoDoCorte =
                            $"a pasta tem mais de {teto} contatos e a leitura parou no teto"
                        ComHelpers.Release(atual)
                        Exit While
                    End If

                    ' UMA PASTA DE CONTATOS TEM DISTLISTITEM TAMBEM, e lista de
                    ' distribuicao nao e contato. Recusa contada, e nao silencio:
                    ' a diferenca entre "nao ha" e "nao li" e o assunto desta fase.
                    Dim c = TryCast(atual, OL.ContactItem)
                    If c Is Nothing Then
                        recusados += 1
                    Else
                        Dim dto = Traduzir(c)
                        If dto Is Nothing Then recusados += 1 Else saida.Items.Add(dto)
                    End If

                    Dim proximo As Object = Nothing
                    Try
                        proximo = itens.GetNext()
                    Catch ex As Exception
                        ' FALHA NO MEIO NAO E FIM DA COLECAO -- a licao que a
                        ' leitura do calendario aprendeu em 28/08.
                        saida.Truncada = True
                        saida.MotivoDoCorte =
                            "a leitura foi interrompida no meio (" & ex.GetType().Name & ")"
                        proximo = Nothing
                    End Try
                    ComHelpers.Release(atual)
                    atual = proximo
                End While

                saida.Skipped = recusados
                Return OperationResult(Of ContactList).Ok(saida)
            Catch ex As COMException
                Return OperationResult(Of ContactList).Fail(
                    ErrorKind.Busy, "falha ao ler contatos: " & ex.GetType().Name)
            Finally
                ComHelpers.Release(itens)
                ComHelpers.Release(destino)
            End Try
        End Function

        ''' <summary>
        ''' Cria um contato <b>na pasta indicada</b>.
        '''
        ''' <c>Items.Add</c> na pasta escolhida, e não criar na padrão e mover:
        ''' um <c>Move</c> que falhe deixaria o contato no catálogo de verdade.
        ''' </summary>
        ''' <param name="marcar">
        ''' Acionado <b>imediatamente antes</b> do primeiro efeito que fica no mundo.
        ''' É o que separa <i>"falhou e nada aconteceu"</i> de <i>"falhou e não se
        ''' sabe"</i> — ver <c>OutlookBroker.MutateAsync</c>, que tem o motivo por
        ''' extenso.
        ''' </param>
        Public Function Create(ns As OL.NameSpace, pasta As FolderKey,
                               rascunho As ContactDraft,
                                                                            Optional marcar As Action = Nothing) _
                                                                            As OperationResult(Of ContactInfo)
            Dim recusa = RecusarRascunho(rascunho)
            If recusa IsNot Nothing Then
                Return OperationResult(Of ContactInfo).Fail(ErrorKind.Denied, recusa)
            End If

            Dim destino As OL.Folder = Nothing
            Dim itens As OL.Items = Nothing
            Dim c As OL.ContactItem = Nothing
            Try
                destino = TryCast(ns.GetFolderFromID(pasta.EntryId, pasta.StoreId), OL.Folder)
                If destino Is Nothing Then
                    Return OperationResult(Of ContactInfo).Fail(ErrorKind.NotFound, "pasta")
                End If

                itens = destino.Items
                c = TryCast(itens.Add(OL.OlItemType.olContactItem), OL.ContactItem)
                If c Is Nothing Then
                    Return OperationResult(Of ContactInfo).Fail(ErrorKind.Unexpected, "Items.Add")
                End If

                c.FullName = rascunho.Nome
                c.Email1Address = If(rascunho.Email, "")
                c.CompanyName = If(rascunho.Empresa, "")

                ' NADA MAIS. Sem Body, sem nota, sem endereco: o que entra e o
                ' que a mensagem ja dizia. E nao ha Forward em lugar nenhum
                ' deste modulo, que e o unico caminho de envio deste objeto.
                If marcar IsNot Nothing Then marcar()
                c.Save()

                ' O SAVE ACONTECEU. Se a releitura falhar daqui para a frente, o
                ' contato EXISTE e a identidade nova se perdeu -- e isso e
                ' ambiguidade, nao sucesso. Devolver Ok(Nothing) dava ao chamador
                ' um resultado formalmente bem-sucedido sem contato e sem EntryID,
                ' e sem como saber que faltava alguma coisa. Achado por revisao
                ' externa em 01/09/2026.
                Dim descrito = Traduzir(c)
                If descrito Is Nothing OrElse descrito.Key Is Nothing OrElse
                   descrito.Key.IsEmpty Then
                    Return OperationResult(Of ContactInfo).Fail(ErrorKind.Ambiguous,
                        "o contato foi criado e a identidade nova nao pode ser lida")
                End If

                Return OperationResult(Of ContactInfo).Ok(descrito)
            Catch ex As COMException
                Return OperationResult(Of ContactInfo).Fail(
                    OutlookFailurePolicy.ClassifyFailure(
                        ex.HResult, isMutation:=True, mutationAttemptStarted:=True),
                    "falha ao criar contato: " & ex.GetType().Name)
            Finally
                ComHelpers.Release(c)
                ComHelpers.Release(itens)
                ComHelpers.Release(destino)
            End Try
        End Function

        ' ==============================================================

        ''' <summary>
        ''' O que impede um rascunho de virar contato. <c>Nothing</c> quando não
        ''' há impedimento. Separada para ter teste sem COM.
        '''
        ''' Exige <b>nome</b>: um contato sem nome vira uma ficha em branco no
        ''' catálogo, e quem a encontrar depois não tem como saber de quem é.
        ''' Não exige e-mail — contato de telefone é contato.
        ''' </summary>
        Friend Function RecusarRascunho(rascunho As ContactDraft) As String
            If rascunho Is Nothing Then Return "rascunho nulo"
            If String.IsNullOrWhiteSpace(rascunho.Nome) Then
                Return "um contato sem nome vira uma ficha anônima no catálogo"
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Relê o contato. <b>Depois</b> do <c>Save</c>, sempre: o
        ''' <c>EntryID</c> pode ter mudado.
        '''
        ''' Cada campo é lido por <see cref="Texto"/>, que devolve
        ''' <c>Nothing</c> quando a leitura falha. Um <c>Catch</c> que devolvesse
        ''' cadeia vazia diria "este contato não tem empresa" sobre um campo que
        ''' ninguém conseguiu ler.
        ''' </summary>
        Private Function Traduzir(c As OL.ContactItem) As ContactInfo
            Try
                Return New ContactInfo With {
                    .Key = New ItemKey(If(Texto(Function() c.EntryID), ""), StoreDe(c)),
                    .Nome = Texto(Function() c.FullName),
                    .Email = Texto(Function() c.Email1Address),
                    .Empresa = Texto(Function() c.CompanyName)
                }
            Catch
                ' Item que nao se deixa ler nao derruba a lista: vira recusa
                ' contada, como no calendario e nas tarefas.
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' O texto, ou <c>Nothing</c> quando não deu para ler.
        '''
        ''' <b>Não devolve cadeia vazia no erro.</b> É a diferença entre "o
        ''' campo está vazio" e "não consegui ler o campo", e este projeto já
        ''' pagou cinco vezes por colapsá-la.
        ''' </summary>
        ''' <b>Friend, e nao Private.</b> O controle negativo desta funcao
        ''' mostrou que ela nao tinha teste nenhum: trocar o Nothing por ""
        ''' nao derrubava nada, porque os testes de leitura montam o
        ''' <c>ContactInfo</c> a mao e nunca passam por aqui. Guarda sem
        ''' teste e guarda que some na proxima refatoracao.
        Friend Function Texto(getter As Func(Of String)) As String
            Try
                Return getter()
            Catch
                Return Nothing
            End Try
        End Function

        Private Function StoreDe(c As OL.ContactItem) As String
            Dim pai As OL.Folder = Nothing
            Try
                pai = TryCast(c.Parent, OL.Folder)
                Return If(pai?.StoreID, "")
            Catch
                Return ""
            Finally
                ComHelpers.Release(pai)
            End Try
        End Function

    End Module

End Namespace

Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' Captura a mensagem inteira <b>numa leitura só</b>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE NUMA LEITURA SÓ</b>
    '''
    ''' Assunto, remetente, destinatários, corpo e <c>PR_CHANGE_KEY</c> obtidos
    ''' em cinco chamadas separadas podem observar cinco estados diferentes de
    ''' uma mensagem que mudou no meio — e a <c>ChangeKey</c> serve justamente
    ''' para prender o corpo à versão que o portão classificou. Se ela vier de
    ''' outra passada, não prende nada.
    '''
    ''' Isto não torna a leitura <b>atômica</b>: o OOM não oferece isso, e a
    ''' §29.2 é a resposta a essa falta. O que se ganha é a janela mais estreita
    ''' possível, e um lugar só onde ela existe.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O CORPO É TEXTO</b>
    '''
    ''' <c>Body</c>, não <c>HTMLBody</c>. O <see cref="ContentPipeline"/> sabe
    ''' converter HTML, e vai saber; aqui o texto simples é o que o Outlook já
    ''' entrega pronto, e pedir HTML seria trazer marcação para depois tirar.
    '''
    ''' <b>Anexo não é lido</b>, e nem contado como conteúdo: a fase não os
    ''' trata, e o portão nega mensagem que tem.
    ''' </summary>
    Friend Module MessageSnapshots

        Private Const DaslChangeKey As String =
            "http://schemas.microsoft.com/mapi/proptag/0x65E20102"

        ''' <summary>
        ''' <b>N corpos numa visita só.</b> É a borda em lote.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O QUE ELA ECONOMIZA, E O QUE NÃO MUDA</b>
        '''
        ''' Chamar <see cref="Read"/> vinte vezes funciona e já funcionava — cada
        ''' chamada atravessa o despachante do broker, entra na STA, pega
        ''' <c>Application</c> e <c>NameSpace</c>, e sai. Vinte idas para vinte
        ''' mensagens que vão para o mesmo lote.
        '''
        ''' Aqui o laço mora <b>dentro</b> de uma ida só. O que <i>não</i> muda é o
        ''' tratamento de cada item: é o mesmo <see cref="Read"/>, com o mesmo
        ''' <c>Finally</c> e a mesma ordem de liberação. Reescrever a leitura aqui
        ''' para "aproveitar" a visita seria duplicar a única rotina desta base que
        ''' já custou quatro violações da R7.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>ITEM QUE FALHA NÃO DERRUBA O LOTE — E NÃO SOME</b>
        '''
        ''' A saída tem <b>uma posição por entrada</b>, na mesma ordem, e o item
        ''' que não deu para ler entra como <c>Nothing</c>. As duas metades
        ''' importam:
        '''
        ''' <list type="bullet">
        ''' <item>falhar o lote inteiro por uma mensagem faria uma pasta com um
        ''' item corrompido nunca ser classificada;</item>
        ''' <item>encolher a lista faria o chamador casar ficha com mensagem
        ''' <b>pelo índice errado</b> — e a ficha é justamente o que diz de quem o
        ''' modelo está falando.</item>
        ''' </list>
        '''
        ''' Por isso o alinhamento posicional é contrato, e não conveniência.
        ''' </summary>
        Public Function ReadMany(ns As OL.NameSpace, items As IReadOnlyList(Of ItemKey)) _
                                 As OperationResult(Of IReadOnlyList(Of MessageSnapshot))
            Dim saida As New List(Of MessageSnapshot)()
            If items Is Nothing Then
                Return OperationResult(Of IReadOnlyList(Of MessageSnapshot)).Ok(saida)
            End If

            For Each item In items
                If item Is Nothing OrElse item.IsEmpty Then
                    ' Posicao preservada: ver o paragrafo do alinhamento.
                    saida.Add(Nothing)
                    Continue For
                End If

                Dim um = Read(ns, item)
                saida.Add(If(um.Succeeded, um.Value, Nothing))
            Next

            Return OperationResult(Of IReadOnlyList(Of MessageSnapshot)).Ok(saida)
        End Function

        Public Function Read(ns As OL.NameSpace, item As ItemKey) _
                             As OperationResult(Of MessageSnapshot)
            ' R7: o objeto e adquirido em UMA variavel, e e ela que o Finally
            ' libera. Escrever TryCast(ns.GetItemFromID(...), MailItem) direto
            ' perdia a referencia quando o item existia e NAO era MailItem — o
            ' RCW ficava sem dono e ninguem o liberava.
            Dim obj As Object = Nothing
            Dim mail As OL.MailItem = Nothing
            Try
                Try
                    obj = ns.GetItemFromID(item.EntryId, item.StoreId)
                Catch ex As COMException
                    Return OperationResult(Of MessageSnapshot).Fail(ErrorKind.NotFound, "item")
                End Try
                If obj Is Nothing Then
                    Return OperationResult(Of MessageSnapshot).Fail(ErrorKind.NotFound, "item")
                End If

                mail = TryCast(obj, OL.MailItem)
                If mail Is Nothing Then
                    Return OperationResult(Of MessageSnapshot).Fail(ErrorKind.NotFound, "item")
                End If

                Dim chave = ChangeKey(mail)
                If String.IsNullOrEmpty(chave) Then
                    ' Sem versao nao da para prender o corpo a leitura que o
                    ' classificou. Falhar aqui e melhor que entregar um
                    ' snapshot que o pipeline vai recusar de qualquer jeito.
                    Return OperationResult(Of MessageSnapshot).Fail(
                        ErrorKind.Unexpected, "sem PR_CHANGE_KEY")
                End If

                ' UMA visita so aos anexos, e o resultado guardado: contar
                ' embutidas e percorrer a colecao, e percorrer duas vezes
                ' pagaria o COM duas vezes pela mesma resposta.
                ' NAO se chama `anexado`: eclipsaria a funcao Anexado -- VB e
                ' case-insensitive, e o erro sai como "tipo nao pode ser
                ' inferido", que nao diz nada sobre o problema. 14a vez
                ' nesta base; a tabela do CLAUDE.md tem a lista.
                Dim osAnexos = Anexado(mail)

                Dim corpo = Texto(Function() mail.Body)
                Dim completo = Not String.IsNullOrEmpty(corpo)

                Return OperationResult(Of MessageSnapshot).Ok(
                    New MessageSnapshot(item, chave,
                                        Texto(Function() mail.Subject),
                                        Texto(Function() mail.SenderName),
                                        Destinatarios(mail),
                                        corpo, ehHtml:=False, corpoCompleto:=completo,
                                        temAnexo:=osAnexos.Real,
                                        embutidas:=osAnexos.Embutidas))
            Finally
                ' `mail` e o MESMO objeto que `obj`: liberar os dois seria
                ' liberar a mesma referencia duas vezes. So `obj` e liberado.
                mail = Nothing
                ComHelpers.Release(obj)
            End Try
        End Function

        ''' <summary>
        ''' Tem anexo — <c>Nothing</c> se não deu para saber.
        '''
        ''' R7: a coleção <c>Attachments</c> é um objeto COM próprio, e
        ''' <c>mail.Attachments.Count</c> deixaria o RCW dela sem dono. Ela é
        ''' adquirida, lida e liberada.
        '''
        ''' Falhar aqui <b>não</b> vira zero: <c>Nothing</c> sobe até o pipeline,
        ''' que recusa. Devolver zero seria transformar "não consegui contar" em
        ''' "não tem" — falha aberta, exatamente a que o 3.0 já custou uma vez.
        ''' </summary>
        ''' <summary>
        ''' Separa anexo de verdade de imagem embutida. A regra mora no
        ''' <see cref="ClassificacaoDeAnexo"/>, e não aqui, porque o portão
        ''' precisa dela também — duas cópias divergem, e quando divergirem o
        ''' portão autoriza por um critério e a captura usa outro.
        ''' </summary>
        Private Function Anexado(mail As OL.MailItem) As (Real As Boolean?, Embutidas As Integer?)
            Dim anexos As OL.Attachments = Nothing
            Try
                anexos = mail.Attachments
                Return ClassificacaoDeAnexo.Contar(anexos)
            Catch
                Return (Nothing, Nothing)
            Finally
                ComHelpers.Release(anexos)
            End Try
        End Function

        ''' <summary>
        ''' Os destinatários, <b>por nome</b>.
        '''
        ''' Nome e não endereço: a guarda do Object Model pode pedir confirmação
        ''' para ler endereço, e um diálogo modal no meio de um resumo é o pior
        ''' momento possível. O nome basta para o modelo entender quem está na
        ''' conversa.
        '''
        ''' Falhar aqui não derruba a captura: lista vazia é menos informação, e
        ''' menos informação é melhor que nenhuma mensagem.
        ''' </summary>
        Private Function Destinatarios(mail As OL.MailItem) As IReadOnlyList(Of String)
            Dim saida As New List(Of String)()
            Dim lista As OL.Recipients = Nothing
            Try
                lista = mail.Recipients
                For i = 1 To lista.Count
                    Dim r As OL.Recipient = Nothing
                    Try
                        r = lista.Item(i)
                        saida.Add(r.Name)
                    Catch
                    Finally
                        ComHelpers.Release(r)
                    End Try
                Next
            Catch
            Finally
                ComHelpers.Release(lista)
            End Try
            Return saida
        End Function

        ''' <summary>
        ''' <c>PR_CHANGE_KEY</c> em hex. R7: o <c>PropertyAccessor</c> é
        ''' adquirido, usado e liberado — não encadeado.
        ''' </summary>
        Private Function ChangeKey(mail As OL.MailItem) As String
            Dim acessor As OL.PropertyAccessor = Nothing
            Try
                acessor = mail.PropertyAccessor
                If acessor Is Nothing Then Return Nothing
                Dim bytes = TryCast(acessor.GetProperty(DaslChangeKey), Byte())
                If bytes Is Nothing OrElse bytes.Length = 0 Then Return Nothing
                Return Convert.ToHexString(bytes)
            Catch
                Return Nothing
            Finally
                ComHelpers.Release(acessor)
            End Try
        End Function

        Private Function Texto(f As Func(Of String)) As String
            Try
                Return If(f(), "")
            Catch
                Return ""
            End Try
        End Function

    End Module

End Namespace

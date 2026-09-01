Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' <b>Em que pasta cada mensagem está, perguntado ao Outlook.</b>
    '''
    ''' Existe para o portão da divulgação poder conferir a pasta contra o
    ''' provedor em vez de contra a afirmação de quem pediu — ver
    ''' <see cref="PastaDoItem"/>, que tem o motivo por extenso.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>NÃO LÊ CORPO, E ISSO É A ORDEM DAS COISAS</b>
    '''
    ''' O portão classifica <b>antes</b> de qualquer corpo ser lido — é o que
    ''' impede ir ao conteúdo sem autorização. Então esta leitura tem de ser
    ''' separada da do corpo, mesmo custando uma visita a mais ao COM: juntá-las
    ''' seria ler o corpo para descobrir se podia lê-lo.
    ''' </summary>
    Friend Module PastasDosItens

        ''' <summary>
        ''' Uma entrada por item pedido. Item que não deu para abrir entra com
        ''' pasta <b>vazia</b>, que o portão trata como negativa — e não some da
        ''' lista, porque sumir faria a contagem do portão não bater com a do
        ''' pedido.
        ''' </summary>
        Public Function Ler(ns As OL.NameSpace, items As IReadOnlyList(Of ItemKey)) _
                            As OperationResult(Of IReadOnlyList(Of PastaDoItem))

            Dim saida As New List(Of PastaDoItem)()
            If items Is Nothing Then
                Return OperationResult(Of IReadOnlyList(Of PastaDoItem)).Ok(saida)
            End If

            For Each item In items
                saida.Add(New PastaDoItem(item, LerUma(ns, item)))
            Next

            Return OperationResult(Of IReadOnlyList(Of PastaDoItem)).Ok(saida)
        End Function

        ''' <summary>
        ''' R7: <c>Parent</c> devolve OUTRO objeto COM, e ele é adquirido numa
        ''' variável própria e liberado no <c>Finally</c>, antes do item — ordem
        ''' inversa à da aquisição.
        '''
        ''' <c>EntryID</c> e <c>StoreID</c> da pasta são escalares e não criam RCW
        ''' próprio; a cadeia proibida seria <c>item.Parent.EntryID</c>, que
        ''' deixaria a pasta sem dono.
        ''' </summary>
        Private Function LerUma(ns As OL.NameSpace, item As ItemKey) As FolderKey
            If item Is Nothing OrElse item.IsEmpty Then Return New FolderKey("", "")

            Dim obj As Object = Nothing
            Dim pai As OL.MAPIFolder = Nothing
            Try
                Try
                    obj = ns.GetItemFromID(item.EntryId, item.StoreId)
                Catch ex As COMException
                    Return New FolderKey("", "")
                End Try
                If obj Is Nothing Then Return New FolderKey("", "")

                Dim mail = TryCast(obj, OL.MailItem)
                If mail Is Nothing Then Return New FolderKey("", "")

                Try
                    pai = TryCast(mail.Parent, OL.MAPIFolder)
                Catch ex As COMException
                    Return New FolderKey("", "")
                End Try
                If pai Is Nothing Then Return New FolderKey("", "")

                Dim entrada = Texto(Function() pai.EntryID)
                Dim loja = Texto(Function() pai.StoreID)
                If entrada.Length = 0 Then Return New FolderKey("", "")

                Return New FolderKey(entrada, loja)
            Catch
                ' Qualquer outra coisa também é "não sei", e "não sei" nega.
                Return New FolderKey("", "")
            Finally
                ComHelpers.Release(pai)
                ComHelpers.Release(obj)
            End Try
        End Function

        Private Function Texto(ler As Func(Of String)) As String
            Try
                Return If(ler(), "")
            Catch
                Return ""
            End Try
        End Function

    End Module

End Namespace

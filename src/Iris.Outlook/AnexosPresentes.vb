Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' <b>Só isto: cada item tem anexo?</b>
    '''
    ''' Separado da leitura de rótulo de propósito. A leitura de rótulo foi
    ''' medida contra esta caixa no 3.0 e tem controle negativo próprio; enfiar
    ''' mais uma propriedade dentro dela para poupar uma visita ao COM colocaria
    ''' em risco a única parte da fase que foi medida de verdade.
    '''
    ''' O custo é uma visita a mais, e uma janela em que o anexo muda entre a
    ''' classificação e a leitura do corpo. Essa janela é <b>fechada em outro
    ''' lugar</b>: <see cref="MessageSnapshots"/> lê o anexo na mesma visita que
    ''' o corpo, e o pipeline recusa. Aqui a leitura serve para o usuário
    ''' receber o motivo certo — "mensagem com anexo" — em vez de uma recusa do
    ''' cofre por contagem que não bate.
    ''' </summary>
    Friend Module AnexosPresentes

        Public Function Ler(ns As OL.NameSpace, items As IReadOnlyList(Of ItemKey)) _
                            As OperationResult(Of IReadOnlyList(Of AttachmentPresence))
            Dim saida As New List(Of AttachmentPresence)()
            For Each item In items
                saida.Add(New AttachmentPresence(item, LerUm(ns, item)))
            Next
            Return OperationResult(Of IReadOnlyList(Of AttachmentPresence)).Ok(saida)
        End Function

        ''' <summary>
        ''' R7 duas vezes: o item e a coleção são adquiridos em variáveis
        ''' próprias e liberados em ordem inversa. <c>obj.Attachments.Count</c>
        ''' deixaria dois RCWs sem dono numa linha só.
        ''' </summary>
        Private Function LerUm(ns As OL.NameSpace, item As ItemKey) As Boolean?
            Dim obj As Object = Nothing
            Dim mail As OL.MailItem = Nothing
            Dim anexos As OL.Attachments = Nothing
            Try
                Try
                    obj = ns.GetItemFromID(item.EntryId, item.StoreId)
                Catch ex As COMException
                    Return Nothing
                End Try
                If obj Is Nothing Then Return Nothing

                mail = TryCast(obj, OL.MailItem)
                If mail Is Nothing Then Return Nothing

                anexos = mail.Attachments
                If anexos Is Nothing Then Return Nothing
                Return anexos.Count > 0
            Catch
                Return Nothing
            Finally
                ComHelpers.Release(anexos)
                ' `mail` e o mesmo objeto que `obj`.
                mail = Nothing
                ComHelpers.Release(obj)
            End Try
        End Function

    End Module

End Namespace

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
                Dim lido = LerUm(ns, item)
                saida.Add(New AttachmentPresence(item, lido.Real, lido.Embutidas))
            Next
            Return OperationResult(Of IReadOnlyList(Of AttachmentPresence)).Ok(saida)
        End Function

        ''' <summary>
        ''' R7 duas vezes: o item e a coleção são adquiridos em variáveis
        ''' próprias e liberados em ordem inversa. <c>obj.Attachments.Count</c>
        ''' deixaria dois RCWs sem dono numa linha só.
        ''' </summary>
        ''' <summary>
        ''' Anexo de verdade e imagens embutidas, pela mesma regra que a captura
        ''' do corpo usa. Ver <see cref="ClassificacaoDeAnexo"/>.
        ''' </summary>
        Private Function LerUm(ns As OL.NameSpace, item As ItemKey) _
                               As (Real As Boolean?, Embutidas As Integer?)
            Dim obj As Object = Nothing
            Dim mail As OL.MailItem = Nothing
            Dim anexos As OL.Attachments = Nothing
            Try
                Try
                    obj = ns.GetItemFromID(item.EntryId, item.StoreId)
                Catch ex As COMException
                    Return (Nothing, Nothing)
                End Try
                If obj Is Nothing Then Return (Nothing, Nothing)

                mail = TryCast(obj, OL.MailItem)
                If mail Is Nothing Then Return (Nothing, Nothing)

                anexos = mail.Attachments
                Return ClassificacaoDeAnexo.Contar(anexos)
            Catch
                Return (Nothing, Nothing)
            Finally
                ComHelpers.Release(anexos)
                ' `mail` e o mesmo objeto que `obj`.
                mail = Nothing
                ComHelpers.Release(obj)
            End Try
        End Function

    End Module

End Namespace

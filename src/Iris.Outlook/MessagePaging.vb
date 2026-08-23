Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' Leitura paginada de mensagens.
    '''
    ''' Tudo aqui roda DENTRO da thread do broker. As regras que a Fase 0
    ''' impôs, e que este arquivo existe para respeitar:
    '''
    '''   • <c>Sort</c> é feito pelo OUTLOOK, nunca em laço nosso. Medido:
    '''     ordenar 770 itens custou 3 ms lá dentro; percorrer os mesmos 770
    '''     custou 12,8 segundos aqui.
    '''   • O corpo NÃO é lido durante a listagem. Uma leitura de corpo no
    '''     meio da paginação bloquearia a fila única da STA (F1-F).
    '''   • Toda referência COM adquirida é liberada. Encadear
    '''     <c>item.Attachments.Count</c> é o erro que já apareceu quatro
    '''     vezes neste projeto.
    ''' </summary>
    Friend Module MessagePaging

        ''' <summary>
        ''' Campo de ordenação no vocabulário do Outlook. Colchetes são a
        ''' sintaxe que o OOM espera.
        ''' </summary>
        Private Function CampoDeOrdenacao(sort As MessageSort) As (Campo As String, Descendente As Boolean)
            Select Case sort
                Case MessageSort.ReceivedAsc : Return ("[ReceivedTime]", False)
                Case MessageSort.SubjectAsc : Return ("[Subject]", False)
                Case MessageSort.SenderAsc : Return ("[SenderName]", False)
                Case Else : Return ("[ReceivedTime]", True)
            End Select
        End Function

        Public Function ReadPage(ns As OL.NameSpace, query As MessageQuery,
                                 offset As Integer, count As Integer) As OperationResult(Of MessagePage)

            If offset < 0 OrElse count <= 0 Then
                Return OperationResult(Of MessagePage).Fail(
                    ErrorKind.Unexpected, "offset ou count invalido")
            End If

            Dim folder As OL.MAPIFolder = Nothing
            Try
                Try
                    folder = TryCast(ns.GetFolderFromID(query.Folder.EntryId, query.Folder.StoreId),
                                     OL.MAPIFolder)
                Catch ex As COMException
                    Return OperationResult(Of MessagePage).Fail(ErrorKind.NotFound, "pasta")
                End Try

                If folder Is Nothing Then
                    Return OperationResult(Of MessagePage).Fail(ErrorKind.NotFound, "pasta")
                End If

                Dim items As OL.Items = Nothing
                Try
                    items = folder.Items

                    Dim ordenacao = CampoDeOrdenacao(query.Sort)
                    Try
                        items.Sort(ordenacao.Campo, ordenacao.Descendente)
                    Catch
                        ' Pasta que não aceita este campo: segue na ordem
                        ' natural em vez de falhar a página inteira.
                    End Try

                    Dim total = items.Count
                    Dim pagina As New MessagePage With {
                        .Generation = query.Generation,
                        .Offset = offset,
                        .TotalAtRead = total
                    }

                    ' Items é 1-based no OOM.
                    Dim primeiro = offset + 1
                    Dim ultimo = Math.Min(offset + count, total)

                    For i = primeiro To ultimo
                        Dim bruto As Object = Nothing
                        Try
                            bruto = items.Item(i)
                            Dim mail = TryCast(bruto, OL.MailItem)
                            If mail Is Nothing Then
                                ' Uma coleção Items não contém apenas
                                ' MailItem (ESCOPO.md seção 5). Convites e
                                ' relatórios de entrega convivem ali.
                                Continue For
                            End If
                            pagina.Items.Add(Summarize(mail, query.Folder.StoreId))
                        Catch ex As COMException
                            ' Item corrompido ou não baixado não pode
                            ' derrubar a página inteira.
                            Continue For
                        Finally
                            ComHelpers.Release(bruto)
                        End Try
                    Next

                    pagina.HasMore = ultimo < total
                    Return OperationResult(Of MessagePage).Ok(pagina)
                Finally
                    ComHelpers.Release(items)
                End Try
            Finally
                ComHelpers.Release(folder)
            End Try
        End Function

        ''' <summary>
        ''' Resumo para a lista. SEM corpo — ver F1-F.
        ''' </summary>
        Private Function Summarize(mail As OL.MailItem, storeId As String) As MailSummary
            Dim anexos = ContarAnexos(mail)

            Dim estado = ContentState.MetadataOnly
            Try
                If mail.DownloadState = OL.OlDownloadState.olFullItem Then
                    ' DownloadState é a PROMESSA do Outlook, não a prova de
                    ' que o corpo pode ser lido — a Fase 0 mediu essa
                    ' diferença. Aqui só o metadado é afirmado.
                    estado = ContentState.BodyAvailable
                End If
            Catch
                estado = ContentState.TransientError
            End Try

            Return New MailSummary With {
                .Key = New ItemKey(Texto(Function() mail.EntryID), storeId),
                .Subject = Texto(Function() mail.Subject),
                .SenderName = Texto(Function() mail.SenderName),
                .ReceivedTime = Data(Function() mail.ReceivedTime),
                .SizeBytes = Numero(Function() mail.Size),
                .HasAttachments = anexos > 0,
                .IsUnread = Booleano(Function() mail.UnRead),
                .IsProtected = Numero(Function() CInt(mail.Permission)) <> 0,
                .MessageClass = Texto(Function() mail.MessageClass),
                .Content = estado
            }
        End Function

        ''' <summary>
        ''' mail.Attachments é objeto COM próprio. Escrever
        ''' mail.Attachments.Count cria um RCW intermediário sem dono.
        ''' </summary>
        Private Function ContarAnexos(mail As OL.MailItem) As Integer
            Dim anexos As OL.Attachments = Nothing
            Try
                anexos = mail.Attachments
                Return anexos.Count
            Catch
                Return 0
            Finally
                ComHelpers.Release(anexos)
            End Try
        End Function

        ' Propriedades COM lançam por item corrompido, offline ou baixado
        ' parcialmente. Um item ruim não pode derrubar a listagem.

        Private Function Texto(getter As Func(Of String)) As String
            Try : Return If(getter(), "") : Catch : Return "" : End Try
        End Function

        Private Function Numero(getter As Func(Of Integer)) As Integer
            Try : Return getter() : Catch : Return 0 : End Try
        End Function

        Private Function Booleano(getter As Func(Of Boolean)) As Boolean
            Try : Return getter() : Catch : Return False : End Try
        End Function

        Private Function Data(getter As Func(Of DateTime)) As DateTimeOffset?
            Try
                Dim valor = getter()
                ' O Outlook devolve hora local sem Kind. Assumir o fuso
                ' local aqui é o que impede a mensagem aparecer com hora
                ' errada quando a Fase 2 persistir isto.
                Return New DateTimeOffset(DateTime.SpecifyKind(valor, DateTimeKind.Local))
            Catch
                Return Nothing
            End Try
        End Function

    End Module

End Namespace

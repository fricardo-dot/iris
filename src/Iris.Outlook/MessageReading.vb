Imports System.Collections.Generic
Imports System.IO
Imports System.Runtime.InteropServices
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' Leitura do conteúdo de uma mensagem.
    '''
    ''' Cabeçalho, corpo e anexos vêm numa chamada só de propósito: obtidos
    ''' em três chamadas separadas, cada uma poderia observar um estado
    ''' diferente de uma mensagem que mudou no meio.
    '''
    ''' <see cref="ContentState.BodyAvailable"/> só é afirmado DEPOIS de o
    ''' corpo ser efetivamente lido. <c>DownloadState</c> é promessa do
    ''' Outlook, e a Fase 0 mediu a diferença entre promessa e fato.
    ''' </summary>
    Friend Module MessageReading

        ''' <summary>
        ''' Teto de corpo entregue à UI. Um e-mail com megabytes de HTML
        ''' citado existe, e jogá-lo inteiro na interface trava a janela
        ''' antes de qualquer renderizador entrar em cena.
        ''' </summary>
        Private Const MaxBodyChars As Integer = 512 * 1024

        Public Function ReadDetail(ns As OL.NameSpace, item As ItemKey) As OperationResult(Of MessageDetail)
            Dim mail As OL.MailItem = Nothing
            Try
                Try
                    mail = TryCast(ns.GetItemFromID(item.EntryId, item.StoreId), OL.MailItem)
                Catch ex As COMException When EhNaoEncontrado(ex.HResult)
                    ' O item pode ter sido movido ou excluído entre a lista e
                    ' o detalhe. Isso é resultado NORMAL, não falha (F1-H).
                    Return OperationResult(Of MessageDetail).Fail(ErrorKind.NotFound, "item")
                End Try

                If mail Is Nothing Then
                    Return OperationResult(Of MessageDetail).Fail(ErrorKind.NotFound, "item")
                End If

                Dim detalhe As New MessageDetail With {
                    .Key = item,
                    .Subject = Texto(Function() mail.Subject),
                    .SenderName = Texto(Function() mail.SenderName),
                    .SenderAddress = Texto(Function() mail.SenderEmailAddress),
                    .ReceivedTime = Data(Function() mail.ReceivedTime),
                    .IsProtected = Numero(Function() CInt(mail.Permission)) <> 0
                }

                LerDestinatarios(mail, detalhe.Recipients)
                LerAnexos(mail, item, detalhe.Attachments)
                LerCorpo(mail, detalhe)

                Return OperationResult(Of MessageDetail).Ok(detalhe)
            Finally
                ComHelpers.Release(mail)
            End Try
        End Function

        ''' <summary>
        ''' Só depois de ler é que se sabe se dava para ler.
        ''' </summary>
        Private Sub LerCorpo(mail As OL.MailItem, detalhe As MessageDetail)
            ' Mensagem protegida por IRM não tem corpo entregue por aqui, e
            ' insistir só produz erro. R11: conteúdo protegido também não vai
            ' para log nem para IA.
            If detalhe.IsProtected Then
                detalhe.Content = ContentState.MetadataOnly
                detalhe.BodyError = ErrorKind.Denied
                Return
            End If

            Try
                Dim corpo = mail.Body
                If corpo Is Nothing Then corpo = ""

                If corpo.Length > MaxBodyChars Then
                    corpo = corpo.Substring(0, MaxBodyChars) &
                            Environment.NewLine & Environment.NewLine &
                            "[...] mensagem truncada pelo Iris."
                End If

                detalhe.TextBody = corpo
                detalhe.Format = BodyFormat.PlainText
                detalhe.Content = If(detalhe.Attachments.Count > 0,
                                     ContentState.AttachmentsAvailable,
                                     ContentState.BodyAvailable)
            Catch ex As COMException When EhConteudoIndisponivel(ex.HResult)
                ' SO os HRESULTs que realmente significam conteudo ausente.
                ' Antes, qualquer COMException virava "nao baixado" — e
                ' engolir a excecao impedia o classificador central do broker
                ' de ver Outlook ocupado, RPC desconectado ou acesso negado,
                ' que sao coisas diferentes e levam a UI a decisoes opostas.
                detalhe.Content = ContentState.TransientError
                detalhe.BodyError = ErrorKind.NotDownloaded
            End Try
        End Sub

        Private Sub LerDestinatarios(mail As OL.MailItem, destino As List(Of RecipientInfo))
            Dim recipients As OL.Recipients = Nothing
            Try
                recipients = mail.Recipients
                For i = 1 To recipients.Count
                    Dim r As OL.Recipient = Nothing
                    Try
                        r = recipients.Item(i)
                        destino.Add(New RecipientInfo With {
                            .DisplayName = Texto(Function() r.Name),
                            .Address = Texto(Function() r.Address),
                            .Kind = TipoDeDestinatario(r),
                            .Resolved = Booleano(Function() r.Resolved)
                        })
                    Catch
                        ' Um destinatário problemático não derruba a leitura.
                    Finally
                        ComHelpers.Release(r)
                    End Try
                Next
            Catch
                ' Ler destinatários é uma das operações que a guarda do
                ' Object Model protege. Falhar aqui não pode custar o corpo.
            Finally
                ComHelpers.Release(recipients)
            End Try
        End Sub

        Private Function TipoDeDestinatario(r As OL.Recipient) As RecipientKind
            Try
                Select Case r.Type
                    Case CInt(OL.OlMailRecipientType.olTo) : Return RecipientKind.To
                    Case CInt(OL.OlMailRecipientType.olCC) : Return RecipientKind.Cc
                    Case CInt(OL.OlMailRecipientType.olBCC) : Return RecipientKind.Bcc
                    Case Else : Return RecipientKind.Unknown
                End Select
            Catch
                Return RecipientKind.Unknown
            End Try
        End Function

        Private Sub LerAnexos(mail As OL.MailItem, dono As ItemKey, destino As List(Of AttachmentInfo))
            Dim anexos As OL.Attachments = Nothing
            Try
                anexos = mail.Attachments
                For i = 1 To anexos.Count
                    Dim a As OL.Attachment = Nothing
                    Try
                        a = anexos.Item(i)
                        Dim nome = Texto(Function() a.FileName)
                        Dim tamanho = Numero(Function() a.Size)

                        destino.Add(New AttachmentInfo With {
                            .Key = New AttachmentKey(dono, i, nome, tamanho),
                            .FileName = nome,
                            .SizeBytes = tamanho,
                            .AttachmentType = Texto(Function() a.Type.ToString()),
                            .ContentId = "",
                            .IsInline = False
                        })
                    Catch
                    Finally
                        ComHelpers.Release(a)
                    End Try
                Next
            Catch
            Finally
                ComHelpers.Release(anexos)
            End Try
        End Sub

        ''' <summary>
        ''' Marca como lida. SÓ se estiver não lida.
        '''
        ''' A verificação não é economia: marcar o que já está lido gera um
        ''' ItemChange à toa, que suja a pasta, que recarrega a lista, que
        ''' muda a seleção, que marca de novo — o laço do F1-G.
        ''' </summary>
        Public Function SetReadState(ns As OL.NameSpace, item As ItemKey,
                                     isRead As Boolean) As OperationResult(Of Boolean)
            Dim mail As OL.MailItem = Nothing
            Try
                Try
                    mail = TryCast(ns.GetItemFromID(item.EntryId, item.StoreId), OL.MailItem)
                Catch ex As COMException When EhNaoEncontrado(ex.HResult)
                    Return OperationResult(Of Boolean).Fail(ErrorKind.NotFound, "item")
                End Try

                If mail Is Nothing Then
                    Return OperationResult(Of Boolean).Fail(ErrorKind.NotFound, "item")
                End If

                If mail.UnRead <> isRead Then
                    ' Já está no estado desejado. Nada a fazer, e nada de
                    ' evento espúrio.
                    Return OperationResult(Of Boolean).Ok(False)
                End If

                mail.UnRead = Not isRead
                mail.Save()
                Return OperationResult(Of Boolean).Ok(True)
            Finally
                ComHelpers.Release(mail)
            End Try
        End Function

        ''' <summary>
        ''' Salva um anexo em disco.
        '''
        ''' Escrita em disco, não leitura: vai por MutateAsync, sem retry.
        ''' O índice é validado contra nome e tamanho antes de gravar, porque
        ''' a coleção pode ter mudado entre listar e salvar.
        ''' </summary>
        Public Function SaveAttachment(ns As OL.NameSpace, key As AttachmentKey,
                                       destino As String, overwrite As Boolean) _
            As OperationResult(Of String)

            If String.IsNullOrWhiteSpace(destino) Then
                Return OperationResult(Of String).Fail(ErrorKind.Unexpected, "destino vazio")
            End If

            If File.Exists(destino) AndAlso Not overwrite Then
                Return OperationResult(Of String).Fail(ErrorKind.Denied, "arquivo ja existe")
            End If

            Dim mail As OL.MailItem = Nothing
            Try
                Try
                    mail = TryCast(ns.GetItemFromID(key.Owner.EntryId, key.Owner.StoreId), OL.MailItem)
                Catch ex As COMException When EhNaoEncontrado(ex.HResult)
                    Return OperationResult(Of String).Fail(ErrorKind.NotFound, "item")
                End Try

                If mail Is Nothing Then
                    Return OperationResult(Of String).Fail(ErrorKind.NotFound, "item")
                End If

                Dim anexos As OL.Attachments = Nothing
                Try
                    anexos = mail.Attachments
                    If key.Index < 1 OrElse key.Index > anexos.Count Then
                        Return OperationResult(Of String).Fail(ErrorKind.NotFound, "anexo")
                    End If

                    Dim a As OL.Attachment = Nothing
                    Try
                        a = anexos.Item(key.Index)

                        ' O indice sozinho e instavel. Confere nome E TAMANHO
                        ' antes de gravar: salvar o anexo errado com o nome
                        ' certo seria pior que falhar. A versao anterior
                        ' prometia os dois no comentario e conferia so o nome.
                        If Not String.Equals(Texto(Function() a.FileName), key.FileName,
                                             StringComparison.Ordinal) OrElse
                           Numero(Function() a.Size) <> key.SizeBytes Then
                            Return OperationResult(Of String).Fail(ErrorKind.Stale, "anexo mudou")
                        End If

                        Return GravarComTemporario(a, destino, overwrite)
                    Finally
                        ComHelpers.Release(a)
                    End Try
                Finally
                    ComHelpers.Release(anexos)
                End Try
            Finally
                ComHelpers.Release(mail)
            End Try
        End Function

        ''' <summary>
        ''' Grava num temporario ao lado do destino e so entao move.
        '''
        ''' Escrever direto no caminho final deixaria um arquivo PARCIAL la
        ''' se o SaveAsFile falhasse no meio — e um anexo truncado com o nome
        ''' certo e pior que anexo nenhum. O move final tambem fecha a janela
        ''' entre o File.Exists e a gravacao, em que outro processo poderia
        ''' criar o arquivo.
        ''' </summary>
        Private Function GravarComTemporario(a As OL.Attachment, destino As String,
                                             overwrite As Boolean) As OperationResult(Of String)
            Dim completo As String
            Try
                completo = Path.GetFullPath(destino)
            Catch
                Return OperationResult(Of String).Fail(ErrorKind.Unexpected, "caminho invalido")
            End Try

            Dim pasta = Path.GetDirectoryName(completo)
            If String.IsNullOrEmpty(pasta) Then
                Return OperationResult(Of String).Fail(ErrorKind.Unexpected, "sem diretorio")
            End If

            Try
                Directory.CreateDirectory(pasta)
            Catch
                Return OperationResult(Of String).Fail(ErrorKind.Denied, "diretorio")
            End Try

            ' No MESMO diretorio: mover entre volumes nao e atomico.
            Dim temporario = Path.Combine(pasta, $".iris-{Guid.NewGuid():N}.tmp")

            Try
                a.SaveAsFile(temporario)
                File.Move(temporario, completo, overwrite)
                Return OperationResult(Of String).Ok(completo)
            Catch ex As IOException
                Limpar(temporario)
                Return OperationResult(Of String).Fail(ErrorKind.Denied, "arquivo ja existe ou em uso")
            Catch ex As UnauthorizedAccessException
                Limpar(temporario)
                Return OperationResult(Of String).Fail(ErrorKind.Denied, "sem permissao")
            Catch
                Limpar(temporario)
                Throw
            End Try
        End Function

        Private Sub Limpar(caminho As String)
            Try
                If File.Exists(caminho) Then File.Delete(caminho)
            Catch
                ' Um temporario orfao nao justifica derrubar a operacao.
            End Try
        End Sub

        ''' <summary>
        ''' MAPI_E_NOT_FOUND, E_INVALIDARG e o "objeto nao esta no cache
        ''' local" do MAPI. Ocupado e desconectado NAO entram aqui.
        ''' </summary>
        Private Function EhConteudoIndisponivel(hresult As Integer) As Boolean
            Return hresult = &H8004010F OrElse hresult = &H80070057 OrElse
                   hresult = &H80040604
        End Function

        Private Function EhNaoEncontrado(hresult As Integer) As Boolean
            Return hresult = &H8004010F OrElse hresult = &H80070057
        End Function

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
                Return New DateTimeOffset(DateTime.SpecifyKind(getter(), DateTimeKind.Local))
            Catch
                Return Nothing
            End Try
        End Function

    End Module

End Namespace

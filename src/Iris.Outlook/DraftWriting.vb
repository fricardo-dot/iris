Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Text
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' Criação, edição e envio de rascunhos.
    '''
    ''' A regra que organiza este arquivo: **o rascunho é do Outlook, não
    ''' nosso**. Responder, responder a todos e encaminhar usam
    ''' <c>Reply</c>, <c>ReplyAll</c> e <c>Forward</c> do próprio OOM.
    ''' Reconstruir destinatários, citação e assinatura à mão seria
    ''' reimplementar regras que o Outlook já aplica — e errar nelas
    ''' significa mandar mensagem para quem não devia.
    '''
    ''' O Iris escreve ACIMA do que o Outlook gerou, nunca por cima.
    ''' </summary>
    Friend Module DraftWriting

        ''' <summary>
        ''' Marca onde termina o texto do usuário e começa a citação.
        '''
        ''' Invisível na mensagem, e é o que permite reabrir um rascunho
        ''' sabendo o que era digitação e o que era citação — sem isso, cada
        ''' salvamento reprocessaria a citação inteira e ela degradaria.
        ''' </summary>
        Private Const MarcaHtml As String = "<!--iris-quote-->"
        Private Const MarcaTexto As String = vbCrLf & "----- mensagem original -----" & vbCrLf

        ' ===================================================================
        ' Criação
        ' ===================================================================

        Public Function CreateNew(app As OL.Application, ns As OL.NameSpace) As OperationResult(Of DraftInfo)
            Dim item As OL.MailItem = Nothing
            Try
                item = TryCast(app.CreateItem(OL.OlItemType.olMailItem), OL.MailItem)
                If item Is Nothing Then
                    Return OperationResult(Of DraftInfo).Fail(ErrorKind.Unexpected, "CreateItem")
                End If

                ' Salvo JÁ: o rascunho precisa existir antes de o usuário
                ' digitar qualquer coisa, para sobreviver a um fechamento
                ' acidental e ter chave estável.
                item.Save()
                Return OperationResult(Of DraftInfo).Ok(Descrever(item, ns))
            Finally
                ComHelpers.Release(item)
            End Try
        End Function

        Public Function CreateReply(ns As OL.NameSpace, origem As ItemKey,
                                    replyAll As Boolean) As OperationResult(Of DraftInfo)
            Dim original As OL.MailItem = Nothing
            Try
                Try
                    original = TryCast(ns.GetItemFromID(origem.EntryId, origem.StoreId), OL.MailItem)
                Catch ex As COMException When EhNaoEncontrado(ex.HResult)
                    Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "original")
                End Try

                If original Is Nothing Then
                    Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "original")
                End If

                Dim resposta As OL.MailItem = Nothing
                Try
                    ' Reply devolve um item NOVO, com dono e liberação próprios.
                    resposta = TryCast(If(replyAll, original.ReplyAll(), original.Reply()), OL.MailItem)
                    If resposta Is Nothing Then
                        Return OperationResult(Of DraftInfo).Fail(ErrorKind.Unexpected, "Reply")
                    End If

                    resposta.Save()
                    Return OperationResult(Of DraftInfo).Ok(Descrever(resposta, ns))
                Finally
                    ComHelpers.Release(resposta)
                End Try
            Finally
                ComHelpers.Release(original)
            End Try
        End Function

        Public Function CreateForward(ns As OL.NameSpace, origem As ItemKey) As OperationResult(Of DraftInfo)
            Dim original As OL.MailItem = Nothing
            Try
                Try
                    original = TryCast(ns.GetItemFromID(origem.EntryId, origem.StoreId), OL.MailItem)
                Catch ex As COMException When EhNaoEncontrado(ex.HResult)
                    Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "original")
                End Try

                If original Is Nothing Then
                    Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "original")
                End If

                Dim encaminhada As OL.MailItem = Nothing
                Try
                    ' Forward preserva os anexos da original, o que
                    ' reconstruir à mão erraria.
                    encaminhada = TryCast(original.Forward(), OL.MailItem)
                    If encaminhada Is Nothing Then
                        Return OperationResult(Of DraftInfo).Fail(ErrorKind.Unexpected, "Forward")
                    End If

                    encaminhada.Save()
                    Return OperationResult(Of DraftInfo).Ok(Descrever(encaminhada, ns))
                Finally
                    ComHelpers.Release(encaminhada)
                End Try
            Finally
                ComHelpers.Release(original)
            End Try
        End Function

        ' ===================================================================
        ' Edição
        ' ===================================================================

        Public Function Update(ns As OL.NameSpace, chave As DraftKey,
                               conteudo As DraftContent) As OperationResult(Of DraftInfo)
            Dim item As OL.MailItem = Nothing
            Try
                Try
                    item = TryCast(ns.GetItemFromID(chave.Item.EntryId, chave.Item.StoreId), OL.MailItem)
                Catch ex As COMException When EhNaoEncontrado(ex.HResult)
                    Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "rascunho")
                End Try

                If item Is Nothing Then
                    Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "rascunho")
                End If

                item.Subject = If(conteudo.Subject, "")
                AplicarDestinatarios(item, conteudo)
                AplicarCorpo(item, conteudo.UserText)

                item.Save()

                ' O EntryID é relido DEPOIS do Save: ele muda quando o item é
                ' movido, e a Fase 0 mediu isso no critério D3.
                Return OperationResult(Of DraftInfo).Ok(Descrever(item, ns))
            Finally
                ComHelpers.Release(item)
            End Try
        End Function

        ''' <summary>
        ''' Escreve o texto do usuário ACIMA do que o Outlook gerou.
        '''
        ''' O formato seguido é o do RASCUNHO, não o nosso: se o Outlook
        ''' montou corpo HTML — com citação e assinatura corporativa —, o
        ''' texto digitado entra como HTML escapado. Forçar texto puro aqui
        ''' apagaria a citação e a assinatura, que é justamente o que torna
        ''' uma resposta utilizável no trabalho.
        ''' </summary>
        Private Sub AplicarCorpo(item As OL.MailItem, userText As String)
            ' NAO chamar esta variavel de "texto": VB e case-insensitive e o
            ' nome eclipsaria a funcao Texto() deste mesmo modulo. Ja
            ' aconteceu com Point, Rect, Path e Dispatcher neste projeto.
            Dim digitado = If(userText, "")

            If item.BodyFormat = OL.OlBodyFormat.olFormatHTML Then
                Dim atual = Texto(Function() item.HTMLBody)
                Dim citacao = SepararCitacaoHtml(atual)
                item.HTMLBody = ParaHtml(digitado) & MarcaHtml & citacao
            Else
                Dim atual = Texto(Function() item.Body)
                Dim citacao = SepararCitacaoTexto(atual)
                item.Body = digitado & MarcaTexto & citacao
            End If
        End Sub

        Private Function SepararCitacaoHtml(corpo As String) As String
            If String.IsNullOrEmpty(corpo) Then Return ""
            Dim pos = corpo.IndexOf(MarcaHtml, StringComparison.Ordinal)
            If pos < 0 Then Return corpo
            Return corpo.Substring(pos + MarcaHtml.Length)
        End Function

        Private Function SepararCitacaoTexto(corpo As String) As String
            If String.IsNullOrEmpty(corpo) Then Return ""
            Dim pos = corpo.IndexOf(MarcaTexto, StringComparison.Ordinal)
            If pos < 0 Then Return corpo
            Return corpo.Substring(pos + MarcaTexto.Length)
        End Function

        ''' <summary>
        ''' O texto vem do usuário e vai para dentro de HTML. Escapar não é
        ''' formalidade: um assunto ou nome com &lt; quebraria a mensagem.
        ''' </summary>
        Private Function ParaHtml(digitado As String) As String
            If String.IsNullOrEmpty(digitado) Then Return "<div></div>"
            Dim sb As New StringBuilder()
            For Each linha In digitado.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
                sb.Append("<div>").Append(WebUtility.HtmlEncode(linha)).Append("</div>")
            Next
            Return sb.ToString()
        End Function

        Private Sub AplicarDestinatarios(item As OL.MailItem, conteudo As DraftContent)
            ' To e Cc como texto: o Outlook resolve na hora do ResolveAll, e
            ' escrever nas propriedades evita mexer na coleção Recipients
            ' item a item, que criaria um RCW por destinatário.
            item.To = If(conteudo.ToLine, "")
            item.CC = If(conteudo.CcLine, "")
        End Sub

        Public Function AddAttachment(ns As OL.NameSpace, chave As DraftKey,
                                      caminho As String) As OperationResult(Of AttachmentInfo)
            If String.IsNullOrWhiteSpace(caminho) OrElse Not File.Exists(caminho) Then
                Return OperationResult(Of AttachmentInfo).Fail(ErrorKind.NotFound, "arquivo")
            End If

            Dim item As OL.MailItem = Nothing
            Try
                item = TryCast(ns.GetItemFromID(chave.Item.EntryId, chave.Item.StoreId), OL.MailItem)
                If item Is Nothing Then
                    Return OperationResult(Of AttachmentInfo).Fail(ErrorKind.NotFound, "rascunho")
                End If

                Dim anexos As OL.Attachments = Nothing
                Try
                    anexos = item.Attachments
                    Dim a As OL.Attachment = Nothing
                    Try
                        a = anexos.Add(caminho)
                        item.Save()
                        Dim indice = anexos.Count
                        Return OperationResult(Of AttachmentInfo).Ok(New AttachmentInfo With {
                            .Key = New AttachmentKey(chave.Item, indice,
                                                     Texto(Function() a.FileName),
                                                     Numero(Function() a.Size)),
                            .FileName = Texto(Function() a.FileName),
                            .SizeBytes = Numero(Function() a.Size),
                            .AttachmentType = Texto(Function() a.Type.ToString())
                        })
                    Finally
                        ComHelpers.Release(a)
                    End Try
                Finally
                    ComHelpers.Release(anexos)
                End Try
            Finally
                ComHelpers.Release(item)
            End Try
        End Function

        Public Function Delete(ns As OL.NameSpace, chave As DraftKey) As OperationResult(Of Boolean)
            Dim item As OL.MailItem = Nothing
            Try
                Try
                    item = TryCast(ns.GetItemFromID(chave.Item.EntryId, chave.Item.StoreId), OL.MailItem)
                Catch ex As COMException When EhNaoEncontrado(ex.HResult)
                    Return OperationResult(Of Boolean).Ok(False)
                End Try

                If item Is Nothing Then Return OperationResult(Of Boolean).Ok(False)
                item.Delete()
                Return OperationResult(Of Boolean).Ok(True)
            Finally
                ComHelpers.Release(item)
            End Try
        End Function

        ' ===================================================================
        ' Envio
        ' ===================================================================

        ''' <summary>
        ''' Prepara a confirmação: resolve destinatários e descobre a conta
        ''' remetente. NÃO envia nada.
        ''' </summary>
        Public Function PrepareSend(ns As OL.NameSpace, chave As DraftKey) As OperationResult(Of SendPreview)
            Dim item As OL.MailItem = Nothing
            Try
                Try
                    item = TryCast(ns.GetItemFromID(chave.Item.EntryId, chave.Item.StoreId), OL.MailItem)
                Catch ex As COMException When EhNaoEncontrado(ex.HResult)
                    Return OperationResult(Of SendPreview).Fail(ErrorKind.NotFound, "rascunho")
                End Try

                If item Is Nothing Then
                    Return OperationResult(Of SendPreview).Fail(ErrorKind.NotFound, "rascunho")
                End If

                Dim previa As New SendPreview With {
                    .Draft = chave,
                    .Subject = Texto(Function() item.Subject),
                    .SendingAccount = ContaRemetente(item)
                }

                Dim recipients As OL.Recipients = Nothing
                Try
                    recipients = item.Recipients
                    ' ResolveAll ANTES de listar: sem isso, "Resolved" seria
                    ' sempre falso e a confirmação mostraria o que o usuário
                    ' digitou, não para quem a mensagem realmente vai.
                    recipients.ResolveAll()

                    For i = 1 To recipients.Count
                        Dim r As OL.Recipient = Nothing
                        Try
                            r = recipients.Item(i)
                            previa.Recipients.Add(New RecipientInfo With {
                                .DisplayName = Texto(Function() r.Name),
                                .Address = EnderecoSmtp(r),
                                .Kind = TipoDeDestinatario(r),
                                .Resolved = Booleano(Function() r.Resolved)
                            })
                        Finally
                            ComHelpers.Release(r)
                        End Try
                    Next
                Finally
                    ComHelpers.Release(recipients)
                End Try

                Return OperationResult(Of SendPreview).Ok(previa)
            Finally
                ComHelpers.Release(item)
            End Try
        End Function

        ''' <summary>
        ''' Resolve o endereço SMTP de verdade.
        '''
        ''' <c>Recipient.Address</c> devolve <c>/O=...</c> para contas
        ''' Exchange, que serve para exibir e NÃO serve para uma confirmação
        ''' de envio: o usuário precisa reconhecer para quem está mandando.
        ''' </summary>
        Private Function EnderecoSmtp(r As OL.Recipient) As String
            Dim entrada As OL.AddressEntry = Nothing
            Try
                entrada = r.AddressEntry
                If entrada Is Nothing Then Return Texto(Function() r.Address)

                Dim tipo = entrada.AddressEntryUserType
                If tipo = OL.OlAddressEntryUserType.olExchangeUserAddressEntry OrElse
                   tipo = OL.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
                    Dim usuario As OL.ExchangeUser = Nothing
                    Try
                        usuario = entrada.GetExchangeUser()
                        If usuario IsNot Nothing Then
                            Dim smtp = Texto(Function() usuario.PrimarySmtpAddress)
                            If Not String.IsNullOrEmpty(smtp) Then Return smtp
                        End If
                    Finally
                        ComHelpers.Release(usuario)
                    End Try
                End If

                Return Texto(Function() r.Address)
            Catch
                Return Texto(Function() r.Address)
            Finally
                ComHelpers.Release(entrada)
            End Try
        End Function

        Private Function ContaRemetente(item As OL.MailItem) As String
            Dim conta As OL.Account = Nothing
            Try
                conta = item.SendUsingAccount
                If conta IsNot Nothing Then Return Texto(Function() conta.SmtpAddress)
                Return ""
            Catch
                Return ""
            Finally
                ComHelpers.Release(conta)
            End Try
        End Function

        ''' <summary>
        ''' Envia. UMA vez.
        '''
        ''' Roda por MutateAsync, com o retry do message filter desligado —
        ''' repetir um Send manda o e-mail duas vezes. Depois de chamado, o
        ''' item não é mais tocado.
        ''' </summary>
        Public Function Send(ns As OL.NameSpace, chave As DraftKey) As OperationResult(Of Boolean)
            Dim item As OL.MailItem = Nothing
            Try
                Try
                    item = TryCast(ns.GetItemFromID(chave.Item.EntryId, chave.Item.StoreId), OL.MailItem)
                Catch ex As COMException When EhNaoEncontrado(ex.HResult)
                    Return OperationResult(Of Boolean).Fail(ErrorKind.NotFound, "rascunho")
                End Try

                If item Is Nothing Then
                    Return OperationResult(Of Boolean).Fail(ErrorKind.NotFound, "rascunho")
                End If

                Dim recipients As OL.Recipients = Nothing
                Try
                    recipients = item.Recipients
                    ' Destinatário não resolvido BLOQUEIA. Ignorar o retorno
                    ' de ResolveAll foi um defeito real do spike da Fase 0.
                    If Not recipients.ResolveAll() Then
                        Return OperationResult(Of Boolean).Fail(
                            ErrorKind.Denied, "destinatario nao resolvido")
                    End If
                Finally
                    ComHelpers.Release(recipients)
                End Try

                item.Send()
                Return OperationResult(Of Boolean).Ok(True)
            Finally
                ComHelpers.Release(item)
            End Try
        End Function

        ' ===================================================================

        Private Function Descrever(item As OL.MailItem, ns As OL.NameSpace) As DraftInfo
            Dim formato = If(Numero(Function() CInt(item.BodyFormat)) = CInt(OL.OlBodyFormat.olFormatHTML),
                             BodyFormat.Html, BodyFormat.PlainText)

            Dim corpo = If(formato = BodyFormat.Html,
                           Texto(Function() item.HTMLBody),
                           Texto(Function() item.Body))

            Dim info As New DraftInfo With {
                .Key = New DraftKey(New ItemKey(Texto(Function() item.EntryID),
                                                StoreIdDe(item))),
                .Subject = Texto(Function() item.Subject),
                .ToLine = Texto(Function() item.To),
                .CcLine = Texto(Function() item.CC),
                .Format = formato,
                .SendingAccount = ContaRemetente(item)
            }

            ' Separa o que o usuário digitou do que o Outlook gerou.
            If formato = BodyFormat.Html Then
                Dim pos = corpo.IndexOf(MarcaHtml, StringComparison.Ordinal)
                If pos >= 0 Then
                    info.UserText = DeHtml(corpo.Substring(0, pos))
                    info.QuotedBody = corpo.Substring(pos + MarcaHtml.Length)
                Else
                    info.QuotedBody = corpo
                End If
            Else
                Dim pos = corpo.IndexOf(MarcaTexto, StringComparison.Ordinal)
                If pos >= 0 Then
                    info.UserText = corpo.Substring(0, pos)
                    info.QuotedBody = corpo.Substring(pos + MarcaTexto.Length)
                Else
                    info.QuotedBody = corpo
                End If
            End If

            LerAnexosDoRascunho(item, info)
            Return info
        End Function

        Private Sub LerAnexosDoRascunho(item As OL.MailItem, info As DraftInfo)
            Dim anexos As OL.Attachments = Nothing
            Try
                anexos = item.Attachments
                For i = 1 To anexos.Count
                    Dim a As OL.Attachment = Nothing
                    Try
                        a = anexos.Item(i)
                        info.Attachments.Add(New AttachmentInfo With {
                            .Key = New AttachmentKey(info.Key.Item, i,
                                                     Texto(Function() a.FileName),
                                                     Numero(Function() a.Size)),
                            .FileName = Texto(Function() a.FileName),
                            .SizeBytes = Numero(Function() a.Size)
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
        ''' Converte de volta o HTML que NÓS geramos — não HTML arbitrário.
        ''' Só desfaz o que ParaHtml fez.
        ''' </summary>
        Private Function DeHtml(html As String) As String
            If String.IsNullOrEmpty(html) Then Return ""
            Dim simples = html.Replace("</div><div>", vbLf).
                               Replace("<div>", "").
                               Replace("</div>", "")
            Return WebUtility.HtmlDecode(simples)
        End Function

        ''' <summary>
        ''' O StoreID vem da pasta PAI: MailItem nao o expoe diretamente. E
        ''' item.Parent devolve um objeto COM, que precisa ser liberado —
        ''' encadear item.Parent.StoreID seria o R7 de novo.
        ''' </summary>
        Private Function StoreIdDe(item As OL.MailItem) As String
            Dim pai As OL.MAPIFolder = Nothing
            Try
                pai = TryCast(item.Parent, OL.MAPIFolder)
                If pai Is Nothing Then Return ""
                Return Texto(Function() pai.StoreID)
            Catch
                Return ""
            Finally
                ComHelpers.Release(pai)
            End Try
        End Function

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

    End Module

End Namespace

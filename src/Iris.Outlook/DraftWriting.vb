Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Text
Imports Iris.Core
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

        ''' <summary>Propriedade de usuário que identifica rascunho do Iris.</summary>
        Private Const PropDoIris As String = "IrisDraft"

        ''' <summary>A mesma marca, por propriedade nomeada MAPI.</summary>
        Private Const PropDoIrisMapi As String =
            "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/IrisDraft"

        ''' <summary>PR_SMTP_ADDRESS_W, para quando o OOM nao entrega o SMTP.</summary>
        Private Const PropSmtpAddress As String =
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
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

                ' Sem citação: é uma mensagem nova.
                PlantarMarca(item, temCitacao:=False)

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

                    ' O corpo que o Reply gerou — citação e assinatura —
                    ' é citação por inteiro.
                    PlantarMarca(resposta, temCitacao:=True)
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

                    PlantarMarca(encaminhada, temCitacao:=True)
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

        ''' <summary>
        ''' Planta a marca de separação no corpo, na CRIAÇÃO do rascunho.
        '''
        ''' Sem ela, a ausência da marca é ambígua, e o leitor teria de
        ''' adivinhar entre dois casos opostos: mensagem nova, em que não há
        ''' citação nenhuma, e resposta recém-montada pelo Outlook, em que a
        ''' citação é o corpo inteiro. Adivinhar "é tudo citação" numa
        ''' mensagem nova embutia o esqueleto HTML vazio do Outlook como se
        ''' fosse mensagem original — em toda mensagem que o usuário
        ''' escrevesse.
        '''
        ''' Resolvido na origem e não no leitor, porque quem cria SABE de
        ''' que caso se trata. O leitor mantém o palpite conservador só para
        ''' rascunho que não foi o Iris que criou: ali, tratar tudo como
        ''' citação preserva o conteúdo intacto.
        ''' </summary>
        Private Sub PlantarMarca(item As OL.MailItem, temCitacao As Boolean)
            MarcarComoDoIris(item)

            If item.BodyFormat = OL.OlBodyFormat.olFormatHTML Then
                ' Comentário HTML: invisível para quem recebe.
                Dim citacao = If(temCitacao, Texto(Function() item.HTMLBody), "")
                item.HTMLBody = ParaHtml("") & MarcaHtml & citacao
            Else
                ' Em texto puro não existe marca invisível — a linha aparece
                ' para quem recebe. Então ela SÓ é plantada quando existe
                ' citação de verdade embaixo dela.
                '
                ' Numa mensagem nova não há marca nenhuma, e o leitor precisa
                ' saber disso: sem marca, ele trataria o corpo inteiro como
                ' citação e o texto do próprio usuário viraria "mensagem
                ' original" não editável.
                If temCitacao Then
                    item.Body = MarcaTexto & Texto(Function() item.Body)
                Else
                    item.Body = ""
                End If
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

        ''' <summary>
        ''' Anexa e devolve o rascunho INTEIRO redescrito.
        '''
        ''' Devolvia so o AttachmentInfo, e isso era um furo na regra de que
        ''' a chave e relida a cada Save: anexar SALVA, o EntryID pode mudar
        ''' ai tambem, e o compositor ficava com a chave velha — o autosave
        ''' seguinte, ou o envio, daria NotFound. Quem salva devolve a
        ''' identidade nova; nao ha excecao.
        ''' </summary>
        Public Function AddAttachment(ns As OL.NameSpace, chave As DraftKey,
                                      caminho As String) As OperationResult(Of DraftInfo)
            If String.IsNullOrWhiteSpace(caminho) OrElse Not File.Exists(caminho) Then
                Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "arquivo")
            End If

            Dim item As OL.MailItem = Nothing
            Try
                item = TryCast(ns.GetItemFromID(chave.Item.EntryId, chave.Item.StoreId), OL.MailItem)
                If item Is Nothing Then
                    Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "rascunho")
                End If

                Dim anexos As OL.Attachments = Nothing
                Try
                    anexos = item.Attachments
                    Dim a As OL.Attachment = Nothing
                    Try
                        a = anexos.Add(caminho)
                    Finally
                        ComHelpers.Release(a)
                    End Try
                Finally
                    ComHelpers.Release(anexos)
                End Try

                item.Save()
                Return OperationResult(Of DraftInfo).Ok(Descrever(item, ns))
            Finally
                ComHelpers.Release(item)
            End Try
        End Function

        ''' <summary>
        ''' Tira um anexo do rascunho.
        '''
        ''' Localiza pelo NOME e pelo tamanho, não só pelo índice. O índice
        ''' guardado envelhece: basta outro anexo ter sido removido antes
        ''' para todos os seguintes andarem uma casa, e aí o índice antigo
        ''' apaga o anexo errado — que é irreversível do ponto de vista do
        ''' usuário, porque o arquivo original pode não existir mais.
        '''
        ''' O índice entra como desempate quando há nomes repetidos.
        ''' </summary>
        Public Function RemoveAttachment(ns As OL.NameSpace, chave As DraftKey,
                                         anexo As AttachmentKey) As OperationResult(Of DraftInfo)
            If anexo Is Nothing Then
                Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "anexo")
            End If

            Dim item As OL.MailItem = Nothing
            Try
                item = TryCast(ns.GetItemFromID(chave.Item.EntryId, chave.Item.StoreId), OL.MailItem)
                If item Is Nothing Then
                    Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "rascunho")
                End If

                ' Duas passadas: achar QUEM é, e só então apagar.
                '
                ' Numa passada só, o primeiro candidato ganhava — e com dois
                ' anexos de mesmo nome e mesmo tamanho, um índice
                ' desatualizado apagava o outro. Apagar o anexo errado é
                ' irreversível do ponto de vista do usuário: o arquivo de
                ' origem pode não existir mais.
                Dim candidatos As New List(Of Integer)()
                Dim anexos As OL.Attachments = Nothing
                Try
                    anexos = item.Attachments
                    For i = 1 To anexos.Count
                        Dim a As OL.Attachment = Nothing
                        Try
                            a = anexos.Item(i)
                            If MesmoArquivo(a, anexo) Then candidatos.Add(i)
                        Catch
                        Finally
                            ComHelpers.Release(a)
                        End Try
                    Next

                    Dim alvo = Escolher(candidatos, anexo.Index)
                    If alvo = 0 Then
                        If candidatos.Count = 0 Then
                            Return OperationResult(Of DraftInfo).Fail(ErrorKind.NotFound, "anexo")
                        End If
                        ' Stale, e NÃO Ambiguous. Neste projeto Ambiguous
                        ' significa "pode ter surtido efeito, nunca repetir",
                        ' e aqui nada aconteceu: a remoção foi recusada ANTES
                        ' do Delete. O que envelheceu foi a chave, e reler o
                        ' rascunho resolve.
                        Return OperationResult(Of DraftInfo).Fail(ErrorKind.Stale,
                                                                  "anexo ambiguo")
                    End If

                    Dim aAlvo As OL.Attachment = Nothing
                    Try
                        aAlvo = anexos.Item(alvo)

                        ' CONFERE DE NOVO O OBJETO QUE VAI APAGAR.
                        '
                        ' A primeira passada validou um Attachment, guardou o
                        ' INDICE e soltou o objeto. Entre as duas passadas a
                        ' colecao pode mudar, e ai o indice aponta para outro
                        ' anexo -- que era exatamente o defeito que as duas
                        ' passadas existem para impedir, sobrevivendo dentro
                        ' delas. A revisao externa pegou.
                        '
                        ' Recusa aqui e Stale pelo mesmo motivo de cima: nada
                        ' aconteceu, e reler o rascunho resolve.
                        If Not MesmoArquivo(aAlvo, anexo) Then
                            Return OperationResult(Of DraftInfo).Fail(
                                ErrorKind.Stale,
                                "o anexo naquele indice mudou entre conferir e apagar")
                        End If

                        aAlvo.Delete()
                    Finally
                        ComHelpers.Release(aAlvo)
                    End Try
                Finally
                    ComHelpers.Release(anexos)
                End Try

                item.Save()
                Return OperationResult(Of DraftInfo).Ok(Descrever(item, ns))
            Finally
                ComHelpers.Release(item)
            End Try
        End Function

        ''' <summary>
        ''' <b>Mesmo nome e mesmo tamanho — e os dois lados conclusivos.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>ESTA COMPARAÇÃO DECIDE UM <c>Delete()</c></b>
        '''
        ''' Ela lia os dois lados com os auxiliares tolerantes — exceção vira
        ''' <c>""</c> e <c>0</c> — e a chave alvo tinha sido montada com os
        ''' mesmos. Se as leituras falhassem nos dois momentos, <c>""/0</c>
        ''' casava com <c>""/0</c>, a comparação dizia "é este" e o anexo errado
        ''' era <b>apagado</b>.
        '''
        ''' É o mesmo defeito que o <c>MessageReading.MesmaIdentidade</c> tinha,
        ''' numa operação destrutiva. Consertar lá e deixar aqui é o erro que
        ''' este projeto já cometeu cinco vezes — a revisão externa pegou este
        ''' perguntando explicitamente pelos irmãos.
        '''
        ''' Agora <b>falha fecha</b>: sem os quatro valores conclusivos, não é o
        ''' mesmo arquivo, e a remoção é recusada antes do <c>Delete</c>.
        '''
        ''' <b>A regra mora num lugar só, de propósito.</b> A primeira versão
        ''' desta função repetia aqui o teste de
        ''' <see cref="AttachmentKey.IdentidadeConhecida"/> — e o controle
        ''' negativo não derrubou nada, porque
        ''' <see cref="MessageReading.MesmaIdentidade"/> já o fazia. Guarda
        ''' duplicada é guarda que ninguém prova: a cópia sai, e o que fica é a
        ''' que tem teste.
        ''' </summary>
        Private Function MesmoArquivo(a As OL.Attachment, alvo As AttachmentKey) As Boolean
            Dim nome As String = Nothing
            Dim tamanho As Integer? = Nothing
            Try
                nome = a.FileName
            Catch
            End Try
            Try
                tamanho = a.Size
            Catch
            End Try

            Return MessageReading.MesmaIdentidade(nome, tamanho, alvo)
        End Function

        ''' <summary>
        ''' Qual dos candidatos apagar. Zero significa "não dá para saber".
        '''
        ''' Candidato único: o índice não importa, porque não há outro que
        ''' possa ser confundido com ele. Vários candidatos: só o índice
        ''' exato serve, e se ele não estiver entre eles a resposta é
        ''' RECUSAR — nunca escolher um. Numa operação que apaga, "o mais
        ''' provável" não é resposta boa o bastante.
        ''' </summary>
        Private Function Escolher(candidatos As List(Of Integer), indice As Integer) As Integer
            If candidatos.Count = 0 Then Return 0
            If candidatos.Count = 1 Then Return candidatos(0)
            If candidatos.Contains(indice) Then Return indice
            Return 0
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
                    .SendingAccount = ContaRemetente(item, ns)
                }

                Dim recipients As OL.Recipients = Nothing
                Dim esperados = 0
                Dim esperadoDepois = -1
                Dim obtidos = 0
                Dim ultimaFalha = ErrorKind.None
                Try
                    recipients = item.Recipients
                    ' ResolveAll ANTES de listar: sem isso, "Resolved" seria
                    ' sempre falso e a confirmação mostraria o que o usuário
                    ' digitou, não para quem a mensagem realmente vai.
                    recipients.ResolveAll()

                    esperados = recipients.Count
                    For i = 1 To esperados
                        Dim r As OL.Recipient = Nothing
                        Try
                            r = recipients.Item(i)
                            ' Resolvido pelo Outlook NAO basta: ele resolve
                            ' um nome interno para /O=..., que e endereco que
                            ' o usuario nao tem como conferir.
                            Dim endereco = EnderecoSmtp(r)
                            previa.Recipients.Add(New RecipientInfo With {
                                .DisplayName = Texto(Function() r.Name),
                                .Address = endereco,
                                .Kind = TipoDeDestinatario(r),
                                .Resolved = Booleano(Function() r.Resolved) AndAlso EhSmtp(endereco)
                            })
                            obtidos += 1
                        Catch ex As COMException
                            ' Um destinatário que não se deixa ler não pode
                            ' virar silêncio: a confirmação mostraria uma
                            ' lista curta como se fosse a lista inteira.
                            ultimaFalha = OutlookFailurePolicy.ClassifyFailure(
                                ex.HResult, isMutation:=False, mutationAttemptStarted:=False)
                        Catch
                            ultimaFalha = ErrorKind.Unexpected
                        Finally
                            ComHelpers.Release(r)
                        End Try
                    Next

                    ' Conta de novo: a coleção pode ter mudado durante o
                    ' percurso, e um snapshot que não vale não pode ser
                    ' apresentado como a lista inteira. Mesmo tratamento que
                    ' MessageReading e LerAnexos — este ponto tinha ficado de
                    ' fora, justamente na tela que confere o envio.
                    esperadoDepois = recipients.Count
                    previa.RecipientsStatus = PartStatus.FromCounts(
                        esperados, esperadoDepois, obtidos, ultimaFalha)
                Catch ex As COMException
                    previa.RecipientsStatus = PartStatus.Missing(
                        OutlookFailurePolicy.ClassifyFailure(ex.HResult, isMutation:=False,
                                                             mutationAttemptStarted:=False))
                Catch
                    previa.RecipientsStatus = PartStatus.Missing(ErrorKind.Unexpected)
                Finally
                    ComHelpers.Release(recipients)
                End Try

                Dim statusDosAnexos As PartStatus = Nothing
                previa.Attachments.AddRange(LerAnexos(item, statusDosAnexos))
                previa.AttachmentsStatus = statusDosAnexos
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
                            If EhSmtp(smtp) Then Return smtp
                        End If
                    Catch
                    Finally
                        ComHelpers.Release(usuario)
                    End Try
                End If

                ' Lista de distribuicao Exchange tem SMTP proprio, e ela e
                ' justamente o destinatario em que errar custa mais caro:
                ' manda para o grupo inteiro.
                If tipo = OL.OlAddressEntryUserType.olExchangeDistributionListAddressEntry Then
                    Dim lista As OL.ExchangeDistributionList = Nothing
                    Try
                        lista = entrada.GetExchangeDistributionList()
                        If lista IsNot Nothing Then
                            Dim smtp = Texto(Function() lista.PrimarySmtpAddress)
                            If EhSmtp(smtp) Then Return smtp
                        End If
                    Catch
                    Finally
                        ComHelpers.Release(lista)
                    End Try
                End If

                ' Ultima tentativa antes de desistir: PR_SMTP_ADDRESS. Cobre
                ' os tipos que nao tem objeto proprio no Object Model —
                ' contatos, agentes, entradas de outros provedores.
                Dim acesso As OL.PropertyAccessor = Nothing
                Try
                    acesso = entrada.PropertyAccessor
                    If acesso IsNot Nothing Then
                        Dim bruto = acesso.GetProperty(PropSmtpAddress)
                        Dim smtp = If(TryCast(bruto, String), "")
                        If EhSmtp(smtp) Then Return smtp
                    End If
                Catch
                Finally
                    ComHelpers.Release(acesso)
                End Try

                ' Sobrou o que o Outlook deu. Pode ser /O=..., e quem decide
                ' se isso serve e o chamador, via EhSmtp: aqui o trabalho e
                ' devolver o melhor que se conseguiu, nao julgar.
                Return Texto(Function() r.Address)
            Catch
                Return Texto(Function() r.Address)
            Finally
                ComHelpers.Release(entrada)
            End Try
        End Function

        ''' <summary>
        ''' Por qual conta a mensagem vai sair.
        '''
        ''' SendUsingAccount é Nothing quando o usuário não escolheu conta —
        ''' que é o caso NORMAL. Devolver vazio aí deixava em branco
        ''' justamente o campo que a tela de confirmação existe para
        ''' mostrar: um "enviando pela conta: " sem nada depois. Quando não
        ''' há escolha explícita, quem manda é o store onde o rascunho mora.
        ''' </summary>
        Private Function ContaRemetente(item As OL.MailItem, ns As OL.NameSpace) As String
            Dim conta As OL.Account = Nothing
            Try
                conta = item.SendUsingAccount
                If conta IsNot Nothing Then
                    Dim escolhida = Texto(Function() conta.SmtpAddress)
                    If escolhida.Length > 0 Then Return escolhida
                End If
            Catch
                ' Cai para o store.
            Finally
                ComHelpers.Release(conta)
            End Try

            Return ContaDoStore(ns, StoreIdDe(item))
        End Function

        Private Function ContaDoStore(ns As OL.NameSpace, storeId As String) As String
            Dim contas As OL.Accounts = Nothing
            Try
                contas = ns.Accounts

                For i = 1 To contas.Count
                    Dim c As OL.Account = Nothing
                    Try
                        c = contas.Item(i)
                        Dim smtp = Texto(Function() c.SmtpAddress)

                        ' Nunca encadear: DeliveryStore sai numa variável
                        ' própria para ser liberado (R7).
                        Dim entrega As OL.Store = Nothing
                        Try
                            entrega = c.DeliveryStore
                            If entrega IsNot Nothing AndAlso
                               String.Equals(Texto(Function() entrega.StoreID), storeId,
                                             StringComparison.Ordinal) Then
                                Return smtp
                            End If
                        Catch
                        Finally
                            ComHelpers.Release(entrega)
                        End Try
                    Catch
                    Finally
                        ComHelpers.Release(c)
                    End Try
                Next

                ' Nenhuma conta entrega neste store. NAO devolver a
                ' primeira: seria apresentar palpite como fato, e a tela
                ' diria "enviando por A" enquanto o Outlook manda por B.
                ' Vazio faz a UI dizer que nao identificou — a verdade.
                Return ""
            Catch
                Return ""
            Finally
                ComHelpers.Release(contas)
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

            ' Em HTML a marca invisível responde sozinha. Em texto puro não
            ' há marca numa mensagem nova, então a pergunta "é rascunho do
            ' Iris?" é o que separa "corpo é do usuário" de "corpo é citação
            ' de um rascunho que veio de outro lugar".
            Dim ehDoIris = EhRascunhoDoIris(item)

            Dim info As New DraftInfo With {
                .Key = New DraftKey(New ItemKey(Texto(Function() item.EntryID),
                                                StoreIdDe(item))),
                .Subject = Texto(Function() item.Subject),
                .ToLine = Texto(Function() item.To),
                .CcLine = Texto(Function() item.CC),
                .Format = formato,
                .SendingAccount = ContaRemetente(item, ns)
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
                ElseIf ehDoIris Then
                    ' Rascunho que o Iris criou, em texto puro, SEM marca:
                    ' é uma mensagem nova, e o corpo inteiro é do usuário.
                    ' Tratar como citação transformaria o que ele digitou em
                    ' "mensagem original" não editável.
                    info.UserText = corpo
                Else
                    ' Rascunho de origem desconhecida: preservar tudo como
                    ' citação é o palpite que não destrói nada.
                    info.QuotedBody = corpo
                End If
            End If

            info.QuotedPreview = If(formato = BodyFormat.Html,
                                    TextoDeHtml(info.QuotedBody),
                                    info.QuotedBody)

            Dim statusDosAnexos As PartStatus = Nothing
            info.Attachments.AddRange(LerAnexos(item, statusDosAnexos))
            info.AttachmentsStatus = statusDosAnexos
            Return info
        End Function

        ''' <summary>
        ''' Marcação HTML para texto legível. É de EXIBIÇÃO, e só: o que
        ''' volta para o Outlook é o HTML original, intacto. Por isso pode
        ''' ser grosseiro — nada aqui é gravado nem enviado.
        ''' </summary>
        Private Function TextoDeHtml(html As String) As String
            If String.IsNullOrEmpty(html) Then Return ""

            Dim sb As New StringBuilder(html.Length)
            Dim dentroDeTag = False
            For Each c In html
                Select Case c
                    Case "<"c : dentroDeTag = True
                    Case ">"c : dentroDeTag = False : sb.Append(" "c)
                    Case Else : If Not dentroDeTag Then sb.Append(c)
                End Select
            Next

            Dim texto = WebUtility.HtmlDecode(sb.ToString())

            ' Tirar tags deixa um rastro de espaços e linhas vazias no lugar
            ' de cada uma. Sem colapsar, a citação vira uma coluna de vácuo.
            Dim linhas = texto.Replace(vbCrLf, vbLf).Split(CChar(vbLf)).
                               Select(Function(l) l.Trim()).
                               Where(Function(l) l.Length > 0)
            Return String.Join(vbCrLf, linhas)
        End Function

        ''' <summary>
        ''' Le os anexos do item. Devolve lista em vez de escrever dentro de
        ''' um DraftInfo porque a confirmacao de envio precisa dos mesmos
        ''' anexos sem ser um rascunho descrito.
        ''' </summary>
        ''' <summary>
        ''' Este rascunho foi criado pelo Iris nesta sessão?
        '''
        ''' Marca gravada numa propriedade de usuário, invisível para quem
        ''' recebe e independente do formato do corpo. Existe porque em texto
        ''' puro não dá para plantar marca no corpo sem que ela apareça na
        ''' mensagem — e sem alguma marca, "corpo sem separador" seria
        ''' ambíguo entre mensagem nova e rascunho de outra origem.
        ''' </summary>
        Private Function EhRascunhoDoIris(item As OL.MailItem) As Boolean
            Dim props As OL.UserProperties = Nothing
            Dim prop As OL.UserProperty = Nothing
            Try
                props = item.UserProperties
                prop = props.Find(PropDoIris)
                If prop IsNot Nothing Then Return True
            Catch
            Finally
                ComHelpers.Release(prop)
                ComHelpers.Release(props)
            End Try

            ' Segunda tentativa, pela propriedade nomeada: UserProperties
            ' pode ser negado por política, e nesse caso a marca foi gravada
            ' pelo outro caminho.
            Dim acesso As OL.PropertyAccessor = Nothing
            Try
                acesso = item.PropertyAccessor
                Dim valor = TryCast(acesso.GetProperty(PropDoIrisMapi), String)
                Return Not String.IsNullOrEmpty(valor)
            Catch
                Return False
            Finally
                ComHelpers.Release(acesso)
            End Try
        End Function

        Private Sub MarcarComoDoIris(item As OL.MailItem)
            Dim props As OL.UserProperties = Nothing
            Dim prop As OL.UserProperty = Nothing
            Try
                props = item.UserProperties
                prop = props.Find(PropDoIris)
                If prop Is Nothing Then
                    prop = props.Add(PropDoIris, OL.OlUserPropertyType.olText)
                End If
                prop.Value = "1"
            Catch
                ' UserProperties pode ser negado por política. Tenta a
                ' propriedade nomeada antes de desistir.
                Dim acesso As OL.PropertyAccessor = Nothing
                Try
                    acesso = item.PropertyAccessor
                    acesso.SetProperty(PropDoIrisMapi, "1")
                Catch
                    ' Sem marca durável. A consequência só aparece ao REABRIR
                    ' um rascunho de texto puro numa sessão futura — e reabrir
                    ' rascunho existente não está implementado. Registrado na
                    ' FASE1 para não virar surpresa quando estiver.
                Finally
                    ComHelpers.Release(acesso)
                End Try
            Finally
                ComHelpers.Release(prop)
                ComHelpers.Release(props)
            End Try
        End Sub

        ''' <summary>
        ''' Lê os anexos e DIZ o quanto conseguiu.
        '''
        ''' Continuava devolvendo só a lista e engolindo as falhas, então
        ''' <c>PrepareSend</c> e <c>Descrever</c> deixavam o status no
        ''' default — afirmando completude que ninguém provou. Numa
        ''' confirmação de envio isso é grave: a tela diz o que vai junto, e
        ''' anexo não lido não deixa rastro.
        ''' </summary>
        Private Function LerAnexos(item As OL.MailItem,
                                   ByRef status As PartStatus) As List(Of AttachmentInfo)
            Dim lista As New List(Of AttachmentInfo)()
            Dim dono = New ItemKey(Texto(Function() item.EntryID), StoreIdDe(item))

            Dim esperados = 0
            Dim esperadoDepois = -1
            Dim obtidos = 0
            Dim ultimaFalha = ErrorKind.None

            Dim anexos As OL.Attachments = Nothing
            Try
                anexos = item.Attachments
                esperados = anexos.Count

                For i = 1 To esperados
                    Dim a As OL.Attachment = Nothing
                    Try
                        a = anexos.Item(i)

                        ' A CHAVE PRECISA SABER SE A IDENTIDADE FOI LIDA.
                        ' Texto() e Numero() devolvem "" e 0 na falha, e esses
                        ' valores entravam na chave como se fossem conhecidos.
                        ' Depois, na REMOCAO, o MesmoArquivo comparava
                        ' fabricado com fabricado, passava, e chamava Delete()
                        ' no anexo errado. E o irmao destrutivo do defeito que
                        ' o MessageReading ja tinha.
                        Dim nome As String = Nothing
                        Dim tamanho As Integer? = Nothing
                        Try
                            nome = a.FileName
                        Catch
                        End Try
                        Try
                            tamanho = a.Size
                        Catch
                        End Try

                        lista.Add(New AttachmentInfo With {
                            .Key = New AttachmentKey(dono, i, If(nome, ""), If(tamanho, 0),
                                                     nome IsNot Nothing AndAlso tamanho.HasValue),
                            .FileName = If(nome, ""),
                            .SizeBytes = If(tamanho, 0)
                        })

                        ' ANEXO COM IDENTIDADE FABRICADA NAO E ANEXO OBTIDO.
                        ' Contar como obtido fazia a lista fechar como
                        ' COMPLETA, e a completude e o que a tela consulta
                        ' antes de encaminhar e de enviar.
                        If nome IsNot Nothing AndAlso tamanho.HasValue Then obtidos += 1
                    Catch ex As COMException
                        ultimaFalha = OutlookFailurePolicy.ClassifyFailure(
                            ex.HResult, isMutation:=False, mutationAttemptStarted:=False)
                    Catch
                        ultimaFalha = ErrorKind.Unexpected
                    Finally
                        ComHelpers.Release(a)
                    End Try
                Next

                esperadoDepois = anexos.Count

            Catch ex As COMException
                status = PartStatus.Missing(OutlookFailurePolicy.ClassifyFailure(
                    ex.HResult, isMutation:=False, mutationAttemptStarted:=False))
                Return lista
            Catch
                status = PartStatus.Missing(ErrorKind.Unexpected)
                Return lista
            Finally
                ComHelpers.Release(anexos)
            End Try

            status = PartStatus.FromCounts(esperados, esperadoDepois, obtidos, ultimaFalha)
            Return lista
        End Function

        ''' <summary>
        ''' A regra mora em <see cref="AddressPolicy"/>, no Core, porque o
        ''' compositor a aplica de novo antes de deixar enviar — e porque
        ''' aqui, dentro da camada COM, nenhum teste a alcançaria.
        ''' </summary>
        Private Function EhSmtp(endereco As String) As Boolean
            Return AddressPolicy.IsUsableSmtp(endereco)
        End Function

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

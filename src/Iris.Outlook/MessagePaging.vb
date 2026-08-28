Imports System.Collections.Generic
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports Iris.Core
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
    '''   • A ordenação é feita pelo OUTLOOK, nunca em laço nosso.
    '''   • O corpo NÃO é lido durante a listagem. Uma leitura de corpo no
    '''     meio da paginação bloquearia a fila única da STA (F1-F).
    '''   • Toda referência COM adquirida é liberada. Encadear
    '''     <c>item.Attachments.Count</c> é o erro que já apareceu quatro
    '''     vezes neste projeto.
    '''
    ''' ------------------------------------------------------------------
    ''' DOIS CAMINHOS, e o porquê
    '''
    ''' <b>ReceivedDesc</b> usa <c>Table</c> + cursor: a Q1 mediu ~18x mais
    ''' rápido no DTO completo, e plano com a profundidade (página de 50 em
    ''' ~27 ms contra ~570 ms).
    '''
    ''' <b>As outras três ordenações</b> continuam na iteração com offset. A
    ''' Q1 mediu ReceivedTime DECRESCENTE e mais nada; filtro keyset sobre
    ''' texto traz ordenação cultural, caixa, acentuação e escaping de
    ''' apóstrofo, que é superfície não medida. Entregar o ganho medido e
    ''' registrar a dívida é melhor que inventar filtro que ninguém testou.
    '''
    ''' Os dois caminhos produzem <c>MailSummary</c> pelo MESMO significado
    ''' de campo — é isso que impede virarem dois produtos diferentes.
    ''' </summary>
    Friend Module MessagePaging

        Private Const DaslRecebido As String = "urn:schemas:httpmail:datereceived"
        Private Const TagAnexo As String = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"

        ''' <summary>
        ''' PR_LONGTERM_ENTRYID_FROM_TABLE.
        '''
        ''' A coluna chamada "EntryID" da Table devolve o EntryID de CURTO
        ''' PRAZO - 24 bytes, valido so dentro da sessao. O MailItem.EntryID
        ''' devolve o de LONGO PRAZO, 70 bytes. Sao identificadores diferentes
        ''' do MESMO item, e nenhum erro aparece: a listagem funciona, so nao
        ''' casa com nada.
        '''
        ''' CUIDADO COM O NOME. "Longo prazo" aqui significa que ela
        ''' sobrevive ao fim da SESSAO, e nada alem disso: a secao 11.1 do
        ''' FASE2 mediu que ela MUDA num Move. Ela e localizador da
        ''' encarnacao atual, NAO identidade do item. Usa-la como chave
        ''' primaria no cache seria exatamente o erro que a Q2 existe para
        ''' impedir.
        '''
        ''' O teste de cruzamento pegou isso: os dois caminhos liam as MESMAS
        ''' 995 mensagens em 34 paginas e ainda assim os conjuntos de chave
        ''' davam intersecao ZERO. Guardar chave de curto prazo num cache que
        ''' sobrevive a sessao seria pior que perder mensagem, porque so
        ''' quebraria na sessao seguinte.
        ''' </summary>
        Private Const TagEntryIdLongo As String = "http://schemas.microsoft.com/mapi/proptag/0x66700102"

        ' Ordem das colunas pedidas à Table. Índice fixo só é seguro porque
        ' TODAS foram confirmadas contra o Outlook real (tools/q1-colunas.ps1):
        ' aceitas E devolvendo valor não nulo em 40 itens.
        Private Const ColEntryId As Integer = 0
        Private Const ColSubject As Integer = 1
        Private Const ColSender As Integer = 2
        Private Const ColRecebido As Integer = 3
        Private Const ColTamanho As Integer = 4
        Private Const ColNaoLida As Integer = 5
        Private Const ColClasse As Integer = 6
        Private Const ColAnexo As Integer = 7

        Private Function CampoDeOrdenacao(sort As MessageSort) As (Campo As String, Descendente As Boolean)
            Select Case sort
                Case MessageSort.ReceivedAsc : Return ("[ReceivedTime]", False)
                Case MessageSort.SubjectAsc : Return ("[Subject]", False)
                Case MessageSort.SenderAsc : Return ("[SenderName]", False)
                Case Else : Return ("[ReceivedTime]", True)
            End Select
        End Function

        Public Function ReadPage(ns As OL.NameSpace, query As MessageQuery,
                                 continuation As String, targetCount As Integer) _
            As OperationResult(Of MessagePage)

            If targetCount <= 0 Then
                Return OperationResult(Of MessagePage).Fail(
                    ErrorKind.Unexpected, "targetCount invalido")
            End If

            ' Cursor de outra consulta é RECUSADO, não reinterpretado. Um
            ' cursor trocado produziria página de outra pasta sem a UI ter
            ' como perceber.
            Dim cursor As MessageCursor = Nothing
            Dim primeiraPagina = String.IsNullOrEmpty(continuation)
            If Not primeiraPagina Then
                If Not MessageCursor.TryDecode(continuation, query, cursor) Then
                    Return OperationResult(Of MessagePage).Fail(
                        ErrorKind.Stale, "cursor de outra consulta ou invalido")
                End If
            End If

            Dim folder As OL.MAPIFolder = Nothing
            Try
                Try
                    folder = TryCast(ns.GetFolderFromID(query.Folder.EntryId, query.Folder.StoreId),
                                     OL.MAPIFolder)
                Catch ex As COMException When EhNaoEncontrado(ex.HResult)
                    ' SO os HRESULTs que realmente significam "nao existe".
                    ' Chamada recusada e acesso negado sao outra coisa e
                    ' levam a UI a decisoes opostas; essas sobem para o
                    ' classificador do broker.
                    Return OperationResult(Of MessagePage).Fail(ErrorKind.NotFound, "pasta")
                End Try

                If folder Is Nothing Then
                    Return OperationResult(Of MessagePage).Fail(ErrorKind.NotFound, "pasta")
                End If

                ' Quem decide o caminho e o MODO DO CURSOR, nao a ordenacao
                ' sozinha. Antes era a ordenacao, e isso quebrava a travessia
                ' exatamente quando o fallback era necessario: a primeira
                ' pagina caia para iteracao e devolvia cursor legado, e a
                ' segunda recusava esse cursor como "de outro modo". A pasta
                ' parava na pagina 2 com Stale.
                Dim usaTabela As Boolean
                If cursor Is Nothing Then
                    usaTabela = (query.Sort = MessageSort.ReceivedDesc)
                ElseIf cursor.Mode = CursorMode.ReceivedDesc Then
                    If query.Sort <> MessageSort.ReceivedDesc Then
                        Return OperationResult(Of MessagePage).Fail(
                            ErrorKind.Stale, "cursor de outra ordenacao")
                    End If
                    usaTabela = True
                Else
                    ' Cursor legado: continua no caminho legado, mesmo que a
                    ' ordenacao seja ReceivedDesc.
                    usaTabela = False
                End If

                Dim pagina As MessagePage
                If usaTabela Then
                    Try
                        pagina = LerPorTabela(folder, query, cursor, targetCount)
                    Catch ex As TablePathUnusableException
                        ' A tabela nao serve para esta pasta. Uma travessia JA
                        ' comecada nao pode trocar de caminho no meio: cursor de
                        ' fronteira nao se traduz em offset, e inventar a
                        ' traducao pularia mensagem. Recarregar resolve.
                        If cursor IsNot Nothing Then
                            Return OperationResult(Of MessagePage).Fail(
                                ErrorKind.Stale, "caminho rapido indisponivel: " & ex.Message)
                        End If
                        pagina = LerPorIteracao(folder, query, Nothing, targetCount)
                    End Try
                Else
                    pagina = LerPorIteracao(folder, query, cursor, targetCount)
                End If

                If primeiraPagina Then pagina.TotalAtStart = ContarItens(folder)
                Return OperationResult(Of MessagePage).Ok(pagina)
            Finally
                ComHelpers.Release(folder)
            End Try
        End Function

        ''' <summary>
        ''' Só na primeira página. Uma contagem por página custaria uma
        ''' chamada COM na fila única da STA para reafirmar um número que já
        ''' nasce desatualizado.
        ''' </summary>
        Private Function ContarItens(folder As OL.MAPIFolder) As Integer?
            Dim colecao As OL.Items = Nothing
            Try
                colecao = folder.Items
                Return colecao.Count
            Catch ex As COMException
                Return Nothing
            Finally
                ComHelpers.Release(colecao)
            End Try
        End Function

        ' ==================================================================
        ' CAMINHO RÁPIDO: Table + cursor
        ' ==================================================================

        Private Function LerPorTabela(folder As OL.MAPIFolder, query As MessageQuery,
                                      cursor As MessageCursor, targetCount As Integer) As MessagePage

            Dim fonte As New TableRowSource(folder)
            Dim fronteira As DateTimeOffset? = If(cursor Is Nothing, Nothing, cursor.Boundary)

            Dim saida = CursorPaging.ReadPage(fonte, fronteira, targetCount)

            Dim pagina As New MessagePage With {
                .Generation = query.Generation,
                .DrainedExtra = saida.DrainedExtra
            }

            Dim descartadas = 0
            For Each linha In saida.Rows
                ' Chave vazia NAO vira item ignorado: e sinal de que a coluna
                ' foi aceita e nao entrega valor — o mesmo falso positivo que
                ' a Q1 teve com Permission, quando eu contei "nao lancou"
                ' como sucesso. A pasta inteira cai para o caminho lento.
                If String.IsNullOrEmpty(linha.EntryId) Then
                    Throw New TablePathUnusableException(
                        "a coluna de EntryID longo foi aceita mas devolveu vazio")
                End If
                ' ReceivedTime nulo tambem derruba a pasta inteira. Qualquer
                ' filtro "< data" exclui linha nula, entao a partir da
                ' segunda pagina esses itens sumiriam em silencio. Medido: 0
                ' de 1792 nesta caixa — mas "nao acontece aqui" nao e
                ' contrato.
                If Not linha.ReceivedTime.HasValue Then
                    Throw New TablePathUnusableException(
                        "item sem ReceivedTime: o filtro por data o excluiria")
                End If
                If Not EhMensagem(linha.MessageClass) Then
                    descartadas += 1
                    Continue For
                End If
                pagina.Items.Add(Resumir(linha, query.Folder.StoreId))
            Next

            pagina.SkippedCount = descartadas
            ' A FABRICACAO SOBE JUNTO COM O DESCARTE, e pelo mesmo motivo:
            ' celula ausente que virou 0 ou False e informacao sobre a
            ' leitura, e sem um numero ela nunca apareceria em lugar nenhum.
            '
            ' Soma so as linhas que ENTRARAM na pagina. As de read-ahead, que o
            ' CursorPaging converteu e devolveu para o lote seguinte, ficam de
            ' fora -- e serao contadas na pagina delas, uma vez so.
            pagina.FabricatedCells = Fabricadas(saida.Rows)
            pagina.NextCursor = If(saida.Ended, Nothing,
                                   MessageCursor.ForBoundary(query, saida.NextBoundary.Value).Encode())
            Return pagina
        End Function

        ''' <summary>
        ''' O caminho por iteração decidia por <c>TryCast(.., MailItem)</c>.
        ''' A tabela não tem TryCast, então a decisão passa a ser por classe.
        '''
        ''' NÃO é a mesma pergunta, e a equivalência está afirmada e não
        ''' provada — está registrada como dívida. O que se sabe desta caixa:
        ''' 1.741 <c>IPM.Note</c>, 49 de reunião e 4 avulsos.
        '''
        ''' A filtragem é feita AQUI e não por restrição DASL de propósito:
        ''' linha excluída pelo servidor não pode ser contada em
        ''' <c>SkippedCount</c>, e "28 de 30" viraria mistério.
        ''' </summary>
        Private Function EhMensagem(classe As String) As Boolean
            If String.IsNullOrEmpty(classe) Then Return False
            Return classe.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function Resumir(linha As TableRow, storeId As String) As MailSummary
            Return New MailSummary With {
                .Key = New ItemKey(linha.EntryId, storeId),
                .Subject = linha.Subject,
                .SenderName = linha.SenderName,
                .ReceivedTime = linha.ReceivedTime,
                .SizeBytes = linha.SizeBytes,
                .HasAttachments = linha.HasAttachments,
                .IsUnread = linha.IsUnread,
                .MessageClass = linha.MessageClass,
                .Content = ContentState.MetadataOnly
            }
        End Function

        ' ==================================================================
        ' CAMINHO LEGADO: iteração com offset
        ' ==================================================================

        ''' <summary>
        ''' O caminho legado: <b>paginação por offset</b> sobre uma coleção
        ''' ordenada. Usado quando a ordenação não é <c>ReceivedDesc</c> ou
        ''' quando a <c>Table</c> não serve para a pasta.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>OFFSET SOBRE COLEÇÃO VIVA REPETE E PULA — POR CONSTRUÇÃO</b>
        '''
        ''' Não é defeito do laço; é o que offset significa. Entre a página N e
        ''' a N+1 a coleção pode mudar, e o <see cref="OffsetPaging"/> tem a
        ''' demonstração em <c>OffsetPagingTests</c>:
        '''
        '''   • item <b>inserido</b> antes do offset → tudo desce uma posição, e
        '''     a próxima página reentrega o último item da anterior.
        '''     <b>Chave repetida.</b>
        '''   • item <b>removido</b> antes do offset → tudo sobe uma posição, e
        '''     um item nunca é visitado. <b>Mensagem perdida em silêncio</b> —
        '''     o pior dos dois, e o sintoma que a Q1 existe para pegar.
        '''   • <b>reordenar sem mudar o conjunto</b> também repete. Numa pasta
        '''     real isso é um <c>Subject</c> mudando com a ordenação em
        '''     <c>SubjectAsc</c>.
        '''
        ''' <b>Eu tinha escrito isto ao contrário</b> — "removido repete,
        ''' inserido pula" — e derivado disso uma regra de teste que também
        ''' estava errada. A aritmética saiu daqui para o
        ''' <see cref="OffsetPaging"/> justamente por isso: onde só a caixa real
        ''' exercita, ninguém demonstra qual direção é a verdadeira.
        '''
        ''' A defesa real seria cursor por chave — o que o caminho por
        ''' <c>Table</c> faz, e o motivo de ele existir. Aqui não dá:
        ''' <c>Items.Sort</c> ordena por campo não único (<c>Subject</c>,
        ''' <c>SenderName</c>) e o OOM não expõe "continue depois desta chave".
        '''
        ''' Observado em 25/08/2026: 993 linhas e 992 chaves distintas numa
        ''' travessia por <c>SubjectAsc</c> da Caixa de Entrada.
        '''
        ''' Quem consome este caminho <b>precisa</b> tolerar repetição. A
        ''' varredura da Fase 2 tolera, porque publica por
        ''' <c>provider_entry_id</c> e recarrega em vez de acumular.
        ''' </summary>
        Private Function LerPorIteracao(folder As OL.MAPIFolder, query As MessageQuery,
                                        cursor As MessageCursor, targetCount As Integer) As MessagePage

            Dim offset = If(cursor Is Nothing, 0, cursor.Offset)
            Dim colecao As OL.Items = Nothing
            Try
                colecao = folder.Items

                ' Sort que falha NAO pode ser engolido: seguir na ordem
                ' natural fingindo que ordenou destroi a coerencia das
                ' paginas. A excecao sobe para o classificador do broker.
                Dim ordenacao = CampoDeOrdenacao(query.Sort)
                colecao.Sort(ordenacao.Campo, ordenacao.Descendente)

                Dim total = colecao.Count
                Dim pagina As New MessagePage With {.Generation = query.Generation}

                ' Items e 1-based no OOM. A aritmetica mora no OffsetPaging,
                ' onde existe teste deterministico dizendo o que cada mutacao
                ' faz - foi por falta disso que eu escrevi a direcao errada.
                Dim janela = OffsetPaging.Janela(offset, targetCount, total)
                Dim primeiro As Integer = janela.Primeiro
                Dim ultimo As Integer = janela.Ultimo
                Dim descartadas = 0
                ' A fabricacao do caminho legado, que ate 28/08 era muda.
                Dim fabricadas = 0

                For i = primeiro To ultimo
                    Dim bruto As Object = Nothing
                    Try
                        bruto = colecao.Item(i)
                        Dim mail = TryCast(bruto, OL.MailItem)
                        If mail Is Nothing Then
                            ' Uma colecao Items nao contem apenas MailItem
                            ' (ESCOPO.md secao 5).
                            descartadas += 1
                            Continue For
                        End If
                        pagina.Items.Add(ResumirDoItem(mail, query.Folder.StoreId, fabricadas))
                    Catch ex As COMException
                        ' Item corrompido ou nao baixado nao derruba a pagina.
                        descartadas += 1
                        Continue For
                    Finally
                        ComHelpers.Release(bruto)
                    End Try
                Next

                pagina.SkippedCount = descartadas
                pagina.FabricatedCells = fabricadas
                ' Avanca por POSICOES EXAMINADAS, nao por DTOs devolvidos:
                ' contar DTOs relia as posicoes puladas e duplicava linha.
                pagina.NextCursor = If(janela.Proximo.HasValue,
                                       MessageCursor.ForOffset(query, janela.Proximo.Value).Encode(),
                                       Nothing)
                Return pagina
            Finally
                ComHelpers.Release(colecao)
            End Try
        End Function

        ''' <summary>
        ''' Quantas células ausentes viraram valor nas linhas dadas.
        '''
        ''' Existe separada, e não embutida, porque é a única parte da contagem
        ''' que dá para provar sem COM — e a propriedade que ela carrega é a
        ''' que já errou duas vezes: <b>conta as linhas que recebeu, e só
        ''' elas</b>.
        ''' </summary>
        Friend Function Fabricadas(linhas As IEnumerable(Of TableRow)) As Integer
            If linhas Is Nothing Then Return 0
            Return linhas.Sum(Function(l) If(l Is Nothing, 0, l.Fabricadas))
        End Function

        Private Function ResumirDoItem(mail As OL.MailItem, storeId As String,
                                       ByRef fabricadas As Integer) As MailSummary
            Dim anexos = ContarAnexos(mail)

            ' ContentState.BodyAvailable significa "corpo LIDO", pela
            ' definicao do enum. A listagem nao le corpo nenhum.
            Return New MailSummary With {
                .Key = New ItemKey(Texto(Function() mail.EntryID, fabricadas), storeId),
                .Subject = Texto(Function() mail.Subject, fabricadas),
                .SenderName = Texto(Function() mail.SenderName, fabricadas),
                .ReceivedTime = Data(Function() mail.ReceivedTime),
                .SizeBytes = Numero(Function() mail.Size, fabricadas),
                .HasAttachments = anexos > 0,
                .IsUnread = Booleano(Function() mail.UnRead, fabricadas),
                .MessageClass = Texto(Function() mail.MessageClass, fabricadas),
                .Content = ContentState.MetadataOnly
            }
        End Function

        ' MAPI_E_NOT_FOUND e o "objeto nao existe" do MAPI; E_INVALIDARG
        ' aparece quando o EntryID nao pertence mais ao store.
        Private Function EhNaoEncontrado(hresult As Integer) As Boolean
            Return hresult = &H8004010F OrElse hresult = &H80070057
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
        '
        ' MAS ELE TAMBEM NAO PODE SUMIR EM SILENCIO. Ate 28/08/2026 estes
        ' quatro transformavam excecao em "", 0, False e Nothing sem contar --
        ' e o caminho por Table ja contava. A revisao externa pegou: a lista
        ' instrumentava o caminho rapido e deixava o legado mudo, entao o zero
        ' que ela mostrava para esta pasta era um zero FABRICADO.
        '
        ' O contador vai por ByRef porque estes auxiliares sao do modulo e o
        ' caminho legado nao tem um objeto onde acumular. Feio, e honesto.

        Private Function Texto(getter As Func(Of String), ByRef fabricadas As Integer) As String
            Try
                Return If(getter(), "")
            Catch
                fabricadas += 1
                Return ""
            End Try
        End Function

        Private Function Numero(getter As Func(Of Integer), ByRef fabricadas As Integer) As Integer
            Try
                Return getter()
            Catch
                fabricadas += 1
                Return 0
            End Try
        End Function

        Private Function Booleano(getter As Func(Of Boolean), ByRef fabricadas As Integer) As Boolean
            Try
                Return getter()
            Catch
                fabricadas += 1
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Data ilegível vira <c>Nothing</c>, e <b>não</b> conta como fabricação.
        '''
        ''' Aqui a ausência é preservada: <c>DateTimeOffset?</c> distingue "não
        ''' sei" de qualquer valor. Contar seria alarmar sobre o único dos cinco
        ''' campos que já faz a coisa certa.
        ''' </summary>
        Private Function Data(getter As Func(Of DateTime)) As DateTimeOffset?
            Try
                Dim valor = getter()
                ' O Outlook devolve hora local sem Kind. Assumir o fuso local
                ' aqui é o que impede a mensagem aparecer com hora errada
                ' quando a Fase 2 persistir isto.
                Return New DateTimeOffset(DateTime.SpecifyKind(valor, DateTimeKind.Local))
            Catch
                Return Nothing
            End Try
        End Function

        ' ==================================================================

        ''' <summary>
        ''' O caminho por tabela nao serve para esta pasta, e a resposta e usar
        ''' o caminho lento — nunca degradar a chave nem seguir com dado que
        ''' some depois. Motivos: coluna de EntryID longo indisponivel, coluna
        ''' aceita devolvendo vazio, ou item sem ReceivedTime.
        '''
        ''' Nao e erro do usuario nem do Outlook: e capacidade ausente.
        ''' </summary>
        Friend NotInheritable Class TablePathUnusableException
            Inherits Exception
            Public Sub New(motivo As String)
                MyBase.New(motivo)
            End Sub
        End Class

        ''' <summary>
        ''' "Esta propriedade nao existe neste provider", e nada alem disso.
        ''' E_INVALIDARG e o que o Outlook devolve para coluna que ele nao
        ''' reconhece — medido: Columns.Add com um DASL invalido da
        ''' "Value does not fall within the expected range".
        ''' </summary>
        Private Function EhColunaRecusada(hresult As Integer) As Boolean
            Return hresult = &H80070057 OrElse hresult = &H8004010F
        End Function

        ''' <summary>Uma linha crua da Table, já fora do COM.</summary>
        Friend NotInheritable Class TableRow
            Public Property EntryId As String = ""
            Public Property Subject As String = ""
            Public Property SenderName As String = ""
            Public Property ReceivedTime As DateTimeOffset?
            Public Property SizeBytes As Integer
            Public Property IsUnread As Boolean
            Public Property HasAttachments As Boolean
            Public Property MessageClass As String = ""

            ''' <summary>
            ''' <b>Quantas células desta linha vieram ausentes e viraram valor.</b>
            '''
            ''' ------------------------------------------------------------------
            ''' <b>POR QUE O NÚMERO MORA NA LINHA, E NÃO NA FONTE</b>
            '''
            ''' Ele já morou na fonte, e teve dois defeitos seguidos por causa
            ''' disso. Primeiro subcontava: eu zerava a cada lote, e uma página
            ''' custa vários. Depois <b>sobrecontava</b>: o <c>CursorPaging</c>
            ''' lê um lote inteiro e para na primeira linha de outro instante, e
            ''' as linhas de <i>read-ahead</i> — que não entram nesta página —
            ''' já tinham sido convertidas e contadas. Na página seguinte elas
            ''' seriam contadas de novo.
            '''
            ''' Contador compartilhado entre "quem converte" e "quem escolhe o
            ''' que entra" erra nas duas direções, e cada conserto de um lado
            ''' abre o outro. Na linha não há o que errar: quem entrar na página
            ''' leva o seu número junto.
            ''' </summary>
            Public Property Fabricadas As Integer
        End Class

        ''' <summary>
        ''' A <c>Table</c> do Outlook vestida de <see cref="IRowSource(Of T)"/>.
        '''
        ''' A ARMADILHA QUE ESTA CLASSE EXISTE PARA NÃO REPETIR: o filtro
        ''' DASL interpreta a data como <b>UTC</b>, e a <c>Table</c> devolve
        ''' <c>ReceivedTime</c> em hora <b>local</b>. Na Q1 essa única
        ''' confusão perdeu 803 de 1.003 mensagens — e a paginação terminou
        ''' parecendo completa. Prova isolada: com string local vieram 938;
        ''' com string UTC, 953; contagem manual, 953.
        ''' </summary>
        Friend NotInheritable Class TableRowSource
            Implements IRowSource(Of TableRow)

            Private ReadOnly _folder As OL.MAPIFolder
            Private _table As OL.Table

            Public Sub New(folder As OL.MAPIFolder)
                _folder = folder
            End Sub

            Public Sub Abrir(fronteira As DateTimeOffset?, inclusiva As Boolean) _
                Implements IRowSource(Of TableRow).Abrir

                Fechar()

                Dim filtro = MontarFiltro(fronteira, inclusiva)
                _table = If(filtro Is Nothing, _folder.GetTable(), _folder.GetTable(filtro))

                Dim colunas As OL.Columns = Nothing
                Try
                    colunas = _table.Columns
                    colunas.RemoveAll()
                    ' A ordem aqui É os índices Col* lá em cima.
                    AdicionarColuna(colunas, TagEntryIdLongo)
                    AdicionarColuna(colunas, "Subject")
                    AdicionarColuna(colunas, "SenderName")
                    AdicionarColuna(colunas, "ReceivedTime")
                    AdicionarColuna(colunas, "Size")
                    AdicionarColuna(colunas, "UnRead")
                    AdicionarColuna(colunas, "MessageClass")
                    AdicionarColuna(colunas, TagAnexo)
                Finally
                    ComHelpers.Release(colunas)
                End Try

                _table.Sort("ReceivedTime", True)
            End Sub

            ''' <summary>
            ''' <c>Columns.Add</c> DEVOLVE um objeto COM. Ignorar o retorno
            ''' deixa um RCW sem dono — a regra R7, que já foi violada quatro
            ''' vezes neste projeto, sempre em código que "só lia".
            ''' </summary>
            Private Shared Sub AdicionarColuna(colunas As OL.Columns, nome As String)
                Dim coluna As OL.Column = Nothing
                Try
                    coluna = colunas.Add(nome)
                Catch ex As COMException When nome = TagEntryIdLongo AndAlso EhColunaRecusada(ex.HResult)
                    ' SO os HRESULTs que significam "esta propriedade nao
                    ' existe aqui". Busy, desconectado e acesso negado sao
                    ' outra coisa e precisam subir para o classificador do
                    ' broker — cair para o caminho lento por causa deles
                    ' esconderia falha de sessao atras de lentidao.
                    Throw New TablePathUnusableException(
                        "PR_LONGTERM_ENTRYID_FROM_TABLE indisponivel neste store")
                Finally
                    ComHelpers.Release(coluna)
                End Try
            End Sub

            Private Shared Function MontarFiltro(fronteira As DateTimeOffset?,
                                                 inclusiva As Boolean) As String
                If Not fronteira.HasValue Then Return Nothing
                Dim operador = If(inclusiva, "<=", "<")
                ' UTC. Ver o comentário da classe.
                Dim quando = fronteira.Value.UtcDateTime.ToString(
                    "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                Return "@SQL=""" & DaslRecebido & """ " & operador & " '" & quando & "'"
            End Function

            Public Function Ler(quantas As Integer) As IReadOnlyList(Of TableRow) _
                Implements IRowSource(Of TableRow).Ler

                Dim vazio = New List(Of TableRow)()
                If _table Is Nothing OrElse _table.EndOfTable Then Return vazio

                Dim bruto = TryCast(_table.GetArray(quantas), Object(,))
                If bruto Is Nothing Then Return vazio

                ' GetArray e 0-based nas DUAS dimensoes.
                Dim primeira = bruto.GetLowerBound(0)
                Dim ultima = bruto.GetUpperBound(0)
                If ultima < primeira Then Return vazio

                Dim linhas As New List(Of TableRow)(ultima - primeira + 1)
                For r = primeira To ultima
                    ' O acumulador vira o da LINHA: zera antes, colhe depois.
                    Fabricadas = 0
                    Dim linha As New TableRow With {
                        .EntryId = ComoEntryId(bruto(r, ColEntryId)),
                        .Subject = ComoTexto(bruto(r, ColSubject)),
                        .SenderName = ComoTexto(bruto(r, ColSender)),
                        .ReceivedTime = ComoData(bruto(r, ColRecebido)),
                        .SizeBytes = ComoInteiro(bruto(r, ColTamanho)),
                        .IsUnread = ComoBooleano(bruto(r, ColNaoLida)),
                        .MessageClass = ComoTexto(bruto(r, ColClasse)),
                        .HasAttachments = ComoBooleano(bruto(r, ColAnexo))
                    }
                    linha.Fabricadas = Fabricadas
                    linhas.Add(linha)
                Next
                Return linhas
            End Function

            Public Sub Fechar() Implements IRowSource(Of TableRow).Fechar
                ComHelpers.Release(_table)
                _table = Nothing
            End Sub

            ''' <summary>
            ''' Instante da linha, para o algoritmo agrupar empate.
            '''
            ''' Truncado ao SEGUNDO porque é a granularidade do filtro DASL:
            ''' agrupar mais fino que o filtro faria a fronteira devolver
            ''' linhas que o algoritmo julga de outro grupo, e a paginação
            ''' andaria em círculo.
            ''' </summary>
            Public Function InstanteDe(linha As TableRow) As DateTimeOffset _
                Implements IRowSource(Of TableRow).InstanteDe

                If Not linha.ReceivedTime.HasValue Then Return DateTimeOffset.MinValue
                Dim q = linha.ReceivedTime.Value
                Return New DateTimeOffset(q.Year, q.Month, q.Day, q.Hour, q.Minute, q.Second, q.Offset)
            End Function

            Public Function ChaveDe(linha As TableRow) As String _
                Implements IRowSource(Of TableRow).ChaveDe
                Return linha.EntryId
            End Function

            ''' <summary>
            ''' Acumulador de <b>uma linha</b>, enquanto ela é convertida.
            '''
            ''' Zerado antes de cada linha e colhido em
            ''' <see cref="TableRow.Fabricadas"/> logo depois — ver o comentário
            ''' lá para os dois defeitos que este campo já teve enquanto foi
            ''' contador de lote e de página.
            ''' </summary>
            Friend Fabricadas As Integer

            ''' <summary>
            ''' Construtor sem pasta, <b>só para a suíte</b>.
            '''
            ''' As conversões são lógica pura e não precisam de COM; o resto da
            ''' classe precisa. Sem esta porta, provar que o contador conta
            ''' exigiria um Outlook aberto com uma célula nula dentro — que é
            ''' justamente o que a medição de 28/08 mostrou não existir nesta
            ''' caixa.
            '''
            ''' Um contador que só pode ser exercitado por um estado que não se
            ''' consegue produzir é um contador sem prova.
            ''' </summary>
            Friend Sub New()
            End Sub

            ''' <summary>
            ''' Texto ausente vira vazio — e conta.
            '''
            ''' Vazio e ausente são coisas diferentes: um assunto em branco é uma
            ''' mensagem sem assunto; um <c>Nothing</c> é o provedor não tendo
            ''' respondido. Colapsar os dois é o que este contador denuncia.
            ''' </summary>
            Friend Function ComoTexto(valor As Object) As String
                Dim s = TryCast(valor, String)
                If s Is Nothing Then
                    Fabricadas += 1
                    Return ""
                End If
                Return s
            End Function

            ''' <summary>
            ''' A coluna de EntryID longo e PT_BINARY (0x..0102), entao volta como
            ''' Byte(). O MailItem.EntryID e hex MAIUSCULO, e a comparacao de
            ''' ItemKey e textual: hex minusculo aqui daria chave que nunca casa,
            ''' com o mesmo sintoma silencioso.
            ''' </summary>
            Friend Function ComoEntryId(valor As Object) As String
                Dim bytes = TryCast(valor, Byte())
                If bytes IsNot Nothing Then Return Convert.ToHexString(bytes)
                Dim s = TryCast(valor, String)
                If s Is Nothing Then
                    Fabricadas += 1
                    Return ""
                End If
                Return s
            End Function

            ''' <summary>Tamanho ausente ou ilegível vira <c>0</c> — e conta.</summary>
            Friend Function ComoInteiro(valor As Object) As Integer
                If valor Is Nothing Then
                    Fabricadas += 1
                    Return 0
                End If
                Try
                    Return Convert.ToInt32(valor, CultureInfo.InvariantCulture)
                Catch
                    Fabricadas += 1
                    Return 0
                End Try
            End Function

            ''' <summary>
            ''' Booleano ausente ou ilegível vira <c>False</c> — e conta.
            '''
            ''' Este é o pior dos três: <c>False</c> em "não lida" e em "tem
            ''' anexo" são afirmações que o usuário lê como fato. O
            ''' <c>MailSummary</c> já não tem <c>IsProtected</c> justamente para
            ''' não afirmar "não é protegida" sem ter medido; aqui a mesma
            ''' afirmação escapava por dentro da conversão.
            ''' </summary>
            Friend Function ComoBooleano(valor As Object) As Boolean
                If valor Is Nothing Then
                    Fabricadas += 1
                    Return False
                End If
                Try
                    Return Convert.ToBoolean(valor, CultureInfo.InvariantCulture)
                Catch
                    Fabricadas += 1
                    Return False
                End Try
            End Function

            Private Shared Function ComoData(valor As Object) As DateTimeOffset?
                If valor Is Nothing Then Return Nothing
                Try
                    Dim quando = Convert.ToDateTime(valor, CultureInfo.InvariantCulture)
                    ' Hora LOCAL — ver o comentário da classe.
                    Return New DateTimeOffset(DateTime.SpecifyKind(quando, DateTimeKind.Local))
                Catch
                    Return Nothing
                End Try
            End Function

        End Class

    End Module

End Namespace

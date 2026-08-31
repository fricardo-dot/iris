Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Assist
Imports Iris.Cache
Imports Iris.Model
Imports Iris.Sync

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>Classificar uma pasta: o que se manda, em que lotes, e o que sobra
    ''' quando um lote não confere.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTA CLASSE DECIDE</b>
    '''
    ''' As decisões de <i>quem entra</i> e <i>o que acontece quando dá errado</i>.
    ''' Ela não sabe o que é COM, não sabe o que é HTTP e não monta envelope: as
    ''' duas bordas entram por delegate, e é isso que a torna testável sem
    ''' Outlook e sem rede.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>SÓ O QUE AINDA NÃO TEM RÓTULO NESTA GERAÇÃO</b>
    '''
    ''' O rótulo pende de (encarnação, geração). Uma varredura nova republica a
    ''' pasta, e <b>tudo</b> volta a não ter rótulo — o que é correto: os corpos
    ''' podem ter mudado, e um rótulo de junho apresentado como atual em agosto é
    ''' exatamente o que a Fase 5 existe para impedir.
    '''
    ''' O efeito prático é que reclassificar depois de varrer custa a pasta
    ''' inteira. Está declarado aqui porque é caro e não é acidente.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>LOTE QUE NÃO CONFERE NÃO DERRUBA A PASTA</b>
    '''
    ''' Ele é contado e a passagem segue. A alternativa — parar tudo — daria a um
    ''' único e-mail hostil o poder de impedir a classificação da caixa inteira:
    ''' basta ele cair no primeiro lote. E lote que não confere <b>não grava
    ''' nada</b>, então seguir não contamina.
    '''
    ''' <b>Não há retentativa.</b> Classificação é leitura e leitura repete, mas
    ''' repetir aqui, dentro da mesma passagem, seria repetir com o mesmo
    ''' conteúdo hostil e obter a mesma recusa — pagando duas vezes. Quem quiser
    ''' tentar de novo roda a passagem de novo, e aí os lotes são outros.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>GRAVA A CADA LOTE, E NÃO NO FIM</b>
    '''
    ''' Duzentas mensagens classificadas e perdidas porque a de número duzentos e
    ''' um derrubou a passagem seria pagar o custo e ficar sem o resultado. Cada
    ''' lote que confere entra sozinho, na sua transação.
    ''' </summary>
    Public NotInheritable Class ClassificarUmaPasta

        ''' <summary>
        ''' <b>Quantas mensagens por lote.</b> Bem abaixo do teto de 200 do
        ''' <see cref="LoteDeClassificacao"/>, e de propósito: quanto maior o
        ''' lote, mais corpos hostis dividem o mesmo contexto do modelo, e é
        ''' nesse compartilhamento que mora o ataque que nenhuma validação de
        ''' formato pega. Vinte é o meio-termo escolhido — não medido, porque
        ''' ainda não há provedor escolhido.
        ''' </summary>
        Public Const PorLote As Integer = 20

        ''' <summary>
        ''' O que a borda do provider tem de devolver: uma parte por ficha, com o
        ''' corpo. <c>Nothing</c> ou lista vazia fazem o lote ser pulado.
        ''' </summary>
        Public Delegate Function Conteudo(
            pedidos As IReadOnlyList(Of PedidoDeParte)) As IReadOnlyList(Of MessagePart)

        ''' <summary>
        ''' O que a borda do transporte tem de devolver: o texto cru da resposta.
        ''' <c>Nothing</c> vale como lote recusado.
        ''' </summary>
        Public Delegate Function Envio(instrucao As String,
                                       partes As IReadOnlyList(Of MessagePart)) As String

        Private ReadOnly _acervo As AcervoDeTodasAsPastas
        Private ReadOnly _cache As RotulosNoCache

        Public Sub New(acervo As AcervoDeTodasAsPastas, cache As RotulosNoCache)
            If acervo Is Nothing Then Throw New ArgumentNullException(NameOf(acervo))
            If cache Is Nothing Then Throw New ArgumentNullException(NameOf(cache))
            _acervo = acervo
            _cache = cache
        End Sub

        ''' <summary>
        ''' Passa uma vez pela pasta.
        ''' </summary>
        ''' <param name="regras">
        ''' O que o dono escreveu. Acima do teto o lote é recusado — e aí a
        ''' passagem inteira devolve <see cref="MotivoDaClassificacao.RegrasDemais"/>
        ''' <b>sem</b> mandar nada: a alternativa seria classificar a caixa dele
        ''' com parte das regras, e ele descobriria pelo resultado.
        ''' </param>
        ''' <param name="quando">
        ''' O instante da observação. Entra como parâmetro porque o resultado é
        ''' gravado com ele, e um relógio lido lá dentro não seria conferível.
        ''' </param>
        Public Function Passar(pasta As Long,
                               regras As IReadOnlyList(Of String),
                               ativacao As String,
                               quando As DateTimeOffset,
                               conteudo As Conteudo,
                               envio As Envio) As ResultadoDaClassificacao

            If conteudo Is Nothing OrElse envio Is Nothing Then
                Return ResultadoDaClassificacao.Parou(MotivoDaClassificacao.SemAsBordas)
            End If

            Dim daPasta = _acervo.Pastas.FirstOrDefault(
                Function(p) p.Chave = pasta)
            If daPasta Is Nothing OrElse Not daPasta.Manifesto.GenerationKey.HasValue Then
                Return ResultadoDaClassificacao.Parou(MotivoDaClassificacao.PastaNaoVarrida)
            End If

            Dim geracao = daPasta.Manifesto.GenerationKey.Value
            Dim jaTem = _cache.Publicados(daPasta.Chave)

            ' SO O QUE ESTA PRESENTE E AINDA NAO TEM ROTULO NESTA GERACAO.
            Dim aFazer = daPasta.Manifesto.Items.
                         Where(Function(i) i.Presence = PresenceState.Presente).
                         Where(Function(i) Not jaTem.ContainsKey(i.ProviderEntryId)).
                         Select(Function(i) New ItemKey(i.ProviderEntryId, daPasta.Store)).
                         ToList()

            If aFazer.Count = 0 Then
                Return ResultadoDaClassificacao.Nada(MotivoDaClassificacao.NadaAFazer)
            End If

            Dim r As New Acumulador(aFazer.Count)

            For inicio = 0 To aFazer.Count - 1 Step PorLote
                Dim chavesDoLote = aFazer.Skip(inicio).Take(PorLote).ToList()
                Dim montado = LoteDeClassificacao.Preparar(chavesDoLote, regras)

                If montado Is Nothing Then
                    ' O UNICO JEITO DE UM LOTE DESTE TAMANHO NAO SER MONTADO e
                    ' regra demais -- as chaves vem do acervo, sao unicas e nao
                    ' sao nulas. Isso nao e falha de um lote: e a configuracao
                    ' do dono, e vale para todos. Parar aqui evita mandar
                    ' dezenas de lotes que vao ser recusados um a um.
                    Return ResultadoDaClassificacao.Parou(MotivoDaClassificacao.RegrasDemais)
                End If

                Dim partes = conteudo(Pedidos(montado, chavesDoLote))
                If partes Is Nothing OrElse partes.Count = 0 Then
                    r.LoteSemConteudo(chavesDoLote.Count)
                    Continue For
                End If

                Dim conferido = montado.Conferir(envio(montado.Instrucao(), partes))
                If Not conferido.IdentidadesConferem Then
                    r.LoteRecusado(chavesDoLote.Count, conferido.Motivo)
                    Continue For
                End If

                Gravar(daPasta.Chave, geracao, ativacao, quando,
                       conferido, regras, r)
            Next

            Return r.Fechar()
        End Function

        ''' <summary>
        ''' A lista que a borda do provider recebe: a chave para ela ir buscar, e
        ''' a ficha para o envelope carregar.
        '''
        ''' <b>A ficha é cunhada aqui e vai junto</b> justamente para a borda não
        ''' ter de inventá-la — se ela inventasse, a resposta seria conferida
        ''' contra fichas que este lote não conhece, e o lote inteiro cairia.
        ''' </summary>
        Private Shared Function Pedidos(montado As LoteDeClassificacao,
                                        chaves As IReadOnlyList(Of ItemKey)) _
                                        As IReadOnlyList(Of PedidoDeParte)
            Return chaves.Select(Function(c) New PedidoDeParte(c, montado.FichaDe(c))).ToList()
        End Function

        ''' <summary>
        ''' Grava o que este lote produziu.
        '''
        ''' <b>Item sem rótulo não entra</b> — ele veio com uma palavra que não
        ''' existe, e gravar a ausência como se fosse classificação faria a
        ''' passagem seguinte pular a mensagem para sempre. Ficando de fora, ela
        ''' volta na próxima.
        ''' </summary>
        Private Sub Gravar(pasta As Long, geracao As Long, ativacao As String,
                           quando As DateTimeOffset,
                           conferido As LoteClassificado,
                           regras As IReadOnlyList(Of String),
                           r As Acumulador)

            Dim rotulos As New Dictionary(Of String, String)(StringComparer.Ordinal)
            Dim confiancas As New Dictionary(Of String, Double?)(StringComparer.Ordinal)

            For Each par In conferido.Rotulos
                rotulos(par.Key.EntryId) = LoteDeClassificacao.NomeDoRotulo(par.Value)
                Dim c As Double
                If conferido.Confiancas.TryGetValue(par.Key, c) Then
                    confiancas(par.Key.EntryId) = c
                End If
            Next

            ' AS REGRAS SO VIAJAM SE HOUVE PERGUNTA. Nothing quer dizer "nao
            ' havia regras nesta varredura"; um dicionario vazio diria "havia, e
            ' nenhuma mensagem casou", que e uma afirmacao diferente.
            Dim casadas As Dictionary(Of String, IReadOnlyList(Of String)) = Nothing
            Dim temRegra = regras IsNot Nothing AndAlso
                           regras.Any(Function(x) Not String.IsNullOrWhiteSpace(x))
            If temRegra Then
                casadas = New Dictionary(Of String, IReadOnlyList(Of String))(StringComparer.Ordinal)
                Dim mudas = New HashSet(Of ItemKey)(conferido.SemRegras)
                For Each par In conferido.Rotulos
                    ' ITEM CUJA RESPOSTA SOBRE REGRAS NAO DEU PARA USAR fica de
                    ' fora do mapa, e assim o cache grava NULO nele -- que e
                    ' "ninguem respondeu", e nao "nenhuma casou".
                    If mudas.Contains(par.Key) Then Continue For
                    Dim minhas As IReadOnlyList(Of String) = Nothing
                    If Not conferido.RegrasCasadas.TryGetValue(par.Key, minhas) Then
                        minhas = Array.Empty(Of String)()
                    End If
                    casadas(par.Key.EntryId) = minhas
                Next
            End If

            Dim feito = _cache.Gravar(pasta, geracao, ativacao, quando,
                                      rotulos, confiancas, casadas)
            r.LoteGravado(feito, conferido)
        End Sub


        ''' <summary>Junta as contas dos lotes. Existe para o laço não inchar.</summary>
        Private NotInheritable Class Acumulador
            Private ReadOnly _pedidos As Integer
            Private _classificados As Integer
            Private _semRotulo As Integer
            Private _semRegras As Integer
            Private _foraDaPasta As Integer
            Private _lotesRecusados As Integer
            Private _naoClassificados As Integer
            Private _primeiraRecusa As String = ""
            Private _geracaoErrada As Boolean

            Public Sub New(pedidos As Integer)
                _pedidos = pedidos
            End Sub

            Public Sub LoteRecusado(quantos As Integer, motivo As String)
                _lotesRecusados += 1
                _naoClassificados += quantos
                If _primeiraRecusa.Length = 0 Then _primeiraRecusa = If(motivo, "")
            End Sub

            Public Sub LoteSemConteudo(quantos As Integer)
                _naoClassificados += quantos
            End Sub

            Public Sub LoteGravado(feito As ResultadoDaGravacao, conferido As LoteClassificado)
                If Not feito.Gravou Then
                    _geracaoErrada = True
                    Return
                End If
                _classificados += feito.Entraram
                _foraDaPasta += feito.ForaDaPasta
                _semRotulo += conferido.SemRotulo.Count
                _semRegras += conferido.SemRegras.Count
            End Sub

            Public Function Fechar() As ResultadoDaClassificacao
                ' GERACAO ERRADA NO MEIO DA PASSAGEM quer dizer que a pasta foi
                ' revarrida enquanto isto rodava. O que ja entrou entrou na
                ' geracao certa; o resto nao vale, e a passagem inteira e
                ' declarada obsoleta em vez de devolver contas de duas geracoes
                ' misturadas.
                If _geracaoErrada Then
                    Return ResultadoDaClassificacao.Parou(MotivoDaClassificacao.PastaRevarrida)
                End If

                Return New ResultadoDaClassificacao(
                    MotivoDaClassificacao.Passou, _pedidos, _classificados,
                    _semRotulo, _semRegras, _foraDaPasta,
                    _lotesRecusados, _naoClassificados, _primeiraRecusa)
            End Function
        End Class

    End Class

    ''' <summary>Uma mensagem que a borda do provider precisa ler, com a ficha dela.</summary>
    Public NotInheritable Class PedidoDeParte
        Public ReadOnly Property Chave As ItemKey
        Public ReadOnly Property Ficha As String

        Friend Sub New(chave As ItemKey, ficha As String)
            Me.Chave = chave
            Me.Ficha = If(ficha, "")
        End Sub
    End Class

    Public Enum MotivoDaClassificacao
        ''' <summary>A passagem rodou. Isso não quer dizer que classificou tudo.</summary>
        Passou = 0
        ''' <summary>Nenhuma mensagem sem rótulo nesta geração.</summary>
        NadaAFazer
        ''' <summary>A pasta não tem geração publicada. Não é pasta vazia.</summary>
        PastaNaoVarrida
        ''' <summary>O dono escreveu regras acima do teto. Nada foi mandado.</summary>
        RegrasDemais
        ''' <summary>A pasta foi varrida de novo no meio. O que sobrou não vale.</summary>
        PastaRevarrida
        ''' <summary>Faltou uma das bordas. Erro de quem montou.</summary>
        SemAsBordas
    End Enum

    ''' <summary>
    ''' O que a passagem fez.
    '''
    ''' <b><see cref="Classificados"/> menor que <see cref="Pedidos"/> é normal</b>,
    ''' e as três razões aparecem separadas de propósito: lote recusado, rótulo
    ''' inventado e mensagem que saiu da pasta são problemas diferentes, e
    ''' somá-los num "faltaram 30" esconderia qual deles está acontecendo.
    ''' </summary>
    Public NotInheritable Class ResultadoDaClassificacao
        Public ReadOnly Property Motivo As MotivoDaClassificacao
        ''' <summary>Quantas mensagens entraram na passagem.</summary>
        Public ReadOnly Property Pedidos As Integer
        ''' <summary>Quantas ficaram com rótulo gravado.</summary>
        Public ReadOnly Property Classificados As Integer
        ''' <summary>Vieram com um rótulo que não existe. Voltam na próxima passagem.</summary>
        Public ReadOnly Property SemRotulo As Integer
        ''' <summary>
        ''' Tiveram rótulo, mas a resposta sobre as regras do dono não deu para
        ''' usar. O rótulo valeu; a pergunta dele ficou sem resposta.
        ''' </summary>
        Public ReadOnly Property SemRegras As Integer
        ''' <summary>Saíram da pasta entre a classificação e a gravação.</summary>
        Public ReadOnly Property ForaDaPasta As Integer
        Public ReadOnly Property LotesRecusados As Integer
        ''' <summary>Quantas mensagens estavam nos lotes que não deram em nada.</summary>
        Public ReadOnly Property NaoClassificados As Integer
        ''' <summary>
        ''' O motivo da <b>primeira</b> recusa. Um só: guardar todos encheria a
        ''' tela com variações da mesma frase, e o primeiro basta para o dono
        ''' saber se foi forma, identidade ou o controle.
        ''' </summary>
        Public ReadOnly Property PrimeiraRecusa As String

        Friend Sub New(motivo As MotivoDaClassificacao, pedidos As Integer,
                       classificados As Integer, semRotulo As Integer,
                       semRegras As Integer, foraDaPasta As Integer,
                       lotesRecusados As Integer, naoClassificados As Integer,
                       primeiraRecusa As String)
            Me.Motivo = motivo
            Me.Pedidos = pedidos
            Me.Classificados = classificados
            Me.SemRotulo = semRotulo
            Me.SemRegras = semRegras
            Me.ForaDaPasta = foraDaPasta
            Me.LotesRecusados = lotesRecusados
            Me.NaoClassificados = naoClassificados
            Me.PrimeiraRecusa = If(primeiraRecusa, "")
        End Sub

        Friend Shared Function Parou(motivo As MotivoDaClassificacao) As ResultadoDaClassificacao
            Return New ResultadoDaClassificacao(motivo, 0, 0, 0, 0, 0, 0, 0, "")
        End Function

        Friend Shared Function Nada(motivo As MotivoDaClassificacao) As ResultadoDaClassificacao
            Return New ResultadoDaClassificacao(motivo, 0, 0, 0, 0, 0, 0, 0, "")
        End Function
    End Class

End Namespace

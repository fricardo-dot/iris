Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
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
            pedidos As IReadOnlyList(Of PedidoDeParte),
            ct As CancellationToken) As IReadOnlyList(Of MessagePart)

        ''' <summary>
        ''' O que a borda do transporte tem de devolver.
        '''
        ''' <b>Não é uma string.</b> Era, e a string apagava a diferença entre
        ''' <i>não saiu</i> e <i>pode ter saído</i>: todo insucesso virava
        ''' <c>Nothing</c>, a passagem contava "lote recusado", e a tela dizia isso
        ''' sobre um lote que talvez tivesse voado. Ver <see cref="RespostaDoLote"/>.
        ''' </summary>
        Public Delegate Function Envio(instrucao As String,
                                       partes As IReadOnlyList(Of MessagePart),
                                       ct As CancellationToken) As RespostaDoLote

        ''' <summary>
        ''' <b>Uma passagem por vez, no programa inteiro.</b>
        '''
        ''' Duas passagens simultâneas sobre o mesmo cache fariam duas coisas
        ''' ruins de uma vez. A primeira é técnica: <c>SqliteConnection</c> não tem
        ''' contrato de uso simultâneo, e uma leitura aberta enquanto a outra abre
        ''' <c>BEGIN IMMEDIATE</c> é erro em tempo de execução — o WAL coordena
        ''' <i>conexões</i>, não torna uma conexão reentrante.
        '''
        ''' A segunda é pior e é anterior ao banco: as duas leem a mesma lista de
        ''' pendentes e <b>mandam os mesmos corpos</b> ao provedor antes de
        ''' qualquer disputa no SQLite. Divulgação duplicada não se desfaz com um
        ''' rollback.
        '''
        ''' A porta é estática porque o recurso disputado é o arquivo, e não esta
        ''' instância: duas <c>ClassificarUmaPasta</c> diferentes sobre o mesmo
        ''' cache disputam do mesmo jeito. Achado por revisão externa em
        ''' 31/08/2026.
        '''
        ''' <b>A porta não é segurada durante a rede</b> — ela é: a passagem
        ''' inteira roda dentro dela, inclusive as chamadas ao provedor. É
        ''' deliberado: o que se quer impedir é justamente a segunda passagem
        ''' mandar os mesmos corpos, e uma porta solta durante a rede não impede
        ''' nada disso. O preço é que a segunda chamada é <i>recusada</i>, e não
        ''' enfileirada — recusar diz a quem chamou o que aconteceu; enfileirar
        ''' esconderia uma espera de minutos atrás de uma chamada que parece
        ''' síncrona.
        ''' </summary>
        ' A PORTA E POR CACHE, e nao por processo.
        '
        ' Um sinalizador unico recusava trabalho legitimo: duas janelas sobre
        ' bancos diferentes nao disputam conexao, arquivo, pasta nem mensagem, e
        ' a segunda levava JaEstaRodando enquanto a primeira esperava minutos de
        ' rede. O recurso disputado e o CACHE. Achado por revisao externa em
        ' 01/09/2026.
        '
        ' Tabela fraca: uma passagem que morra sem soltar nao segura o cache vivo
        ' na memoria por causa da porta.
        Private Shared ReadOnly _porta As New Object()
        Private Shared ReadOnly _rodandoEm As New Runtime.CompilerServices.ConditionalWeakTable(Of Object, Object)()

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
                               envio As Envio,
                               Optional ct As CancellationToken = Nothing) _
                               As ResultadoDaClassificacao

            If conteudo Is Nothing OrElse envio Is Nothing Then
                Return ResultadoDaClassificacao.Parou(MotivoDaClassificacao.SemAsBordas)
            End If

            Dim dono = _cache.Cache
            SyncLock _porta
                Dim ja As Object = Nothing
                If _rodandoEm.TryGetValue(dono, ja) Then
                    Return ResultadoDaClassificacao.Parou(MotivoDaClassificacao.JaEstaRodando)
                End If
                _rodandoEm.Add(dono, dono)
            End SyncLock

            Try
                Return Correr(pasta, regras, ativacao, quando, conteudo, envio, ct)
            Finally
                SyncLock _porta
                    _rodandoEm.Remove(dono)
                End SyncLock
            End Try
        End Function

        Private Function Correr(pasta As Long,
                                regras As IReadOnlyList(Of String),
                                ativacao As String,
                                quando As DateTimeOffset,
                                conteudo As Conteudo,
                                envio As Envio,
                                ct As CancellationToken) As ResultadoDaClassificacao

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
                ' O PEDIDO DE PARADA E OLHADO ENTRE LOTES.
                '
                ' Uma passagem sobre uma pasta grande e vinte idas a rede, e
                ' nenhuma delas era interrompivel: fechar a janela, trocar de
                ' pasta ou perder o Outlook nao impedia os lotes seguintes de
                ' saírem. Parar entre lotes -- e nao no meio de um -- e o unico
                ' lugar onde parar nao deixa duvida sobre o que saiu.
                '
                ' O que ja foi gravado CONTINUA valendo, e as contagens vao
                ' junto no desfecho. Achado por revisao externa em 01/09/2026.
                If ct.IsCancellationRequested Then Return r.Fechar(MotivoDaClassificacao.Parada)

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

                Dim lidas = conteudo(Pedidos(montado, chavesDoLote), ct)
                If lidas Is Nothing OrElse lidas.Count = 0 Then
                    r.LoteSemConteudo(chavesDoLote.Count)
                    Continue For
                End If

                ' O CONTROLE VAI JUNTO, E QUEM O ACRESCENTA E DAQUI.
                '
                ' Ele nao estava sendo mandado: a instrucao anunciava a ficha do
                ' controle e nenhuma mensagem com essa ficha ia no pedido, entao o
                ' modelo fabricava a linha a partir da propria instrucao. O rotulo
                ' certo nao provava nada, porque o controle nao participava do
                ' conjunto que um "classifique todas" atinge. Achado por revisao
                ' externa em 31/08/2026.
                '
                ' A borda nao o monta porque a borda le o Outlook, e o controle nao
                ' esta no Outlook.
                Dim partes As New List(Of MessagePart)(lidas)
                partes.Add(montado.ParteDoControle())

                ' A GERACAO E RECONFERIDA AQUI, ANTES DE DIVULGAR.
                '
                ' A corrida com a varredura so era percebida no Gravar -- isto e,
                ' depois de os corpos terem ido para o provedor. A gravacao era
                ' recusada corretamente, e o conteudo ja tinha saido: o retrato em
                ' que a passagem se baseou fora substituido, e ninguem soube antes
                ' de pagar. Achado por revisao externa em 01/09/2026.
                '
                ' Entre esta conferencia e o envio ainda ha uma fresta, e ela nao
                ' se fecha sem parar a varredura -- mas ela e de milissegundos, e
                ' nao dos minutos que os lotes anteriores levaram.
                If Not GeracaoAindaVale(daPasta.Chave, geracao) Then
                    r.Revarrida(chavesDoLote.Count)
                    Exit For
                End If

                ' QUAIS SAIRAM DE VERDADE. O pipeline da borda recusa item a
                ' item -- anexo, corpo pela metade, referencia embutida --, e
                ' cobrar resposta sobre uma mensagem que o provedor nunca viu
                ' derrubava o lote inteiro, sempre no mesmo lote.
                Dim enviadas = lidas.Where(Function(p) p IsNot Nothing AndAlso
                                                       p.Item IsNot Nothing).
                                     Select(Function(p) p.Ficha).ToList()

                ' O QUE NAO SAIU ENTRA NA CONTA AQUI. Sem isto, Pedidos e
                ' Classificados divergiam e nada explicava a diferenca.
                r.RecusadasPeloConteudo(chavesDoLote.Count - enviadas.Count)

                Dim veio = envio(montado.Instrucao(), partes, ct)

                ' AMBIGUO NAO E RECUSADO, E A PASSAGEM PARA AQUI.
                '
                ' "Pode ter saido e nao se sabe o que aconteceu" e o desfecho que
                ' este projeto trata com mais cuidado que qualquer outro, e ele
                ' estava sendo dobrado em "lote recusado" -- que quer dizer NADA
                ' SAIU. A tela dizia a coisa oposta do que aconteceu, e o diario,
                ' que sabia a verdade, ninguem le.
                '
                ' E para: uma resposta incerta quase sempre quer dizer que a rede
                ' ou o provedor estao num estado ruim, e seguir mandando gasta
                ' dinheiro e divulga mais enquanto o dono ainda nao sabe do
                ' primeiro. Nao e retry -- os lotes seguintes sao outras mensagens
                ' --, e parar e o que da a ele a chance de olhar. Achado por
                ' revisao externa em 01/09/2026.
                If veio IsNot Nothing AndAlso veio.Incerta Then
                    r.LoteIncerto(chavesDoLote.Count, veio.Motivo)
                    Return r.Fechar(MotivoDaClassificacao.Incerta)
                End If

                Dim conferido = montado.Conferir(If(veio Is Nothing, Nothing, veio.Texto),
                                                 enviadas)
                If Not conferido.IdentidadesConferem Then
                    ' O MOTIVO DA BORDA GANHA DO MOTIVO DO PARSER. "A resposta nao
                    ' e JSON valido" e verdade sobre um Nothing e nao explica nada;
                    ' "o portao negou" explica.
                    Dim porque = If(veio IsNot Nothing AndAlso veio.Motivo.Length > 0,
                                    veio.Motivo, conferido.Motivo)
                    r.LoteRecusado(chavesDoLote.Count, porque)
                    Continue For
                End If

                Gravar(daPasta.Chave, geracao, ativacao, quando,
                       conferido, regras, r)

                ' REVARRIDA NO MEIO PARA O LACO AQUI, e nao no fim.
                '
                ' Antes o laco seguia ate o ultimo lote e so entao declarava a
                ' passagem obsoleta: os corpos dos lotes seguintes eram lidos e
                ' MANDADOS, e todas as gravacoes eram recusadas do mesmo jeito.
                ' Custo e divulgacao depois de a passagem ja saber que nada mais
                ' pode valer. Achado por revisao externa em 31/08/2026.
                If r.Obsoleta Then Exit For
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
        ''' <summary>
        ''' A geração em que esta passagem se baseou ainda é a publicada?
        '''
        ''' Lida do acervo já recarregado, e não do banco: quem recarrega é o
        ''' dreno, e ir ao SQLite a cada lote acrescentaria uma leitura por lote
        ''' para responder o que o retrato em memória já sabe.
        ''' </summary>
        Private Function GeracaoAindaVale(pasta As Long, geracao As Long) As Boolean
            Dim agora = _acervo.Pastas.FirstOrDefault(Function(p) p.Chave = pasta)
            If agora Is Nothing OrElse Not agora.Manifesto.GenerationKey.HasValue Then
                Return False
            End If
            Return agora.Manifesto.GenerationKey.Value = geracao
        End Function

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
            Private _recusadasPeloConteudo As Integer
            Private _lotesIncertos As Integer
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

            ''' <summary>
            ''' Mensagens que a borda leu e o <b>pipeline recusou</b> — anexo, corpo
            ''' pela metade, referência embutida, HTML ilegível.
            '''
            ''' Conta separada porque é a única das quatro que o dono pode
            ''' <i>entender</i>: "três têm anexo" é uma frase acionável, e some-la a
            ''' "o lote foi recusado" produziria um número que não sugere nada.
            ''' </summary>
            Public Sub RecusadasPeloConteudo(quantas As Integer)
                If quantas <= 0 Then Return
                _recusadasPeloConteudo += quantas
                _naoClassificados += quantas
            End Sub

            ''' <summary>
            ''' A pasta foi republicada, e este lote não chegou a sair. As mensagens
            ''' dele entram como não classificadas — elas não foram, e some-las ao
            ''' que ficou de fora por outro motivo seria misturar duas coisas.
            ''' </summary>
            ''' <summary>
            ''' <b>Um lote cujo desfecho não se sabe.</b> Conta separado de
            ''' <c>LotesRecusados</c> porque as duas afirmações são opostas: recusado
            ''' é <i>nada saiu</i>, incerto é <i>pode ter saído</i>.
            ''' </summary>
            Public Sub LoteIncerto(quantos As Integer, motivo As String)
                _lotesIncertos += 1
                _naoClassificados += quantos
                If _primeiraRecusa.Length = 0 Then _primeiraRecusa = If(motivo, "")
            End Sub

            Public Sub Revarrida(quantos As Integer)
                _geracaoErrada = True
                _naoClassificados += quantos
            End Sub

            ''' <summary>
            ''' A pasta foi revarrida no meio: nada mais vale, e o laço para.
            ''' </summary>
            Public ReadOnly Property Obsoleta As Boolean
                Get
                    Return _geracaoErrada
                End Get
            End Property

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

            ''' <summary>
            ''' <b>O desfecho, e ele nunca apaga o que já aconteceu.</b>
            '''
            ''' ------------------------------------------------------------------
            ''' <b>A REVARREDURA ZERAVA AS CONTAGENS</b>
            '''
            ''' O ramo da geração errada devolvia <c>Parou(PastaRevarrida)</c>, que é
            ''' <c>Pedidos=0, Classificados=0</c>. Cenário real: o primeiro lote grava
            ''' vinte rótulos, a pasta é republicada, e a passagem reporta que não fez
            ''' nada — sobre uma passagem que gravou vinte linhas no cache.
            '''
            ''' É a mesma família de defeito da cerca do transmissor: dizer que nada
            ''' aconteceu quando algo aconteceu. Aqui o custo é menor — ninguém
            ''' reenvia por causa disto —, mas a frase na tela mentia, e a próxima
            ''' passagem não repetiria o trabalho que a frase disse não ter sido
            ''' feito. Achado por revisão externa em 01/09/2026.
            '''
            ''' O motivo muda; as contagens vão junto sempre.
            ''' </summary>
            Public Function Fechar(Optional motivo As MotivoDaClassificacao =
                                       MotivoDaClassificacao.Passou) _
                                   As ResultadoDaClassificacao
                Dim qual = If(_geracaoErrada, MotivoDaClassificacao.PastaRevarrida, motivo)

                Return New ResultadoDaClassificacao(
                    qual, _pedidos, _classificados,
                    _semRotulo, _semRegras, _foraDaPasta,
                    _lotesRecusados, _naoClassificados, _primeiraRecusa,
                    _recusadasPeloConteudo, _lotesIncertos)
            End Function
        End Class

    End Class

    ''' <summary>
    ''' <b>O que a borda do transporte devolve — e o que ela não pode esconder.</b>
    '''
    ''' Era uma <c>String</c>, e <c>Nothing</c> significava "não deu". Isso dobrava
    ''' num só valor coisas que levam a ações opostas: <i>o portão negou</i> (nada
    ''' saiu, e o dono precisa assinar a ativação), <i>o provedor recusou</i> (nada
    ''' saiu, e o dono precisa olhar a credencial) e <i>a rede caiu depois do
    ''' primeiro byte</i> (<b>pode ter saído</b>, e o dono precisa saber disso).
    '''
    ''' A última é a que não podia estar aqui dentro. Dizer "lote recusado" sobre
    ''' ela é dizer que nada saiu quando algo pode ter saído — o defeito que este
    ''' projeto persegue desde o começo. Achado por revisão externa em 01/09/2026.
    ''' </summary>
    Public NotInheritable Class RespostaDoLote

        ''' <summary>O texto cru do modelo. Vazio quando não houve resposta.</summary>
        Public ReadOnly Property Texto As String

        ''' <summary>
        ''' <b>Pode ter saído, e não se sabe o que aconteceu.</b> Nunca é o mesmo
        ''' que recusado, e nunca é somado a ele.
        ''' </summary>
        Public ReadOnly Property Incerta As Boolean

        ''' <summary>Em português, para a tela. Vazio quando houve resposta.</summary>
        Public ReadOnly Property Motivo As String

        Private Sub New(texto As String, incerta As Boolean, motivo As String)
            Me.Texto = If(texto, "")
            Me.Incerta = incerta
            Me.Motivo = If(motivo, "")
        End Sub

        Public Shared Function Respondeu(texto As String) As RespostaDoLote
            Return New RespostaDoLote(texto, False, "")
        End Function

        ''' <summary>Nada saiu, e sabe-se por quê.</summary>
        Public Shared Function Recusada(motivo As String) As RespostaDoLote
            Return New RespostaDoLote("", False, motivo)
        End Function

        ''' <summary>Pode ter saído. Ver <see cref="Incerta"/>.</summary>
        Public Shared Function NaoSeSabe(motivo As String) As RespostaDoLote
            Return New RespostaDoLote("", True, motivo)
        End Function

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
        ''' <summary>
        ''' Já havia uma passagem em andamento. <b>Não é erro</b>: é a recusa que
        ''' impede duas passagens de mandarem os mesmos corpos ao provedor.
        ''' </summary>
        JaEstaRodando

        ''' <summary>
        ''' <b>Alguém pediu para parar</b> — hoje, o fechamento da janela. Os lotes
        ''' que já rodaram valem, e as contagens dizem quantos foram. <b>Não é
        ''' erro</b>, e não é "nada aconteceu".
        ''' </summary>
        Parada

        ''' <summary>
        ''' <b>Um lote pode ter saído e não se sabe o que aconteceu com ele.</b>
        '''
        ''' A passagem para aqui, de propósito: seguir mandando gastaria dinheiro e
        ''' divulgaria mais enquanto o dono ainda não sabe do primeiro. O diário tem
        ''' a linha; a tela tem de dizer, porque ninguém lê o diário.
        ''' </summary>
        Incerta
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

        ''' <summary>
        ''' Mensagens que a borda leu e o pipeline recusou — anexo, corpo pela
        ''' metade, referência embutida. Elas <b>não saíram da máquina</b>.
        ''' </summary>
        Public ReadOnly Property RecusadasPeloConteudo As Integer

        ''' <summary>
        ''' Lotes cujo desfecho <b>não se sabe</b> — podem ter saído. O oposto de
        ''' <see cref="LotesRecusados"/>, e nunca somado a ele.
        ''' </summary>
        Public ReadOnly Property LotesIncertos As Integer

        Friend Sub New(motivo As MotivoDaClassificacao, pedidos As Integer,
                       classificados As Integer, semRotulo As Integer,
                       semRegras As Integer, foraDaPasta As Integer,
                       lotesRecusados As Integer, naoClassificados As Integer,
                       primeiraRecusa As String,
                       Optional recusadasPeloConteudo As Integer = 0,
                       Optional lotesIncertos As Integer = 0)
            Me.Motivo = motivo
            Me.RecusadasPeloConteudo = recusadasPeloConteudo
            Me.LotesIncertos = lotesIncertos
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

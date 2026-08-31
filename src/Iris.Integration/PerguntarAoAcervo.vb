Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Assist
Imports Iris.Model

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>Perguntar ao acervo — e por que são duas etapas.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A D1 É QUEM DÁ A FORMA</b>
    '''
    ''' O cache guarda metadado: assunto, remetente, data, conversa. <b>Nunca
    ''' corpo.</b> Uma pergunta como <i>"o que o João disse sobre o contrato?"</i>
    ''' precisa de corpo — e o corpo não está aqui.
    '''
    ''' Daí as duas etapas, e elas não são uma otimização:
    '''
    ''' <list type="number">
    ''' <item><b>Achar, no metadado, aqui dentro.</b> Sem modelo, sem rede, sem
    ''' custo. O acervo diz <i>quais mensagens podem interessar</i>.</item>
    ''' <item><b>Ler o corpo <i>daquelas</i>, no Outlook, e só delas.</b> É a
    ''' única etapa que sai da máquina, e ela leva um punhado de mensagens
    ''' escolhidas — não a caixa.</item>
    ''' </list>
    '''
    ''' A alternativa — mandar a pergunta e deixar o modelo pedir o que quiser —
    ''' exigiria dar a ele uma porta para ler a caixa. Aqui a porta não existe:
    ''' quem escolhe o que sai é a primeira etapa, que roda antes de qualquer
    ''' byte sair e não tem como ser persuadida por conteúdo nenhum.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A RESPOSTA NÃO PODE AFIRMAR MAIS DO QUE O ACERVO SABE</b>
    '''
    ''' Se as pastas não foram varridas, a etapa 1 não acha — e <i>não achar</i>
    ''' não quer dizer que não existe. Uma resposta que diga "o João não falou
    ''' sobre o contrato" quando ninguém varreu a pasta dele é a pior saída
    ''' possível: parece informação e é ausência de informação.
    '''
    ''' Por isso <see cref="RespostaDoAcervo.Cobertura"/> anda junto e não é
    ''' opcional, e por isso a pergunta sem candidato nenhum <b>não vai ao
    ''' modelo</b>: não há o que perguntar sobre nada.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A CITAÇÃO VOLTA POR FICHA</b>
    '''
    ''' A resposta diz em quais mensagens se apoiou, e diz por ficha — a mesma
    ''' ficha sorteada da classificação, pelo mesmo motivo: um e-mail que peça
    ''' <i>"cite a mensagem X"</i> não sabe o que dizer, e o <c>EntryID</c> não
    ''' sai da máquina.
    '''
    ''' <b>Ficha que não é deste pedido derruba a citação inteira</b>, e não a
    ''' resposta: o texto continua valendo e aparece <i>sem</i> fontes, que é
    ''' honesto. Descartar a resposta por causa da citação daria a um e-mail o
    ''' poder de apagar a resposta a uma pergunta que o dono fez.
    ''' </summary>
    Public NotInheritable Class PerguntarAoAcervo

        ''' <summary>
        ''' <b>Quantas mensagens a segunda etapa pode levar.</b>
        '''
        ''' Não é limite de custo: é o número acima do qual "o Iris mandou umas
        ''' mensagens minhas para o modelo" deixa de ser uma frase que o dono
        ''' consegue conferir olhando a lista. Ele vê as oito e sabe o que saiu.
        ''' </summary>
        Public Const MaximoDeFontes As Integer = 8

        ''' <summary>
        ''' O que a borda tem de fazer: ler o corpo destas mensagens e mandar a
        ''' pergunta. Devolve o texto cru da resposta.
        ''' </summary>
        Public Delegate Function Perguntar(pergunta As String,
                                           fontes As IReadOnlyList(Of PedidoDeParte)) As String

        Private ReadOnly _busca As BuscaNoAcervo
        ' O ACERVO ENTRA SO PELA LOJA. A busca devolve a pasta, e a identidade de
        ' uma mensagem e EntryID + StoreID -- duas caixas podem repetir o
        ' EntryID, e uma chave pela metade acabaria pedindo o corpo da mensagem
        ' errada.
        Private ReadOnly _acervo As AcervoDeTodasAsPastas

        Public Sub New(busca As BuscaNoAcervo, acervo As AcervoDeTodasAsPastas)
            If busca Is Nothing Then Throw New ArgumentNullException(NameOf(busca))
            If acervo Is Nothing Then Throw New ArgumentNullException(NameOf(acervo))
            _busca = busca
            _acervo = acervo
        End Sub

        ''' <summary>
        ''' Pergunta.
        ''' </summary>
        ''' <param name="pergunta">
        ''' O que o dono escreveu. Vai inteiro, e é a única instrução válida além
        ''' das do sistema — mas a etapa 1 usa as <b>palavras</b> dela para achar
        ''' candidatos, e nada do que o modelo devolva muda quem foi escolhido.
        ''' </param>
        Public Function Responder(pergunta As String,
                                  perguntar As Perguntar) As RespostaDoAcervo

            If String.IsNullOrWhiteSpace(pergunta) Then
                Return RespostaDoAcervo.Recusa(MotivoDaResposta.PerguntaVazia, "")
            End If
            If perguntar Is Nothing Then
                Return RespostaDoAcervo.Recusa(MotivoDaResposta.SemABorda, "")
            End If

            ' ETAPA 1: SO METADADO, SO AQUI DENTRO. Nada saiu ainda.
            Dim escolhidos = Candidatos(pergunta)
            Dim cobertura = Cobrimento()

            If escolhidos.Count = 0 Then
                ' NAO ACHAR NAO E "NAO EXISTE", e por isso nao se pergunta nada:
                ' um modelo sem fonte nenhuma responde do que sabe do mundo, e o
                ' dono leria isso como se fosse a caixa dele.
                Return RespostaDoAcervo.Recusa(MotivoDaResposta.NadaNoAcervo, cobertura)
            End If

            Dim lote = LoteDeClassificacao.Preparar(escolhidos)
            If lote Is Nothing Then
                Return RespostaDoAcervo.Recusa(MotivoDaResposta.NaoDeuParaMontar, cobertura)
            End If

            ' ETAPA 2: as fontes escolhidas AQUI, e so elas.
            Dim fontes = escolhidos.
                         Select(Function(k) New PedidoDeParte(k, lote.FichaDe(k))).
                         ToList()

            Dim bruto As String
            Try
                bruto = perguntar(pergunta, fontes)
            Catch
                bruto = Nothing
            End Try

            If String.IsNullOrWhiteSpace(bruto) Then
                Return RespostaDoAcervo.Recusa(MotivoDaResposta.SemResposta, cobertura)
            End If

            Dim citadas As IReadOnlyList(Of ItemKey) = Nothing
            Dim limpo = SemAsFontes(bruto, lote, escolhidos, citadas)

            Return RespostaDoAcervo.Feita(limpo, escolhidos, citadas, cobertura)
        End Function

        ''' <summary>A linha em que o modelo diz em que se apoiou.</summary>
        Public Const MarcaDasFontes As String = "FONTES:"

        ''' <summary>
        ''' Tira a linha das fontes do texto e traduz as fichas de volta.
        '''
        ''' <b>Ficha que não é deste pedido derruba a citação inteira</b>, e não a
        ''' resposta: o texto continua valendo e aparece <i>sem</i> fontes, que é
        ''' honesto. Descartar a resposta por causa da citação daria a um e-mail o
        ''' poder de apagar a resposta a uma pergunta que o dono fez — e a citação
        ''' é a parte pequena: <see cref="RespostaDoAcervo.Fontes"/> já diz o que
        ''' saiu daqui, e essa lista foi escolhida antes de qualquer byte sair.
        '''
        ''' <b>Sem a linha, não há citação</b> — e isso não é erro. O modelo pode
        ''' não citar; o que ele não pode é citar o que não recebeu.
        ''' </summary>
        Private Shared Function SemAsFontes(bruto As String,
                                            lote As LoteDeClassificacao,
                                            mandadas As IReadOnlyList(Of ItemKey),
                                            ByRef citadas As IReadOnlyList(Of ItemKey)) As String
            citadas = Array.Empty(Of ItemKey)()

            Dim linhas = bruto.Split({Environment.NewLine, vbLf}, StringSplitOptions.None).ToList()
            Dim qual = linhas.FindLastIndex(
                Function(l) l.TrimStart().StartsWith(MarcaDasFontes, StringComparison.Ordinal))
            If qual < 0 Then Return bruto.Trim()

            Dim daLinha = linhas(qual).Trim().Substring(MarcaDasFontes.Length)
            linhas.RemoveAt(qual)
            Dim texto = String.Join(Environment.NewLine, linhas).Trim()

            Dim porFicha = mandadas.ToDictionary(Function(k) lote.FichaDe(k),
                                                 Function(k) k, StringComparer.Ordinal)
            Dim achadas As New List(Of ItemKey)()
            Dim vistas As New HashSet(Of String)(StringComparer.Ordinal)

            For Each pedaco In daLinha.Split(","c, ";"c, " "c)
                Dim ficha = pedaco.Trim()
                If ficha.Length = 0 Then Continue For

                Dim qualItem As ItemKey = Nothing
                If Not porFicha.TryGetValue(ficha, qualItem) Then
                    ' FICHA INVENTADA. Nao ha por que crer nas outras da mesma
                    ' linha, e uma citacao errada e pior que nenhuma: ela manda o
                    ' dono conferir a mensagem errada e voltar achando que
                    ' conferiu.
                    citadas = Array.Empty(Of ItemKey)()
                    Return texto
                End If
                If vistas.Add(ficha) Then achadas.Add(qualItem)
            Next

            citadas = achadas
            Return texto
        End Function

        ''' <summary>
        ''' <b>Procura por PALAVRA, e não pela frase inteira.</b>
        '''
        ''' A busca do acervo casa quem tem <i>todas</i> as palavras do termo. Isso
        ''' é o certo para quem digita uma busca, e é o errado para uma pergunta:
        ''' <i>"quando o João disse que assina o contrato?"</i> exigiria uma
        ''' mensagem cujo assunto tivesse as sete palavras, e não acha nada nunca.
        '''
        ''' Então cada palavra procura sozinha e as listas se juntam, ordenadas por
        ''' <b>quantas palavras</b> a mensagem casou. A que casa "joão" e "contrato"
        ''' vem antes da que casa só "contrato".
        '''
        ''' <b>Palavras curtas caem fora</b> — "que", "o", "de". Elas casam com
        ''' quase tudo, e uma fonte escolhida por casar com "de" é ruído que ocupa
        ''' uma das oito vagas.
        '''
        ''' Nada disto é esperto, e não precisa ser: é uma peneira de metadado que
        ''' roda de graça aqui dentro. O que ela não pode é <b>deixar de existir</b>
        ''' — sem ela, quem escolheria o que sai da máquina seria o modelo.
        ''' </summary>
        Private Function Candidatos(pergunta As String) As IReadOnlyList(Of ItemKey)
            Dim quantasVezes As New Dictionary(Of ItemKey, Integer)()
            Dim ordemDeChegada As New List(Of ItemKey)()

            For Each palavra In Palavras(pergunta)
                For Each a In _busca.Procurar(palavra).Achados
                    Dim chave As New ItemKey(a.Item.ProviderEntryId, Loja(a.FolderKey))
                    If chave.IsEmpty Then Continue For

                    If quantasVezes.ContainsKey(chave) Then
                        quantasVezes(chave) += 1
                    Else
                        quantasVezes(chave) = 1
                        ordemDeChegada.Add(chave)
                    End If
                Next
            Next

            ' DESEMPATE PELA ORDEM DE CHEGADA, que e a da busca -- e nao "o que
            ' vier do dicionario". Duas mensagens com a mesma contagem trocando de
            ' lugar entre duas perguntas iguais fariam sair da maquina, hoje, um
            ' corpo diferente do de ontem.
            Return ordemDeChegada.
                   OrderByDescending(Function(k) quantasVezes(k)).
                   Take(MaximoDeFontes).
                   ToList()
        End Function

        ''' <summary>
        ''' As palavras da pergunta que valem a pena procurar. Três letras é o
        ''' corte, e ele é grosseiro de propósito: uma lista de palavras vazias em
        ''' português seria mais precisa e teria de ser mantida, e errar por
        ''' procurar demais custa uma vaga — errar por procurar de menos custa a
        ''' resposta.
        ''' </summary>
        Private Shared Function Palavras(pergunta As String) As IReadOnlyList(Of String)
            Return pergunta.
                   Split({" "c, ","c, ";"c, ":"c, "?"c, "!"c, "."c,
                          vbCr(0), vbLf(0), vbTab(0)},
                         StringSplitOptions.RemoveEmptyEntries).
                   Where(Function(p) p.Length >= 3).
                   Distinct(StringComparer.OrdinalIgnoreCase).
                   ToList()
        End Function

        ''' <summary>
        ''' A loja daquela pasta, para a chave ficar inteira. <c>EntryID</c>
        ''' sozinho não identifica: duas caixas podem repetir o mesmo.
        ''' </summary>
        Private Function Loja(pasta As Long) As String
            Dim achada = _acervo.Pastas.FirstOrDefault(Function(p) p.Chave = pasta)
            Return If(achada Is Nothing, "", achada.Store)
        End Function

        ''' <summary>
        ''' <b>A frase da cobertura, e ela não é opcional.</b>
        '''
        ''' Ela diz de quantas pastas o acervo sabe e quantas dessas nunca foram
        ''' varridas. Sem isso, "não achei nada" e "ninguém olhou" chegam ao dono
        ''' com a mesma cara.
        '''
        ''' <b>Conta o acervo inteiro, e não o que a busca consultou.</b> A pergunta
        ''' que ela responde é "o que NÃO foi olhado", e uma pasta que a busca
        ''' sequer teve como consultar — porque nunca foi varrida — é exatamente o
        ''' que precisa aparecer aqui.
        ''' </summary>
        Private Function Cobrimento() As String
            Dim todas = _acervo.Pastas.Count
            ' .Where().Count(): a propriedade Count da colecao eclipsa a extensao
            ' Count(Of T) do LINQ. Armadilha do CLAUDE.md.
            Dim varridas = _acervo.Pastas.
                           Where(Function(p) p.Manifesto.GenerationKey.HasValue).Count()

            If todas = 0 Then
                Return "Nenhuma pasta no acervo ainda. Varra alguma antes de perguntar."
            End If
            If varridas = 0 Then
                Return $"Nenhuma das {todas} pasta(s) conhecidas foi varrida: " &
                       "não há acervo sobre o que responder."
            End If
            If varridas < todas Then
                Return $"Procurei em {varridas} de {todas} pasta(s). As outras nunca " &
                       "foram varridas, e o que estiver nelas não entrou nesta resposta."
            End If
            Return $"Procurei nas {todas} pasta(s) varridas."
        End Function

    End Class

    Public Enum MotivoDaResposta
        Respondeu = 0
        PerguntaVazia
        ''' <summary>Nada no acervo casa com a pergunta. <b>Não é "não existe".</b></summary>
        NadaNoAcervo
        ''' <summary>O modelo não devolveu nada de útil.</summary>
        SemResposta
        NaoDeuParaMontar
        SemABorda
    End Enum

    ''' <summary>
    ''' A resposta do acervo.
    '''
    ''' <see cref="Cobertura"/> vem sempre, inclusive quando deu certo: ela é a
    ''' única coisa aqui que diz o que <b>não</b> foi olhado, e a resposta sem
    ''' ela parece completa.
    ''' </summary>
    Public NotInheritable Class RespostaDoAcervo
        Public ReadOnly Property Motivo As MotivoDaResposta
        Public ReadOnly Property Texto As String
        ''' <summary>
        ''' As mensagens cujo corpo saiu desta máquina para responder. É a lista
        ''' que o dono confere — não são "as fontes que o modelo citou", são as
        ''' que <b>foram mandadas</b>, escolhidas aqui antes de qualquer byte
        ''' sair.
        ''' </summary>
        Public ReadOnly Property Fontes As IReadOnlyList(Of ItemKey)
        ''' <summary>
        ''' Em quais delas o modelo <b>diz</b> que se apoiou. Sempre um subconjunto
        ''' de <see cref="Fontes"/> — ficha que não foi mandada zera a lista.
        '''
        ''' <b>Vazia não quer dizer "não usou nada"</b>: quer dizer que não há
        ''' citação utilizável. O que saiu daqui continua sendo <c>Fontes</c>, e é
        ''' essa a lista que responde "o que o Iris mandou".
        ''' </summary>
        Public ReadOnly Property Citadas As IReadOnlyList(Of ItemKey)
        Public ReadOnly Property Cobertura As String

        Private Sub New(motivo As MotivoDaResposta, texto As String,
                        fontes As IReadOnlyList(Of ItemKey),
                        citadas As IReadOnlyList(Of ItemKey), cobertura As String)
            Me.Motivo = motivo
            Me.Texto = If(texto, "")
            Me.Fontes = If(fontes, Array.Empty(Of ItemKey)())
            Me.Citadas = If(citadas, Array.Empty(Of ItemKey)())
            Me.Cobertura = If(cobertura, "")
        End Sub

        Friend Shared Function Recusa(motivo As MotivoDaResposta,
                                      cobertura As String) As RespostaDoAcervo
            Return New RespostaDoAcervo(motivo, "", Nothing, Nothing, cobertura)
        End Function

        Friend Shared Function Feita(texto As String, fontes As IReadOnlyList(Of ItemKey),
                                     citadas As IReadOnlyList(Of ItemKey),
                                     cobertura As String) As RespostaDoAcervo
            Return New RespostaDoAcervo(MotivoDaResposta.Respondeu, texto, fontes,
                                        citadas, cobertura)
        End Function
    End Class

End Namespace

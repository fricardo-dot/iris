Imports System.Collections.Generic
Imports System.Linq

Namespace Global.Iris.Model

    ''' <summary>
    ''' <b>Quem falou por último em cada conversa, e há quantos dias.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A V1 NÃO DIZ "DÍVIDA", E ISSO É A DECISÃO PRINCIPAL</b>
    '''
    ''' Sem classificação, o Iris sabe com segurança <i>quem falou por último e
    ''' quando</i>. Ele <b>não</b> sabe se aquilo pede resposta: um "obrigado,
    ''' recebido" é a última mensagem da conversa e não deve nada a ninguém.
    '''
    ''' Chamar isso de dívida seria afirmar o que não se sabe, e uma fila que
    ''' erra a afirmação perde a confiança inteira — inclusive nas linhas em que
    ''' acertou. Por isso o estado é <see cref="EstadoDaConversa.PossivelResposta"/>,
    ''' e não "precisa de mim": aquilo vem com os rótulos, depois.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>TRÊS MANEIRAS DE NÃO TER RESPOSTA, E ELAS SÃO DIFERENTES</b>
    '''
    ''' <list type="bullet">
    ''' <item><b>Sem os enviados varridos</b> — a fila recusa. Uma conversa cuja
    ''' última mensagem é do outro só é pendência se o Iris pôde ver as
    ''' <i>suas</i> respostas; sem isso ela cobraria exatamente o que já foi
    ''' feito, com a confiança de quem mediu.</item>
    ''' <item><b>Nada classificável</b> — havia conversas e nenhuma deu para
    ''' dizer de quem é a vez. Acontece quando as identidades do dono estão
    ''' incompletas, e o resultado seria uma tela dizendo "não há nada para hoje"
    ''' sobre uma caixa cheia. <b>Contar não bastava:</b> um número ao lado de
    ''' uma lista vazia não protege a afirmação principal.</item>
    ''' <item><b>Fila vazia mesmo</b> — olhou e não há nada esperando.</item>
    ''' </list>
    '''
    ''' As três são <see cref="MotivoDaFila"/>, e a tela diz coisas diferentes
    ''' para cada uma. Achado por revisão externa em 31/08/2026: a segunda
    ''' estava se passando pela terceira.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DIA É DIA DE CALENDÁRIO, E NÃO BLOCO DE 24 HORAS</b>
    '''
    ''' Uma mensagem de sexta às 23h30, aberta na segunda de manhã, esperou
    ''' <b>três dias</b> para quem a mandou — e 56 horas para o relógio. A fila
    ''' existe para ser lida por uma pessoa de manhã, então ela conta como a
    ''' pessoa conta: datas no fuso dela.
    '''
    ''' O fuso entra por parâmetro porque teste não pode depender da máquina, e
    ''' porque <c>DateTimeOffset</c> sozinho não carrega regra de horário de
    ''' verão — carrega só o deslocamento daquele instante.
    ''' </summary>
    Public NotInheritable Class FilaDeRespostas

        ''' <summary>
        ''' Monta as duas filas a partir das mensagens conhecidas.
        ''' </summary>
        ''' <param name="coberturaDosEnviados">
        ''' <b>Por caixa, até quando as respostas do dono são conhecidas</b> — o
        ''' instante em que a pasta de enviados daquela caixa foi publicada.
        '''
        ''' Vazio recusa a fila inteira. Caixa ausente do dicionário tem as
        ''' mensagens descartadas: sem ver os enviados <i>daquela</i> caixa, toda
        ''' conversa já respondida nela apareceria como pendente.
        ''' </param>
        ''' <param name="dispensadas">
        ''' Conversas que o dono marcou como "não exige resposta". Elas somem da
        ''' fila e a contagem <b>de conversas</b> volta no resultado — some da
        ''' vista, não do conhecimento.
        ''' </param>
        Public Shared Function Montar(mensagens As IEnumerable(Of MensagemNaFila),
                                      eu As MinhasIdentidades,
                                      agora As DateTimeOffset,
                                      fuso As TimeZoneInfo,
                                      coberturaDosEnviados As IReadOnlyDictionary(Of String, DateTimeOffset),
                                      dispensadas As IEnumerable(Of String),
                                      Optional remetentesIgnorados As MinhasIdentidades = Nothing) As ResultadoDaFila

            Dim cobertura = If(coberturaDosEnviados,
                               CType(New Dictionary(Of String, DateTimeOffset)(),
                                     IReadOnlyDictionary(Of String, DateTimeOffset)))
            If cobertura.Count = 0 Then Return ResultadoDaFila.SemOsEnviados()

            Dim oFuso = If(fuso, TimeZoneInfo.Utc)
            Dim fora As New ForaDaFila()
            Dim dispensa = New HashSet(Of String)(
                If(dispensadas, Enumerable.Empty(Of String)()), StringComparer.Ordinal)
            Dim dispensadasVistas As New HashSet(Of String)(StringComparer.Ordinal)

            Dim porConversa As New Dictionary(Of String, List(Of MensagemNaFila))(StringComparer.Ordinal)

            For Each m In If(mensagens, Enumerable.Empty(Of MensagemNaFila)())
                If m Is Nothing Then Continue For

                ' CONVERSA DESCONHECIDA NÃO É CONVERSA PRÓPRIA. Juntar as vazias
                ' faria de todas as mensagens ilegíveis uma conversa só, com dez
                ' pessoas diferentes dentro. Só espaço em branco conta como
                ' vazia: uma chave de espaços agruparia do mesmo jeito.
                If String.IsNullOrWhiteSpace(m.Conversa) Then
                    fora.MensagensSemConversa += 1
                    Continue For
                End If

                If Not m.Quando.HasValue Then
                    fora.MensagensSemData += 1
                    Continue For
                End If

                ' CAIXA SEM ENVIADOS VARRIDOS NAO ENTRA.
                '
                ' Uma pasta de enviados varrida em QUALQUER caixa liberava a fila
                ' inteira, inclusive as caixas cujas respostas ninguem viu -- e
                ' toda conversa ja respondida nelas virava pendencia. Achado por
                ' revisao externa em 31/08/2026.
                If Not cobertura.ContainsKey(m.Caixa) Then
                    fora.MensagensSemCoberturaDaCaixa += 1
                    Continue For
                End If

                ' A CHAVE E CAIXA + CONVERSA, e nao a conversa sozinha: o mesmo
                ' ConversationID pode aparecer em duas caixas, e junta-las faria a
                ' mensagem mais nova de uma decidir de quem e a vez na outra.
                Dim chave = m.Caixa & ControlChars.NullChar & m.Conversa

                ' A DISPENSA VALE POR CAIXA, e nao por ConversationID solto.
                '
                ' O agrupamento ja separava as caixas e a dispensa nao: dispensar
                ' uma conversa na caixa compartilhada apagava tambem a conversa de
                ' mesmo id na caixa pessoal -- e o dono nao teria como notar, porque
                ' o que some some. Achado por revisao externa em 01/09/2026.
                '
                ' A LINHA ANTIGA, so com o id, continua valendo para TODAS as
                ' caixas: e o que o dono ja escreveu no arquivo, e reinterpreta-la
                ' como "so na caixa tal" faria conversas dispensadas voltarem sem
                ' ninguem pedir.
                If dispensa.Contains(chave) OrElse dispensa.Contains(m.Conversa) Then
                    dispensadasVistas.Add(chave)
                    Continue For
                End If

                Dim lista As List(Of MensagemNaFila) = Nothing
                If Not porConversa.TryGetValue(chave, lista) Then
                    lista = New List(Of MensagemNaFila)()
                    porConversa(chave) = lista
                End If
                lista.Add(m)
            Next

            fora.ConversasDispensadas = dispensadasVistas.Count

            ' O REMETENTE IGNORADO E CONFERIDO NA LINHA, E NAO NA MENSAGEM.
            '
            ' Tirar as mensagens dele antes de agrupar mudaria QUEM FALOU POR
            ' ULTIMO: uma conversa que ele encerrou passaria a parecer
            ' encerrada por outra pessoa, e a fila trocaria de lado sozinha.
            ' A regra que o dono escreveu e "mensagem deste remetente
            ' normalmente nao exige resposta", e isso e sobre a LINHA.
            Dim ignorados = If(remetentesIgnorados, New MinhasIdentidades({}))

            Dim linhas As New List(Of LinhaDaFila)()
            For Each par In porConversa
                Dim doGrupo = par.Value(0)
                Dim linha = Decidir(doGrupo.Conversa, par.Value, eu, agora, oFuso)

                ' ALEM DA COBERTURA, O IRIS NAO SABE DE QUEM E A VEZ.
                '
                ' Se a ultima mensagem da conversa e POSTERIOR a varredura dos
                ' enviados desta caixa, uma resposta pode existir e nao ter sido
                ' vista. Era o defeito mais grave da fase: Enviados varrida em 1o,
                ' pergunta no dia 29, resposta pelo OWA no dia 30, Entrada varrida
                ' no dia 31 -- e a tela dizia "esperando ha 2 dias" sobre uma
                ' conversa ja respondida, com toda a cara de dado fresco.
                '
                ' Nao ha prazo de frescor inventado aqui: a regra e o proprio
                ' instante da varredura, que e medido.
                If linha IsNot Nothing AndAlso linha.Quando > cobertura(doGrupo.Caixa) Then
                    fora.ConversasAlemDaCobertura += 1
                    Continue For
                End If

                If linha Is Nothing Then
                    fora.ConversasSemDirecao += 1
                ElseIf ignorados.DirecaoDe(linha.RemetenteDaUltima) = Direcao.Minha Then
                    ' "Minha" aqui quer dizer "esta no conjunto" -- o mesmo
                    ' casamento de endereco, com outro conjunto. Reaproveitar
                    ' MinhasIdentidades traz de graca a normalizacao, o X.500 e
                    ' a exigencia de forma.
                    fora.ConversasDeRemetenteIgnorado += 1
                Else
                    linhas.Add(linha)
                End If
            Next

            ' A MAIS ANTIGA PRIMEIRO, PELO INSTANTE E NÃO PELO NÚMERO DE DIAS.
            '
            ' Ordenar por Dias empataria 8 dias e 23 horas com 8 dias e 1 hora, e
            ' o desempate cairia no assunto — a fila prometia "a mais antiga
            ' primeiro" e entregava ordem alfabética dentro do dia. O terceiro
            ' critério é a conversa, para a ordem ser estável mesmo com assunto
            ' repetido: lista que troca de ordem sozinha entre duas aberturas
            ' ensina a não confiar nela.
            Return New ResultadoDaFila(
                linhas.OrderBy(Function(l) l.Quando).
                       ThenBy(Function(l) l.Assunto, StringComparer.Ordinal).
                       ThenBy(Function(l) l.Conversa, StringComparer.Ordinal).ToList(),
                fora, porConversa.Count)
        End Function

        ''' <summary>
        ''' A linha de uma conversa — ou <c>Nothing</c> quando não dá para dizer
        ''' quem falou por último.
        '''
        ''' <b>Empate é "não sei", e não desempate.</b> Duas mensagens no mesmo
        ''' instante com direções diferentes acontecem — cópia de sistema,
        ''' relógio de servidor, importação em lote — e escolher uma delas seria
        ''' inventar a resposta. Escolher <i>sempre a minha</i> esconderia
        ''' pendência; escolher sempre a do outro criaria pendência falsa. A
        ''' terceira saída é dizer que não se sabe.
        '''
        ''' <b>Empate de mesma direção escolhe pela chave.</b> A direção é
        ''' segura, mas a mensagem que a tela vai abrir não pode depender da
        ''' ordem em que o leitor devolveu as linhas — o plano da consulta
        ''' mudaria a linha mostrada sem nada no código mudar.
        ''' </summary>
        Private Shared Function Decidir(conversa As String,
                                        mensagens As List(Of MensagemNaFila),
                                        eu As MinhasIdentidades,
                                        agora As DateTimeOffset,
                                        fuso As TimeZoneInfo) As LinhaDaFila

            Dim maisNova = mensagens.Max(Function(m) m.Quando.Value)
            Dim ultimas = mensagens.Where(Function(m) m.Quando.Value = maisNova).ToList()

            Dim direcoes = ultimas.Select(Function(m) eu.DirecaoDe(m.Remetente)).
                           Distinct().ToList()

            ' DIREÇÃO DESCONHECIDA NÃO VIRA LINHA. Uma linha que não sabe de
            ' quem é a vez não diz nada e ocupa a fila; a conta dela aparece no
            ' resultado, e o MOTIVO da fila muda quando ela domina.
            If direcoes.Count <> 1 Then Return Nothing
            Dim direcao = direcoes(0)
            If direcao = Direcao.Desconhecida Then Return Nothing

            Dim escolhida = ultimas.OrderBy(Function(m) m.Chave.EntryId,
                                            StringComparer.Ordinal).First()

            Return New LinhaDaFila(conversa, escolhida.Chave, escolhida.Assunto,
                                   escolhida.QuemEscreveu, escolhida.Remetente, maisNova,
                                   DiasDeCalendario(maisNova, agora, fuso), direcao)
        End Function

        ''' <summary>
        ''' <b>Quantas datas se passaram</b>, no fuso do dono — e não quantos
        ''' blocos de 24 horas.
        '''
        ''' Sexta 23h30 até segunda 08h são 56 horas e <b>três dias</b> para quem
        ''' está esperando. A fila é lida por uma pessoa de manhã, e conta como
        ''' ela conta.
        '''
        ''' Futuro vira zero: relógio de servidor adiantado ordenaria antes de
        ''' tudo e diria "esperando há -3 dias", que não quer dizer nada.
        ''' </summary>
        Friend Shared Function DiasDeCalendario(quando As DateTimeOffset,
                                                agora As DateTimeOffset,
                                                fuso As TimeZoneInfo) As Integer
            Dim de = TimeZoneInfo.ConvertTime(quando, fuso).Date
            Dim ate = TimeZoneInfo.ConvertTime(agora, fuso).Date
            Dim dias = CInt((ate - de).TotalDays)
            Return If(dias < 0, 0, dias)
        End Function

    End Class

    ''' <summary>
    ''' Uma mensagem, reduzida ao que a fila precisa. <b>Sem corpo</b> — a fila
    ''' inteira se responde com metadado, e é isso que a mantém fora do portão
    ''' de divulgação.
    ''' </summary>
    Public NotInheritable Class MensagemNaFila

        Public ReadOnly Property Chave As ItemKey
        Public ReadOnly Property Conversa As String
        Public ReadOnly Property Assunto As String
        ''' <summary>O nome de exibição — o que vai para a tela.</summary>
        Public ReadOnly Property QuemEscreveu As String
        ''' <summary>O endereço — o que decide a direção.</summary>
        Public ReadOnly Property Remetente As String
        Public ReadOnly Property Quando As DateTimeOffset?

        ''' <summary>
        ''' A caixa de onde ela veio. <b>Faz parte da identidade da conversa</b>:
        ''' o mesmo <c>ConversationID</c> pode aparecer em duas caixas — cópia,
        ''' importação, caixa compartilhada — e juntá-las deixaria a mensagem
        ''' mais nova de uma decidir de quem é a vez na outra.
        ''' </summary>
        Public ReadOnly Property Caixa As String

        Public Sub New(chave As ItemKey, conversa As String, assunto As String,
                       quemEscreveu As String, remetente As String,
                       quando As DateTimeOffset?,
                       Optional caixa As String = Nothing)
            Me.Chave = If(chave, New ItemKey("", ""))
            Me.Conversa = If(conversa, "")
            Me.Assunto = If(assunto, "")
            Me.QuemEscreveu = If(quemEscreveu, "")
            Me.Remetente = If(remetente, "")
            Me.Quando = quando
            Me.Caixa = If(caixa, If(Me.Chave.StoreId, ""))
        End Sub

    End Class

    ''' <summary>Uma conversa na fila.</summary>
    Public NotInheritable Class LinhaDaFila

        Public ReadOnly Property Conversa As String
        ''' <summary>A última mensagem — é ela que a tela abre.</summary>
        Public ReadOnly Property Chave As ItemKey
        Public ReadOnly Property Assunto As String
        ''' <summary>O nome de exibição de quem falou por último.</summary>
        Public ReadOnly Property Quem As String
        ''' <summary>
        ''' O <b>endereço</b> de quem falou por último. Existe para a regra
        ''' "ignorar este remetente" ter em que se apoiar: nome de exibição
        ''' repete e muda, e uma regra sobre nome atingiria quem não devia.
        ''' </summary>
        Public ReadOnly Property RemetenteDaUltima As String
        Public ReadOnly Property Quando As DateTimeOffset
        Public ReadOnly Property Dias As Integer
        Public ReadOnly Property Direcao As Direcao

        Friend Sub New(conversa As String, chave As ItemKey, assunto As String,
                       quem As String, remetenteDaUltima As String,
                       quando As DateTimeOffset, dias As Integer,
                       direcao As Direcao)
            Me.Conversa = conversa
            Me.Chave = chave
            Me.Assunto = assunto
            Me.Quem = quem
            Me.RemetenteDaUltima = If(remetenteDaUltima, "")
            Me.Quando = quando
            Me.Dias = dias
            Me.Direcao = direcao
        End Sub

        ''' <summary>
        ''' <b>Possível resposta</b> quando a última é do outro;
        ''' <b>aguardando</b> quando é minha.
        '''
        ''' "Possível", e não "pendente": o Iris sabe quem falou por último, e
        ''' não se aquilo pede resposta.
        ''' </summary>
        Public ReadOnly Property Estado As EstadoDaConversa
            Get
                Select Case Direcao
                    Case Iris.Model.Direcao.DoOutro : Return EstadoDaConversa.PossivelResposta
                    Case Iris.Model.Direcao.Minha : Return EstadoDaConversa.Aguardando
                    Case Else : Return EstadoDaConversa.NaoSei
                End Select
            End Get
        End Property

        ''' <summary>
        ''' A faixa, por dias. São cortes sobre um número, e o número continua
        ''' visível ao lado — quem discorda do corte ainda consegue conferir.
        ''' </summary>
        Public ReadOnly Property Faixa As FaixaDeEspera
            Get
                If Dias > 14 Then Return FaixaDeEspera.Critico
                If Dias >= 7 Then Return FaixaDeEspera.Atrasado
                If Dias >= 3 Then Return FaixaDeEspera.Atencao
                Return FaixaDeEspera.Normal
            End Get
        End Property

    End Class

    Public Enum EstadoDaConversa
        NaoSei = 0
        ''' <summary>A última é do outro. <b>Pode</b> ser a sua vez.</summary>
        PossivelResposta
        ''' <summary>A última é sua. A espera é do outro.</summary>
        Aguardando
    End Enum

    Public Enum FaixaDeEspera
        Normal = 0
        Atencao
        Atrasado
        Critico
    End Enum

    ''' <summary>
    ''' Por que a fila está como está. <b>Zero é "não respondeu"</b>, então é o
    ''' que aparece em campo esquecido — e o que a tela trata como recusa.
    ''' </summary>
    Public Enum MotivoDaFila
        ''' <summary>Faltou a varredura dos enviados. A fila não foi montada.</summary>
        SemOsEnviados = 0
        ''' <summary>
        ''' Havia conversas e <b>nenhuma</b> deu para classificar. Não é fila
        ''' vazia: é uma caixa cheia sobre a qual o Iris não sabe dizer nada, e
        ''' quase sempre quer dizer que as identidades do dono estão incompletas.
        ''' </summary>
        NadaClassificavel
        ''' <summary>A fila foi montada. Pode estar vazia — vazia é uma resposta.</summary>
        Respondida
    End Enum

    ''' <summary>
    ''' <b>O que não coube na fila, e por quê.</b>
    '''
    ''' Cada contador diz a <b>unidade</b> no nome, e não há total somando-os:
    ''' somar mensagens com conversas produzia um número sem unidade —
    ''' cem mensagens dispensadas de uma conversa só valiam cem, e cem mensagens
    ''' sem direção da mesma conversa valiam uma. Achado por revisão externa.
    ''' </summary>
    Public NotInheritable Class ForaDaFila
        ''' <summary>Mensagens sem conversa legível. Não dá para atribuí-las a nada.</summary>
        Public Property MensagensSemConversa As Integer
        ''' <summary>Mensagens sem data. Sem ela não há conta de dias.</summary>
        Public Property MensagensSemData As Integer
        ''' <summary>Conversas em que não deu para dizer de quem é a vez.</summary>
        Public Property ConversasSemDirecao As Integer
        ''' <summary>Conversas que o dono marcou como "não exige resposta".</summary>
        Public Property ConversasDispensadas As Integer
        ''' <summary>
        ''' Conversas cuja última mensagem é de um remetente que o dono mandou
        ''' ignorar.
        ''' </summary>
        Public Property ConversasDeRemetenteIgnorado As Integer

        ''' <summary>
        ''' Mensagens de uma caixa cujos enviados nunca foram varridos. Sem ver
        ''' as respostas <b>daquela</b> caixa, nada nela pode ser afirmado.
        ''' </summary>
        Public Property MensagensSemCoberturaDaCaixa As Integer

        ''' <summary>
        ''' Conversas cuja última mensagem é <b>mais nova</b> que a varredura dos
        ''' enviados: uma resposta pode existir e não ter sido vista.
        ''' </summary>
        Public Property ConversasAlemDaCobertura As Integer
    End Class

    ''' <summary>As duas filas, e a ressalva.</summary>
    Public NotInheritable Class ResultadoDaFila

        Public ReadOnly Property Linhas As IReadOnlyList(Of LinhaDaFila)
        Public ReadOnly Property Fora As ForaDaFila
        Public ReadOnly Property Motivo As MotivoDaFila
        ''' <summary>
        ''' Quantas conversas foram consideradas — o denominador. Sem ele,
        ''' "3 conversas sem direção" não diz se sobraram trezentas ou nenhuma.
        ''' </summary>
        Public ReadOnly Property ConversasVistas As Integer

        Friend Sub New(linhas As IReadOnlyList(Of LinhaDaFila), fora As ForaDaFila,
                       conversasVistas As Integer)
            Me.Linhas = linhas
            Me.Fora = fora
            Me.ConversasVistas = conversasVistas

            ' NADA CLASSIFICAVEL NAO E FILA VAZIA. Sem esta distincao, uma caixa
            ' cheia com as identidades do dono incompletas produzia a tela
            ' dizendo "nao ha nada para hoje" -- a afirmacao mais errada que
            ' esta fila consegue fazer.
            Motivo = If(linhas.Count = 0 AndAlso fora.ConversasSemDirecao > 0,
                        MotivoDaFila.NadaClassificavel,
                        MotivoDaFila.Respondida)
        End Sub

        Private Sub New()
            Linhas = Array.Empty(Of LinhaDaFila)()
            Fora = New ForaDaFila()
            Motivo = MotivoDaFila.SemOsEnviados
        End Sub

        ''' <summary>
        ''' A fila foi montada? Falso quando faltou a varredura ou quando nada
        ''' deu para classificar — as duas são "não sei", e nenhuma é vazio.
        ''' </summary>
        Public ReadOnly Property Respondeu As Boolean
            Get
                Return Motivo = MotivoDaFila.Respondida
            End Get
        End Property

        ''' <summary>
        ''' Sem os enviados varridos, a fila recusa. Ver o cabeçalho de
        ''' <see cref="FilaDeRespostas"/>.
        ''' </summary>
        Friend Shared Function SemOsEnviados() As ResultadoDaFila
            Return New ResultadoDaFila()
        End Function

        ''' <summary>Só as que podem ser a sua vez.</summary>
        Public Function Minhas() As IReadOnlyList(Of LinhaDaFila)
            Return Linhas.Where(
                Function(l) l.Estado = EstadoDaConversa.PossivelResposta).ToList()
        End Function

        ''' <summary>Só as em que você está esperando outra pessoa.</summary>
        Public Function Deles() As IReadOnlyList(Of LinhaDaFila)
            Return Linhas.Where(
                Function(l) l.Estado = EstadoDaConversa.Aguardando).ToList()
        End Function

    End Class

End Namespace

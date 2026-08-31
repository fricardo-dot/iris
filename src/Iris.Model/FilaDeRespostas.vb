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
    ''' <b>SEM VER OS ENVIADOS, NÃO HÁ FILA</b>
    '''
    ''' Uma conversa cuja última mensagem é do outro só é pendência se o Iris
    ''' pôde ver as <i>suas</i> respostas. Sem Itens Enviados varrido, toda
    ''' conversa respondida por você apareceria como pendente — a fila cobraria
    ''' exatamente o que já foi feito, com a confiança de quem mediu.
    '''
    ''' Então ela não mostra uma lista incompleta: ela <b>recusa</b> e diz por
    ''' quê. É a mesma regra que a busca no acervo já aplica — pasta não varrida
    ''' não é pasta vazia.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE FICA DE FORA, E A CONTA DISSO APARECE</b>
    '''
    ''' Mensagem sem conversa, sem data, ou de remetente que não dá para
    ''' identificar não entra na fila — e o número de cada caso volta no
    ''' resultado. Descartar em silêncio faria a fila parecer completa quando
    ''' metade da caixa não coube nela.
    ''' </summary>
    Public NotInheritable Class FilaDeRespostas

        ''' <summary>
        ''' Monta as duas filas a partir das mensagens conhecidas.
        ''' </summary>
        ''' <param name="viuOsEnviados">
        ''' A pasta de itens enviados foi varrida? <b>Falso recusa a fila
        ''' inteira</b>, e não devolve uma lista parcial.
        ''' </param>
        ''' <param name="dispensadas">
        ''' Conversas que o dono marcou como "não exige resposta". Elas somem da
        ''' fila e a contagem delas volta no resultado — some da vista, não do
        ''' conhecimento.
        ''' </param>
        Public Shared Function Montar(mensagens As IEnumerable(Of MensagemNaFila),
                                      eu As MinhasIdentidades,
                                      agora As DateTimeOffset,
                                      viuOsEnviados As Boolean,
                                      dispensadas As IEnumerable(Of String)) As ResultadoDaFila

            If Not viuOsEnviados Then Return ResultadoDaFila.SemOsEnviados()

            Dim fora As New ForaDaFila()
            Dim dispensa = New HashSet(Of String)(
                If(dispensadas, Enumerable.Empty(Of String)()), StringComparer.Ordinal)

            Dim porConversa As New Dictionary(Of String, List(Of MensagemNaFila))(StringComparer.Ordinal)

            For Each m In If(mensagens, Enumerable.Empty(Of MensagemNaFila)())
                If m Is Nothing Then Continue For

                ' CONVERSA DESCONHECIDA NÃO É CONVERSA PRÓPRIA. Juntar as vazias
                ' faria de todas as mensagens ilegíveis uma conversa só, com dez
                ' pessoas diferentes dentro.
                If String.IsNullOrEmpty(m.Conversa) Then
                    fora.SemConversa += 1
                    Continue For
                End If

                If Not m.Quando.HasValue Then
                    fora.SemData += 1
                    Continue For
                End If

                If dispensa.Contains(m.Conversa) Then
                    fora.Dispensadas += 1
                    Continue For
                End If

                Dim lista As List(Of MensagemNaFila) = Nothing
                If Not porConversa.TryGetValue(m.Conversa, lista) Then
                    lista = New List(Of MensagemNaFila)()
                    porConversa(m.Conversa) = lista
                End If
                lista.Add(m)
            Next

            Dim linhas As New List(Of LinhaDaFila)()
            For Each par In porConversa
                Dim linha = Decidir(par.Key, par.Value, eu, agora)
                If linha Is Nothing Then
                    fora.SemDirecao += 1
                Else
                    linhas.Add(linha)
                End If
            Next

            ' A MAIS ANTIGA PRIMEIRO. É a ordem que a tela existe para mostrar:
            ' o que está esperando há mais tempo é o que se perde de vista.
            Return New ResultadoDaFila(
                linhas.OrderByDescending(Function(l) l.Dias).
                       ThenBy(Function(l) l.Assunto, StringComparer.Ordinal).ToList(),
                fora)
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
        ''' </summary>
        Private Shared Function Decidir(conversa As String,
                                        mensagens As List(Of MensagemNaFila),
                                        eu As MinhasIdentidades,
                                        agora As DateTimeOffset) As LinhaDaFila

            Dim maisNova = mensagens.Max(Function(m) m.Quando.Value)
            Dim ultimas = mensagens.Where(Function(m) m.Quando.Value = maisNova).ToList()

            Dim direcoes = ultimas.Select(Function(m) eu.DirecaoDe(m.Remetente)).
                           Distinct().ToList()

            Dim direcao As Direcao
            If direcoes.Count = 1 Then
                direcao = direcoes(0)
            Else
                direcao = Direcao.Desconhecida
            End If

            ' DIREÇÃO DESCONHECIDA NÃO VIRA LINHA. Uma linha que não sabe de
            ' quem é a vez não diz nada e ocupa a fila; a conta dela aparece no
            ' resultado, que é onde ela é útil.
            If direcao = Direcao.Desconhecida Then Return Nothing

            Dim escolhida = ultimas(0)
            Dim dias = CInt(Math.Floor((agora - maisNova).TotalDays))
            If dias < 0 Then dias = 0

            Return New LinhaDaFila(conversa, escolhida.Chave, escolhida.Assunto,
                                   escolhida.QuemEscreveu, maisNova, dias, direcao)
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

        Public Sub New(chave As ItemKey, conversa As String, assunto As String,
                       quemEscreveu As String, remetente As String,
                       quando As DateTimeOffset?)
            Me.Chave = chave
            Me.Conversa = If(conversa, "")
            Me.Assunto = If(assunto, "")
            Me.QuemEscreveu = If(quemEscreveu, "")
            Me.Remetente = If(remetente, "")
            Me.Quando = quando
        End Sub

    End Class

    ''' <summary>Uma conversa na fila.</summary>
    Public NotInheritable Class LinhaDaFila

        Public ReadOnly Property Conversa As String
        ''' <summary>A última mensagem — é ela que a tela abre.</summary>
        Public ReadOnly Property Chave As ItemKey
        Public ReadOnly Property Assunto As String
        Public ReadOnly Property Quem As String
        Public ReadOnly Property Quando As DateTimeOffset
        Public ReadOnly Property Dias As Integer
        Public ReadOnly Property Direcao As Direcao

        Friend Sub New(conversa As String, chave As ItemKey, assunto As String,
                       quem As String, quando As DateTimeOffset, dias As Integer,
                       direcao As Direcao)
            Me.Conversa = conversa
            Me.Chave = chave
            Me.Assunto = assunto
            Me.Quem = quem
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
    ''' <b>O que não coube na fila, e por quê.</b>
    '''
    ''' Existe porque descartar em silêncio faria a fila parecer completa quando
    ''' metade da caixa não coube nela. Cada número é uma pergunta que a tela
    ''' pode responder ao dono.
    ''' </summary>
    Public NotInheritable Class ForaDaFila
        Public Property SemConversa As Integer
        Public Property SemData As Integer
        Public Property SemDirecao As Integer
        Public Property Dispensadas As Integer

        Public ReadOnly Property Total As Integer
            Get
                Return SemConversa + SemData + SemDirecao + Dispensadas
            End Get
        End Property
    End Class

    ''' <summary>As duas filas, e a ressalva.</summary>
    Public NotInheritable Class ResultadoDaFila

        Public ReadOnly Property Linhas As IReadOnlyList(Of LinhaDaFila)
        Public ReadOnly Property Fora As ForaDaFila
        ''' <summary>
        ''' <b>Falso quando a fila não pôde ser montada.</b> Diferente de fila
        ''' vazia: uma diz "não há nada esperando", a outra diz "não sei".
        ''' </summary>
        Public ReadOnly Property Respondeu As Boolean

        Friend Sub New(linhas As IReadOnlyList(Of LinhaDaFila), fora As ForaDaFila)
            Me.Linhas = linhas
            Me.Fora = fora
            Respondeu = True
        End Sub

        Private Sub New()
            Linhas = Array.Empty(Of LinhaDaFila)()
            Fora = New ForaDaFila()
            Respondeu = False
        End Sub

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

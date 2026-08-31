Imports System.Collections.Generic
Imports System.Linq

Namespace Global.Iris.Model

    ''' <summary>
    ''' <b>Quem merece um rascunho pronto — e o que um rascunho pronto NÃO é.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>NADA É ENVIADO, E NADA É SEQUER ESCRITO SEM UM CLIQUE</b>
    '''
    ''' A regra do projeto é que nada sai por e-mail. Aqui ela precisa de uma
    ''' segunda metade, porque um rascunho automático é a primeira coisa neste
    ''' programa que produz <i>texto em nome do dono</i>: ele também não entra no
    ''' compositor sozinho.
    '''
    ''' Escrever no compositor sem clique seria mutação local sem volta — o dono
    ''' abre a resposta, encontra um texto que não escreveu, e não tem como saber
    ''' o que havia antes. A tela mostra; ele decide.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O RASCUNHO NÃO É GRAVADO EM LUGAR NENHUM</b>
    '''
    ''' Ele nasce do corpo do e-mail e cita o corpo do e-mail. A D1 diz que o
    ''' cache guarda metadado e nunca corpo — e um rascunho gravado seria corpo
    ''' entrando pela porta dos fundos, com outro nome.
    '''
    ''' Então ele vive na sessão e morre com ela. O custo é real: fechar o Iris
    ''' perde os rascunhos, e refazê-los custa outra vez. Está escolhido, e o
    ''' preço fica escrito.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>QUEM MERECE</b>
    '''
    ''' Só quem está esperando: rótulo <c>precisa_de_mim</c>. Redigir para uma
    ''' newsletter é queimar dinheiro; redigir para o que o dono já respondeu é
    ''' pior, porque produz um texto que <i>parece</i> pendência.
    '''
    ''' E há um teto. Sem ele, mandar varrer uma pasta grande dispararia
    ''' centenas de redações de uma vez — e a conta chegaria depois.
    ''' </summary>
    Public NotInheritable Class RascunhosAutomaticos

        ''' <summary>
        ''' <b>Quantos rascunhos por rodada.</b> Número escolhido, não medido: é
        ''' quantas respostas uma pessoa revisa numa sentada antes de parar de
        ''' ler o que está revisando. Acima disso, a fila de rascunhos vira uma
        ''' segunda caixa de entrada.
        ''' </summary>
        Public Const PorRodada As Integer = 10

        ''' <summary>
        ''' O rótulo que merece rascunho, e <b>só</b> ele.
        ''' </summary>
        Public Const RotuloQueMerece As String = "precisa_de_mim"

        ''' <summary>
        ''' Quem merece rascunho, na ordem em que deve ser redigido.
        ''' </summary>
        ''' <param name="jaTem">
        ''' Mensagens que já têm rascunho <b>desta sessão</b>. Refazer o que já
        ''' existe gastaria dinheiro para produzir a mesma coisa — e, pior,
        ''' substituiria um texto que o dono pode já ter lido.
        ''' </param>
        ''' <param name="dispensadas">
        ''' Mensagens em que ele disse que não quer. A recusa dele vale mais do
        ''' que o rótulo, e vale para sempre nesta sessão.
        ''' </param>
        Public Shared Function Escolher(
                mensagens As IReadOnlyList(Of MensagemNaFila),
                rotulos As IReadOnlyDictionary(Of ItemKey, String),
                jaTem As IReadOnlyCollection(Of ItemKey),
                dispensadas As IReadOnlyCollection(Of ItemKey),
                Optional teto As Integer = PorRodada) As IReadOnlyList(Of MensagemNaFila)

            If mensagens Is Nothing OrElse rotulos Is Nothing Then
                Return Array.Empty(Of MensagemNaFila)()
            End If

            Dim feitas = New HashSet(Of ItemKey)(If(jaTem, CType(Array.Empty(Of ItemKey)(),
                                                                IReadOnlyCollection(Of ItemKey))))
            Dim recusadas = New HashSet(Of ItemKey)(If(dispensadas,
                                                       CType(Array.Empty(Of ItemKey)(),
                                                             IReadOnlyCollection(Of ItemKey))))

            Dim quantos = Math.Max(0, Math.Min(teto, PorRodada))

            ' MAIS VELHA PRIMEIRO, e nao mais nova. O rascunho existe para
            ' desatolar o que esta parado; comecar pelo recente atenderia
            ' primeiro quem menos esperou, e o teto faria o resto nunca chegar.
            '
            ' (Comentario AQUI, e nao no meio da cadeia de metodos: em VB a
            ' continuacao implicita nao aceita uma linha so de comentario.)
            Return mensagens.
                   Where(Function(m) m IsNot Nothing AndAlso m.Chave IsNot Nothing).
                   Where(Function(m) Merece(m, rotulos)).
                   Where(Function(m) Not feitas.Contains(m.Chave)).
                   Where(Function(m) Not recusadas.Contains(m.Chave)).
                   OrderBy(Function(m) m.Quando.GetValueOrDefault(DateTimeOffset.MaxValue)).
                   ThenBy(Function(m) m.Assunto, StringComparer.Ordinal).
                   Take(quantos).
                   ToList()
        End Function

        Private Shared Function Merece(m As MensagemNaFila,
                                       rotulos As IReadOnlyDictionary(Of ItemKey, String)) As Boolean
            Dim rotulo As String = Nothing
            If Not rotulos.TryGetValue(m.Chave, rotulo) Then Return False
            Return String.Equals(rotulo, RotuloQueMerece, StringComparison.Ordinal)
        End Function

    End Class

    ''' <summary>
    ''' <b>Os rascunhos da sessão.</b> Em memória, e é decisão: ver o cabeçalho
    ''' de <see cref="RascunhosAutomaticos"/>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>UM RASCUNHO VELHO É PIOR QUE NENHUM</b>
    '''
    ''' Ele é escrito a partir de uma <i>versão</i> da mensagem. Se a mensagem
    ''' mudou — o remetente mandou outra, o corpo foi baixado inteiro, a thread
    ''' andou —, o rascunho responde a um texto que não está mais lá, e o dono
    ''' não tem como perceber isso lendo o rascunho.
    '''
    ''' Então ele guarda a versão junto e <see cref="Pegar"/> <b>não devolve</b>
    ''' o que não corresponde. Não apaga: só não entrega. Apagar esconderia que
    ''' houve um rascunho, e o dono que se lembra de tê-lo visto merece a frase
    ''' "aquele rascunho era de uma versão anterior" em vez do silêncio.
    ''' </summary>
    Public NotInheritable Class RascunhosDaSessao

        Private ReadOnly _porChave As New Dictionary(Of ItemKey, RascunhoPronto)()
        Private ReadOnly _dispensadas As New HashSet(Of ItemKey)()
        Private ReadOnly _trava As New Object()

        ''' <summary>
        ''' Guarda. <paramref name="versao"/> é a <c>PR_CHANGE_KEY</c> da leitura
        ''' que produziu este texto.
        ''' </summary>
        Public Sub Guardar(chave As ItemKey, versao As String, texto As String)
            If chave Is Nothing OrElse String.IsNullOrWhiteSpace(texto) Then Return
            SyncLock _trava
                _porChave(chave) = New RascunhoPronto(If(versao, ""), texto)
            End SyncLock
        End Sub

        ''' <summary>
        ''' O rascunho desta mensagem <b>nesta versão</b>, ou <c>Nothing</c>.
        ''' </summary>
        Public Function Pegar(chave As ItemKey, versao As String) As RascunhoPronto
            If chave Is Nothing Then Return Nothing
            SyncLock _trava
                Dim achado As RascunhoPronto = Nothing
                If Not _porChave.TryGetValue(chave, achado) Then Return Nothing
                If Not String.Equals(achado.Versao, If(versao, ""),
                                     StringComparison.Ordinal) Then Return Nothing
                Return achado
            End SyncLock
        End Function

        ''' <summary>
        ''' <b>Existe um rascunho para esta mensagem, em alguma versão?</b>
        '''
        ''' É o que a escolha da rodada pergunta, e é de propósito que ela não
        ''' pergunte pela versão: um rascunho velho não deve ser <i>mostrado</i>,
        ''' mas também não deve ser refeito automaticamente — refazer sozinho
        ''' gastaria dinheiro toda vez que uma mensagem mudasse de versão, sem o
        ''' dono ter pedido nada.
        ''' </summary>
        Public Function Tem(chave As ItemKey) As Boolean
            If chave Is Nothing Then Return False
            SyncLock _trava
                Return _porChave.ContainsKey(chave)
            End SyncLock
        End Function

        Public Function Feitos() As IReadOnlyCollection(Of ItemKey)
            SyncLock _trava
                Return _porChave.Keys.ToList()
            End SyncLock
        End Function

        ''' <summary>
        ''' O dono disse que não quer rascunho para esta. Vale pela sessão, e
        ''' <b>apaga o que já havia</b>: deixar o texto lá depois de ele recusar
        ''' seria manter na tela justamente o que ele mandou tirar.
        ''' </summary>
        Public Sub Dispensar(chave As ItemKey)
            If chave Is Nothing Then Return
            SyncLock _trava
                _dispensadas.Add(chave)
                _porChave.Remove(chave)
            End SyncLock
        End Sub

        Public Function Dispensadas() As IReadOnlyCollection(Of ItemKey)
            SyncLock _trava
                Return _dispensadas.ToList()
            End SyncLock
        End Function

        ''' <summary>
        ''' Esquece tudo. Chamado quando a sessão de IA é descartada — o mesmo
        ''' momento em que o resto da memória do assistente é jogado fora.
        ''' </summary>
        Public Sub Esquecer()
            SyncLock _trava
                _porChave.Clear()
                _dispensadas.Clear()
            End SyncLock
        End Sub

    End Class

    ''' <summary>Um rascunho, com a versão da mensagem que o originou.</summary>
    Public NotInheritable Class RascunhoPronto
        ''' <summary>A <c>PR_CHANGE_KEY</c> da leitura que produziu este texto.</summary>
        Public ReadOnly Property Versao As String
        Public ReadOnly Property Texto As String

        Friend Sub New(versao As String, texto As String)
            Me.Versao = If(versao, "")
            Me.Texto = If(texto, "")
        End Sub
    End Class

End Namespace

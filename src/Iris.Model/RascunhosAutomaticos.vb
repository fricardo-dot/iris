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
            ' CHAVE VAZIA NAO E MENSAGEM. Um ItemKey sem EntryId nao identifica
            ' nada, e duas dessas colidem entre si: a rodada gastaria uma redacao
            ' e a guardaria por cima da anterior.
            '
            ' E A MESMA CHAVE DUAS VEZES SO CONTA UMA. A lista vem do acervo e
            ' nao promete unicidade; sem isto, uma duplicata gastava dois pedidos
            ' para produzir o mesmo texto -- e o segundo sobrescrevia o primeiro.
            ' Achado por revisao externa em 31/08/2026.
            '
            ' O DESEMPATE VAI ATE A CHAVE, e ela e o unico campo que garante
            ' ordem. Data e assunto empatam com facilidade -- tres "RE: reuniao"
            ' do mesmo minuto nao sao raros --, e o que vem do banco nao tem
            ' ordem prometida. Sem isso, a MESMA caixa rendia um lote de
            ' rascunhos diferente a cada abertura, e o dono via a lista mudar
            ' sozinha. Achado por revisao externa em 01/09/2026.
            '
            ' (O comentario fica AQUI e nao no meio da cadeia: em VB o "." no
            ' fim da linha continua a expressao, e um comentario a quebra.)
            Return mensagens.
                   Where(Function(m) m IsNot Nothing AndAlso
                                     m.Chave IsNot Nothing AndAlso
                                     Not m.Chave.IsEmpty).
                   Where(Function(m) Merece(m, rotulos)).
                   Where(Function(m) Not feitas.Contains(m.Chave)).
                   Where(Function(m) Not recusadas.Contains(m.Chave)).
                   GroupBy(Function(m) m.Chave).
                   Select(Function(g) g.First()).
                   OrderBy(Function(m) m.Quando.GetValueOrDefault(DateTimeOffset.MaxValue)).
                   ThenBy(Function(m) m.Assunto, StringComparer.Ordinal).
                   ThenBy(Function(m) m.Chave.EntryId, StringComparer.Ordinal).
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
    '''
    ''' ------------------------------------------------------------------
    ''' <b>REDIGIR DEMORA, E O DONO NÃO ESPERA</b>
    '''
    ''' Entre pedir a redação e guardá-la passam segundos, e nesses segundos ele
    ''' pode dispensar aquela mensagem ou fechar o painel. A versão anterior
    ''' guardava assim mesmo: o <c>SyncLock</c> protegia os dicionários e não
    ''' protegia a <i>decisão</i>, então a dispensa entrava, a redação terminava,
    ''' e o texto que ele mandou tirar voltava. Achado por revisão externa em
    ''' 31/08/2026.
    '''
    ''' Agora quem vai redigir <see cref="Reservar"/> antes. A reserva carrega a
    ''' <b>geração</b> da sessão, e <see cref="Guardar"/> recusa se a geração
    ''' mudou (houve <see cref="Esquecer"/>) ou se a chave foi dispensada no
    ''' meio. O trabalho pago se perde, e é o lado certo de perder: o outro lado
    ''' é o programa desfazendo o que o dono acabou de mandar fazer.
    ''' </summary>
    Public NotInheritable Class RascunhosDaSessao

        Private ReadOnly _porChave As New Dictionary(Of ItemKey, RascunhoPronto)()
        Private ReadOnly _dispensadas As New HashSet(Of ItemKey)()
        ' EM VOO: a chave E A GERACAO em que ela foi reservada.
        '
        ' Era um conjunto de chaves, e isso deixava a rodada VELHA apagar a
        ' reserva da rodada NOVA: A reserva na geracao 0, Esquecer() limpa tudo
        ' e vai para a 1, B reserva a mesma mensagem na 1, A termina e remove a
        ' chave -- que agora e de B. Ai uma terceira rodada reserva de novo o que
        ' B ainda esta redigindo, e paga duas vezes pelo mesmo texto.
        '
        ' A trava protegia a estrutura e nao a identidade da decisao. E a mesma
        ' familia do defeito que a revisao da Fase 8 achou. Achado por revisao
        ' externa em 31/08/2026.
        Private ReadOnly _emVoo As New Dictionary(Of ItemKey, Long)()
        Private ReadOnly _trava As New Object()
        Private _geracao As Long = 1

        ''' <summary>
        ''' <b>Reserva a mensagem antes de mandar redigir.</b>
        '''
        ''' Devolve a geração a apresentar depois, ou <c>Nothing</c> quando esta
        ''' mensagem não deve ser redigida agora — já tem rascunho, já foi
        ''' dispensada, ou já há uma redação dela em voo.
        '''
        ''' A parte do "em voo" é o que impede duas rodadas simultâneas de pedirem
        ''' a mesma redação duas vezes e uma sobrescrever a outra.
        ''' </summary>
        Public Function Reservar(chave As ItemKey) As Long?
            If chave Is Nothing OrElse chave.IsEmpty Then Return Nothing
            SyncLock _trava
                If _porChave.ContainsKey(chave) Then Return Nothing
                If _dispensadas.Contains(chave) Then Return Nothing
                If _emVoo.ContainsKey(chave) Then Return Nothing
                _emVoo(chave) = _geracao
                Return _geracao
            End SyncLock
        End Function

        ''' <summary>
        ''' A redação não deu. Solta a reserva para ela poder voltar — <b>se a
        ''' reserva ainda for esta</b>.
        '''
        ''' Soltar sem conferir deixava uma rodada velha liberar a mensagem que
        ''' outra, mais nova, já estava redigindo.
        ''' </summary>
        Public Sub Soltar(chave As ItemKey, Optional reserva As Long? = Nothing)
            If chave Is Nothing Then Return
            SyncLock _trava
                SoltarSeForMinha(chave, reserva)
            End SyncLock
        End Sub

        ''' <summary>Sob a trava do chamador. Devolve se a reserva era mesmo esta.</summary>
        Private Function SoltarSeForMinha(chave As ItemKey, reserva As Long?) As Boolean
            Dim daVez As Long
            If Not _emVoo.TryGetValue(chave, daVez) Then Return False
            If reserva.HasValue AndAlso daVez <> reserva.Value Then Return False
            _emVoo.Remove(chave)
            Return True
        End Function

        ''' <summary>
        ''' Guarda, se ainda for para guardar. Devolve <c>False</c> quando o
        ''' trabalho perdeu a validade no meio do caminho.
        '''
        ''' <paramref name="versao"/> é a <c>PR_CHANGE_KEY</c> da leitura que
        ''' produziu este texto, e <b>vazia não serve</b>: sem ela, duas ausências
        ''' passariam por igualdade de versão, e o rascunho continuaria sendo
        ''' entregue depois de a mensagem mudar. Ausência não prova que nada
        ''' mudou; ela prova que ninguém sabe.
        ''' </summary>
        ''' <summary>
        ''' <b>Sem reserva, não se solta reserva de ninguém.</b>
        '''
        ''' O ramo sem identidade chamava a soltura sem conferir, e aí uma chamada
        ''' sem reserva usurpava a vaga de uma rodada em voo: ela gravava por cima,
        ''' liberava a mensagem para uma terceira rodada, e quando a rodada em voo
        ''' terminasse a reserva dela já não existia. A identidade só protegia quem
        ''' a apresentava. Achado por revisão externa em 01/09/2026.
        ''' </summary>
        Public Function Guardar(chave As ItemKey, versao As String, texto As String,
                                Optional reserva As Long? = Nothing) As Boolean
            If chave Is Nothing OrElse chave.IsEmpty Then Return False
            If String.IsNullOrWhiteSpace(texto) Then Return False
            If String.IsNullOrWhiteSpace(versao) Then Return False

            SyncLock _trava
                ' A RESERVA E CONFERIDA ANTES DE SER SOLTA. Se ela nao e mais
                ' minha, quem esta em voo e outra rodada -- e nao posso nem gravar
                ' nem liberar a vaga dela.
                If reserva.HasValue Then
                    If Not SoltarSeForMinha(chave, reserva) Then Return False
                    If reserva.Value <> _geracao Then Return False
                ElseIf _emVoo.ContainsKey(chave) Then
                    ' HA UMA RODADA EM VOO E ESTA CHAMADA NAO E ELA. Gravar aqui
                    ' seria passar na frente de um trabalho ja pago, e liberar a
                    ' vaga dele para uma terceira rodada.
                    Return False
                End If

                If _dispensadas.Contains(chave) Then Return False
                _porChave(chave) = New RascunhoPronto(versao, texto)
                Return True
            End SyncLock
        End Function

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
                ' A reserva NAO e solta: se ela fosse, a redacao em voo poderia
                ' voltar e guardar. O Guardar confere a dispensa de novo, sob a
                ' mesma trava, e e ele quem recusa.
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
        '''
        ''' <b>E vira a geração</b>, para uma redação que já estava em voo não
        ''' ressuscitar dentro de uma sessão que foi descartada.
        ''' </summary>
        Public Sub Esquecer()
            SyncLock _trava
                _porChave.Clear()
                _dispensadas.Clear()
                _emVoo.Clear()
                _geracao += 1
            End SyncLock
        End Sub

        ''' <summary>
        ''' As que já têm rascunho <b>ou estão sendo redigidas agora</b>. É o que a
        ''' escolha da rodada precisa: pedir de novo o que está em voo gastaria dois
        ''' pedidos pelo mesmo texto.
        ''' </summary>
        Public Function FeitosOuEmVoo() As IReadOnlyCollection(Of ItemKey)
            SyncLock _trava
                Return _porChave.Keys.Concat(_emVoo.Keys).Distinct().ToList()
            End SyncLock
        End Function

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

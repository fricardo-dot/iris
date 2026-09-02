Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports Iris.Assist
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>A borda em lote — o que faltava para seis das dez etapas rodarem.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ELA É, E O QUE ELA DELIBERADAMENTE NÃO É</b>
    '''
    ''' <see cref="ClassificarUmaPasta"/> sabe tudo sobre <i>o que</i> mandar —
    ''' quais mensagens faltam rótulo, em lotes de quantas, com quais regras do
    ''' dono e com que controle. O que ele não tem é como <b>ler os corpos</b> e
    ''' como <b>falar com o provedor</b>: as duas coisas moram na camada de
    ''' aplicação, e ele as recebe como dois delegates.
    '''
    ''' Esta classe é esses dois delegates, e nada mais. Ela não decide o que
    ''' classificar, não grava rótulo, não interpreta resposta.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>NÃO HÁ CAMINHO NOVO DE DIVULGAÇÃO — E ISSO É O PONTO</b>
    '''
    ''' A tentação óbvia era montar aqui um envelope próprio e mandar. Seria um
    ''' <b>segundo</b> lugar onde o portão pode ser esquecido, e o portão é a
    ''' única coisa entre o corpo dos e-mails do dono e uma API de terceiro.
    '''
    ''' Então: o mesmo <see cref="ContextoDoOutlook"/> que serve o resumo por
    ''' mensagem, o mesmo <see cref="AssistTransmitter"/>, o mesmo cofre, o mesmo
    ''' diário. O que muda é a <b>seleção</b> — o lote em vez da mensagem aberta
    ''' — e a existência de fichas. Toda garantia já testada do caminho por
    ''' mensagem vale aqui sem precisar ser testada de novo:
    '''
    ''' <list type="bullet">
    ''' <item>o portão classifica cada item antes de qualquer leitura de corpo;</item>
    ''' <item>anexo, corpo incompleto e referência embutida recusam a mensagem;</item>
    ''' <item>o cofre confere que o envelope é exatamente o que foi aprovado;</item>
    ''' <item>a intenção vai ao diário antes do primeiro byte.</item>
    ''' </list>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O ACOPLAMENTO TEMPORAL, DECLARADO</b>
    '''
    ''' <see cref="Conteudo"/> anota o lote corrente; <see cref="Envio"/> o usa.
    ''' Chamar o segundo sem o primeiro é erro de quem monta, e por isso ele
    ''' <b>recusa</b> em vez de mandar com a seleção anterior — mandar o lote
    ''' passado com a instrução do lote novo é divulgação que ninguém pediu.
    '''
    ''' A ordem é garantida por <c>ClassificarUmaPasta.Passar</c>, que chama os
    ''' dois em sequência e roda uma passagem por vez, por cache. Depender disso
    ''' em silêncio seria depender de alguém não mudar de ideia; a recusa é a
    ''' verificação.
    ''' </summary>
    Friend NotInheritable Class BordaEmLote

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _transmissor As AssistTransmitter
        Private ReadOnly _destino As AssistDestination
        Private ReadOnly _pasta As FolderKey

        ''' <summary>
        ''' O lote que <see cref="Conteudo"/> acabou de ler. <c>Nothing</c> antes
        ''' da primeira chamada e entre passagens.
        ''' </summary>
        Private _lote As IReadOnlyList(Of PedidoDeParte)

        ''' <summary>As fichas do lote corrente, por chave.</summary>
        Private _fichas As Dictionary(Of ItemKey, String)

        ''' <summary>
        ''' <b>Quantos BYTES de corpo cabem numa mensagem quando vinte dividem o
        ''' envelope.</b>
        '''
        ''' O envelope inteiro tem <c>EnvelopeBuilder.TetoPadrao</c>; a divisão é por
        ''' <c>ClassificarUmaPasta.PorLote</c>, com 20% de margem para a instrução, o
        ''' esqueleto JSON e o controle. Uma mensagem acima disto é recusada sozinha
        ''' — sem ela, o lote inteiro seria truncado, o cofre recusaria, e o mesmo
        ''' grupo voltaria a falhar em toda passagem.
        '''
        ''' <b>Bytes, e não caracteres.</b> A primeira versão contava caracteres e
        ''' dividia por dois, com um comentário chamando a conta de "conservadora em
        ''' português" — só que um emoji pesa quatro bytes, e português tem emoji
        ''' como qualquer outra língua. Vinte corpos "dentro do teto" somavam o dobro
        ''' do orçamento, o envelope saía truncado, e aquele grupo nunca era
        ''' classificado. Achado por revisão externa em 02/09/2026.
        ''' </summary>
        Friend Shared ReadOnly Property TetoDoCorpoNoLote As Integer
            Get
                Return CInt(EnvelopeBuilder.TetoPadrao * 0.8) \ ClassificarUmaPasta.PorLote
            End Get
        End Property

        Friend Sub New(broker As IOutlookBroker, transmissor As AssistTransmitter,
                       destino As AssistDestination, pasta As FolderKey)
            If broker Is Nothing Then Throw New ArgumentNullException(NameOf(broker))
            If transmissor Is Nothing Then Throw New ArgumentNullException(NameOf(transmissor))
            _broker = broker
            _transmissor = transmissor
            _destino = destino
            _pasta = If(pasta, New FolderKey("", ""))
        End Sub

        ''' <summary>
        ''' <b>Os corpos do lote, lidos numa visita só ao Outlook.</b>
        '''
        ''' Mensagem que o pipeline recusa não entra, e isso não é perda
        ''' silenciosa: <c>ClassificarUmaPasta</c> conta a diferença entre o que
        ''' pediu e o que voltou, e ela aparece como <i>não classificados</i>.
        '''
        ''' Devolver <c>Nothing</c> ou lista vazia faz o lote ser pulado — o que é
        ''' o desfecho certo quando nenhuma mensagem sobreviveu ao pipeline: um
        ''' envelope com o controle e nenhuma mensagem gastaria uma chamada para
        ''' perguntar sobre ninguém.
        ''' </summary>
        Friend Function Conteudo(pedidos As IReadOnlyList(Of PedidoDeParte),
                                 ct As CancellationToken) _
                                 As IReadOnlyList(Of MessagePart)
            ' PARAR ANTES DE LER e melhor que parar antes de mandar: o corpo nem
            ' chega a sair do Outlook.
            If ct.IsCancellationRequested Then Return Nothing
            _lote = Nothing
            _fichas = Nothing
            If pedidos Is Nothing OrElse pedidos.Count = 0 Then Return Nothing

            ' O MAPA E MONTADO ANTES DA LEITURA, e nao durante.
            '
            ' O contexto pergunta a ficha de cada item enquanto monta as partes;
            ' um mapa construido no meio disso seria estado sendo escrito e lido
            ' no mesmo laco.
            Dim mapa As New Dictionary(Of ItemKey, String)()
            For Each p In pedidos
                If p Is Nothing OrElse p.Chave Is Nothing Then Continue For
                mapa(p.Chave) = p.Ficha
            Next

            Dim itens = pedidos.Where(Function(p) p IsNot Nothing AndAlso p.Chave IsNot Nothing).
                                Select(Function(p) p.Chave).ToList()
            If itens.Count = 0 Then Return Nothing

            _lote = pedidos
            _fichas = mapa

            Return ContextoDo(itens, mapa).Partes()
        End Function

        ''' <summary>
        ''' <b>O lote atravessa o portão, o cofre e o diário — e volta o texto.</b>
        '''
        ''' <c>Nothing</c> vale como lote recusado, e é o que sai de todos os
        ''' desfechos que não são resposta: portão negou, cofre recusou, diário
        ''' não fechou, rede caiu. <c>ClassificarUmaPasta</c> conta o lote
        ''' recusado e <b>segue</b> para o próximo — uma pasta não fica sem
        ''' classificação inteira porque um lote no meio não voou.
        '''
        ''' <b>As partes vêm de fora</b>, e não são relidas aqui. É deliberado:
        ''' entre elas está o <i>controle do lote</i>, que
        ''' <c>ClassificarUmaPasta</c> acrescenta e que não corresponde a mensagem
        ''' nenhuma. Remontar a lista aqui o perderia — e o controle é a única
        ''' coisa que separa uma resposta do modelo de uma resposta ditada por um
        ''' e-mail hostil que estava no lote.
        ''' </summary>
        Friend Function Envio(instrucao As String,
                              partes As IReadOnlyList(Of MessagePart),
                              ct As CancellationToken) As RespostaDoLote
            ' O LOTE E CONSUMIDO, e nao so lido.
            '
            ' Ele ficava guardado depois do envio, e um segundo Envio sem um
            ' Conteudo novo -- por refatoracao, por retry, por uso direto --
            ' reusaria em silencio as fichas e a autorizacao do lote anterior.
            ' A porta da passagem impede isso hoje; depender de uma invariante
            ' de outra classe para nao divulgar errado e depender de alguem nao
            ' mudar de ideia. Achado por revisao externa em 01/09/2026.
            Dim lote = _lote
            Dim fichas = _fichas
            _lote = Nothing
            _fichas = Nothing

            If lote Is Nothing Then Return RespostaDoLote.Recusada("o lote não foi lido")
            If ct.IsCancellationRequested Then
                Return RespostaDoLote.Recusada("a classificação foi interrompida")
            End If
            If partes Is Nothing OrElse partes.Count = 0 Then
                Return RespostaDoLote.Recusada("nenhuma mensagem do lote pôde ser lida")
            End If

            ' O CONTEXTO E MONTADO COM AS CHAVES DAS PARTES QUE VAO SAIR, e
            ' nao com as do lote pedido.
            '
            ' ------------------------------------------------------------------
            ' UMA MENSAGEM COM ANEXO MATAVA O LOTE INTEIRO, E PARA SEMPRE
            '
            ' O portao aprovava as vinte chaves do lote; o pipeline recusava a que
            ' tem anexo -- corretamente -- e o envelope saía com dezenove. A
            ' capability exige o conjunto EXATO que aprovou, entao ela recusava, e
            ' nada era mandado.
            '
            ' Isso e certo para um resumo de conversa: uma thread com um membro
            ' faltando nao e a thread. Para um lote de classificacao e fatal, e de
            ' um jeito que nao se percebe: os lotes se formam sempre na mesma
            ' ordem a partir de "presentes e sem rotulo", entao a mesma mensagem
            ' com anexo cai no mesmo lote em toda passagem. Aquelas vinte nunca
            ' seriam classificadas -- e uma caixa de verdade tem anexo em toda
            ' parte. Achado pelo teste do lote parcialmente recusado, em
            ' 01/09/2026.
            '
            ' A correcao e de camada, e nao de politica: quem decide o que sai e o
            ' pipeline, e o portao tem de ser perguntado sobre EXATAMENTE isso. A
            ' mensagem recusada nao e divulgada nem classificada -- ela some do
            ' pedido inteiro, que e o que "recusada" quer dizer.
            '
            ' A parte sem item -- o controle -- fica de fora: ela nao tem chave
            ' para o portao classificar, e a isencao dela vive no envelope.
            Dim itens = partes.Where(Function(p) p IsNot Nothing AndAlso p.Item IsNot Nothing).
                               Select(Function(p) p.Item).ToList()
            If itens.Count = 0 Then
                Return RespostaDoLote.Recusada("nenhuma mensagem do lote pôde ser lida")
            End If

            Dim contexto = ContextoDo(itens, fichas)

            Dim desfecho = _transmissor.Executar(
                contexto.Pedido(AssistOperation.Classificar),
                AddressOf contexto.Classificar,
                Function() New EnvelopeBuilder().Montar(
                    AssistOperation.Classificar, instrucao, partes),
                ct)

            Return Traduzir(desfecho)
        End Function

        ''' <summary>
        ''' <b>O desfecho do transmissor, sem dobrar o incerto no recusado.</b>
        '''
        ''' Isto era <c>If Kind &lt;&gt; Respondeu Then Return Nothing</c>, e o
        ''' <c>Nothing</c> engolia <see cref="AssistOutcomeKind.Ambiguo"/> junto com
        ''' todo o resto. A passagem contava "lote recusado" — que quer dizer
        ''' <b>nada saiu</b> — sobre um lote que pode ter voado. Era a afirmação
        ''' oposta à verdade, na única categoria em que este projeto não pode errar.
        ''' Achado por revisão externa em 01/09/2026.
        '''
        ''' Os motivos são os que o dono pode <b>agir sobre</b>: assinar a ativação,
        ''' olhar a credencial, ou ir conferir o que saiu. Um "não deu" único não
        ''' distingue nenhum dos três.
        ''' </summary>
        Private Shared Function Traduzir(d As AssistOutcome) As RespostaDoLote
            If d Is Nothing Then
                Return RespostaDoLote.Recusada("o envio não devolveu desfecho")
            End If

            Select Case d.Kind
                Case AssistOutcomeKind.Respondeu
                    Return RespostaDoLote.Respondeu(d.Texto)

                Case AssistOutcomeKind.Negado
                    Return RespostaDoLote.Recusada(
                        "a autorização não cobre isto (" & d.MotivoDoPortao.ToString() & ")")

                Case AssistOutcomeKind.NaoComecou
                    Return RespostaDoLote.Recusada(
                        "o provedor não estava pronto — confira a credencial")

                Case AssistOutcomeKind.SemDiario
                    ' NADA SAIU, e sabe-se: a transmissao nao chegou a ser tentada
                    ' porque o diario nao pode registrar a intencao. Sem registro
                    ' nao se transmite.
                    Return RespostaDoLote.Recusada(
                        "o diário do que sai não pôde registrar — nada foi mandado")

                Case AssistOutcomeKind.Ambiguo,
                     AssistOutcomeKind.AmbiguoSemFechamentoDoDiario
                    Return RespostaDoLote.NaoSeSabe(
                        "um lote PODE ter saído e não dá para saber o que aconteceu " &
                        "com ele" & Codigo(d))

                Case Else
                    ' Recusado e Desconhecido. Recusado e o cofre negando, e nada
                    ' saiu; Desconhecido e o zero do enum, que nunca e sucesso.
                    Return RespostaDoLote.Recusada(
                        "o pedido foi recusado antes de sair" & Codigo(d))
            End Select
        End Function

        Private Shared Function Codigo(d As AssistOutcome) As String
            If Not d.CodigoHttp.HasValue Then Return ""
            Return " (HTTP " & d.CodigoHttp.Value.ToString() & ")"
        End Function

        ''' <summary>
        ''' O contexto desta seleção. Novo a cada chamada de propósito: ele é um
        ''' apanhado de funções sobre uma seleção, e guardar um entre lotes seria
        ''' guardar a seleção junto.
        ''' </summary>
        Private Function ContextoDo(itens As IReadOnlyList(Of ItemKey),
                                    fichas As Dictionary(Of ItemKey, String)) _
                                    As ContextoDoOutlook
            Return New ContextoDoOutlook(
                _broker, _destino,
                Function() (Pasta:=_pasta, Itens:=itens),
                Function(chave) FichaDe(fichas, chave),
                TetoDoCorpoNoLote)
        End Function

        ''' <summary>
        ''' O mapa vem por parâmetro, e não do campo: o campo é consumido em
        ''' <see cref="Envio"/>, e o contexto pode perguntar depois disso.
        ''' </summary>
        Private Shared Function FichaDe(fichas As Dictionary(Of ItemKey, String),
                                        chave As ItemKey) As String
            If fichas Is Nothing OrElse chave Is Nothing Then Return Nothing
            Dim ficha As String = Nothing
            If fichas.TryGetValue(chave, ficha) Then Return ficha
            Return Nothing
        End Function

    End Class

End Namespace

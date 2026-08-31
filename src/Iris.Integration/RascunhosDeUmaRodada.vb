Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports Iris.Model

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>Uma rodada de rascunhos automáticos.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>UM PEDIDO POR MENSAGEM, E NÃO UM LOTE</b>
    '''
    ''' A classificação vai em lotes porque o resultado dela é uma palavra de um
    ''' enum: mesmo que um corpo hostil contamine o vizinho, o estrago é um
    ''' rótulo errado.
    '''
    ''' Aqui o resultado é <b>texto escrito em nome do dono</b>. Deixar dois
    ''' corpos hostis dividirem o mesmo contexto significaria que o conteúdo de
    ''' um e-mail pode influenciar a resposta que ele vai mandar para outra
    ''' pessoa — e essa é a única classe de contaminação que não dá para
    ''' consertar com uma superfície fechada, porque a superfície aqui é prosa
    ''' livre.
    '''
    ''' Um pedido por mensagem custa mais e é a única forma de isolamento que
    ''' existe. É a mesma conta que o cabeçalho do <c>LoteDeClassificacao</c> faz
    ''' e resolve do outro jeito, pelo motivo oposto.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>NADA É ENVIADO, E NADA ENTRA NO COMPOSITOR</b>
    '''
    ''' A rodada guarda o texto na sessão. Quem o põe no compositor é o dono,
    ''' com um clique, no botão que já existe. Escrever sozinho seria mutação
    ''' local sem volta.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>FALHA NUMA MENSAGEM NÃO DERRUBA A RODADA</b>
    '''
    ''' Pela mesma razão do lote de classificação: a mensagem que falha pode ser
    ''' justamente a hostil, e deixá-la parar a rodada daria a ela o poder de
    ''' impedir os rascunhos de todas as outras.
    ''' </summary>
    Public NotInheritable Class RascunhosDeUmaRodada

        ''' <summary>
        ''' O que a borda tem de devolver para <b>uma</b> mensagem: o texto do
        ''' rascunho e a versão da leitura que o produziu. Texto vazio, ou
        ''' <c>Nothing</c>, vale como "não deu".
        ''' </summary>
        Public Delegate Function Redigir(mensagem As MensagemNaFila) As RedacaoFeita

        Private ReadOnly _sessao As RascunhosDaSessao

        Public Sub New(sessao As RascunhosDaSessao)
            If sessao Is Nothing Then Throw New ArgumentNullException(NameOf(sessao))
            _sessao = sessao
        End Sub

        ''' <summary>
        ''' Redige para quem merece, até o teto.
        ''' </summary>
        ''' <param name="parar">
        ''' O dono fechou o painel, ou trocou de pasta. A rodada para <b>entre</b>
        ''' mensagens: o que já foi escrito fica guardado, porque foi pago.
        ''' </param>
        Public Function Passar(mensagens As IReadOnlyList(Of MensagemNaFila),
                               rotulos As IReadOnlyDictionary(Of ItemKey, String),
                               redigir As Redigir,
                               Optional teto As Integer = RascunhosAutomaticos.PorRodada,
                               Optional parar As CancellationToken = Nothing) _
                               As ResultadoDaRodada

            If redigir Is Nothing Then Return New ResultadoDaRodada(0, 0, 0, 0, False)

            Dim escolhidas = RascunhosAutomaticos.Escolher(
                mensagens, rotulos, _sessao.FeitosOuEmVoo(), _sessao.Dispensadas(), teto)

            Dim tentadas = 0
            Dim escritos = 0
            Dim falharam = 0

            For Each m In escolhidas
                If parar.IsCancellationRequested Then
                    Return New ResultadoDaRodada(escolhidas.Count, tentadas,
                                                 escritos, falharam, True)
                End If

                ' RESERVA ANTES DE PEDIR. Entre pedir e guardar passam segundos, e
                ' neles o dono pode dispensar a mensagem ou fechar o painel -- a
                ' reserva carrega a geracao da sessao, e o Guardar recusa se ela
                ' mudou. Tambem e o que impede duas rodadas simultaneas de pedirem
                ' a mesma redacao. Achado por revisao externa em 31/08/2026.
                Dim reserva = _sessao.Reservar(m.Chave)
                If Not reserva.HasValue Then Continue For

                tentadas += 1
                Dim feita As RedacaoFeita = Nothing
                Try
                    feita = redigir(m)
                Catch ex As OperationCanceledException When parar.IsCancellationRequested
                    ' SO E INTERRUPCAO SE FOR A DESTA RODADA.
                    '
                    ' Antes, qualquer OperationCanceledException virava "o dono
                    ' interrompeu" -- inclusive o timeout interno do provedor, e
                    ' inclusive uma que a mensagem hostil conseguisse provocar. Ai
                    ' uma mensagem so derrubava a rodada inteira, que e exatamente
                    ' o que o Catch de baixo existe para impedir.
                    _sessao.Soltar(m.Chave)
                    Return New ResultadoDaRodada(escolhidas.Count, tentadas,
                                                 escritos, falharam, True)
                Catch
                    ' A MENSAGEM QUE FALHA PODE SER A HOSTIL. Deixa-la parar a
                    ' rodada daria a ela o poder de impedir os rascunhos de todas
                    ' as outras.
                    feita = Nothing
                End Try

                If feita Is Nothing OrElse String.IsNullOrWhiteSpace(feita.Texto) Then
                    _sessao.Soltar(m.Chave)
                    falharam += 1
                    Continue For
                End If

                ' O GUARDAR PODE RECUSAR, e a recusa e falha: a redacao aconteceu,
                ' custou, e nao virou rascunho. Contar como escrita faria a tela
                ' dizer que ha um rascunho que ninguem vai achar.
                If _sessao.Guardar(m.Chave, feita.Versao, feita.Texto, reserva) Then
                    escritos += 1
                Else
                    falharam += 1
                End If
            Next

            Return New ResultadoDaRodada(escolhidas.Count, tentadas,
                                         escritos, falharam, False)
        End Function

    End Class

    ''' <summary>O que a borda devolve por mensagem.</summary>
    Public NotInheritable Class RedacaoFeita
        Public ReadOnly Property Texto As String
        ''' <summary>
        ''' A <c>PR_CHANGE_KEY</c> da leitura que produziu este texto. É ela que
        ''' faz um rascunho de versão anterior não ser entregue depois.
        ''' </summary>
        Public ReadOnly Property Versao As String

        Public Sub New(texto As String, versao As String)
            Me.Texto = If(texto, "")
            Me.Versao = If(versao, "")
        End Sub
    End Class

    ''' <summary>
    ''' O que a rodada fez.
    '''
    ''' <see cref="Falharam"/> aparece separado de propósito: uma rodada que
    ''' escolheu dez e escreveu zero é um problema, e somá-los num "escreveu 0 de
    ''' 10" não diria se ninguém merecia ou se tudo deu errado.
    '''
    ''' <b><see cref="Escolhidas"/> e <see cref="Tentadas"/> são coisas
    ''' diferentes</b>, e havia só a primeira. Numa rodada interrompida ela dizia
    ''' "dez" tendo tocado numa — e <c>Escritos + Falharam</c> não fechava com
    ''' nada, sem ninguém conseguir dizer se as outras nove tinham dado errado ou
    ''' simplesmente não tinham sido tentadas. Achado por revisão externa em
    ''' 31/08/2026.
    ''' </summary>
    Public NotInheritable Class ResultadoDaRodada
        ''' <summary>Quantas mereciam e cabiam no teto.</summary>
        Public ReadOnly Property Escolhidas As Integer
        ''' <summary>
        ''' Quantas foram efetivamente pedidas. Sempre igual a
        ''' <c>Escritos + Falharam</c>.
        ''' </summary>
        Public ReadOnly Property Tentadas As Integer
        Public ReadOnly Property Escritos As Integer
        Public ReadOnly Property Falharam As Integer
        ''' <summary>
        ''' O dono interrompeu — e só ele: um cancelamento vindo de dentro do
        ''' provedor conta como falha daquela mensagem. O que já foi escrito
        ''' <b>fica</b>: foi pago.
        ''' </summary>
        Public ReadOnly Property Interrompida As Boolean

        Friend Sub New(escolhidas As Integer, tentadas As Integer, escritos As Integer,
                       falharam As Integer, interrompida As Boolean)
            Me.Escolhidas = escolhidas
            Me.Tentadas = tentadas
            Me.Escritos = escritos
            Me.Falharam = falharam
            Me.Interrompida = interrompida
        End Sub
    End Class

End Namespace

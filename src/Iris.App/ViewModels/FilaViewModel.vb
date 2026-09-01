Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Linq
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Integration
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>A fila de respostas na tela.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ESTA TELA AFIRMA COISAS SOBRE O TRABALHO DO DONO</b>
    '''
    ''' "Fulano está esperando você há 20 dias" é uma afirmação forte, lida de
    ''' manhã, que muda o que a pessoa faz no dia. Por isso <b>nenhuma regra de
    ''' quem-fala-por-último mora aqui</b>: elas ficam inteiras em
    ''' <see cref="FilaDeRespostas"/>, e o que este arquivo decide é
    ''' apresentação — as três frases, a ressalva, e o que fazer quando uma ação
    ''' falha.
    '''
    ''' A frase é a parte que mais importa. Uma lista vazia sem frase deixa o
    ''' dono concluir "não tenho nada", e esse é o único desfecho que a fila não
    ''' pode produzir por engano.
    ''' </summary>
    Public NotInheritable Class FilaViewModel
        Inherits ObservableObject

        Private ReadOnly _montar As Func(Of MinhasIdentidades, DateTimeOffset, TimeZoneInfo,
                                         IEnumerable(Of String), MinhasIdentidades,
                                         ResultadoDaFila)
        Private ReadOnly _dispensas As DispensasDaFila
        Private ReadOnly _relogio As Func(Of DateTimeOffset)
        Private ReadOnly _fuso As TimeZoneInfo
        Private ReadOnly _identidades As Func(Of MinhasIdentidades)
        Private ReadOnly _abrir As Action(Of ItemKey)
        Private ReadOnly _rotulo As Func(Of ItemKey, String)
        Private ReadOnly _regrasCasadas As Func(Of ItemKey, Integer)
        ' AS DUAS PARCELAS QUE FALTAVAM CHEGAR AQUI.
        '
        ' A pontuacao sabia contar prazo e pessoa proxima, e a fila nunca as
        ' passava: duas parcelas que nenhum caminho conseguia produzir. Parcela
        ' inalcancavel e o mesmo defeito de guardar um dado que ninguem le.
        Private ReadOnly _pessoaProxima As Func(Of ItemKey, Boolean)
        Private ReadOnly _prazo As Func(Of ItemKey, DateTimeOffset?)
        ' Enderecos que enviaram dos Enviados e nao estao em identidades.txt.
        Private ReadOnly _quemFalta As Func(Of MinhasIdentidades, IReadOnlyList(Of String))

        Public Sub New(montar As Func(Of MinhasIdentidades, DateTimeOffset, TimeZoneInfo,
                                       IEnumerable(Of String), MinhasIdentidades,
                                       ResultadoDaFila),
                       dispensas As DispensasDaFila,
                       identidades As Func(Of MinhasIdentidades),
                       relogio As Func(Of DateTimeOffset),
                       fuso As TimeZoneInfo,
                       Optional abrir As Action(Of ItemKey) = Nothing,
                       Optional rotulo As Func(Of ItemKey, String) = Nothing,
                       Optional regrasCasadas As Func(Of ItemKey, Integer) = Nothing,
                       Optional pessoaProxima As Func(Of ItemKey, Boolean) = Nothing,
                       Optional prazo As Func(Of ItemKey, DateTimeOffset?) = Nothing,
                       Optional quemFalta As Func(Of MinhasIdentidades, IReadOnlyList(Of String)) = Nothing)
            _montar = montar
            _dispensas = If(dispensas, New DispensasDaFila())
            _identidades = If(identidades, Function() New MinhasIdentidades({}))
            _relogio = If(relogio, Function() DateTimeOffset.Now)
            _fuso = If(fuso, TimeZoneInfo.Local)
            _abrir = abrir
            _rotulo = rotulo
            _regrasCasadas = regrasCasadas
            _pessoaProxima = pessoaProxima
            _prazo = prazo
            _quemFalta = quemFalta

            AtualizarCommand = New AsyncRelayCommand(AddressOf Atualizar)
        End Sub

        Public ReadOnly Property AtualizarCommand As IRelayCommand

        ''' <summary>
        ''' <b>Ordenar pela nota, ou pela idade?</b> Desligado por padrão: a idade
        ''' é a ordem que não depende de opinião nenhuma, e a nota é feita de pesos
        ''' que ninguém mediu.
        ''' </summary>
        Public Property PorPrioridade As Boolean
            Get
                Return _porPrioridade
            End Get
            Set(value As Boolean)
                ' REORDENA O QUE ESTA NA TELA, e nao rele o acervo.
                '
                ' Chamar Atualizar aqui montava a fila de novo, e entre a primeira
                ' carga e o clique pode ter chegado mensagem, mudado o relogio ou
                ' entrado uma dispensa: linhas apareciam, sumiam ou mudavam de
                ' idade. O botao dizia "reordena" e trocava o conteudo. Achado por
                ' revisao externa em 31/08/2026.
                If SetProperty(_porPrioridade, value) Then Repovoar()
            End Set
        End Property
        Private _porPrioridade As Boolean

        ' O ULTIMO RETRATO LIDO. E dele que a reordenacao parte.
        Private _ultimo As ResultadoDaFila

        ''' <summary>As conversas em que pode ser a vez do dono.</summary>
        Public ReadOnly Property Minhas As New ObservableCollection(Of LinhaNaTela)()

        ''' <summary>As conversas em que o dono está esperando outra pessoa.</summary>
        Public ReadOnly Property Deles As New ObservableCollection(Of LinhaNaTela)()

        Private _frase As String = ""
        ''' <summary>
        ''' <b>O que a fila conseguiu dizer</b> — e as três recusas não se
        ''' parecem.
        '''
        ''' Lista vazia sem frase deixa o dono concluir "não tenho nada", e é o
        ''' único desfecho que esta tela não pode produzir por engano.
        ''' </summary>
        Public Property Frase As String
            Get
                Return _frase
            End Get
            Private Set(value As String)
                SetProperty(_frase, If(value, ""))
            End Set
        End Property

        Private _ressalva As String = ""
        ''' <summary>
        ''' O que ficou de fora, e por quê. Vazio quando nada ficou — ressalva
        ''' que aparece sempre vira ruído, e ruído ensina a não ler.
        ''' </summary>
        Public Property Ressalva As String
            Get
                Return _ressalva
            End Get
            Private Set(value As String)
                SetProperty(_ressalva, If(value, ""))
                OnPropertyChanged(NameOf(TemRessalva))
            End Set
        End Property

        Public ReadOnly Property TemRessalva As Boolean
            Get
                Return _ressalva.Length > 0
            End Get
        End Property

        ''' <summary>
        ''' <b>A fila pôde ser montada?</b> Falso esconde as listas — mostrar
        ''' zero linhas ao lado de "não sei" seria a própria contradição na tela.
        ''' </summary>
        Private _respondeu As Boolean
        Public Property Respondeu As Boolean
            Get
                Return _respondeu
            End Get
            Private Set(value As Boolean)
                SetProperty(_respondeu, value)
            End Set
        End Property

        ''' <summary>
        ''' Relê o acervo e remonta as duas filas.
        '''
        ''' <b>Falha do leitor não derruba a janela.</b> Isto roda dentro de um
        ''' comando do WPF, no dispatcher: uma exceção do cache subiria sem
        ''' ninguém para pegá-la, e o programa fecharia porque uma lista não
        ''' carregou. A fila some e diz por quê, que é o que ela já faz nas
        ''' outras três recusas.
        ''' </summary>
        ''' <summary>
        ''' <b>A leitura sai da thread da tela.</b>
        '''
        ''' Ela percorre o acervo inteiro — todas as pastas, todas as conversas —,
        ''' e rodava dentro do clique. A janela parava de pintar, de rolar e de
        ''' responder até o SQLite, o agrupamento e a montagem das linhas
        ''' terminarem. Achado por revisão externa em 01/09/2026.
        '''
        ''' <b>Só a leitura.</b> O que mexe nas coleções continua depois do
        ''' <c>Await</c>, e portanto de volta no contexto de quem chamou — que na
        ''' janela é a thread da tela. <c>ObservableCollection</c> mexida de fora
        ''' dela é exceção na hora do binding.
        '''
        ''' <b>Uma por vez.</b> Dois cliques seguidos fariam duas leituras
        ''' concorrentes, e a segunda a terminar mandaria — que pode ser a que
        ''' começou antes. A segunda chamada simplesmente não faz nada.
        ''' </summary>
        Public Async Function Atualizar() As Task
            If _lendo Then Return
            _lendo = True
            Try
                Await Ler()
            Finally
                _lendo = False
            End Try
        End Function
        Private _lendo As Boolean

        Private Async Function Ler() As Task
            ' O RETRATO BOM SO E APAGADO QUANDO HA OUTRO PARA POR NO LUGAR.
            '
            ' As tres saidas de falha limpavam as listas antes de qualquer coisa,
            ' entao uma leitura que estourasse trocava dados validos por uma tela
            ' vazia. A frase explicava a falha e o dono perdia a referencia que
            ' ja tinha -- e a fila e a tela que ele abre de manha.
            '
            ' Manter o retrato velho tem um preco: ele fica na tela sem dizer que
            ' e velho. A frase passa a dizer. Achado por revisao externa em
            ' 01/09/2026.
            Dim eu = _identidades()
            Dim agora = _relogio()
            Dim dispensadas = _dispensas.Conversas()
            Dim ignorados = _dispensas.Remetentes()

            Dim r As ResultadoDaFila
            Try
                ' AS ENTRADAS SAO LIDAS AQUI, e nao la dentro: o relogio e as
                ' dispensas sao estado da tela, e le-los de outra thread seria
                ' a mesma corrida que este Await existe para evitar.
                r = Await Task.Run(Function() _montar(eu, agora, _fuso,
                                                     dispensadas, ignorados))
            Catch ex As Exception
                Frase = "Não consegui ler o acervo agora (" & ex.GetType().Name &
                        "). " & OQueSobrou()
                Return
            End Try

            If r Is Nothing Then
                Frase = "O acervo não respondeu. " & OQueSobrou()
                Return
            End If

            Minhas.Clear()
            Deles.Clear()

            Respondeu = r.Respondeu
            Frase = FraseDe(r)
            Ressalva = Juntar(RessalvaDe(r), FaltamIdentidades())

            _ultimo = r
            If Not r.Respondeu Then Return

            Povoar(r)
        End Function

        ''' <summary>
        ''' Redesenha as duas listas a partir do retrato que já está na mão.
        ''' <b>Sem tocar no acervo</b> — ver o motivo em <see cref="PorPrioridade"/>.
        ''' </summary>
        ''' <summary>
        ''' A segunda metade da frase de falha: <b>o que ficou na tela</b>.
        '''
        ''' Com lista vazia, "a fila não vale" basta. Com linhas na tela, o dono
        ''' precisa saber que o que ele está vendo é de antes — senão ele lê uma
        ''' fila desatualizada como se fosse a de agora, que é pior do que não ver
        ''' fila nenhuma.
        ''' </summary>
        Private Function OQueSobrou() As String
            If Minhas.Count = 0 AndAlso Deles.Count = 0 Then
                Return "A fila não vale enquanto isso."
            End If
            Return "O que está na tela é da leitura anterior."
        End Function

        Private Sub Repovoar()
            ' SEM RETRATO, NAO FAZ NADA -- e nao vai buscar um.
            '
            ' Ir ao acervo daqui era o que fazia o botao de ordenar reler tudo;
            ' depois que a leitura virou assincrona, dispara-la daqui seria
            ' dispara-la sem ninguem para esperar por ela.
            If _ultimo Is Nothing Then Return
            If Not _ultimo.Respondeu Then Return
            Minhas.Clear()
            Deles.Clear()
            Povoar(_ultimo)
        End Sub

        Private Sub Povoar(r As ResultadoDaFila)
            For Each par In NaOrdem(r.Minhas())
                Minhas.Add(New LinhaNaTela(par.Linha, par.Nota, Me))
            Next
            For Each par In NaOrdem(r.Deles())
                Deles.Add(New LinhaNaTela(par.Linha, par.Nota, Me))
            Next
        End Sub

        ''' <summary>
        ''' <b>A ordem por prioridade REORDENA, e não esconde.</b>
        '''
        ''' Ligar a prioridade muda quem aparece primeiro; não muda quem aparece,
        ''' e não tira os dias de linha nenhuma. Uma ordenação que também filtrasse
        ''' esconderia justamente o caso em que a nota errou — e o dono não teria
        ''' como descobrir que ela errou.
        '''
        ''' Desligada, a ordem é a do <see cref="FilaDeRespostas"/>: mais antiga
        ''' primeiro. Ela é a ordem que não depende de opinião nenhuma, e por isso
        ''' é o padrão.
        ''' </summary>
        Private Function NaOrdem(linhas As IEnumerable(Of LinhaDaFila)) _
                                  As IEnumerable(Of LinhaComNota)

            ' A NOTA E CALCULADA UMA VEZ SO, e a MESMA viaja para a tela.
            '
            ' Antes ela era calculada duas: uma na chave da ordenacao e outra no
            ' construtor da linha. E a mesma funcao em codigo, e nao a mesma
            ' avaliacao -- as fontes do rotulo e das regras sao delegates, e nada
            ' promete que devolvem o mesmo duas vezes. A linha podia ser ordenada
            ' com 22 pontos e mostrar 2, que e exatamente a divergencia entre
            ' ordem e explicacao que esta fase existe para impedir. Achado por
            ' revisao externa em 31/08/2026.
            Dim comNota = linhas.Select(Function(l) New LinhaComNota(l, Nota(l))).ToList()
            If Not PorPrioridade Then Return comNota

            ' DESEMPATE ATE O FIM. Nota, dias, assunto -- e a CONVERSA, que e
            ' unica. Sem o ultimo criterio, duas linhas iguais nos tres primeiros
            ' campos dependiam da ordem em que o acervo as enumerou, e trocavam de
            ' lugar entre duas atualizacoes -- justamente o que o comentario
            ' anterior dizia impedir.
            Return comNota.OrderByDescending(Function(x) x.Nota.Total).
                           ThenByDescending(Function(x) x.Linha.Dias).
                           ThenBy(Function(x) x.Linha.Assunto, StringComparer.Ordinal).
                           ThenBy(Function(x) x.Linha.Conversa, StringComparer.Ordinal)
        End Function

        ''' <summary>Uma linha e a nota dela, calculada uma vez.</summary>
        Private NotInheritable Class LinhaComNota
            Public ReadOnly Property Linha As LinhaDaFila
            Public ReadOnly Property Nota As Prioridade

            Public Sub New(linha As LinhaDaFila, nota As Prioridade)
                Me.Linha = linha
                Me.Nota = nota
            End Sub
        End Class

        ''' <summary>
        ''' A nota de uma linha. <b>A mesma função que ordena é a que explica</b> —
        ''' duas contas separadas divergiriam, e a divergência apareceria como uma
        ''' tela cuja explicação não bate com a própria ordem.
        ''' </summary>
        Friend Function Nota(linha As LinhaDaFila) As Prioridade
            Return PrioridadeDaFila.Pontuar(
                linha.Dias,
                If(_rotulo Is Nothing, "", _rotulo(linha.Chave)),
                If(_regrasCasadas Is Nothing, 0, _regrasCasadas(linha.Chave)),
                _pessoaProxima IsNot Nothing AndAlso _pessoaProxima(linha.Chave),
                If(_prazo Is Nothing, Nothing, _prazo(linha.Chave)),
                _relogio())
        End Function

        ''' <summary>
        ''' <b>Uma frase por motivo, e elas dizem coisas diferentes.</b>
        '''
        ''' A do meio é a que existe por causa de um defeito: sem ela, uma caixa
        ''' cheia com as identidades incompletas produzia a tela dizendo que o
        ''' dia estava limpo.
        ''' </summary>
        Friend Shared Function FraseDe(r As ResultadoDaFila) As String
            Select Case r.Motivo
                Case MotivoDaFila.SemOsEnviados
                    Return "Não dá para montar a fila sem ter varrido os Itens " &
                           "Enviados: sem ver as suas respostas, toda conversa que " &
                           "você já respondeu apareceria aqui como pendente."

                Case MotivoDaFila.NadaClassificavel
                    Return $"Vi {r.ConversasVistas} conversa(s) e não consegui dizer de " &
                           "quem é a vez em nenhuma delas. Quase sempre isso quer dizer " &
                           "que falta um endereço seu em identidades.txt — e não que " &
                           "não há nada esperando."

                Case Else
                    If r.Linhas.Count = 0 Then
                        Return "Nada esperando. Olhei " & r.ConversasVistas &
                               " conversa(s)."
                    End If
                    ' "COM A ULTIMA PALAVRA DO OUTRO LADO", e nao "com alguem
                    ' esperando": o Iris sabe quem escreveu por ultimo, e nao se
                    ' aquilo ainda espera resposta. Pode ter sido resolvido por
                    ' telefone, em reuniao, ou num encaminhamento que virou outra
                    ' conversa -- nada disso aparece aqui.
                    Return $"{r.Linhas.Count} de {r.ConversasVistas} conversa(s) " &
                           "com a última palavra dita."
            End Select
        End Function

        ''' <summary>
        ''' O que ficou de fora, com a <b>unidade</b> junto: mensagens e conversas
        ''' são coisas diferentes, e somar as duas produziria um número sem
        ''' significado.
        ''' </summary>
        ''' <summary>
        ''' <b>A ressalva das identidades que faltam.</b>
        '''
        ''' Numa pasta de enviados, quem envia é o dono. Um endereço que aparece
        ''' ali e não está em <c>identidades.txt</c> faz toda conversa dele entrar
        ''' na fila como se fosse de outra pessoa — e agora também pode disparar um
        ''' rascunho automático para algo que ele já respondeu.
        '''
        ''' A frase mostra <b>os endereços</b>, e não só a contagem: sem eles o dono
        ''' não sabe o que escrever no arquivo. Três é o corte — a lista existe para
        ''' ele reconhecer o alias, não para ser completa.
        '''
        ''' <b>Falha aqui não derruba a fila</b>: a ressalva é diagnóstico, e uma
        ''' fila boa não pode sumir porque o diagnóstico estourou.
        ''' </summary>
        Private Function FaltamIdentidades() As String
            If _quemFalta Is Nothing Then Return ""

            Dim faltam As IReadOnlyList(Of String)
            Try
                faltam = _quemFalta(_identidades())
            Catch
                Return ""
            End Try
            If faltam Is Nothing OrElse faltam.Count = 0 Then Return ""

            Dim alguns = String.Join(", ", faltam.Take(3))
            Dim eOutros = If(faltam.Count > 3, $" e mais { faltam.Count - 3}", "")
            Return $"{faltam.Count} endereço(s) enviaram dos seus Itens Enviados e não " &
                   $"estão em identidades.txt ({alguns}{eOutros}) — enquanto faltarem, " &
                   "as conversas deles aparecem como se fossem de outra pessoa"
        End Function

        Private Shared Function Juntar(a As String, b As String) As String
            If a.Length = 0 Then Return b
            If b.Length = 0 Then Return a
            Return a & ". " & b
        End Function

        Friend Shared Function RessalvaDe(r As ResultadoDaFila) As String
            If Not r.Respondeu Then Return ""

            Dim partes As New List(Of String)()
            With r.Fora
                If .ConversasSemDirecao > 0 Then _
                    partes.Add($"{ .ConversasSemDirecao} conversa(s) sem saber de quem é a vez")
                If .ConversasDispensadas > 0 Then _
                    partes.Add($"{ .ConversasDispensadas} dispensada(s) por você")
                If .ConversasDeRemetenteIgnorado > 0 Then _
                    partes.Add($"{ .ConversasDeRemetenteIgnorado} de remetente ignorado")
                ' AS DUAS DE COBERTURA VEM COM O MOTIVO JUNTO. Elas nao sao
                ' descarte: sao o Iris dizendo que naquele pedaco ele nao pode
                ' afirmar, e o dono conserta varrendo.
                If .ConversasAlemDaCobertura > 0 Then _
                    partes.Add($"{ .ConversasAlemDaCobertura} conversa(s) mais novas que a " &
                               "última varredura dos Itens Enviados — varra-os para saber")
                If .MensagensSemCoberturaDaCaixa > 0 Then _
                    partes.Add($"{ .MensagensSemCoberturaDaCaixa} mensagem(ns) de caixa cujos " &
                               "Itens Enviados nunca foram varridos")
                If .MensagensSemConversa > 0 Then _
                    partes.Add($"{ .MensagensSemConversa} mensagem(ns) sem conversa legível")
                If .MensagensSemData > 0 Then _
                    partes.Add($"{ .MensagensSemData} mensagem(ns) sem data")
            End With

            If partes.Count = 0 Then Return ""
            Return "Fora da fila: " & String.Join("; ", partes) & "."
        End Function

        ' ==============================================================
        ' AS TRES ACOES

        ''' <summary>
        ''' A janela não conseguiu abrir a mensagem — ela não está na lista
        ''' carregada. <b>Dizer é obrigatório:</b> clicar em Abrir e não
        ''' acontecer nada ensina que o botão não funciona.
        ''' </summary>
        Public Sub NaoDeuParaAbrir()
            Frase = "Essa mensagem não está na lista aberta agora. Selecione a " &
                    "pasta dela e tente de novo."
        End Sub

        Friend Sub Abrir(chave As ItemKey)
            _abrir?.Invoke(chave)
        End Sub

        ''' <summary>
        ''' Dispensa a conversa e a tira da tela <b>na hora</b>. Se a gravação
        ''' falhar, a linha fica — dizer que dispensou sem ter dispensado
        ''' deixaria o dono achando que resolveu.
        ''' </summary>
        ''' <summary>
        ''' <b>Dispensa a conversa <i>daquela caixa</i>.</b>
        '''
        ''' Gravava só o <c>ConversationID</c>, e o mesmo id existe em duas caixas:
        ''' dispensar na compartilhada apagava também a da caixa pessoal.
        ''' </summary>
        Friend Sub Dispensar(linha As LinhaNaTela)
            If Not _dispensas.DispensarConversa(
                   linha.Caixa & ControlChars.NullChar & linha.Conversa) Then
                Frase = "Não consegui gravar a dispensa. A conversa continua na fila."
                Return
            End If
            Atualizar()
        End Sub

        Friend Sub IgnorarRemetente(linha As LinhaNaTela)
            If Not _dispensas.IgnorarRemetente(linha.Endereco) Then
                Frase = "Não consegui gravar a regra. O remetente continua aparecendo."
                Return
            End If
            Atualizar()
        End Sub

    End Class

    ''' <summary>
    ''' Uma linha da fila, com as três ações. <b>Sem regra</b>: tudo o que ela
    ''' mostra veio decidido de <see cref="LinhaDaFila"/>.
    ''' </summary>
    Public NotInheritable Class LinhaNaTela

        Private ReadOnly _linha As LinhaDaFila
        Private ReadOnly _dona As FilaViewModel

        ''' <summary>
        ''' A nota desta linha. <b>Só é produto acompanhada de
        ''' <see cref="PorQue"/></b>: um número sozinho na tela é um palpite com
        ''' cara de conta, e o dono que discordar não terá do que discordar.
        ''' </summary>
        Public ReadOnly Property Pontos As Double
        ''' <summary>A conta aberta, em português, uma parcela por linha.</summary>
        Public ReadOnly Property PorQue As String

        Friend Sub New(linha As LinhaDaFila, nota As Prioridade, dona As FilaViewModel)
            _linha = linha
            _dona = dona

            ' A NOTA VEM PRONTA, e e a MESMA que ordenou. Recalcula-la aqui
            ' deixaria a ordem e a explicacao virem de duas avaliacoes
            ' diferentes.
            Pontos = nota.Total
            PorQue = nota.Explicar()

            ResponderCommand = New RelayCommand(Sub() dona.Abrir(linha.Chave))
            DispensarCommand = New RelayCommand(Sub() dona.Dispensar(Me))
            IgnorarRemetenteCommand = New RelayCommand(Sub() dona.IgnorarRemetente(Me),
                                                       Function() Endereco.Length > 0)
        End Sub

        Public ReadOnly Property Quem As String
            Get
                Return If(_linha.Quem.Length > 0, _linha.Quem, "(sem remetente)")
            End Get
        End Property

        Public ReadOnly Property Assunto As String
            Get
                Return If(_linha.Assunto.Length > 0, _linha.Assunto, "(sem assunto)")
            End Get
        End Property

        Public ReadOnly Property Dias As Integer
            Get
                Return _linha.Dias
            End Get
        End Property

        ''' <summary>
        ''' "há 20 dias", e o número aparece por extenso porque é o dado —
        ''' a faixa é só um corte sobre ele.
        ''' </summary>
        Public ReadOnly Property Espera As String
            Get
                Select Case _linha.Dias
                    Case 0 : Return "hoje"
                    Case 1 : Return "há 1 dia"
                    Case Else : Return $"há {_linha.Dias} dias"
                End Select
            End Get
        End Property

        Public ReadOnly Property Faixa As FaixaDeEspera
            Get
                Return _linha.Faixa
            End Get
        End Property

        ''' <summary>
        ''' <b>"Possível resposta"</b>, e não "precisa de você": o Iris sabe quem
        ''' falou por último, e não se aquilo pede resposta.
        ''' </summary>
        Public ReadOnly Property Estado As String
            Get
                Select Case _linha.Estado
                    Case EstadoDaConversa.PossivelResposta : Return "Possível resposta"
                    Case EstadoDaConversa.Aguardando : Return "Aguardando"
                    Case Else : Return "Não sei"
                End Select
            End Get
        End Property

        Friend ReadOnly Property Conversa As String
            Get
                Return _linha.Conversa
            End Get
        End Property

        ''' <summary>
        ''' A caixa desta linha. É metade da identidade da conversa — o mesmo
        ''' <c>ConversationID</c> existe em duas caixas —, e é o que faz a
        ''' dispensa valer só onde o dono clicou.
        ''' </summary>
        Friend ReadOnly Property Caixa As String
            Get
                Return If(_linha.Chave?.StoreId, "")
            End Get
        End Property

        Friend ReadOnly Property Endereco As String
            Get
                Return _linha.RemetenteDaUltima
            End Get
        End Property

        Public ReadOnly Property ResponderCommand As IRelayCommand
        Public ReadOnly Property DispensarCommand As IRelayCommand
        Public ReadOnly Property IgnorarRemetenteCommand As IRelayCommand

    End Class

End Namespace

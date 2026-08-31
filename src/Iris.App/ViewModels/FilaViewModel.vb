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

        Public Sub New(montar As Func(Of MinhasIdentidades, DateTimeOffset, TimeZoneInfo,
                                       IEnumerable(Of String), MinhasIdentidades,
                                       ResultadoDaFila),
                       dispensas As DispensasDaFila,
                       identidades As Func(Of MinhasIdentidades),
                       relogio As Func(Of DateTimeOffset),
                       fuso As TimeZoneInfo,
                       Optional abrir As Action(Of ItemKey) = Nothing)
            _montar = montar
            _dispensas = If(dispensas, New DispensasDaFila())
            _identidades = If(identidades, Function() New MinhasIdentidades({}))
            _relogio = If(relogio, Function() DateTimeOffset.Now)
            _fuso = If(fuso, TimeZoneInfo.Local)
            _abrir = abrir

            AtualizarCommand = New RelayCommand(AddressOf Atualizar)
        End Sub

        Public ReadOnly Property AtualizarCommand As IRelayCommand

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
        Public Sub Atualizar()
            Dim r As ResultadoDaFila
            Try
                r = _montar(_identidades(), _relogio(), _fuso,
                            _dispensas.Conversas(), _dispensas.Remetentes())
            Catch ex As Exception
                Minhas.Clear()
                Deles.Clear()
                Respondeu = False
                Ressalva = ""
                Frase = "Não consegui ler o acervo agora (" & ex.GetType().Name &
                        "). A fila não vale enquanto isso."
                Return
            End Try

            If r Is Nothing Then
                Minhas.Clear()
                Deles.Clear()
                Respondeu = False
                Ressalva = ""
                Frase = "O acervo não respondeu. A fila não vale enquanto isso."
                Return
            End If

            Minhas.Clear()
            Deles.Clear()

            Respondeu = r.Respondeu
            Frase = FraseDe(r)
            Ressalva = RessalvaDe(r)

            If Not r.Respondeu Then Return

            For Each l In r.Minhas()
                Minhas.Add(New LinhaNaTela(l, Me))
            Next
            For Each l In r.Deles()
                Deles.Add(New LinhaNaTela(l, Me))
            Next
        End Sub

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
        Friend Sub Dispensar(linha As LinhaNaTela)
            If Not _dispensas.DispensarConversa(linha.Conversa) Then
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

        Friend Sub New(linha As LinhaDaFila, dona As FilaViewModel)
            _linha = linha
            _dona = dona

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

Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Linq
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>A caixa dividida na tela.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A COBERTURA É A PRIMEIRA COISA QUE ELA DIZ</b>
    '''
    ''' Uma caixa dividida em gavetas bonitas dá a impressão de que <i>tudo</i>
    ''' foi olhado. Se a varredura classificou quarenta de novecentas, isso é
    ''' falso — e o dono só descobriria abrindo a última gaveta.
    '''
    ''' Então a frase vem antes das gavetas, e ela é sempre um número: <i>"893 de
    ''' 900 ainda não classificadas"</i>. Não é rodapé nem aviso: é a legenda da
    ''' tela, e sem ela a tela mente por omissão.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE MORA AQUI E O QUE NÃO MORA</b>
    '''
    ''' Aqui: as frases, e o que fazer quando não há nada. A divisão inteira é do
    ''' <see cref="CaixasSeparadas"/>, que é puro — inclusive a decisão de que
    ''' uma mensagem aparece numa gaveta só.
    ''' </summary>
    Public NotInheritable Class CaixasViewModel
        Inherits ObservableObject

        Private ReadOnly _mensagens As Func(Of IReadOnlyList(Of MensagemNaFila))
        Private ReadOnly _rotulos As Func(Of IReadOnlyDictionary(Of ItemKey, String))
        Private ReadOnly _regrasCasadas As Func(Of IReadOnlyDictionary(Of ItemKey, IReadOnlyList(Of String)))
        Private ReadOnly _regrasDoDono As Func(Of IReadOnlyList(Of String))
        Private ReadOnly _abrir As Action(Of ItemKey)

        Public Sub New(mensagens As Func(Of IReadOnlyList(Of MensagemNaFila)),
                       rotulos As Func(Of IReadOnlyDictionary(Of ItemKey, String)),
                       regrasCasadas As Func(Of IReadOnlyDictionary(Of ItemKey, IReadOnlyList(Of String))),
                       regrasDoDono As Func(Of IReadOnlyList(Of String)),
                       Optional abrir As Action(Of ItemKey) = Nothing)
            _mensagens = mensagens
            _rotulos = rotulos
            _regrasCasadas = regrasCasadas
            _regrasDoDono = regrasDoDono
            _abrir = abrir

            AtualizarCommand = New RelayCommand(AddressOf Atualizar)
        End Sub

        Public ReadOnly Property AtualizarCommand As IRelayCommand

        Public ReadOnly Property Gavetas As New ObservableCollection(Of GavetaNaTela)()

        ''' <summary>
        ''' <b>A legenda da tela.</b> Nunca vazia enquanto houver mensagem: uma
        ''' caixa dividida sem esta frase deixa o dono concluir que tudo foi
        ''' olhado.
        ''' </summary>
        Public Property Cobertura As String
            Get
                Return _cobertura
            End Get
            Private Set(value As String)
                SetProperty(_cobertura, If(value, ""))
            End Set
        End Property
        Private _cobertura As String = ""

        ''' <summary>
        ''' <b>Mostrar as gavetas vazias?</b> Ligado, porque "nenhuma mensagem
        ''' espera você" e "esta divisão não existe" são coisas diferentes, e uma
        ''' tela que só mostra gaveta com conteúdo muda de forma a cada
        ''' varredura.
        '''
        ''' Fica desligável porque com dez regras do dono a lista fica longa, e
        ''' aí a escolha é dele — <b>e a legenda continua contando as não
        ''' classificadas</b> mesmo com a gaveta escondida, senão esconder viraria
        ''' um jeito de a tela mentir.
        ''' </summary>
        Public Property MostrarVazias As Boolean
            Get
                Return _mostrarVazias
            End Get
            Set(value As Boolean)
                If SetProperty(_mostrarVazias, value) Then Atualizar()
            End Set
        End Property
        Private _mostrarVazias As Boolean = True

        Public Sub Atualizar()
            Gavetas.Clear()

            Dim todas As IReadOnlyList(Of MensagemNaFila) = Nothing
            Dim divididas As IReadOnlyList(Of Gaveta) = Nothing

            Try
                todas = If(_mensagens Is Nothing, Nothing, _mensagens())
                divididas = CaixasSeparadas.Dividir(
                    todas,
                    If(_rotulos Is Nothing, Nothing, _rotulos()),
                    If(_regrasCasadas Is Nothing, Nothing, _regrasCasadas()),
                    If(_regrasDoDono Is Nothing, Nothing, _regrasDoDono()))
            Catch
                ' UMA LEITURA QUE FALHA NAO PODE VIRAR UMA CAIXA VAZIA COM CARA
                ' DE CAIXA LIMPA. A frase diz o que aconteceu, e nao ha gaveta
                ' nenhuma para o dono confundir com resultado.
                Cobertura = "Não deu para ler a classificação agora."
                Return
            End Try

            If divididas Is Nothing Then
                Cobertura = "Não deu para ler a classificação agora."
                Return
            End If

            For Each g In divididas
                If g.Vazia AndAlso Not MostrarVazias Then Continue For
                Gavetas.Add(New GavetaNaTela(g, _abrir))
            Next

            Cobertura = Legenda(divididas)
        End Sub

        ''' <summary>
        ''' A legenda, e ela sempre dá o número.
        '''
        ''' <b>"Todas classificadas" só é dito quando é verdade</b> — e uma caixa
        ''' sem mensagem nenhuma não é isso: é uma caixa não varrida, ou vazia, e
        ''' as duas merecem outra frase.
        ''' </summary>
        Private Shared Function Legenda(gavetas As IReadOnlyList(Of Gaveta)) As String
            Dim total = gavetas.Sum(Function(g) g.Quantas)
            If total = 0 Then Return "Nada aqui ainda. Varra uma pasta para começar."

            Dim resto = gavetas.
                        Where(Function(g) g.Nome = CaixasSeparadas.NomeDasNaoClassificadas).
                        Sum(Function(g) g.Quantas)

            If resto = 0 Then
                Return $"{total} de {total} classificadas."
            End If
            If resto = total Then
                Return $"Nenhuma das {total} foi classificada ainda."
            End If
            Return $"{resto} de {total} ainda não classificadas."
        End Function

    End Class

    ''' <summary>Uma gaveta, pronta para a tela.</summary>
    Public NotInheritable Class GavetaNaTela

        Public ReadOnly Property Nome As String
        Public ReadOnly Property Quantas As Integer
        Public ReadOnly Property Vazia As Boolean
        ''' <summary>
        ''' Esta gaveta nasceu de uma regra que o dono escreveu. A tela precisa:
        ''' a regra dele ele pode corrigir; a categoria que o programa inventou,
        ''' não.
        ''' </summary>
        Public ReadOnly Property DoDono As Boolean
        Public ReadOnly Property Mensagens As New ObservableCollection(Of MensagemNaGaveta)()

        Friend Sub New(gaveta As Gaveta, abrir As Action(Of ItemKey))
            Nome = gaveta.Nome
            Quantas = gaveta.Quantas
            Vazia = gaveta.Vazia
            DoDono = gaveta.DoDono

            For Each m In gaveta.Mensagens
                Mensagens.Add(New MensagemNaGaveta(m, abrir))
            Next
        End Sub

    End Class

    Public NotInheritable Class MensagemNaGaveta

        Public ReadOnly Property Assunto As String
        Public ReadOnly Property Quem As String
        Public ReadOnly Property Quando As DateTimeOffset?
        Public ReadOnly Property AbrirCommand As IRelayCommand

        Friend Sub New(m As MensagemNaFila, abrir As Action(Of ItemKey))
            Assunto = If(m.Assunto, "")
            Quem = If(String.IsNullOrWhiteSpace(m.QuemEscreveu), m.Remetente, m.QuemEscreveu)
            Quando = m.Quando

            Dim chave = m.Chave
            AbrirCommand = New RelayCommand(
                Sub() If abrir IsNot Nothing Then abrir(chave),
                Function() abrir IsNot Nothing)
        End Sub

    End Class

End Namespace

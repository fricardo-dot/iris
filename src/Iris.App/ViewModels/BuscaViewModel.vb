Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Linq
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Integration

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>A busca do acervo na tela.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ESTE ARQUIVO EXISTE JUNTO COM A BUSCA, E NÃO DEPOIS</b>
    '''
    ''' O erro mais comum deste projeto, seis vezes contadas pela revisão
    ''' externa da Fase 3, é <b>proteção que existe e não está ligada a
    ''' nada</b>: o método existe, os testes o chamam, e na aplicação ninguém
    ''' chama. Entregar <see cref="BuscaNoAcervo"/> sem tela seria a sétima.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A RESSALVA NÃO É DECORAÇÃO, E POR ISSO NÃO SOME</b>
    '''
    ''' <see cref="Ressalva"/> aparece sempre que houve uma busca — inclusive
    ''' quando ela achou. Uma ressalva que só aparece no resultado vazio
    ''' ensina o usuário a lê-la como "não achei", e ela não é isso: ela diz
    ''' <b>onde</b> se procurou, com que alcance, e que o corpo da mensagem
    ''' não é procurável. Isso vale igual quando há dez achados.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>SÍNCRONO, E ISSO É MEDIDO</b>
    '''
    ''' A busca lê o manifesto de cada pasta e casa em memória. Sobre o acervo
    ''' medido em 28/08/2026 — 1.123 linhas — isso é instantâneo, e um
    ''' <c>Async</c> aqui traria geração, cancelamento e guarda de descarte
    ''' para administrar um trabalho que termina antes de a tela repintar.
    '''
    ''' O dia em que deixar de ser instantâneo, isto muda — e aí o custo do
    ''' ciclo de vida se paga. O gatilho está escrito em
    ''' <see cref="BuscaNoAcervo"/>.
    ''' </summary>
    Public NotInheritable Class BuscaViewModel
        Inherits ObservableObject

        Private ReadOnly _busca As Func(Of String, ResultadoDaBusca)

        Private _termo As String = ""
        Private _ressalva As String = ""
        Private _procurou As Boolean

        ''' <summary>
        ''' A função de busca é injetada, e não construída aqui.
        '''
        ''' Duas razões. A primeira é a §26.2 e a <c>ArchitectureTests</c>: a
        ''' camada de apresentação não instancia leitor de cache. A segunda é
        ''' que sem isto não haveria teste — o construtor abriria banco.
        ''' </summary>
        Public Sub New(busca As Func(Of String, ResultadoDaBusca))
            _busca = busca
            ProcurarCommand = New RelayCommand(AddressOf Procurar, Function() _busca IsNot Nothing)
            LimparCommand = New RelayCommand(AddressOf Limpar, Function() Procurou)
        End Sub

        Public ReadOnly Property Achados As New ObservableCollection(Of LinhaAchada)()
        Public ReadOnly Property ProcurarCommand As IRelayCommand
        Public ReadOnly Property LimparCommand As IRelayCommand

        Public Property Termo As String
            Get
                Return _termo
            End Get
            Set(value As String)
                SetProperty(_termo, If(value, ""))
            End Set
        End Property

        ''' <summary>
        ''' O que a busca <b>deve</b> dizer sobre si mesma. Nunca "não existe".
        ''' </summary>
        Public Property Ressalva As String
            Get
                Return _ressalva
            End Get
            Private Set(value As String)
                If SetProperty(_ressalva, If(value, "")) Then
                    OnPropertyChanged(NameOf(TemRessalva))
                End If
            End Set
        End Property

        Public ReadOnly Property TemRessalva As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(_ressalva)
            End Get
        End Property

        ''' <summary>Alguém já procurou alguma coisa nesta sessão de tela.</summary>
        Public Property Procurou As Boolean
            Get
                Return _procurou
            End Get
            Private Set(value As Boolean)
                If SetProperty(_procurou, value) Then
                    OnPropertyChanged(NameOf(SemAchados))
                    LimparCommand.NotifyCanExecuteChanged()
                End If
            End Set
        End Property

        ''' <summary>
        ''' Procurou e não achou — que <b>não</b> é o mesmo que não existir, e
        ''' é por isso que a ressalva fica visível junto.
        ''' </summary>
        ''' <summary>
        ''' Quantos dos achados vieram do segundo passe.
        '''
        ''' A tela diz o número, e não só marca as linhas, porque a pergunta
        ''' que o usuário faz é sobre o conjunto: "isto que eu estou vendo é o
        ''' que eu pedi, ou é o que se pareceu com o que eu pedi?".
        ''' </summary>
        Public Property Aproximados As Integer

        Public ReadOnly Property TemAproximados As Boolean
            Get
                Return Aproximados > 0
            End Get
        End Property

        ''' <summary>
        ''' <b>A frase que impede o palpite de passar por achado.</b>
        '''
        ''' Diz o que a aproximação é — mesma raiz, ou uma letra — para o
        ''' usuário poder julgar sozinho se aquilo responde a pergunta dele.
        ''' "Resultados aproximados", sem dizer aproximados COMO, seria
        ''' ressalva decorativa.
        ''' </summary>
        Public ReadOnly Property FraseDosAproximados As String
            Get
                If Aproximados <= 0 Then Return ""
                Return $"{Aproximados} destes casaram por aproximação — mesma " &
                       "raiz de palavra, ou uma letra de diferença. Estão " &
                       "marcados, e vêm depois dos exatos."
            End Get
        End Property

        Public ReadOnly Property SemAchados As Boolean
            Get
                Return _procurou AndAlso Achados.Count = 0
            End Get
        End Property

        Private Sub Procurar()
            If _busca Is Nothing Then Return

            Dim r As ResultadoDaBusca
            Try
                r = _busca(Termo)
            Catch ex As Exception
                ' Banco travado, arquivo sumindo. Nao pode derrubar a janela, e
                ' tambem nao pode virar "nao achei": as duas coisas sao
                ' diferentes, e confundi-las e o defeito que a §23 persegue.
                Achados.Clear()
                Procurou = True
                Ressalva = "Não consegui procurar no acervo agora. " &
                           "Isso não diz nada sobre o que existe na caixa."
                OnPropertyChanged(NameOf(SemAchados))
                Return
            End Try

            ''' EXATOS PRIMEIRO, E NÃO POR ESTÉTICA.
            '''
            ''' A busca ganhou um segundo passe tolerante a erro de digitação e
            ''' a flexão (medido em 29/08: 0,4% para 93,8%, e 0% para 100%). Um
            ''' achado aproximado é um palpite bom, e palpite no topo da lista
            ''' empurra a certeza para baixo da dobra: quem olha os três
            ''' primeiros resultados veria três palpites.
            '''
            ''' Estável dentro de cada grupo — a ordem das pastas e do
            ''' manifesto continua valendo, e não é esta linha que a decide.
            Achados.Clear()
            For Each a In r.Achados.OrderBy(Function(x) If(x.Grau = GrauDoAchado.Exato, 0, 1))
                Achados.Add(New LinhaAchada(a))
            Next

            Procurou = True
            Ressalva = r.Ressalva
            Aproximados = r.Achados.Where(Function(x) x.Grau <> GrauDoAchado.Exato).Count()
            OnPropertyChanged(NameOf(SemAchados))
            OnPropertyChanged(NameOf(Aproximados))
            OnPropertyChanged(NameOf(TemAproximados))
            OnPropertyChanged(NameOf(FraseDosAproximados))
        End Sub

        Private Sub Limpar()
            Termo = ""
            Achados.Clear()
            Procurou = False
            Ressalva = ""

            ''' A CONTAGEM DE APROXIMADOS TAMBÉM, e ela tinha sido esquecida.
            '''
            ''' Sem esta linha, limpar uma busca que tinha palpites deixava na
            ''' tela "3 destes casaram por aproximação" sobre uma lista vazia —
            ''' uma ressalva verdadeira ontem, falsa agora, e do tipo que
            ''' ninguém confere porque ressalva parece inofensiva.
            Aproximados = 0

            OnPropertyChanged(NameOf(SemAchados))
            OnPropertyChanged(NameOf(Aproximados))
            OnPropertyChanged(NameOf(TemAproximados))
            OnPropertyChanged(NameOf(FraseDosAproximados))
        End Sub

    End Class

    ''' <summary>
    ''' Uma linha achada, já no formato da tela.
    '''
    ''' Existe para a apresentação não formatar dado do domínio ponto a ponto
    ''' no XAML — e para a data ser tratada uma vez só, num lugar onde a falha
    ''' de conversão é decidida em vez de virar exceção de binding.
    ''' </summary>
    Public NotInheritable Class LinhaAchada

        Public ReadOnly Property Assunto As String
        Public ReadOnly Property Remetente As String
        Public ReadOnly Property Quando As String
        Public ReadOnly Property Pasta As String

        ''' <summary>
        ''' A presença conforme o acervo. <b>Suspeito</b> quer dizer que a
        ''' mensagem pode já ter sido apagada — e mostrar isso é o que impede
        ''' o resultado de parecer o estado corrente da caixa.
        ''' </summary>
        Public ReadOnly Property Aviso As String

        ''' <summary>
        ''' <b>Esta linha casou por aproximação.</b>
        '''
        ''' Marca na própria linha, e não só no rodapé: quem lê uma lista lê
        ''' linha a linha, e um contador no rodapé não diz <i>qual</i> das
        ''' quinze é o palpite.
        ''' </summary>
        Public ReadOnly Property Aproximado As Boolean

        Friend Sub New(a As AchadoDaBusca)
            Aproximado = a.Grau <> GrauDoAchado.Exato
            Assunto = If(String.IsNullOrWhiteSpace(a.Item.Subject), "(sem assunto)", a.Item.Subject)
            Remetente = If(a.Item.SenderName, "")
            Pasta = If(a.NomeDaPasta, "")

            ' TryParse fora do If ternario: o VB avalia os dois ramos do If
            ' e a variavel ainda nao esta atribuida quando ele monta o
            ' segundo. Escrito assim, a conversao acontece uma vez e a falha
            ' e uma DECISAO -- data ilegivel vira vazio, e nao excecao de
            ' binding num lugar onde ninguem esta olhando.
            '
            ' E o local NAO pode se chamar `quando`: ele eclipsaria a
            ' propriedade Quando, e a mensagem seria "String nao converte para
            ' DateTimeOffset" numa linha que nao tem String nenhuma. Quarta vez
            ' nesta sessao -- e a tabela do CLAUDE.md ja tem doze entradas.
            Dim recebida As DateTimeOffset = Nothing
            If DateTimeOffset.TryParse(a.Item.ReceivedAt, recebida) Then
                Quando = recebida.ToLocalTime().ToString("dd/MM/yyyy HH:mm")
            Else
                Quando = ""
            End If

            Aviso = If(a.Item.Presence = Iris.Sync.PresenceState.Suspeito,
                       "pode já ter sido apagada", "")
        End Sub

        Public ReadOnly Property TemAviso As Boolean
            Get
                Return Not String.IsNullOrEmpty(Aviso)
            End Get
        End Property

    End Class

End Namespace

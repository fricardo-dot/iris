Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Core
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>A pasta pessoal de Contatos — e a frase que ela nunca pode dizer.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>"NENHUM CONTATO" É MENTIRA NESTA CAIXA</b>
    '''
    ''' Medido em 28/08/2026: a pasta padrão de Contatos tem <b>0 itens</b>, e a
    ''' organização inteira é endereçável — porque os contatos corporativos
    ''' vivem no <b>GAL</b>, que a §8 põe fora de escopo.
    '''
    ''' Uma tela que mostrasse a lista vazia e calasse estaria afirmando
    ''' ausência a partir de não ter olhado, e desta vez sobre <i>pessoas</i>. É
    ''' a mesma família de defeito que esta base corrigiu em cinco lugares, no
    ''' pior lugar possível para tê-la.
    '''
    ''' Por isso a ressalva vem do <c>ContactWriting</c>, junto com o resultado,
    ''' e não é montada aqui: ressalva que depende de uma tela lembrar de
    ''' escrever é ressalva que some na próxima tela.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E O AVISO DE DUPLICATA NÃO BLOQUEIA</b>
    '''
    ''' Contato repetido não some sozinho: o catálogo fica com duas fichas da
    ''' mesma pessoa, com dados diferentes, e quem completar endereço no Outlook
    ''' acerta uma das duas por sorte.
    '''
    ''' Mas o aviso é <b>aviso</b>, e não recusa — e a razão está no parágrafo
    ''' de cima. A busca só enxerga os contatos que esta leitura trouxe: não a
    ''' pasta inteira quando houve truncamento, e nunca o GAL. Bloquear com essa
    ''' base seria transformar "não encontrei" em "não existe", que é
    ''' exatamente o erro que a fase inteira evita — só que virado do avesso.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E O CONTATO NÃO VAI PARA O ASSISTENTE</b>
    '''
    ''' Não há caminho daqui para a IA, e é decisão. Um contato é dado pessoal
    ''' de <i>terceiro</i> — alguém que não é o usuário e não escolheu nada. A
    ''' cerimônia de ativação autoriza operações nomeadas sobre pastas de
    ''' mensagens escolhidas; ela não autoriza mandar o catálogo de endereços
    ''' para um provedor externo, e reusá-la para isso alargaria uma permissão
    ''' que ninguém deu.
    ''' </summary>
    Public NotInheritable Class ContatosViewModel
        Inherits ObservableObject
        Implements IDisposable

        Private Const Teto As Integer = 200

        Private ReadOnly _broker As IContatosBroker

        Private _pasta As FolderKey
        Private _carregando As Boolean
        Private _gravando As Boolean
        Private _erro As String = ""
        Private _resumo As String = ""
        Private _ressalva As String = ""
        Private _aviso As String = ""
        Private _novoNome As String = ""
        Private _novoEmail As String = ""
        Private _novaEmpresa As String = ""
        Private _disposed As Boolean
        Private _geracao As Integer

        Public Sub New(broker As IContatosBroker)
            _broker = broker

            AbrirCommand = New AsyncRelayCommand(AddressOf AbrirAsync,
                                                 Function() Not _carregando)
            AtualizarCommand = New AsyncRelayCommand(AddressOf CarregarAsync,
                                                     Function() _pasta IsNot Nothing AndAlso Not _carregando)
            CriarCommand = New AsyncRelayCommand(AddressOf CriarAsync, Function() PodeCriar)
        End Sub

        Public ReadOnly Property Contatos As New ObservableCollection(Of LinhaDeContato)()
        Public ReadOnly Property AbrirCommand As IAsyncRelayCommand
        Public ReadOnly Property AtualizarCommand As IAsyncRelayCommand
        Public ReadOnly Property CriarCommand As IAsyncRelayCommand

        Public ReadOnly Property TemPasta As Boolean
            Get
                Return _pasta IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property Carregando As Boolean
            Get
                Return _carregando
            End Get
        End Property

        ''' <summary>
        ''' Criar exige <b>nome</b>, como o <c>ContactWriting</c> exige. A tela
        ''' não repete a guarda para substituí-la: repete para não oferecer um
        ''' botão que vai ser recusado.
        ''' </summary>
        Public ReadOnly Property PodeCriar As Boolean
            Get
                Return _pasta IsNot Nothing AndAlso Not _gravando AndAlso
                       Not _carregando AndAlso Not String.IsNullOrWhiteSpace(_novoNome)
            End Get
        End Property

        Public Property NovoNome As String
            Get
                Return _novoNome
            End Get
            Set(value As String)
                If SetProperty(_novoNome, If(value, "")) Then
                    OnPropertyChanged(NameOf(PodeCriar))
                    OnPropertyChanged(NameOf(TemProposta))
                    CriarCommand.NotifyCanExecuteChanged()
                End If
            End Set
        End Property

        Public Property NovoEmail As String
            Get
                Return _novoEmail
            End Get
            Set(value As String)
                If SetProperty(_novoEmail, If(value, "")) Then
                    OnPropertyChanged(NameOf(TemProposta))
                    ReavaliarDuplicata()
                End If
            End Set
        End Property

        Public Property NovaEmpresa As String
            Get
                Return _novaEmpresa
            End Get
            Set(value As String)
                SetProperty(_novaEmpresa, If(value, ""))
            End Set
        End Property

        Public ReadOnly Property TemProposta As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(_novoNome) OrElse
                       Not String.IsNullOrWhiteSpace(_novoEmail)
            End Get
        End Property

        ''' <summary>
        ''' O que a leitura <b>não alcança</b>. Vem do escritor, e some da tela
        ''' só quando não há leitura nenhuma para ressalvar.
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
                Return Not String.IsNullOrWhiteSpace(_ressalva)
            End Get
        End Property

        ''' <summary>
        ''' O aviso de contato repetido. Não impede criar — ver o comentário da
        ''' classe.
        ''' </summary>
        Public Property Aviso As String
            Get
                Return _aviso
            End Get
            Private Set(value As String)
                SetProperty(_aviso, If(value, ""))
                OnPropertyChanged(NameOf(TemAviso))
            End Set
        End Property

        Public ReadOnly Property TemAviso As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(_aviso)
            End Get
        End Property

        Public Property Erro As String
            Get
                Return _erro
            End Get
            Private Set(value As String)
                SetProperty(_erro, If(value, ""))
                OnPropertyChanged(NameOf(TemErro))
            End Set
        End Property

        Public ReadOnly Property TemErro As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(_erro)
            End Get
        End Property

        Public Property Resumo As String
            Get
                Return _resumo
            End Get
            Private Set(value As String)
                SetProperty(_resumo, If(value, ""))
            End Set
        End Property

        ''' <summary>
        ''' <b>ETAPA UM: propor.</b> Copia o remetente para o formulário, e
        ''' <b>para aí</b>.
        '''
        ''' Não cria. É a mesma separação da Fase 5, pelo mesmo motivo: a única
        ''' coisa que impede o catálogo de encher de ficha que ninguém pediu é
        ''' haver um segundo passo, feito por uma pessoa.
        '''
        ''' Determinístico e local — nome e endereço já vieram na mensagem. O
        ''' assistente não participa, e o contato não sai da máquina.
        ''' </summary>
        Public Sub ProporDoRemetente(nome As String, email As String)
            NovoNome = If(nome, "").Trim()
            NovoEmail = If(email, "").Trim()
            NovaEmpresa = ""
            Erro = ""

            If String.IsNullOrWhiteSpace(NovoNome) AndAlso Not String.IsNullOrWhiteSpace(NovoEmail) Then
                ' Sem nome de exibicao, o endereco e o unico nome que existe --
                ' e uma ficha com endereco no nome e melhor que uma ficha em
                ' branco, que o escritor recusaria depois do clique.
                NovoNome = NovoEmail
            End If

            ReavaliarDuplicata()
        End Sub

        Public Async Function AbrirAsync() As Task
            If _disposed OrElse _carregando Then Return

            Erro = ""
            _carregando = True
            OnPropertyChanged(NameOf(Carregando))
            AvisarComandos()

            Dim r As OperationResult(Of FolderKey)
            Try
                r = Await _broker.GetDefaultContactsFolderAsync(CancellationToken.None)
            Finally
                _carregando = False
                OnPropertyChanged(NameOf(Carregando))
                AvisarComandos()
            End Try

            If _disposed Then Return

            If Not r.Succeeded Then
                Erro = "não consegui achar a pasta de Contatos (" & r.Kind.ToString() & ")."
                Return
            End If

            _pasta = r.Value
            OnPropertyChanged(NameOf(TemPasta))
            OnPropertyChanged(NameOf(PodeCriar))
            AvisarComandos()

            Await CarregarAsync()
        End Function

        Public Async Function CarregarAsync() As Task
            If _pasta Is Nothing OrElse _disposed Then Return

            Dim minha = Interlocked.Increment(_geracao)
            _carregando = True
            OnPropertyChanged(NameOf(Carregando))
            AvisarComandos()

            Try
                Dim r = Await _broker.GetContactsAsync(_pasta, Teto, CancellationToken.None)

                ' SO A LEITURA MAIS NOVA ESCREVE NA TELA. Resposta atrasada de
                ' uma leitura anterior sobrescreveria a atual.
                If _disposed OrElse minha <> _geracao Then Return

                If Not r.Succeeded Then
                    Erro = "não consegui ler os contatos (" & r.Kind.ToString() & ")."
                    Resumo = ""
                    ' A RESSALVA FICA. Falha na leitura e o momento em que menos
                    ' se pode deixar de dizer que o GAL esta fora do alcance.
                    Ressalva = RegrasDeContato.ForaDoAlcance
                    Return
                End If

                Erro = ""
                Contatos.Clear()
                For Each c In r.Value.Items
                    Contatos.Add(New LinhaDeContato(c))
                Next
                Resumo = Descrever(r.Value)
                Ressalva = r.Value.ForaDoAlcance
                ReavaliarDuplicata()
            Catch ex As Exception
                If _disposed OrElse minha <> _geracao Then Return
                Erro = "não consegui ler os contatos (" & ex.GetType().Name & ")."
                Ressalva = RegrasDeContato.ForaDoAlcance
            Finally
                If minha = _geracao Then
                    _carregando = False
                    OnPropertyChanged(NameOf(Carregando))
                    AvisarComandos()
                End If
            End Try
        End Function

        ''' <summary>
        ''' <b>ETAPA DOIS: criar.</b> Um contato, deste formulário, com este
        ''' clique. Nunca em lote.
        '''
        ''' <b>Pública, e não privada</b>, pelo mesmo motivo que as mutações da
        ''' Fase 5: o <c>AsyncRelayCommand</c> serializa o que passa por ele, e
        ''' uma guarda alcançável só pelo botão é uma guarda que o teste não
        ''' consegue exercitar — prova o toolkit, e não esta classe.
        ''' </summary>
        Public Async Function CriarAsync() As Task
            If Not PodeCriar OrElse _disposed Then Return

            Dim rascunho As New ContactDraft With {
                .Nome = _novoNome,
                .Email = _novoEmail,
                .Empresa = _novaEmpresa
            }

            _gravando = True
            AvisarComandos()
            Erro = ""

            Try
                Dim r = Await _broker.CreateContactAsync(_pasta, rascunho, CancellationToken.None)

                ' O contato criado nao se desfaz -- nem deve. O que nao pode e a
                ' continuacao mexer numa tela que ja foi embora.
                If _disposed Then Return

                If Not r.Succeeded Then
                    Erro = If(String.IsNullOrWhiteSpace(r.Detail),
                              "não consegui criar o contato (" & r.Kind.ToString() & ").",
                              r.Detail)
                    Return
                End If

                NovoNome = ""
                NovoEmail = ""
                NovaEmpresa = ""
                Aviso = ""
            Catch ex As Exception
                If _disposed Then Return
                Erro = "não consegui criar o contato (" & ex.GetType().Name & ")."
                Return
            Finally
                _gravando = False
                AvisarComandos()
            End Try

            Await CarregarAsync()
        End Function

        ''' <summary>
        ''' <b>O resumo, que nunca afirma ausência.</b>
        '''
        ''' Zero contatos LIDOS, e não "nenhum contato" — a ressalva ao lado
        ''' explica a diferença, e o número aqui não pode contradizê-la.
        ''' </summary>
        Private Shared Function Descrever(lista As ContactList) As String
            Dim partes As New List(Of String)()

            If lista.Items.Count = 0 Then
                partes.Add("nenhum contato LIDO nesta pasta")
            Else
                partes.Add($"{lista.Items.Count} contato(s) lido(s)")
            End If

            If lista.Truncada Then
                partes.Add(If(String.IsNullOrWhiteSpace(lista.MotivoDoCorte),
                              "a leitura foi truncada", lista.MotivoDoCorte))
            End If

            ' NOTHING E ZERO SAO COISAS DIFERENTES: "nao contei" nao vira
            ' "contei e nao houve".
            If Not lista.Skipped.HasValue Then
                partes.Add("não sei quantos itens foram recusados")
            ElseIf lista.Skipped.Value > 0 Then
                partes.Add($"{lista.Skipped.Value} item(ns) recusado(s)")
            End If

            Return String.Join(" — ", partes)
        End Function

        ''' <summary>
        ''' Refaz o aviso de repetido. Chamado quando o endereço muda e quando a
        ''' lista é relida — as duas coisas mudam a resposta.
        ''' </summary>
        Private Sub ReavaliarDuplicata()
            If String.IsNullOrWhiteSpace(_novoEmail) Then
                Aviso = ""
                Return
            End If

            Dim achado = RegrasDeContato.Procurar(
                Contatos.Select(Function(l) l.Origem), _novoEmail)

            If achado Is Nothing Then
                Aviso = ""
                Return
            End If

            ' AVISO, E NAO RECUSA. E "ja ha um contato lido com este endereco",
            ' e nao "este contato ja existe": a busca so viu o que esta leitura
            ' trouxe.
            Aviso = "já há um contato lido com este endereço (" &
                    If(achado.Nome, "sem nome legível") & "). Criar de novo " &
                    "deixa duas fichas da mesma pessoa no catálogo."
        End Sub

        Private Sub AvisarComandos()
            OnPropertyChanged(NameOf(PodeCriar))
            AbrirCommand.NotifyCanExecuteChanged()
            AtualizarCommand.NotifyCanExecuteChanged()
            CriarCommand.NotifyCanExecuteChanged()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            _disposed = True
            Interlocked.Increment(_geracao)
        End Sub
    End Class

    ''' <summary>Uma linha da lista de contatos.</summary>
    Public NotInheritable Class LinhaDeContato

        Public Sub New(origem As ContactInfo)
            Me.Origem = origem
        End Sub

        ''' <summary>
        ''' O contato como veio. Guardado porque a busca por repetido precisa do
        ''' <c>Email</c> anulável de verdade — e não do texto que a tela mostra,
        ''' onde ausência e ilegibilidade já viraram a mesma frase.
        ''' </summary>
        Public ReadOnly Property Origem As ContactInfo

        Public ReadOnly Property Nome As String
            Get
                Return Legivel(Origem?.Nome, "sem nome")
            End Get
        End Property

        Public ReadOnly Property Email As String
            Get
                Return Legivel(Origem?.Email, "sem e-mail")
            End Get
        End Property

        Public ReadOnly Property Empresa As String
            Get
                Return Legivel(Origem?.Empresa, "")
            End Get
        End Property

        ''' <summary>
        ''' <c>Nothing</c> é "não consegui ler"; vazio é "está vazio". A tela
        ''' diz as duas coisas com palavras diferentes, porque são diferentes.
        ''' </summary>
        Private Shared Function Legivel(valor As String, vazio As String) As String
            If valor Is Nothing Then Return "não consegui ler"
            If String.IsNullOrWhiteSpace(valor) Then Return vazio
            Return valor
        End Function
    End Class

End Namespace

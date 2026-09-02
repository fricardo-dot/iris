Imports System.Diagnostics
Imports System.IO
Imports System.Threading
Imports CommunityToolkit.Mvvm.ComponentModel
Imports CommunityToolkit.Mvvm.Input
Imports Iris.Update

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>"Verificar atualizações" — e nada além de verificar.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ELE NÃO INSTALA, E ISSO É DESENHO</b>
    '''
    ''' Um atualizador que se substitui sozinho precisa encerrar o programa que
    ''' está rodando, trocar o executável que o Windows tem aberto, e subir de
    ''' novo — e no meio disso há um instante em que não existe Iris nenhum no
    ''' disco. Uma queda de energia ali deixa a máquina sem o programa. Dá para
    ''' fazer direito, com um segundo executável que faz a troca; é bem mais
    ''' máquina do que uma pessoa com duas máquinas precisa.
    '''
    ''' Aqui ele baixa, <b>confere a assinatura e o hash</b>, e diz onde o
    ''' arquivo ficou. O clique duplo é do dono. O que se perde é conveniência; o
    ''' que não se perde é que <b>nada recebe o nome final</b> sem o hash bater
    ''' com o manifesto que o dono assinou. O pacote inteiro chega ao disco antes
    ''' disso — num arquivo temporário de nome imprevisível, que é apagado em
    ''' qualquer recusa.
    '''
    ''' <b>E a garantia é pontual.</b> Ela vale no instante da última conferência.
    ''' Depois que o caminho aparece na tela, o arquivo é um arquivo como outro
    ''' qualquer, e o Iris não o vigia — <i>escolhe</i> não vigiar; daria para
    ''' reconferir antes de mostrar, ou segurar o arquivo aberto, e nenhuma das
    ''' duas coisas alcançaria o duplo clique, que é onde importa. O que
    ''' alcançaria é assinatura Authenticode no executável, e isso é outro
    ''' assunto — está em LANCAR.md, na seção do que a assinatura não compra.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E ELE NÃO PERGUNTA SOZINHO</b>
    '''
    ''' Não há verificação no arranque nem temporizador. Um programa que fala com
    ''' um servidor sem ninguém pedir é um programa que anuncia, a cada abertura,
    ''' que aquela máquina existe e está com o Iris ligado. Aqui a rede só é
    ''' tocada quando alguém clica.
    ''' </summary>
    Public NotInheritable Class AtualizacaoViewModel
        Inherits ObservableObject
        Implements IDisposable

        ''' <summary><c>Nothing</c> quando a chave de assinatura ainda não foi
        ''' configurada — ver <c>ChaveDeAtualizacao</c>.</summary>
        Private ReadOnly _procura As ProcuraDeVersao
        Private ReadOnly _pasta As String

        ''' <summary>
        ''' O manifesto da última procura BEM-SUCEDIDA, e só enquanto ele ainda
        ''' descreve o que a tela está mostrando. Guardá-lo além disso deixaria o
        ''' botão "Baixar" apontando para uma versão que a frase não menciona
        ''' mais.
        ''' </summary>
        Private _oferta As ManifestoDeVersao

        ''' <summary>
        ''' <b>Cancela o que estiver em voo quando a janela fecha.</b>
        '''
        ''' Sem isto, uma procura ou um download em andamento continuava vivo
        ''' depois do descarte e voltava para escrever <c>Frase</c> e
        ''' <c>Baixado</c> num ViewModel que já não é de ninguém. Não derrubava o
        ''' programa — a janela já tinha ido — mas mantinha um download de 60 MB
        ''' correndo depois de o dono mandar fechar.
        ''' </summary>
        Private ReadOnly _ateFechar As New CancellationTokenSource()
        Private _descartado As Boolean

        Public Sub New(procuraDeVersao As ProcuraDeVersao, pastaDeDestino As String)
            _procura = procuraDeVersao
            _pasta = If(pastaDeDestino, "")

            ' "E NAO ESTA OCUPADO" NOS DOIS.
            '
            ' Cada AsyncRelayCommand ja impede a propria reentrancia, e os dois
            ' sao comandos DIFERENTES: dava para clicar "Verificar" no meio de um
            ' download. Esquecer() limpava a oferta enquanto o download dela
            ' continuava, e a tela terminava mostrando a frase de uma procura ao
            ' lado do arquivo da outra.
            ' "E NAO ESTA DESCARTADA" nos tres: os comandos sao objetos publicos,
            ' e uma referencia guardada por um binding do WPF sobrevive ao
            ' Dispose. Sem isto, ExecuteAsync depois do descarte tocava
            ' _ateFechar.Token -- que LANCA ObjectDisposedException -- e antes
            ' disso ja tinha escrito Ocupado e Frase num objeto de ninguem.
            VerificarCommand = New AsyncRelayCommand(
                AddressOf VerificarAsync,
                Function() _procura IsNot Nothing AndAlso Not Ocupado AndAlso Not _descartado)
            BaixarCommand = New AsyncRelayCommand(
                AddressOf BaixarAsync,
                Function() _oferta IsNot Nothing AndAlso Not Ocupado AndAlso Not _descartado)
            MostrarNaPastaCommand = New RelayCommand(
                AddressOf MostrarNaPasta,
                Function() Baixado.Length > 0 AndAlso Not _descartado)

            _frase = If(_procura Is Nothing,
                        "A verificação de versões ainda não foi configurada nesta " &
                        "compilação: falta a chave pública de assinatura.",
                        "")
        End Sub

        ''' <summary>A versão deste executável, que a tela mostra sempre.</summary>
        Public ReadOnly Property Instalada As String =
            "Versão " & ProcuraDeVersao.VersaoInstalada().ToString(3)

        Private _frase As String = ""

        ''' <summary>O que a última procura concluiu, em português.</summary>
        Public Property Frase As String
            Get
                Return _frase
            End Get
            Set
                SetProperty(_frase, If(Value, ""))
            End Set
        End Property

        Private _notas As String = ""

        ''' <summary>O que mudou na versão oferecida. Vem do manifesto assinado.</summary>
        Public Property Notas As String
            Get
                Return _notas
            End Get
            Set
                SetProperty(_notas, If(Value, ""))
                OnPropertyChanged(NameOf(TemNotas))
            End Set
        End Property

        Public ReadOnly Property TemNotas As Boolean
            Get
                Return Notas.Length > 0
            End Get
        End Property

        Private _haVersaoNova As Boolean

        Public Property HaVersaoNova As Boolean
            Get
                Return _haVersaoNova
            End Get
            Set
                SetProperty(_haVersaoNova, Value)
            End Set
        End Property

        Private _ocupado As Boolean

        Public Property Ocupado As Boolean
            Get
                Return _ocupado
            End Get
            Set
                If Not SetProperty(_ocupado, Value) Then Return
                ' OS DOIS COMANDOS DEPENDEM DISTO, entao os dois sao reavaliados.
                ' Sem o aviso, o CanExecute continuaria devolvendo o valor de
                ' antes ate algo mais mexer na tela.
                VerificarCommand?.NotifyCanExecuteChanged()
                BaixarCommand?.NotifyCanExecuteChanged()
            End Set
        End Property

        Private _baixado As String = ""

        ''' <summary>Onde o pacote conferido ficou. Vazio enquanto não há um.</summary>
        Public Property Baixado As String
            Get
                Return _baixado
            End Get
            Set
                SetProperty(_baixado, If(Value, ""))
                OnPropertyChanged(NameOf(TemBaixado))
                MostrarNaPastaCommand.NotifyCanExecuteChanged()
            End Set
        End Property

        Public ReadOnly Property TemBaixado As Boolean
            Get
                Return Baixado.Length > 0
            End Get
        End Property

        Public ReadOnly Property VerificarCommand As AsyncRelayCommand
        Public ReadOnly Property BaixarCommand As AsyncRelayCommand
        Public ReadOnly Property MostrarNaPastaCommand As RelayCommand

        ''' <summary>
        ''' <b>Toda procura começa apagando a anterior.</b>
        '''
        ''' Sem isto, uma segunda procura que falhe na rede deixaria na tela a
        ''' frase nova ao lado do botão "Baixar" da procura antiga — e o botão
        ''' funcionaria, baixando a versão de antes. O estado da tela tem de
        ''' descrever <i>uma</i> procura, e é sempre a última.
        ''' </summary>
        Private Async Function VerificarAsync() As Task
            If _procura Is Nothing OrElse _descartado Then Return

            Esquecer()
            Ocupado = True
            Frase = "Perguntando…"
            Try
                Dim r = Await _procura.Procurar(_ateFechar.Token)
                If _descartado Then Return
                Frase = r.Frase

                If r.Desfecho = DesfechoDaProcura.HaVersaoNova AndAlso
                   r.Manifesto IsNot Nothing Then
                    _oferta = r.Manifesto
                    Notas = r.Manifesto.Notas
                    HaVersaoNova = True
                End If
            Finally
                ' O Finally RODA MESMO DEPOIS DO Return la em cima -- e por isso
                ' ele tambem precisa da guarda. Sem ela, Ocupado = False disparava
                ' PropertyChanged e NotifyCanExecuteChanged num ViewModel ja
                ' descartado, que e exatamente o que o Dispose diz impedir.
                If Not _descartado Then
                    Ocupado = False
                    BaixarCommand.NotifyCanExecuteChanged()
                End If
            End Try
        End Function

        Private Async Function BaixarAsync() As Task
            Dim oQue = _oferta
            If oQue Is Nothing OrElse _descartado Then Return

            Ocupado = True
            Frase = "Baixando a versão " & oQue.Versao.ToString(3) & "…"
            Try
                Dim pacote = Await _procura.Baixar(oQue, _pasta, _ateFechar.Token)
                If _descartado Then Return
                If pacote.Veio Then
                    Baixado = pacote.Caminho
                    ' A FRASE DIZ O QUE FALTA FAZER. "Baixado com sucesso" e
                    ' verdade e nao ajuda: o programa novo so passa a existir
                    ' quando alguem executa o arquivo, e quem le a tela tem de
                    ' saber que essa parte e dele.
                    Frase = "O pacote está conferido e salvo. Feche o Iris e " &
                            "execute o arquivo para atualizar."
                Else
                    Frase = "Não deu para baixar: " & pacote.Motivo
                End If
            Finally
                If Not _descartado Then Ocupado = False
            End Try
        End Function

        Private Sub Esquecer()
            _oferta = Nothing
            HaVersaoNova = False
            Notas = ""
            Baixado = ""
            BaixarCommand.NotifyCanExecuteChanged()
        End Sub

        ''' <summary>
        ''' Abre o Explorer com o arquivo selecionado. <b>Nunca o executa</b>:
        ''' um botão que roda o instalador seria o passo que este desenho
        ''' deliberadamente não dá.
        ''' </summary>
        Private Sub MostrarNaPasta()
            If Baixado.Length = 0 Then Return
            Try
                Dim argumento = "/select," & Chr(34) & Baixado & Chr(34)
                Process.Start(New ProcessStartInfo("explorer.exe", argumento) With {
                    .UseShellExecute = True
                })
            Catch
                ' O caminho ja esta na tela; nao poder abrir a pasta nao e
                ' motivo para derrubar o programa.
                Frase = "O arquivo está em " & Baixado
            End Try
        End Sub

        ''' <summary>
        ''' Cancela o que estiver em voo e marca a tela como descartada, para as
        ''' continuações que já estiverem depois do <c>Await</c> não escreverem
        ''' num ViewModel que ninguém mais mostra.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            If _descartado Then Return
            _descartado = True
            Try
                _ateFechar.Cancel()
            Catch
            End Try
            _ateFechar.Dispose()
        End Sub

        ''' <summary>
        ''' Onde os pacotes ficam: a pasta de Downloads do usuário, que é onde
        ''' alguém procura um instalador baixado. Se ela não existir — perfil
        ''' redirecionado, máquina gerenciada — cai em Documentos, que existe
        ''' sempre.
        ''' </summary>
        Public Shared Function PastaPadrao() As String
            Dim perfil = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Dim ondeBaixa = Path.Combine(perfil, "Downloads")
            If Directory.Exists(ondeBaixa) Then Return ondeBaixa
            Return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End Function

    End Class

End Namespace

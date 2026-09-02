Imports System.Windows
Imports Iris.App.ViewModels
Imports Iris.Core

''' <summary>
''' Composition root.
'''
''' É o ÚNICO lugar do Iris.App autorizado a conhecer Iris.Outlook. Todo o
''' resto — janelas e ViewModels — fala apenas com IOutlookBroker, definido
''' em Iris.Core. O teste arquitetural cobra isso, e a fronteira existe
''' porque a Fase 0 mostrou como é fácil um RCW escapar.
''' </summary>
Class Application

    Private _broker As Iris.Outlook.OutlookBroker
    Private _viewModel As MainViewModel
    Private _log As ILog
    ' A verificacao de versoes. Nasce so quando ha chave publica configurada,
    ' e morre no Exit junto com o resto.
    '
    ' O HttpClient e DELA, e nao daqui: monta-lo neste arquivo faria o
    ' assembly do Iris.App referenciar System.Net.Http, e a lista de quem
    ' pode abrir socket passaria de dois para tres. Quem contou foi o
    ' EgressArquiteturaTests, na primeira tentativa.
    Private _procuraDeVersoes As Update.ProcuraDeVersao

    Private Sub Application_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
        ' O LOG NAO PODE SER O QUE IMPEDE O PROGRAMA DE ABRIR.
        '
        ' O construtor do FileLog cria a pasta, e criar pasta falha: perfil
        ' temporario defeituoso, %LOCALAPPDATA% redirecionado e fora do ar, ACL
        ' corporativa, disco cheio. Fora do Try, isso derrubava o processo ANTES
        ' da janela e antes da caixa de erro -- e a caixa de erro manda consultar
        ' um log que, nesse caminho, nem chegou a existir.
        '
        ' Sem log o Iris roda pior; sem abrir, nao roda. Achado por revisao
        ' externa em 01/09/2026.
        Try
            _log = New FileLog(FileLog.DefaultPath())
        Catch
            _log = New NullLog()
        End Try
        _log.Write(LogLevel.Info, "app.startup", "iniciando")

        ' Tier 0 = renderizacao por software: TUDO fica travado, e nenhuma
        ' otimizacao de codigo compensa. Vale saber antes de culpar o codigo.
        _log.Write(LogLevel.Info, "app.render",
                   $"tier={Media.RenderCapability.Tier >> 16}")

        _broker = New Iris.Outlook.OutlookBroker(_log)

        Try
            _broker.Start()
        Catch ex As Exception
            ' Sem broker não há aplicativo. Falhar aqui em silêncio deixaria
            ' uma janela morta na tela.
            _log.Write(LogLevel.Error, "app.startup", ex.GetType().Name)
            MessageBox.Show(
                "Não foi possível iniciar o serviço interno do Iris." & Environment.NewLine &
                "Detalhes no log do aplicativo.",
                "Iris", MessageBoxButton.OK, MessageBoxImage.Error)
            Shutdown(1)
            Return
        End Try

        ' A COMPOSICAO TAMBEM PRECISA DE CERCA, e ela nao tinha.
        '
        ' O Try acima terminava no Start() do broker. Montar o MainViewModel, a
        ' janela e mostra-la ficava fora, e ali ha coisa que falha em maquina de
        ' verdade: o diario de buscas cria um MUTEX NOMEADO no construtor, e
        ' perfil corporativo com ACL restritiva no namespace de objetos do Windows
        ' recusa. Sem cerca, o Iris morria com excecao nao tratada -- sem a caixa
        ' de erro que este bloco inteiro existe para dar, e com o broker JA
        ' INICIADO, deixando a STA e os RCW vivos: OUTLOOK.EXE orfao (R7).
        '
        ' Achado por revisao externa em 02/09/2026.
        Try
            ' A partir daqui só o contrato circula.
            Dim broker As IOutlookBroker = _broker
            _viewModel = New MainViewModel(broker, Dispatcher,
                                           New WindowsSaveFileService(),
                                           New WindowsPickFileService(),
                                           log:=_log,
                                           atualizacao:=MontarAtualizacao())

            Dim janela As New MainWindow With {.DataContext = _viewModel}
            janela.Show()
        Catch ex As Exception
            _log.Write(LogLevel.Error, "app.startup",
                       "a janela nao pode ser montada: " & ex.GetType().Name)

            ' O QUE JA FOI CRIADO SAI ANTES DA CAIXA DE ERRO. Descartar depois
            ' seria descartar depois de o dono clicar em OK -- e ele pode demorar,
            ' ou fechar a caixa pelo X.
            Try
                _viewModel?.Dispose()
            Catch
            End Try
            _viewModel = Nothing

            MessageBox.Show(
                "Não foi possível abrir a janela do Iris." & Environment.NewLine &
                "Detalhes no log do aplicativo.",
                "Iris", MessageBoxButton.OK, MessageBoxImage.Error)

            ' O broker sai pelo Application_Exit, que roda depois do Shutdown:
            ' ele libera os COM na propria STA, e e la que isso tem de acontecer.
            Shutdown(1)
            Return
        End Try

        ' Primeira conexão sem bloquear a abertura da janela: se o Outlook
        ' estiver ocupado, o usuário vê a tela e o estado, não uma janela
        ' congelada.
        _viewModel.Connection.Observe(_viewModel.InitializeAsync(), "app.initialize")
    End Sub

    ''' <summary>
    ''' <b>Monta a verificação de versões — ou não monta nenhuma.</b>
    '''
    ''' Sem chave pública e sem endereço não há o que verificar, e a tela diz
    ''' isso. A alternativa seria montar assim mesmo e deixar toda resposta virar
    ''' "a assinatura não confere" — a frase de <i>alguém trocou o arquivo</i>
    ''' dita quando o que houve foi <i>ninguém configurou ainda</i>.
    '''
    ''' <b>Aqui, e não no ViewModel</b>: construir um <c>HttpClient</c> é adquirir
    ''' a capacidade de rede, e a regra da base é que isso aconteça no composition
    ''' root, onde está visível. Ver <c>EgressArquiteturaTests</c>.
    ''' </summary>
    Private Function MontarAtualizacao() As ViewModels.AtualizacaoViewModel
        Dim pasta = ViewModels.AtualizacaoViewModel.PastaPadrao()

        If Not ChaveDeAtualizacao.Configurada Then
            _log.Write(LogLevel.Info, "app.startup",
                       "verificacao de versoes desligada: falta a chave publica")
            Return New ViewModels.AtualizacaoViewModel(Nothing, pasta)
        End If

        Dim chave = ChaveDeAtualizacao.Bytes()
        If chave.Length = 0 Then
            ' CHAVE PREENCHIDA E ILEGIVEL E PIOR QUE VAZIA, porque alguem
            ' acreditou ter configurado. Vai para o log como erro, e nao como
            ' informacao.
            _log.Write(LogLevel.Error, "app.startup",
                       "a chave publica de atualizacao esta preenchida e nao e legivel")
            Return New ViewModels.AtualizacaoViewModel(Nothing, pasta)
        End If

        _procuraDeVersoes = New Update.ProcuraDeVersao(
            ChaveDeAtualizacao.Endereco, chave)

        Return New ViewModels.AtualizacaoViewModel(_procuraDeVersoes, pasta)
    End Function

    Private Sub Application_Exit(sender As Object, e As ExitEventArgs) Handles Me.Exit
        ' Encerramento ordenado: o broker libera os objetos COM na própria
        ' thread STA, com o message pump ainda vivo. Pular isto é o caminho
        ' para um OUTLOOK.EXE órfão (R7).
        ' DOIS Try, e nao um. Estavam juntos, e uma excecao no descarte da
        ' janela pulava o do broker inteiro -- deixando a STA e os RCW vivos,
        ' que e exatamente o OUTLOOK.EXE orfao que este bloco existe para
        ' evitar (R7). O descarte do broker nao pode depender de o da janela
        ' ter corrido bem.
        Try
            _viewModel?.Dispose()
        Catch ex As Exception
            _log?.Write(LogLevel.Error, "app.exit", "descarte da janela falhou: " & ex.GetType().Name)
        End Try
        Try
            _procuraDeVersoes?.Dispose()
        Catch ex As Exception
            _log?.Write(LogLevel.Error, "app.exit",
                        "descarte da procura de versoes falhou: " & ex.GetType().Name)
        End Try
        Try
            _broker?.Dispose()
        Catch ex As Exception
            _log?.Write(LogLevel.Error, "app.exit", "descarte do broker falhou: " & ex.GetType().Name)
        End Try
        _log?.Write(LogLevel.Info, "app.exit", "encerrado")
    End Sub

End Class

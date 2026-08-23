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

    Private Sub Application_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
        _log = New FileLog(FileLog.DefaultPath())
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

        ' A partir daqui só o contrato circula.
        Dim broker As IOutlookBroker = _broker
        _viewModel = New MainViewModel(broker, Dispatcher,
                                       New WindowsSaveFileService(), New WindowsPickFileService())

        Dim janela As New MainWindow With {.DataContext = _viewModel}
        janela.Show()

        ' Primeira conexão sem bloquear a abertura da janela: se o Outlook
        ' estiver ocupado, o usuário vê a tela e o estado, não uma janela
        ' congelada.
        _viewModel.Connection.Observe(_viewModel.InitializeAsync(), "app.initialize")
    End Sub

    Private Sub Application_Exit(sender As Object, e As ExitEventArgs) Handles Me.Exit
        ' Encerramento ordenado: o broker libera os objetos COM na própria
        ' thread STA, com o message pump ainda vivo. Pular isto é o caminho
        ' para um OUTLOOK.EXE órfão (R7).
        Try
            _viewModel?.Dispose()
            _broker?.Dispose()
        Catch
        End Try
        _log?.Write(LogLevel.Info, "app.exit", "encerrado")
    End Sub

End Class

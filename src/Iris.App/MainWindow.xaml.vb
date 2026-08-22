Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Interop
Imports System.Windows.Media

''' <summary>
''' Barra de título própria.
'''
''' Nenhuma lógica de aplicação vive aqui — a janela é ligada ao
''' ConnectionViewModel por DataContext. O que existe é a plumbing Win32 que
''' o chrome customizado exige, e que não tem como ser feita em XAML.
'''
''' Decisões que valem registrar:
'''
'''   • <c>SystemCommands</c> em vez de mexer em <c>WindowState</c> na mão:
'''     é o que preserva a semântica do sistema.
'''   • <c>WM_NCHITTEST</c> devolvendo <c>HTMAXBUTTON</c> sobre o botão
'''     maximizar. Sem isso, o Snap Layouts do Windows 11 — o painel que
'''     aparece ao pousar o mouse ali — simplesmente não abre. Em troca, o
'''     WPF deixa de receber os eventos de mouse do botão, então o clique e
'''     o realce precisam ser tratados aqui.
'''   • <c>WM_GETMINMAXINFO</c> por monitor. Maximizar sem isso faz a janela
'''     invadir a barra de tarefas. Usar SystemParameters.WorkArea seria
'''     mais simples e erraria em monitor secundário e em DPI misto.
''' </summary>
Class MainWindow

    Private Const WM_NCHITTEST As Integer = &H84
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const WM_NCLBUTTONUP As Integer = &HA2
    Private Const WM_NCMOUSELEAVE As Integer = &H2A2
    Private Const WM_GETMINMAXINFO As Integer = &H24

    Private Const HTCLIENT As Integer = 1
    Private Const HTMAXBUTTON As Integer = 9

    Private Const MONITOR_DEFAULTTONEAREST As Integer = 2

    ' Glifos do Segoe Fluent Icons / Segoe MDL2 Assets
    Private Const GlifoMaximizar As String = ChrW(&HE922)
    Private Const GlifoRestaurar As String = ChrW(&HE923)

    Private Const DarkModeAttribute As Integer = 20
    Private Const DarkModeAttributeLegacy As Integer = 19

    ' NativePoint, e nao POINT: VB e case-insensitive, entao o nome
    ' POINT eclipsaria System.Windows.Point dentro desta classe.
    <StructLayout(LayoutKind.Sequential)>
    Private Structure NativePoint
        Public x As Integer
        Public y As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure MINMAXINFO
        Public ptReserved As NativePoint
        Public ptMaxSize As NativePoint
        Public ptMaxPosition As NativePoint
        Public ptMinTrackSize As NativePoint
        Public ptMaxTrackSize As NativePoint
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure MONITORINFO
        Public cbSize As Integer
        Public rcMonitor As RECT
        Public rcWork As RECT
        Public dwFlags As Integer
    End Structure

    <DllImport("user32.dll")>
    Private Shared Function MonitorFromWindow(hwnd As IntPtr, flags As Integer) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetMonitorInfo(monitor As IntPtr, ByRef info As MONITORINFO) As Boolean
    End Function

    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attribute As Integer,
                                                  ByRef value As Integer, size As Integer) As Integer
    End Function

    Private _maximizeHover As Boolean

    ' ===================================================================
    ' Ciclo de vida da janela
    ' ===================================================================

    Private Sub MainWindow_SourceInitialized(sender As Object, e As EventArgs) Handles Me.SourceInitialized
        Dim hwnd = New WindowInteropHelper(Me).Handle

        ' A moldura ainda existe fora da área cliente; mantê-la escura evita
        ' uma linha clara em volta de um app escuro no Windows 11.
        Try
            Dim ligado As Integer = 1
            If DwmSetWindowAttribute(hwnd, DarkModeAttribute, ligado, 4) <> 0 Then
                DwmSetWindowAttribute(hwnd, DarkModeAttributeLegacy, ligado, 4)
            End If
        Catch
        End Try

        HwndSource.FromHwnd(hwnd)?.AddHook(AddressOf WndProc)
    End Sub

    Private Sub MainWindow_StateChanged(sender As Object, e As EventArgs) Handles Me.StateChanged
        Dim maximizada = WindowState = WindowState.Maximized
        MaximizeButton.Content = If(maximizada, GlifoRestaurar, GlifoMaximizar)
        MaximizeButton.ToolTip = If(maximizada, "Restaurar", "Maximizar")
        ' O leitor de tela precisa anunciar a ação atual, não a inicial.
        Automation.AutomationProperties.SetName(MaximizeButton, If(maximizada, "Restaurar", "Maximizar"))
    End Sub

    ' ===================================================================
    ' Botões de legenda
    ' ===================================================================

    Private Sub MinimizeButton_Click(sender As Object, e As RoutedEventArgs) Handles MinimizeButton.Click
        SystemCommands.MinimizeWindow(Me)
    End Sub

    Private Sub MaximizeButton_Click(sender As Object, e As RoutedEventArgs) Handles MaximizeButton.Click
        AlternarMaximizado()
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As RoutedEventArgs) Handles CloseButton.Click
        SystemCommands.CloseWindow(Me)
    End Sub

    Private Sub AlternarMaximizado()
        If WindowState = WindowState.Maximized Then
            SystemCommands.RestoreWindow(Me)
        Else
            SystemCommands.MaximizeWindow(Me)
        End If
    End Sub

    ''' <summary>
    ''' Clique direito na área de arrasto abre o menu do sistema, como faz
    ''' qualquer janela do Windows. Sem isto, o chrome próprio perderia um
    ''' comportamento que as pessoas usam sem pensar.
    ''' </summary>
    Private Sub MainWindow_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) _
        Handles Me.MouseRightButtonUp
        Dim ponto = PointToScreen(e.GetPosition(Me))
        If ponto.Y - Left > 0 Then
            SystemCommands.ShowSystemMenu(Me, ponto)
            e.Handled = True
        End If
    End Sub

    ' ===================================================================
    ' Win32
    ' ===================================================================

    Private Function WndProc(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr,
                             ByRef handled As Boolean) As IntPtr
        Select Case msg

            Case WM_GETMINMAXINFO
                AjustarLimitesDeMaximizacao(hwnd, lParam)

            Case WM_NCHITTEST
                ' O Snap Layouts do Windows 11 só aparece se o hit test
                ' informar que aquele ponto é o botão maximizar.
                If SobreBotaoMaximizar(lParam) Then
                    DefinirRealceMaximizar(True)
                    handled = True
                    Return New IntPtr(HTMAXBUTTON)
                End If
                DefinirRealceMaximizar(False)

            Case WM_NCMOUSELEAVE
                DefinirRealceMaximizar(False)

            Case WM_NCLBUTTONDOWN
                ' Devolver HTMAXBUTTON tira os eventos de mouse do WPF, então
                ' o clique precisa ser tratado aqui — senão o botão vira
                ' decoração que não faz nada.
                If wParam.ToInt32() = HTMAXBUTTON Then
                    handled = True
                End If

            Case WM_NCLBUTTONUP
                If wParam.ToInt32() = HTMAXBUTTON Then
                    AlternarMaximizado()
                    handled = True
                End If

        End Select

        Return IntPtr.Zero
    End Function

    ''' <summary>
    ''' Maximizar sem isto invade a área da barra de tarefas. O cálculo é
    ''' por MONITOR: usar SystemParameters.WorkArea usaria sempre o monitor
    ''' primário e erraria em monitor secundário e em escalas diferentes.
    ''' </summary>
    Private Sub AjustarLimitesDeMaximizacao(hwnd As IntPtr, lParam As IntPtr)
        Try
            Dim monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
            If monitor = IntPtr.Zero Then Return

            Dim info As New MONITORINFO With {.cbSize = Marshal.SizeOf(Of MONITORINFO)()}
            If Not GetMonitorInfo(monitor, info) Then Return

            Dim mmi = Marshal.PtrToStructure(Of MINMAXINFO)(lParam)
            mmi.ptMaxPosition.x = info.rcWork.Left - info.rcMonitor.Left
            mmi.ptMaxPosition.y = info.rcWork.Top - info.rcMonitor.Top
            mmi.ptMaxSize.x = info.rcWork.Right - info.rcWork.Left
            mmi.ptMaxSize.y = info.rcWork.Bottom - info.rcWork.Top
            Marshal.StructureToPtr(mmi, lParam, True)
        Catch
            ' Falhar aqui só devolve o comportamento padrão do Windows.
        End Try
    End Sub

    ''' <summary>
    ''' O retângulo do botão é recalculado a cada consulta, e não guardado:
    ''' ele muda com resize, maximização, mudança de DPI e troca de monitor.
    ''' </summary>
    Private Function SobreBotaoMaximizar(lParam As IntPtr) As Boolean
        If MaximizeButton Is Nothing OrElse Not MaximizeButton.IsVisible Then Return False

        Dim bruto = lParam.ToInt32()
        Dim tela As New Point(CShort(bruto And &HFFFF), CShort((bruto >> 16) And &HFFFF))

        Try
            Dim canto = MaximizeButton.PointToScreen(New Point(0, 0))
            Dim dpi = VisualTreeHelper.GetDpi(MaximizeButton)
            Dim largura = MaximizeButton.ActualWidth * dpi.DpiScaleX
            Dim altura = MaximizeButton.ActualHeight * dpi.DpiScaleY

            Return tela.X >= canto.X AndAlso tela.X < canto.X + largura AndAlso
                   tela.Y >= canto.Y AndAlso tela.Y < canto.Y + altura
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Com HTMAXBUTTON o WPF não recebe IsMouseOver, então o realce do
    ''' botão precisa ser aplicado à mão — senão ele fica morto justamente
    ''' quando o Snap Layouts está prestes a abrir.
    ''' </summary>
    Private Sub DefinirRealceMaximizar(ativo As Boolean)
        If _maximizeHover = ativo OrElse MaximizeButton Is Nothing Then Return
        _maximizeHover = ativo

        If ativo Then
            MaximizeButton.Background = TryCast(TryFindResource("Brush.Surface.3"), Brush)
            MaximizeButton.Foreground = TryCast(TryFindResource("Brush.Text.Primary"), Brush)
        Else
            MaximizeButton.Background = Brushes.Transparent
            MaximizeButton.Foreground = TryCast(TryFindResource("Brush.Text.Secondary"), Brush)
        End If
    End Sub

End Class

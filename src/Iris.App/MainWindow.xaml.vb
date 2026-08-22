Imports System.Runtime.InteropServices

''' <summary>
''' Sem lógica de aplicação: a janela é ligada ao ConnectionViewModel por
''' DataContext, definido no composition root.
'''
''' O único code-behind é a barra de título escura, que não tem como ser
''' feita em XAML: a moldura da janela pertence ao Windows, não ao WPF, e
''' sem isto uma barra branca fica gritando em cima de um app escuro.
''' </summary>
Class MainWindow

    ' DWMWA_USE_IMMERSIVE_DARK_MODE. O valor virou 20 no Windows 10 20H1;
    ' em builds anteriores era 19. Tentar os dois é mais barato que
    ' detectar versão.
    Private Const DarkModeAttribute As Integer = 20
    Private Const DarkModeAttributeLegacy As Integer = 19

    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attribute As Integer,
                                                  ByRef value As Integer, size As Integer) As Integer
    End Function

    Private Sub MainWindow_SourceInitialized(sender As Object, e As EventArgs) Handles Me.SourceInitialized
        Try
            Dim hwnd = New Global.System.Windows.Interop.WindowInteropHelper(Me).Handle
            Dim ligado As Integer = 1
            If DwmSetWindowAttribute(hwnd, DarkModeAttribute, ligado, 4) <> 0 Then
                DwmSetWindowAttribute(hwnd, DarkModeAttributeLegacy, ligado, 4)
            End If
        Catch
            ' Barra clara não impede o uso do aplicativo.
        End Try
    End Sub

End Class

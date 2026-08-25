Imports System.Collections.Generic
Imports System.Net
Imports System.Text
Imports System.Threading

''' <summary>
''' Um servidor HTTP <b>local</b> que faz o que o teste mandar.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE UM SERVIDOR DE VERDADE</b>
'''
''' Um <c>HttpMessageHandler</c> falso provaria o código em volta do
''' <c>HttpClient</c> e não o <c>HttpClient</c> — e as propriedades que o 3.4
''' precisa provar são justamente dele: redirect não seguido, timeout,
''' cancelamento, teto de resposta, corpo que chega.
'''
''' Aqui há socket, requisição e resposta. O que fica de fora é a rede: tudo
''' em <c>127.0.0.1</c>, e nenhum byte sai da máquina.
''' </summary>
Friend NotInheritable Class ServidorFalso
    Implements IDisposable

    Private ReadOnly _ouvinte As HttpListener
    Private ReadOnly _thread As Thread
    Private _parando As Boolean

    ''' <summary>Tudo o que chegou, para o teste conferir.</summary>
    Friend ReadOnly Recebidos As New List(Of Recebido)()

    ''' <summary>Como responder. O teste troca antes de chamar.</summary>
    Friend Property Codigo As Integer = 200
    Friend Property Corpo As String = "resposta do modelo"
    ''' <summary>Segura a resposta este tanto — para provar timeout.</summary>
    Friend Property Demora As TimeSpan = TimeSpan.Zero
    ''' <summary>Manda um redirect para cá, em vez de responder.</summary>
    Friend Property RedirecionarPara As String
    ''' <summary>Devolve este tanto de bytes, para provar o teto.</summary>
    Friend Property TamanhoDaResposta As Integer = 0

    Friend NotInheritable Class Recebido
        Public Property Corpo As Byte()
        Public Property Metodo As String
        Public Property Caminho As String
        Public Property Cabecalhos As Dictionary(Of String, String)
    End Class

    Friend ReadOnly Property Endereco As String

    Friend Sub New()
        Dim porta = PortaLivre()
        Endereco = $"http://127.0.0.1:{porta}/assist"
        _ouvinte = New HttpListener()
        _ouvinte.Prefixes.Add($"http://127.0.0.1:{porta}/")
        _ouvinte.Start()

        _thread = New Thread(AddressOf Atender) With {.IsBackground = True}
        _thread.Start()
    End Sub

    Private Shared Function PortaLivre() As Integer
        Dim l As New Sockets.TcpListener(IPAddress.Loopback, 0)
        l.Start()
        Dim p = CType(l.LocalEndpoint, IPEndPoint).Port
        l.Stop()
        Return p
    End Function

    Private Sub Atender()
        While Not _parando
            Dim ctx As HttpListenerContext
            Try
                ctx = _ouvinte.GetContext()
            Catch
                Return
            End Try

            Try
                ' 'corpo' eclipsaria a propriedade Corpo — VB e
                ' case-insensitive, e o CLAUDE.md ja lista sete ocorrencias.
                Dim recebido = LerTudo(ctx.Request)
                Dim cab As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                For Each nome In ctx.Request.Headers.AllKeys
                    cab(nome) = ctx.Request.Headers(nome)
                Next

                SyncLock Recebidos
                    Recebidos.Add(New Recebido With {
                        .Corpo = recebido, .Metodo = ctx.Request.HttpMethod,
                        .Caminho = ctx.Request.Url.AbsolutePath, .Cabecalhos = cab})
                End SyncLock

                If Demora > TimeSpan.Zero Then Thread.Sleep(Demora)

                If Not String.IsNullOrEmpty(RedirecionarPara) Then
                    ctx.Response.StatusCode = 302
                    ctx.Response.Headers("Location") = RedirecionarPara
                    ctx.Response.Close()
                    Continue While
                End If

                Dim texto As String = Corpo
                If TamanhoDaResposta > 0 Then texto = New String("z"c, TamanhoDaResposta)
                Dim bytes = Encoding.UTF8.GetBytes(texto)
                ctx.Response.StatusCode = Codigo
                ctx.Response.ContentLength64 = bytes.Length
                ctx.Response.OutputStream.Write(bytes, 0, bytes.Length)
                ctx.Response.Close()
            Catch
                Try
                    ctx.Response.Abort()
                Catch
                End Try
            End Try
        End While
    End Sub

    Private Shared Function LerTudo(r As HttpListenerRequest) As Byte()
        Using m As New IO.MemoryStream()
            r.InputStream.CopyTo(m)
            Return m.ToArray()
        End Using
    End Function

    Friend ReadOnly Property Ultimo As Recebido
        Get
            SyncLock Recebidos
                Return If(Recebidos.Count = 0, Nothing, Recebidos(Recebidos.Count - 1))
            End SyncLock
        End Get
    End Property

    Public Sub Dispose() Implements IDisposable.Dispose
        _parando = True
        Try
            _ouvinte.Stop()
            _ouvinte.Close()
        Catch
        End Try
    End Sub

End Class

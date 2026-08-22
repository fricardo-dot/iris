Imports System.Globalization
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports Iris.Model

Namespace Global.Iris.Core

    Public Enum LogLevel
        Debug
        Info
        Warn
        [Error]
    End Enum

    ''' <summary>
    ''' Política de log do Iris, valendo desde a primeira linha escrita —
    ''' não como faxina depois (F1-M).
    '''
    ''' O log NUNCA registra: corpo de mensagem, assunto completo, endereço
    ''' completo, nome de anexo ou caminho de arquivo do usuário. O log é uma
    ''' segunda cópia de dados corporativos fora do OST se deixarmos, e
    ''' herda todas as obrigações do R11 e do R14 sem nenhuma das proteções.
    '''
    ''' O que ele registra: operação, duração, HRESULT, tipo de erro, chave
    ''' encurtada e, quando é preciso correlacionar, um hash curto.
    ''' </summary>
    Public Module Redact

        ''' <summary>
        ''' Assunto vira comprimento e hash. Dá para correlacionar duas
        ''' ocorrências da mesma mensagem sem registrar o texto.
        ''' </summary>
        Public Function Subject(value As String) As String
            If String.IsNullOrEmpty(value) Then Return "assunto(vazio)"
            Return $"assunto(len={value.Length},h={ShortHash(value)})"
        End Function

        ''' <summary>
        ''' Endereço vira só o domínio. "fulano@empresa.com" → "@empresa.com".
        ''' O domínio basta para diagnosticar; a parte local identifica uma
        ''' pessoa.
        ''' </summary>
        Public Function Address(value As String) As String
            If String.IsNullOrEmpty(value) Then Return "endereco(vazio)"
            Dim at = value.LastIndexOf("@"c)
            If at < 0 Then Return $"endereco(h={ShortHash(value)})"
            Return "endereco(*" & value.Substring(at) & ")"
        End Function

        ''' <summary>Caminho vira só a extensão e o tamanho.</summary>
        Public Function FilePath(value As String) As String
            If String.IsNullOrEmpty(value) Then Return "arquivo(vazio)"
            Return $"arquivo(ext={Path.GetExtension(value)})"
        End Function

        Public Function ShortHash(value As String) As String
            If value Is Nothing Then value = String.Empty
            Using sha = SHA256.Create()
                Dim bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value))
                Return Convert.ToHexString(bytes, 0, 4).ToLowerInvariant()
            End Using
        End Function
    End Module

    Public Interface ILog
        Sub Write(level As LogLevel, operation As String, message As String)
    End Interface

    ''' <summary>
    ''' Log em arquivo, uma linha por evento. Rotação simples por tamanho.
    ''' </summary>
    Public NotInheritable Class FileLog
        Implements ILog

        Private Const MaxBytes As Long = 2 * 1024 * 1024

        Private ReadOnly _path As String
        Private ReadOnly _gate As New Object()

        ' Parametro NAO se chama "path": VB e case-insensitive e o nome
        ' eclipsaria a classe System.IO.Path dentro do construtor.
        Public Sub New(logPath As String)
            _path = logPath
            Directory.CreateDirectory(Path.GetDirectoryName(logPath))
        End Sub

        Public Shared Function DefaultPath() As String
            Dim dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "logs")
            Return Path.Combine(dir, "iris.log")
        End Function

        Public Sub Write(level As LogLevel, operation As String, message As String) Implements ILog.Write
            Dim line = String.Format(CultureInfo.InvariantCulture,
                                     "{0:yyyy-MM-dd HH:mm:ss.fff} {1,-5} {2,-28} {3}",
                                     DateTime.Now, level.ToString().ToUpperInvariant(),
                                     operation, message)
            SyncLock _gate
                Try
                    Rotate()
                    File.AppendAllText(_path, line & Environment.NewLine, Encoding.UTF8)
                Catch
                    ' Log nunca derruba o aplicativo. Nem mesmo com disco
                    ' cheio — esta máquina já provou que isso acontece.
                End Try
            End SyncLock
        End Sub

        Private Sub Rotate()
            Dim info As New FileInfo(_path)
            If Not info.Exists OrElse info.Length < MaxBytes Then Return
            Dim backup = _path & ".1"
            If File.Exists(backup) Then File.Delete(backup)
            File.Move(_path, backup)
        End Sub
    End Class

    ''' <summary>Descarta tudo. Para testes.</summary>
    Public NotInheritable Class NullLog
        Implements ILog
        Public Sub Write(level As LogLevel, operation As String, message As String) Implements ILog.Write
        End Sub
    End Class

End Namespace

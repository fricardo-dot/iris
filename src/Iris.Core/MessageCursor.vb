Imports System.Globalization
Imports System.Text
Imports Iris.Model

Namespace Global.Iris.Core

    Public Enum CursorMode
        ''' <summary>Keyset por ReceivedTime decrescente, via Table.</summary>
        ReceivedDesc
        ''' <summary>Iteração com offset. É o caminho lento, ainda usado
        ''' pelas ordenações que a Q1 não mediu.</summary>
        LegacyOffset
    End Enum

    ''' <summary>
    ''' Continuação de paginação, opaca para quem chama.
    '''
    ''' O cursor tem TAMANHO FIXO e não guarda chave nenhuma. Isso não é
    ''' economia: é consequência do algoritmo. A página drena o grupo do
    ''' último instante por inteiro, então quando ela termina não sobra
    ''' empatado para trás, e não há o que lembrar. Um cursor que precisasse
    ''' carregar chaves emitidas seria sinal de que o algoritmo mudou.
    '''
    ''' Ele carrega uma IMPRESSÃO da consulta — pasta, store, ordenação e
    ''' geração. Cursor de outra consulta é recusado em vez de produzir
    ''' página de outra pasta: a paginação é volátil por natureza, e somar a
    ''' isso um cursor trocado daria erro invisível.
    ''' </summary>
    Public NotInheritable Class MessageCursor

        Private Const Versao As String = "iris1"
        Private Const MaxTexto As Integer = 1024

        Public ReadOnly Property Mode As CursorMode
        ''' <summary>Fronteira, em UTC. Só no modo ReceivedDesc.</summary>
        Public ReadOnly Property Boundary As DateTimeOffset?
        ''' <summary>Posição bruta. Só no modo LegacyOffset.</summary>
        Public ReadOnly Property Offset As Integer

        Private ReadOnly _impressao As String

        Private Sub New(mode As CursorMode, boundary As DateTimeOffset?,
                        offset As Integer, impressao As String)
            Me.Mode = mode
            Me.Boundary = boundary
            Me.Offset = offset
            _impressao = impressao
        End Sub

        ''' <summary>
        ''' Impressão da consulta. Inclui a geração: recarregar invalida o
        ''' cursor antigo, que é o comportamento desejado.
        ''' </summary>
        Public Shared Function Fingerprint(query As MessageQuery) As String
            If query Is Nothing Then Return ""
            Dim pasta = If(query.Folder Is Nothing, "", query.Folder.EntryId & "\" & query.Folder.StoreId)
            Dim cru = $"{pasta}|{CInt(query.Sort)}|{query.Generation}"
            Using sha = Security.Cryptography.SHA256.Create()
                Dim bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(cru))
                Return Convert.ToHexString(bytes, 0, 8)
            End Using
        End Function

        Public Shared Function ForBoundary(query As MessageQuery, boundary As DateTimeOffset) As MessageCursor
            Return New MessageCursor(CursorMode.ReceivedDesc, boundary.ToUniversalTime(),
                                     0, Fingerprint(query))
        End Function

        Public Shared Function ForOffset(query As MessageQuery, offset As Integer) As MessageCursor
            Return New MessageCursor(CursorMode.LegacyOffset, Nothing,
                                     Math.Max(0, offset), Fingerprint(query))
        End Function

        Public Function Encode() As String
            Dim carga = If(Mode = CursorMode.ReceivedDesc,
                           Boundary.Value.UtcTicks.ToString(CultureInfo.InvariantCulture),
                           Offset.ToString(CultureInfo.InvariantCulture))
            Dim letra = If(Mode = CursorMode.ReceivedDesc, "D", "L")
            Dim cru = $"{Versao}|{letra}|{_impressao}|{carga}"
            Return Convert.ToBase64String(Encoding.UTF8.GetBytes(cru))
        End Function

        ''' <summary>
        ''' Decodifica e CONFERE que o cursor pertence a esta consulta.
        '''
        ''' Falha devolve False, não exceção nem cursor "quase certo": um
        ''' cursor de outra pasta produziria página de outra pasta, e a UI
        ''' não teria como perceber.
        ''' </summary>
        Public Shared Function TryDecode(text As String, query As MessageQuery,
                                         ByRef cursor As MessageCursor) As Boolean
            cursor = Nothing
            If String.IsNullOrEmpty(text) Then Return False
            If text.Length > MaxTexto Then Return False

            Dim cru As String
            Try
                cru = Encoding.UTF8.GetString(Convert.FromBase64String(text))
            Catch ex As FormatException
                Return False
            Catch ex As DecoderFallbackException
                Return False
            End Try

            Dim partes = cru.Split("|"c)
            If partes.Length <> 4 Then Return False
            If partes(0) <> Versao Then Return False
            If partes(2) <> Fingerprint(query) Then Return False

            Select Case partes(1)
                Case "D"
                    Dim ticks As Long
                    If Not Long.TryParse(partes(3), NumberStyles.Integer,
                                         CultureInfo.InvariantCulture, ticks) Then Return False
                    If ticks < 0 OrElse ticks > DateTime.MaxValue.Ticks Then Return False
                    cursor = New MessageCursor(CursorMode.ReceivedDesc,
                                               New DateTimeOffset(ticks, TimeSpan.Zero),
                                               0, partes(2))
                    Return True
                Case "L"
                    Dim posicao As Integer
                    If Not Integer.TryParse(partes(3), NumberStyles.Integer,
                                            CultureInfo.InvariantCulture, posicao) Then Return False
                    If posicao < 0 Then Return False
                    cursor = New MessageCursor(CursorMode.LegacyOffset, Nothing, posicao, partes(2))
                    Return True
                Case Else
                    Return False
            End Select
        End Function

    End Class

End Namespace

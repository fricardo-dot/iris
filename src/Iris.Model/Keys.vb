''' <summary>
''' Identificadores. Todos imutáveis, com igualdade por valor, e compostos
''' apenas de String e Integer.
'''
''' REGRA DO PROJETO: nenhum tipo em Iris.Model pode expor Object, dynamic
''' ou delegate. Um RCW é atribuível a Object, então bastaria uma
''' propriedade dessas para o COM atravessar a fronteira sem que este
''' assembly sequer conheça os tipos do Interop.
''' </summary>
Namespace Global.Iris.Model

    ''' <summary>
    ''' Identidade de um item no Outlook.
    '''
    ''' ATENÇÃO, medido na Fase 0 (critério D3): <c>EntryId</c> MUDA quando o
    ''' item é movido, mesmo dentro do mesmo store. Portanto esta chave
    ''' identifica um item *agora*, e não sobrevive a movimentação. A chave
    ''' interna estável é entregável da Fase 2.
    ''' </summary>
    Public NotInheritable Class ItemKey
        Implements IEquatable(Of ItemKey)

        Public ReadOnly Property EntryId As String
        Public ReadOnly Property StoreId As String

        Public Sub New(entryId As String, storeId As String)
            Me.EntryId = If(entryId, String.Empty)
            Me.StoreId = If(storeId, String.Empty)
        End Sub

        Public ReadOnly Property IsEmpty As Boolean
            Get
                Return EntryId.Length = 0
            End Get
        End Property

        Public Overloads Function Equals(other As ItemKey) As Boolean Implements IEquatable(Of ItemKey).Equals
            If other Is Nothing Then Return False
            Return String.Equals(EntryId, other.EntryId, StringComparison.Ordinal) AndAlso
                   String.Equals(StoreId, other.StoreId, StringComparison.Ordinal)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, ItemKey))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return HashCode.Combine(EntryId, StoreId)
        End Function

        Public Overrides Function ToString() As String
            Return $"item:{Shorten(EntryId)}@{Shorten(StoreId)}"
        End Function

        ''' <summary>
        ''' Só um prefixo. O EntryID inteiro identifica uma mensagem real e
        ''' não deve ir para log (F1-M).
        ''' </summary>
        Friend Shared Function Shorten(value As String) As String
            If String.IsNullOrEmpty(value) Then Return "(vazio)"
            Return If(value.Length <= 8, value, value.Substring(0, 8) & "…")
        End Function
    End Class

    ''' <summary>Identidade de uma pasta.</summary>
    Public NotInheritable Class FolderKey
        Implements IEquatable(Of FolderKey)

        Public ReadOnly Property EntryId As String
        Public ReadOnly Property StoreId As String

        Public Sub New(entryId As String, storeId As String)
            Me.EntryId = If(entryId, String.Empty)
            Me.StoreId = If(storeId, String.Empty)
        End Sub

        Public Overloads Function Equals(other As FolderKey) As Boolean Implements IEquatable(Of FolderKey).Equals
            If other Is Nothing Then Return False
            Return String.Equals(EntryId, other.EntryId, StringComparison.Ordinal) AndAlso
                   String.Equals(StoreId, other.StoreId, StringComparison.Ordinal)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, FolderKey))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return HashCode.Combine(EntryId, StoreId)
        End Function

        Public Overrides Function ToString() As String
            Return $"folder:{ItemKey.Shorten(EntryId)}"
        End Function
    End Class

    ''' <summary>
    ''' Identidade de um anexo.
    '''
    ''' O índice sozinho é instável: a coleção pode mudar entre obter e usar.
    ''' Nome e tamanho acompanham para o broker VALIDAR que o anexo no índice
    ''' ainda é o mesmo antes de salvar.
    ''' </summary>
    Public NotInheritable Class AttachmentKey
        Public ReadOnly Property Owner As ItemKey
        Public ReadOnly Property Index As Integer
        Public ReadOnly Property FileName As String
        Public ReadOnly Property SizeBytes As Integer

        Public Sub New(owner As ItemKey, index As Integer, fileName As String, sizeBytes As Integer)
            Me.Owner = owner
            Me.Index = index
            Me.FileName = If(fileName, String.Empty)
            Me.SizeBytes = sizeBytes
        End Sub

        Public Overrides Function ToString() As String
            Return $"anexo[{Index}] {FileName} ({SizeBytes} bytes)"
        End Function
    End Class

    ''' <summary>
    ''' Identidade de um rascunho. Tipo próprio, e não ItemKey, para que o
    ''' compilador impeça passar uma mensagem qualquer para SendDraftAsync.
    ''' </summary>
    Public NotInheritable Class DraftKey
        Public ReadOnly Property Item As ItemKey

        Public Sub New(item As ItemKey)
            Me.Item = item
        End Sub

        Public Overrides Function ToString() As String
            Return $"rascunho:{Item}"
        End Function
    End Class

End Namespace

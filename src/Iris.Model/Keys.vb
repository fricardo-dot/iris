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
        Implements IEquatable(Of AttachmentKey)

        Public ReadOnly Property Owner As ItemKey
        Public ReadOnly Property Index As Integer
        Public ReadOnly Property FileName As String
        Public ReadOnly Property SizeBytes As Integer

        ''' <summary>
        ''' Se <see cref="FileName"/> e <see cref="SizeBytes"/> foram <b>lidos</b>,
        ''' ou se são o vazio e o zero de quem não conseguiu ler.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>A GUARDA DE IDENTIDADE OLHAVA SÓ UM DOS DOIS LADOS</b>
        '''
        ''' Antes de gravar um anexo, o leitor confere nome e tamanho, porque o
        ''' índice é instável. A conferência passou a fechar quando a leitura
        ''' <i>de agora</i> falha — e continuava cega para a leitura <i>de
        ''' antes</i>, gravada aqui: uma chave montada com <c>""/0</c> por falha
        ''' casava com qualquer anexo que hoje leia <c>""/0</c>, e com um que
        ''' leia <c>"x.dat"/0</c> se o nome tiver sido o único lido.
        '''
        ''' Sem este campo, "identidade conferida" queria dizer "os dois valores
        ''' batem", e não "os dois valores são conhecidos e batem".
        ''' </summary>
        Public ReadOnly Property IdentidadeConhecida As Boolean

        Public Sub New(owner As ItemKey, index As Integer, fileName As String, sizeBytes As Integer,
                       Optional identidadeConhecida As Boolean = True)
            Me.Owner = owner
            Me.Index = index
            Me.FileName = If(fileName, String.Empty)
            Me.SizeBytes = sizeBytes
            Me.IdentidadeConhecida = identidadeConhecida
        End Sub

        ''' <summary>
        ''' SEM o nome do arquivo. A politica de log proibe registrar nome de
        ''' anexo, e alguem inevitavelmente vai passar a chave para o log — a
        ''' versao anterior deste ToString entregava o nome de bandeja.
        ''' </summary>
        Public Overrides Function ToString() As String
            Return $"anexo[{Index}] {SizeBytes}b"
        End Function

        Public Overloads Function Equals(other As AttachmentKey) As Boolean _
            Implements IEquatable(Of AttachmentKey).Equals
            If other Is Nothing Then Return False
            ' O "EU SEI" ENTRA NA IGUALDADE, e ficou de fora quando ele
            ' nasceu. Duas chaves com os mesmos quatro campos e confiancas
            ' DIFERENTES nao sao a mesma chave: o MesmaIdentidade trata uma
            ' como conferivel e a outra nao, e um comparador que as iguala
            ' esconde exatamente essa mudanca de confianca.
            Return Index = other.Index AndAlso
                   SizeBytes = other.SizeBytes AndAlso
                   IdentidadeConhecida = other.IdentidadeConhecida AndAlso
                   String.Equals(FileName, other.FileName, StringComparison.Ordinal) AndAlso
                   Equals(Owner, other.Owner)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, AttachmentKey))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return HashCode.Combine(Owner, Index, FileName, SizeBytes, IdentidadeConhecida)
        End Function
    End Class

    ''' <summary>
    ''' Identidade de um compromisso. Tipo próprio pelo mesmo motivo do
    ''' <see cref="DraftKey"/>: o compilador impede passar uma mensagem para
    ''' uma operação de calendário — e aqui isso vale mais, porque a operação
    ''' do outro lado <b>apaga</b>.
    ''' </summary>
    Public NotInheritable Class AppointmentKey
        Implements IEquatable(Of AppointmentKey)

        Public ReadOnly Property Item As ItemKey

        Public Sub New(item As ItemKey)
            If item Is Nothing Then Throw New ArgumentNullException(NameOf(item))
            Me.Item = item
        End Sub

        Public Overrides Function ToString() As String
            Return "compromisso " & Item.ToString()
        End Function

        Public Overloads Function Equals(other As AppointmentKey) As Boolean _
            Implements IEquatable(Of AppointmentKey).Equals
            If other Is Nothing Then Return False
            Return Equals(Item, other.Item)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, AppointmentKey))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Item.GetHashCode()
        End Function
    End Class

    ''' <summary>
    ''' Identidade de um rascunho. Tipo próprio, e não ItemKey, para que o
    ''' compilador impeça passar uma mensagem qualquer para SendDraftAsync.
    ''' </summary>
    Public NotInheritable Class DraftKey
        Implements IEquatable(Of DraftKey)

        Public ReadOnly Property Item As ItemKey

        Public Sub New(item As ItemKey)
            Me.Item = item
        End Sub

        ' Igualdade por valor tambem aqui: sem ela, usar a chave num
        ' dicionario, num Set ou na selecao do WPF compararia por
        ' referencia e nunca casaria depois de reconstruida.
        Public Overloads Function Equals(other As DraftKey) As Boolean _
            Implements IEquatable(Of DraftKey).Equals
            If other Is Nothing Then Return False
            Return Equals(Item, other.Item)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, DraftKey))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return If(Item Is Nothing, 0, Item.GetHashCode())
        End Function

        Public Overrides Function ToString() As String
            Return $"rascunho:{Item}"
        End Function
    End Class

End Namespace

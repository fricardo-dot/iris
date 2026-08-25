Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq

Namespace Global.Iris.Model

    ''' <summary>
    ''' Um registro de rótulo dentro do valor de <c>MSIP_Labels</c> — os campos
    ''' de <b>um</b> GUID.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE POR GUID, E NÃO UM CAMPO SOLTO POR LEITURA</b>
    '''
    ''' A medição do 3.0 mostrou que o valor traz <c>Enabled</c>, <c>Name</c>,
    ''' <c>SetDate</c>, <c>Method</c>, <c>SiteId</c> e <c>ContentBits</c> — e
    ''' que um item pode carregar <b>mais de um</b> registro, inclusive de
    ''' rótulo já desligado.
    '''
    ''' Guardar um <c>ContentBits</c> único por leitura colapsaria registros
    ''' diferentes num campo só, e o portão decidiria sobre uma mistura. O
    ''' <c>ContentBits</c> pertence ao registro do GUID; é ali que ele mora.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTE TIPO NÃO AFIRMA</b>
    '''
    ''' Que <c>ContentBits</c> signifique alguma coisa. O 3.0 mediu que o campo
    ''' <b>existe</b> — não que reflita a proteção corrente, que seja autêntico,
    ''' que sua ausência queira dizer "sem proteção", nem que cubra toda forma
    ''' de IRM. E a P16 vale para ele igual: vem no mesmo cabeçalho.
    '''
    ''' Aqui o valor é preservado, não interpretado.
    ''' </summary>
    Public NotInheritable Class LabelRecord

        ''' <summary>O GUID, normalizado em minúsculas com hífens.</summary>
        Public ReadOnly Property Id As String

        ''' <summary>
        ''' O campo <c>Enabled</c>. <c>Nothing</c> quando o campo <b>não
        ''' aparece</b> — que é diferente de aparecer com <c>False</c>.
        ''' </summary>
        Public ReadOnly Property Enabled As Boolean?

        ''' <summary>
        ''' O <c>ContentBits</c> como inteiro, quando ele aparece e é um inteiro.
        '''
        ''' <c>Nothing</c> cobre dois casos que o
        ''' <see cref="ContentBitsIlegivel"/> separa: o campo não veio, ou veio
        ''' e não deu para ler. Os dois negam, mas por motivos diferentes, e a
        ''' diferença é o que aparece para o usuário.
        ''' </summary>
        Public ReadOnly Property ContentBits As Integer?

        ''' <summary>O campo <c>ContentBits</c> veio e não é inteiro.</summary>
        Public ReadOnly Property ContentBitsIlegivel As Boolean

        ''' <summary>
        ''' Nomes dos campos vistos neste registro. <b>Nomes, não valores</b> —
        ''' o valor de <c>Name</c> é texto escolhido pela empresa e pode ele
        ''' próprio ser sensível.
        ''' </summary>
        Public ReadOnly Property Campos As IReadOnlyList(Of String)

        Public Sub New(id As String, enabled As Boolean?, contentBits As Integer?,
                       contentBitsIlegivel As Boolean, campos As IReadOnlyList(Of String))
            Me.Id = If(id, "")
            Me.Enabled = enabled
            Me.ContentBits = contentBits
            Me.ContentBitsIlegivel = contentBitsIlegivel
            Me.Campos = If(campos, CType(Array.Empty(Of String)(), IReadOnlyList(Of String)))
        End Sub

        ''' <summary>Este registro está ligado — <c>Enabled=True</c>, explícito.</summary>
        Public ReadOnly Property Ativo As Boolean
            Get
                Return Enabled.HasValue AndAlso Enabled.Value
            End Get
        End Property

    End Class

    ''' <summary>
    ''' Constrói <see cref="LabelRecord"/> a partir dos pares crus, sem decidir
    ''' nada sobre eles.
    ''' </summary>
    Public NotInheritable Class LabelRecordBuilder

        Private ReadOnly _id As String
        Private ReadOnly _campos As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Private _enabled As Boolean?
        Private _bits As Integer?
        Private _bitsIlegivel As Boolean
        Private _enabledIlegivel As Boolean
        Private _contradicao As Boolean

        Public Sub New(id As Guid)
            _id = id.ToString("D", CultureInfo.InvariantCulture)
        End Sub

        Public ReadOnly Property Id As String
            Get
                Return _id
            End Get
        End Property

        ''' <summary>
        ''' O mesmo campo apareceu duas vezes <b>dizendo coisas diferentes</b>.
        '''
        ''' Separado de <see cref="Ilegivel"/> de propósito: valor que se
        ''' contradiz e valor que não dá para ler são coisas distintas, e
        ''' juntá-las faria "Enabled=Talvez" sair como conflito — o que
        ''' descreveria mal o que aconteceu.
        ''' </summary>
        Public ReadOnly Property Contraditorio As Boolean
            Get
                Return _contradicao
            End Get
        End Property

        ''' <summary>
        ''' Algum campo veio com valor que não dá para interpretar —
        ''' <c>Enabled</c> que não é <c>True</c> nem <c>False</c>.
        ''' </summary>
        Public ReadOnly Property Ilegivel As Boolean
            Get
                Return _enabledIlegivel
            End Get
        End Property

        Public Sub Aceitar(campo As String, valor As String)
            _campos.Add(campo)

            If String.Equals(campo, "Enabled", StringComparison.OrdinalIgnoreCase) Then
                Dim lido As Boolean?
                If String.Equals(valor, "True", StringComparison.OrdinalIgnoreCase) Then
                    lido = True
                ElseIf String.Equals(valor, "False", StringComparison.OrdinalIgnoreCase) Then
                    lido = False
                Else
                    _enabledIlegivel = True
                    Return
                End If
                ' O MESMO campo dizendo duas coisas no mesmo registro. Era isto
                ' que saia como Present na primeira versao do parser.
                If _enabled.HasValue AndAlso _enabled.Value <> lido.Value Then _contradicao = True
                _enabled = lido
                Return
            End If

            If String.Equals(campo, "ContentBits", StringComparison.OrdinalIgnoreCase) Then
                Dim n As Integer
                If Integer.TryParse(valor, NumberStyles.Integer, CultureInfo.InvariantCulture, n) Then
                    If _bits.HasValue AndAlso _bits.Value <> n Then _contradicao = True
                    _bits = n
                Else
                    _bitsIlegivel = True
                End If
            End If
        End Sub

        Public Function Construir() As LabelRecord
            Return New LabelRecord(_id, _enabled, _bits, _bitsIlegivel, _campos.ToList())
        End Function

    End Class

End Namespace

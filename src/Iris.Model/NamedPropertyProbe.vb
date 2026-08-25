Imports System.Collections.Generic

Namespace Global.Iris.Model

    ''' <summary>
    ''' Como o <c>PropertyAccessor</c> se comporta com uma named property —
    ''' <b>medido</b>, não suposto.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO PRECISOU EXISTIR</b>
    '''
    ''' A primeira rodada da medição do 3.0 devolveu <c>Blank</c> para
    ''' <b>120 de 120</b> itens. Duas leituras cabem nesse número:
    '''
    '''   1. ninguém nesta caixa tem rótulo, e a propriedade existe vazia;
    '''   2. esta propriedade <b>nunca</b> falha — ela devolve string vazia
    '''      quando não existe, e "vazio" não quer dizer nada.
    '''
    ''' Se for a segunda, <c>Blank</c> é <b>indistinguível de ausente</b>, e um
    ''' portão que tratasse <c>Blank</c> como conclusivo estaria decidindo
    ''' sobre ruído. Não dá para escolher entre as duas olhando só o rótulo.
    '''
    ''' O controle é ler, no MESMO item e pelo MESMO caminho, uma propriedade
    ''' que <b>comprovadamente não existe</b>. O que ela devolver é o que
    ''' "ausente" parece nesta conta.
    '''
    ''' <b>Conjunto fixo, decidido aqui.</b> Nada de nome vindo do chamador:
    ''' named property tem mapeamento próprio no store, e gerar candidatos em
    ''' massa mexe nesse mapeamento.
    ''' </summary>
    Public NotInheritable Class NamedPropertyProbe

        ''' <summary>Um DASL, e o que ele devolveu.</summary>
        Public NotInheritable Class Tentativa
            Public ReadOnly Property Rotulo As String
            Public ReadOnly Property Dasl As String
            Public ReadOnly Property Lancou As Boolean
            Public ReadOnly Property HResult As Integer?
            Public ReadOnly Property Excecao As String
            ''' <summary>Nome do tipo devolvido, quando não lançou.</summary>
            Public ReadOnly Property TipoDoValor As String
            Public ReadOnly Property Comprimento As Integer?

            Public Sub New(rotulo As String, dasl As String, lancou As Boolean,
                           hresult As Integer?, excecao As String,
                           tipoDoValor As String, comprimento As Integer?)
                Me.Rotulo = rotulo
                Me.Dasl = dasl
                Me.Lancou = lancou
                Me.HResult = hresult
                Me.Excecao = excecao
                Me.TipoDoValor = tipoDoValor
                Me.Comprimento = comprimento
            End Sub

            ''' <summary>Como esta tentativa aparece no relatório.</summary>
            Public Overrides Function ToString() As String
                If Lancou Then
                    Return $"{Rotulo}: LANÇOU {Excecao}" &
                           If(HResult.HasValue, $" 0x{HResult.Value:X8}", "")
                End If
                Return $"{Rotulo}: devolveu {If(TipoDoValor, "Nothing")}" &
                       If(Comprimento.HasValue, $", {Comprimento.Value} car.", "")
            End Function
        End Class

        Public ReadOnly Property Item As ItemKey
        Public ReadOnly Property Tentativas As IReadOnlyList(Of Tentativa)

        Public Sub New(item As ItemKey, tentativas As IReadOnlyList(Of Tentativa))
            Me.Item = item
            Me.Tentativas = If(tentativas, CType(Array.Empty(Of Tentativa)(), IReadOnlyList(Of Tentativa)))
        End Sub

    End Class

End Namespace

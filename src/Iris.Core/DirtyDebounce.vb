Namespace Global.Iris.Core

    ''' <summary>
    ''' Decide QUANDO uma pasta suja deve ser recarregada.
    '''
    ''' Máquina pura: recebe o instante de fora e não conhece timer,
    ''' dispatcher nem relógio. Vive separada do <c>FolderWatcher</c> porque
    ''' lá dentro a única forma de testá-la seria esperar segundos de relógio
    ''' real — e a suíte deste projeto já teve um teste intermitente, que é
    ''' pior que teste nenhum.
    '''
    ''' Duas regras, e as duas já estiveram erradas aqui:
    '''
    '''   • O atraso conta do ÚLTIMO evento, não do primeiro. A versão
    '''     original recarregava 450 ms depois do primeiro evento, viesse o
    '''     que viesse depois — ou seja, não era debounce, era temporizador.
    '''   • O teto existe para a rajada que não para. Sem ele, uma
    '''     sincronização longa do Exchange adiaria a recarga
    '''     indefinidamente, e a lista ficaria velha justamente quando mais
    '''     muda. O defeito original aqui foi escrever
    '''     <c>desde &lt; 450 AndAlso desde &lt; 2000</c> — a MESMA variável
    '''     nas duas comparações, o que se reduz a <c>desde &lt; 450</c> e
    '''     deixava o teto como letra morta. Duas grandezas diferentes
    '''     precisam de dois nomes diferentes.
    ''' </summary>
    Public NotInheritable Class DirtyDebounce

        ''' <summary>Silêncio necessário depois do último evento.</summary>
        Public Const DebounceMsPadrao As Integer = 450

        ''' <summary>Espera máxima desde o primeiro evento da rajada.</summary>
        Public Const TetoMsPadrao As Integer = 2000

        Private ReadOnly _debounceMs As Integer
        Private ReadOnly _tetoMs As Integer

        Private _suja As Boolean
        Private _primeiro As DateTimeOffset
        Private _ultimo As DateTimeOffset

        Public Sub New(Optional debounceMs As Integer = DebounceMsPadrao,
                       Optional tetoMs As Integer = TetoMsPadrao)
            _debounceMs = debounceMs
            _tetoMs = tetoMs
        End Sub

        Public ReadOnly Property IsDirty As Boolean
            Get
                Return _suja
            End Get
        End Property

        ''' <summary>
        ''' Chegou um evento. O primeiro da rajada marca o início; todos
        ''' adiam o silêncio.
        ''' </summary>
        Public Sub Mark(agora As DateTimeOffset)
            If Not _suja Then
                _suja = True
                _primeiro = agora
            End If
            _ultimo = agora
        End Sub

        ''' <summary>
        ''' Já dá para recarregar?
        '''
        ''' Consulta pura: NÃO muda estado. Quem decidir recarregar chama
        ''' <see cref="Clear"/> — separar as duas coisas é o que permite ao
        ''' teste perguntar várias vezes sem alterar o que está medindo.
        ''' </summary>
        Public Function ShouldFlush(agora As DateTimeOffset) As Boolean
            If Not _suja Then Return False

            Dim silencio = (agora - _ultimo).TotalMilliseconds
            Dim espera = (agora - _primeiro).TotalMilliseconds

            ' Forma positiva: o silêncio libera, e o teto libera sozinho.
            ' Escrever isto como "não recarregar enquanto ambos faltarem"
            ' também estaria correto — e era assim que estava no watcher —
            ' mas na forma positiva a independência das duas condições fica
            ' à vista, e foi confundi-las que produziu o defeito original.
            Return silencio >= _debounceMs OrElse espera >= _tetoMs
        End Function

        Public Sub Clear()
            _suja = False
        End Sub

    End Class

End Namespace

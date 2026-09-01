Imports System.Collections.Generic
Imports Iris.Integration
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' <b>Os rótulos do acervo, lidos uma vez por retrato.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO NÃO É UM CACHE "PARA IR MAIS RÁPIDO"</b>
    '''
    ''' A fila e a caixa dividida pedem o rótulo <b>por linha</b>: elas recebem
    ''' funções <c>ItemKey → rótulo</c>. Ligá-las direto ao banco faria uma
    ''' consulta SQL por linha desenhada — e a consulta lê a pasta inteira, então
    ''' uma fila de trinta linhas leria o acervo trinta vezes.
    '''
    ''' Então a leitura acontece uma vez e fica na mão. O que decide quando ela
    ''' envelheceu é o <b>carimbo do retrato</b> do acervo, que sobe a cada
    ''' publicação — e publicação é exatamente o momento em que os rótulos podem
    ''' ter mudado.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>FALHA DE LEITURA NÃO VIRA "NADA CLASSIFICADO"</b>
    '''
    ''' Se o banco não responder, isto devolve uma leitura vazia — e vazia aqui é
    ''' indistinguível de "ninguém classificou nada". A distinção existe em
    ''' <see cref="LeituraDeRotulos.PastasLidas"/>, e quem mostra a tela precisa
    ''' dela: zero pasta lida com pastas no acervo é falha, não é caixa limpa.
    '''
    ''' <b>E ela não repete a falha a cada linha.</b> O carimbo é gravado mesmo
    ''' quando a leitura falha; senão cada linha desenhada tentaria abrir o banco
    ''' de novo, e uma falha barata viraria trinta.
    ''' </summary>
    Friend NotInheritable Class RotulosNaMao

        Private ReadOnly _ler As Func(Of LeituraDeRotulos)
        Private ReadOnly _carimbo As Func(Of Integer)
        Private ReadOnly _trava As New Object()

        Private _lida As LeituraDeRotulos
        Private _de As Integer = -1

        Public Sub New(ler As Func(Of LeituraDeRotulos), carimbo As Func(Of Integer))
            _ler = ler
            _carimbo = carimbo
        End Sub

        Public Function Atual() As LeituraDeRotulos
            SyncLock _trava
                Dim agora = If(_carimbo Is Nothing, 0, _carimbo())
                If _lida IsNot Nothing AndAlso agora = _de Then Return _lida

                Try
                    _lida = If(_ler Is Nothing, Nothing, _ler())
                Catch
                    _lida = Nothing
                End Try

                If _lida Is Nothing Then _lida = LeituraDeRotulos.Vazia()
                _de = agora
                Return _lida
            End SyncLock
        End Function

        ''' <summary>O rótulo desta mensagem, ou vazio.</summary>
        Public Function Rotulo(chave As ItemKey) As String
            If chave Is Nothing Then Return ""
            Dim achado As String = Nothing
            Return If(Atual().Rotulos.TryGetValue(chave, achado), achado, "")
        End Function

        ''' <summary>
        ''' Quantas regras do dono esta mensagem satisfez.
        '''
        ''' <b>Ausência e lista vazia dão zero aqui, e isso é uma perda</b> — a
        ''' pontuação só sabe contar. A distinção sobrevive em
        ''' <see cref="LeituraDeRotulos.RegrasCasadas"/>, que é o que a caixa
        ''' dividida recebe inteiro.
        ''' </summary>
        Public Function QuantasRegras(chave As ItemKey) As Integer
            If chave Is Nothing Then Return 0
            Dim achadas As IReadOnlyList(Of String) = Nothing
            If Not Atual().RegrasCasadas.TryGetValue(chave, achadas) Then Return 0
            Return If(achadas Is Nothing, 0, achadas.Count)
        End Function

    End Class

End Namespace

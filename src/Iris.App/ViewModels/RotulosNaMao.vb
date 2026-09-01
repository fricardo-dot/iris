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
    ''' <b>E a falha não fica guardada.</b> Uma leitura que estoura devolve vazio
    ''' <i>desta vez</i> e não carimba nada — a próxima tenta de novo. A versão
    ''' anterior carimbava a falha para não repeti-la a cada linha desenhada, e o
    ''' preço era pior do que o problema: uma falha transitória — o banco ocupado
    ''' por um lote de classificação, por exemplo — congelava "nenhum rótulo" até
    ''' a próxima publicação, que pode não vir nunca.
    '''
    ''' O custo de tentar de novo por linha é uma consulta a mais numa tela que
    ''' já está errada; o custo de congelar é uma tela que mente até o programa
    ''' ser reaberto. Achado por revisão externa em 01/09/2026.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O CARIMBO NÃO VÊ A CLASSIFICAÇÃO</b>
    '''
    ''' Ele é o contador de recargas do acervo, e gravar rótulo <b>não</b>
    ''' republica pasta nenhuma. Uma passagem de classificação que termine com a
    ''' tela aberta não move o carimbo, e a leitura guardada continuaria valendo
    ''' — os rótulos novos só apareceriam numa recarga futura por outro motivo.
    '''
    ''' Por isso existe <see cref="Esquecer"/>, e quem grava rótulo tem de
    ''' chamá-lo. Não é elegante depender de quem chama; a alternativa era o
    ''' cache de rótulos saber avisar a tela, e isso é uma dependência de baixo
    ''' para cima que este projeto não tem em lugar nenhum.
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

                Dim nova As LeituraDeRotulos = Nothing
                Try
                    nova = If(_ler Is Nothing, Nothing, _ler())
                Catch
                    nova = Nothing
                End Try

                ' FALHA NAO CARIMBA. Ver o cabecalho.
                If nova Is Nothing Then Return LeituraDeRotulos.Vazia()

                _lida = nova
                _de = agora
                Return _lida
            End SyncLock
        End Function

        ''' <summary>
        ''' Joga fora a leitura guardada. <b>Quem grava rótulo chama isto</b> — o
        ''' carimbo do acervo não enxerga gravação de rótulo, porque gravar rótulo
        ''' não republica pasta nenhuma.
        ''' </summary>
        Public Sub Esquecer()
            SyncLock _trava
                _lida = Nothing
                _de = -1
            End SyncLock
        End Sub

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

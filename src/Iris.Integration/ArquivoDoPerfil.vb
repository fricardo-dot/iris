Imports System.Collections.Generic
Imports System.IO
Imports System.Text

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>Lê um arquivo de configuração do perfil sem confiar no tamanho dele.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO EXISTE</b>
    '''
    ''' As regras do dono e as identidades dele moram em
    ''' <c>%LOCALAPPDATA%\Iris\</c>, e eram lidas com <c>File.ReadAllLines</c> —
    ''' que materializa o arquivo inteiro, todas as cadeias e depois as coleções
    ''' do LINQ, <b>antes</b> de qualquer conferência.
    '''
    ''' Isso não é hipótese acadêmica: é uma pasta do perfil, gravável por
    ''' qualquer processo que rode como o dono. Um arquivo de dois gigabytes no
    ''' lugar de um de dez linhas derruba o Iris por falta de memória — e o
    ''' <c>Catch</c> que devolveria "sem regras" nem chega a rodar, porque o
    ''' estouro acontece antes. Achado por revisão externa em 02/09/2026.
    '''
    ''' Não é defesa contra código hostil local: contra esse, nada nesta máquina
    ''' é. É defesa contra <b>acidente</b> — um arquivo trocado, um log que foi
    ''' parar no lugar errado, uma cópia mal feita — que é o que de fato acontece
    ''' com arquivo em pasta de perfil.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>TRÊS TETOS, E CADA UM PEGA UMA COISA</b>
    '''
    ''' O do <b>arquivo</b> é conferido antes de abrir e pega o caso grosseiro. O
    ''' de <b>linhas</b> pega o arquivo de tamanho razoável com um milhão de linhas
    ''' vazias. O de <b>caracteres por linha</b> pega o arquivo de uma linha só com
    ''' tudo dentro — que passaria pelos outros dois.
    '''
    ''' Estourar qualquer um devolve <b>nada</b>, e não "o que deu para ler": meia
    ''' lista de regras classificaria a caixa com parte das regras do dono sem
    ''' dizer qual ficou de fora, e meia lista de identidades faria as mensagens
    ''' dele virarem "do outro" — os dois defeitos que estes arquivos já tiveram.
    ''' </summary>
    Friend Module ArquivoDoPerfil

        ''' <summary>
        ''' Teto do arquivo inteiro. Um megabyte são dezenas de milhares de linhas
        ''' de regra; quem passa disso não está escrevendo regras.
        ''' </summary>
        Friend Const MaximoDeBytes As Long = 1024L * 1024L

        ''' <summary>Linhas úteis. Acima disto, o arquivo não é o que se pensa.</summary>
        Friend Const MaximoDeLinhas As Integer = 20_000

        ''' <summary>
        ''' Caracteres por linha. Uma regra é uma frase; um endereço é curto. Uma
        ''' linha maior que isto é um arquivo de outra coisa.
        ''' </summary>
        Friend Const MaximoPorLinha As Integer = 4_000

        ''' <summary>
        ''' As linhas úteis — sem vazias, sem comentário, já aparadas.
        '''
        ''' <c>Nothing</c> quer dizer <b>não deu para ler com confiança</b>: arquivo
        ''' ausente, grande demais, linhas demais, linha longa demais, ou falha de
        ''' E/S. Quem chama decide o que dizer; o que não pode é receber metade.
        ''' </summary>
        Friend Function Linhas(caminho As String) As List(Of String)
            Try
                Dim info As New FileInfo(caminho)
                If Not info.Exists Then Return Nothing
                If info.Length > MaximoDeBytes Then Return Nothing

                Dim uteis As New List(Of String)()
                ' STREAM, e nao ReadAllLines: o teto acima e do arquivo no disco, e
                ' entre a medida e a leitura ele pode crescer. Ler linha a linha faz
                ' o teto de linhas ser a segunda rede, e nao a primeira esperanca.
                Using leitor As New StreamReader(caminho, Encoding.UTF8)
                    Dim lidas = 0
                    While True
                        Dim linha = leitor.ReadLine()
                        If linha Is Nothing Then Exit While

                        lidas += 1
                        If lidas > MaximoDeLinhas Then Return Nothing
                        If linha.Length > MaximoPorLinha Then Return Nothing

                        Dim aparada = linha.Trim()
                        If aparada.Length = 0 Then Continue While
                        If aparada.StartsWith("#", StringComparison.Ordinal) Then Continue While
                        uteis.Add(aparada)
                    End While
                End Using

                Return uteis
            Catch
                Return Nothing
            End Try
        End Function

    End Module

End Namespace

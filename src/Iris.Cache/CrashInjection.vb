Imports System.Threading

Namespace Global.Iris.Cache

    ''' <summary>
    ''' Pontos onde o processo pode morrer, nomeados para que o teste possa
    ''' escolher um.
    '''
    ''' Existe porque "sobrevive a crash" não é verificável olhando o código:
    ''' o intervalo perigoso entre gravar as linhas, avançar o checkpoint e
    ''' publicar é justamente onde não há nada para ler. A única forma de
    ''' provar é morrer ali de propósito e reabrir o arquivo.
    ''' </summary>
    Public NotInheritable Class CrashInjection

        ' Nomes usados pelos testes. Constantes e nao enum porque o harness
        ' de crash recebe o ponto como argumento de linha de comando.
        Public Const AntesDeGravarPagina As String = "antes-de-gravar-pagina"
        Public Const DentroDaPaginaAntesDoCommit As String = "dentro-da-pagina-antes-do-commit"
        Public Const DepoisDoCommitDaPagina As String = "depois-do-commit-da-pagina"
        Public Const DentroDaPublicacaoAntesDoCommit As String = "dentro-da-publicacao-antes-do-commit"
        Public Const DepoisDoCommitDaPublicacao As String = "depois-do-commit-da-publicacao"

        ''' <summary>
        ''' Entre o consumidor RECEBER e a dívida ser marcada como drenada.
        '''
        ''' É a janela que torna a entrega <i>ao menos uma vez</i>: morrer aqui
        ''' deixa a UI já tendo agido e o disco ainda dizendo que ela não
        ''' recebeu. Na reabertura, a mesma geração é entregue de novo — e é por
        ''' isso que o consumidor tem de ser idempotente.
        ''' </summary>
        ''' <summary>
        ''' Entre as duas atualizações da reconciliação do diário de divulgação.
        '''
        ''' Existe porque ali havia <b>duas escritas independentes</b>: morrer no
        ''' meio deixava as ambíguas gravadas e o aviso perdido para sempre — na
        ''' abertura seguinte a transição não pegava mais nada, e o usuário
        ''' nunca ficava sabendo que pode ter saído conteúdo.
        ''' </summary>
        Public Const EntreAsDuasReconciliacoes As String =
            "entre-as-duas-reconciliacoes"

        Public Const DepoisDeReceberAntesDeMarcarDrenada As String =
            "depois-de-receber-antes-de-marcar-drenada"

        Private Shared _ponto As String
        Private Shared _acao As Action

        ''' <summary>
        ''' Arma um ponto. <paramref name="acao"/> é o que "morrer" significa
        ''' — lançar (prova atomicidade) ou matar o processo (prova
        ''' durabilidade). São coisas diferentes e o teste escolhe qual.
        ''' </summary>
        Public Shared Sub Armar(ponto As String, acao As Action)
            Volatile.Write(_ponto, ponto)
            _acao = acao
        End Sub

        Public Shared Sub Desarmar()
            Volatile.Write(_ponto, Nothing)
            _acao = Nothing
        End Sub

        Public Shared Sub Talvez(ponto As String)
            If Not String.Equals(Volatile.Read(_ponto), ponto, StringComparison.Ordinal) Then Return
            Dim a = _acao
            Desarmar()
            a?.Invoke()
        End Sub

    End Class

End Namespace

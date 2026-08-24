Namespace Global.Iris.Cache

    ''' <summary>
    ''' Defeitos ligáveis, para dar controle negativo ao teste de crash.
    '''
    ''' Um teste que só confirma "depois do crash está tudo consistente" passa
    ''' igualzinho num escritor correto e num escritor que simplesmente não
    ''' grava nada. O controle negativo é o que separa os dois: ligo o defeito,
    ''' o mesmo teste tem de acusar perda.
    '''
    ''' É a regra do CLAUDE.md — "quando o controle negativo for barato,
    ''' confirme desfazendo a correção e vendo o teste falhar" — feita
    ''' executável, em vez de feita uma vez à mão e esquecida.
    ''' </summary>
    Public NotInheritable Class CacheWriterDefects

        ''' <summary>
        ''' Avança o checkpoint numa transação PRÓPRIA, antes de gravar as
        ''' linhas. É o desenho ingênuo: "salva o progresso primeiro".
        '''
        ''' Morrer no intervalo deixa o cursor dizendo que a página foi lida
        ''' sem que nenhuma linha dela exista. A retomada começa na página
        ''' seguinte e aquelas mensagens NUNCA MAIS são lidas — perda
        ''' silenciosa, sem erro, sem log, e invisível para qualquer contagem
        ''' que confie no cursor.
        ''' </summary>
        Public Shared Property CheckpointAntesDasLinhas As Boolean

        Public Shared Sub Limpar()
            CheckpointAntesDasLinhas = False
        End Sub

    End Class

End Namespace

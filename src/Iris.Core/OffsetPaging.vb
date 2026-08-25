Namespace Global.Iris.Core

    ''' <summary>
    ''' A aritmética da paginação por <b>offset</b>, isolada do COM.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO SAIU DE DENTRO DO ADAPTADOR</b>
    '''
    ''' Eu escrevi, com confiança, que offset sobre coleção viva repete quando
    ''' algo é <b>removido</b> e pula quando algo é <b>inserido</b>. É o
    ''' contrário, e o Codex mostrou com quatro linhas de exemplo.
    '''
    ''' A afirmação errada não era decorativa: dela eu tinha derivado uma regra
    ''' de teste ("repetir com as pontas concordando é defeito, porque deslocar
    ''' offset exige remoção, e remoção apareceria na ponta"). A regra caiu
    ''' junto com a premissa.
    '''
    ''' O conserto não é reescrever o comentário. É tirar a aritmética de um
    ''' lugar onde só a caixa real a exercita — e onde, portanto, ninguém nunca
    ''' vai <b>demonstrar</b> qual das duas direções é a verdadeira.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ACONTECE DE FATO</b>
    '''
    ''' Com página de 2 sobre <c>[A B C D]</c>, a primeira devolve A e B e o
    ''' cursor guarda <c>offset = 2</c>. A segunda lê as posições 3 e 4.
    '''
    '''   <b>Inserção</b> antes do offset — <c>[X A B C D]</c>: as posições 3 e
    '''   4 agora são <b>B</b> e C. <b>B sai repetido.</b>
    '''
    '''   <b>Remoção</b> antes do offset — <c>[B C D]</c>: as posições 3 e 4
    '''   agora são D e nada. <b>C some em silêncio.</b>
    '''
    ''' Perder em silêncio é o pior dos dois, e é o que a Q1 existe para pegar.
    ''' A defesa real seria cursor por chave — o que o caminho por
    ''' <c>Table</c> faz, e o motivo de ele existir. Aqui não dá:
    ''' <c>Items.Sort</c> ordena por campo não único (<c>Subject</c>,
    ''' <c>SenderName</c>) e o OOM não expõe "continue depois desta chave".
    ''' </summary>
    Public Module OffsetPaging

        ''' <summary>
        ''' A janela 1-based que esta página deve ler, e onde a próxima começa.
        '''
        ''' <c>Proximo</c> avança por <b>posições examinadas</b>, não por itens
        ''' devolvidos: contar devolvidos releria as posições puladas e
        ''' duplicaria linha.
        ''' </summary>
        Public Function Janela(offset As Integer, quantas As Integer, total As Integer) _
                              As (Primeiro As Integer, Ultimo As Integer, Proximo As Integer?)
            Dim primeiro = offset + 1
            Dim ultimo = Math.Min(offset + quantas, total)
            Dim proximo As Integer? = If(ultimo < total, CType(ultimo, Integer?), Nothing)
            Return (primeiro, ultimo, proximo)
        End Function

    End Module

End Namespace

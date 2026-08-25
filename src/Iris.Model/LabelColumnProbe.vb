Namespace Global.Iris.Model

    ''' <summary>
    ''' O rótulo é projetável como <b>coluna de <c>Table</c></b>?
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE A PERGUNTA IMPORTA</b>
    '''
    ''' A Fase 0 mediu ~16 ms por item para montar o DTO de uma mensagem, e
    ''' foi esse número que tornou o cache obrigatório. Classificar por item
    ''' custa a mesma ordem de grandeza. Se o rótulo vier por coluna, dá para
    ''' classificar uma pasta inteira barato; se não vier, cada decisão custa
    ''' uma ida ao COM, e isso muda o que a UI pode oferecer.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A ARMADILHA QUE ESTE TIPO EXISTE PARA REGISTRAR</b>
    '''
    ''' <c>Columns.Add</c> pode <b>aceitar</b> a coluna e mesmo assim entregar
    ''' erro ou nada nas linhas. Por isso "aceitou" e "veio valor" são campos
    ''' <b>separados</b>: um probe que só olhasse o <c>Add</c> concluiria que
    ''' o caminho barato existe, e ele não existiria.
    ''' </summary>
    Public NotInheritable Class LabelColumnProbe

        ''' <summary><c>Columns.Add</c> não lançou.</summary>
        Public Property ColunaAceita As Boolean

        ''' <summary>HRESULT do <c>Add</c>, quando ele recusou.</summary>
        Public Property HResultDoAdd As Integer?

        Public Property LinhasLidas As Integer
        ''' <summary>Linhas em que a célula trouxe valor não vazio.</summary>
        Public Property LinhasComValor As Integer
        ''' <summary>Linhas em que a célula veio nula ou vazia.</summary>
        Public Property LinhasSemValor As Integer
        ''' <summary>Linhas em que ler a célula lançou.</summary>
        Public Property LinhasComErro As Integer

        ''' <summary>Quanto a travessia custou, para comparar com o por-item.</summary>
        Public Property MilissegundosTotais As Double

        ''' <summary>
        ''' O caminho barato existe de verdade: a coluna foi aceita <b>e</b>
        ''' alguma linha entregou valor sem erro.
        '''
        ''' Aceita e sem nenhum valor é ambíguo — pode ser "ninguém tem
        ''' rótulo" ou "a coluna não funciona". A medição registra os dois
        ''' números para que a ambiguidade fique visível em vez de virar
        ''' conclusão.
        ''' </summary>
        Public ReadOnly Property Serve As Boolean
            Get
                Return ColunaAceita AndAlso LinhasComErro = 0 AndAlso LinhasComValor > 0
            End Get
        End Property

    End Class

End Namespace

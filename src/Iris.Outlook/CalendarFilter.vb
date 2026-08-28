Imports System.Globalization

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' <b>O filtro <c>Restrict</c> de uma janela de calendário — só texto.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO É PÚBLICO E MORA SOZINHO</b>
    '''
    ''' <c>Iris.Outlook</c> <b>não</b> abre os internos para a suíte, e isso é
    ''' deliberado: a camada COM não deve ser cutucada por dentro. Mas o
    ''' formato deste filtro é lógica pura — nenhuma linha aqui toca em COM —
    ''' e é exatamente o tipo de coisa que tem de ser testável sem Outlook
    ''' aberto.
    '''
    ''' Então ele sai de <c>CalendarReading</c>, que é <c>Friend</c> e cheio de
    ''' RCW, e vira um tipo público de duas funções. A alternativa seria abrir
    ''' os internos do projeto inteiro para testar duas concatenações.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O FORMATO DA DATA NÃO É ESCOLHA</b>
    '''
    ''' O <c>Restrict</c> do OOM <b>não entende ISO</b>. Ele espera a data no
    ''' formato que a documentação prescreve, e usar a cultura da máquina faria
    ''' o filtro depender do idioma do Office: numa máquina em português,
    ''' <c>28/08/2026</c> em vez de <c>08/28/2026</c> — e o Outlook não muda de
    ''' idéia junto.
    '''
    ''' Já é a segunda vez neste projeto que a cultura ambiente entra onde não
    ''' devia. A primeira foi um teste que media a máquina em vez do código, e
    ''' custou uma tarde.
    ''' </summary>
    Public NotInheritable Class CalendarFilter

        Private Sub New()
        End Sub

        ''' <summary>
        ''' O filtro de uma janela semiaberta <c>[de, ate)</c>.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>A COMPARAÇÃO É CRUZADA, E ISSO IMPORTA</b>
        '''
        ''' <c>[End] &gt; de AND [Start] &lt; ate</c> — <i>termina depois do
        ''' começo da janela E começa antes do fim dela</i>. É a definição de
        ''' intersecção, e não de contenção.
        '''
        ''' Filtrar só por <c>[Start]</c> seria mais simples e esconderia a
        ''' reunião que começou ontem e termina hoje: ela pertence à janela de
        ''' hoje, e some. É o mesmo formato de erro que a Q1 pegou na paginação
        ''' de mensagens — a lista continuava plausível.
        '''
        ''' Semiaberta nas pontas para que duas janelas adjacentes não contem o
        ''' mesmo compromisso duas vezes, pela mesma razão que o cursor da
        ''' paginação é semiaberto.
        ''' </summary>
        Public Shared Function Janela(de As DateTimeOffset, ate As DateTimeOffset) As String
            Return $"[End] > '{Data(de)}' AND [Start] < '{Data(ate)}'"
        End Function

        ''' <summary>
        ''' A data como o <c>Restrict</c> a quer: hora <b>local</b>, formato
        ''' americano, cultura invariante.
        '''
        ''' Local e não UTC porque o OOM compara contra a hora local do item —
        ''' foi essa confusão que, na Q1, fez o filtro DASL perder 803 de 1.003
        ''' mensagens <b>terminando parecendo completa</b>.
        ''' </summary>
        Public Shared Function Data(d As DateTimeOffset) As String
            Return d.LocalDateTime.ToString("MM/dd/yyyy hh:mm tt", CultureInfo.InvariantCulture)
        End Function

    End Class

End Namespace

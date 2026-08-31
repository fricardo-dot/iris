Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq

Namespace Global.Iris.Model

    ''' <summary>
    ''' <b>A ordem sugerida da fila — e por que cada linha está onde está.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>UMA NOTA QUE NINGUÉM CONSEGUE CONFERIR É UM PALPITE COM CARA DE
    ''' CONTA</b>
    '''
    ''' Esta é a decisão que dá forma ao arquivo inteiro. Ordenar a fila por uma
    ''' pontuação muda o que o dono faz de manhã, e uma pontuação que ele não
    ''' consegue destrinchar é pior do que nenhuma ordenação: ele passa a
    ''' obedecer a um número que não sabe de onde vem, e quando ele discordar não
    ''' terá do que discordar.
    '''
    ''' Então a nota <b>não é um número</b>. Ela é uma lista de parcelas com
    ''' nome, valor e frase, e o total é a soma delas — conferível com papel e
    ''' caneta. <see cref="Prioridade.Explicar"/> devolve isso em português.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>OS DIAS NÃO SOMEM NUNCA</b>
    '''
    ''' A pontuação <i>reordena</i>; ela não substitui. "Fulano está esperando há
    ''' 20 dias" continua na tela em toda linha, inclusive nas que a nota jogou
    ''' para baixo — porque é o único dado aqui que não é opinião de ninguém.
    '''
    ''' Uma tela que troque os dias pela nota transforma um fato em juízo, e
    ''' esconde justamente o caso em que a nota errou.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>OS PESOS SÃO ESCOLHIDOS, E NÃO MEDIDOS</b>
    '''
    ''' Nenhum deles saiu de dado nenhum: não há histórico de "o que o dono
    ''' respondeu primeiro" para calibrar nada. São palpites com justificativa,
    ''' e é assim que estão escritos. Quando houver medição, eles mudam — e o
    ''' teste que os congela vai falhar e obrigar alguém a dizer por quê.
    ''' </summary>
    Public NotInheritable Class PrioridadeDaFila

        ''' <summary>
        ''' <b>Um ponto por dia de espera.</b> É a âncora da escala: todos os
        ''' outros pesos são lidos como "equivale a tantos dias".
        '''
        ''' Linear, e não exponencial. Uma curva que dispara faria a linha de 60
        ''' dias esmagar todas as outras para sempre — e uma pendência que nunca
        ''' sai de cima é uma pendência que o dono aprende a ignorar.
        ''' </summary>
        Public Const PorDia As Double = 1.0

        ''' <summary>
        ''' <b>Vinte dias por "espera você".</b> Vale mais que qualquer espera
        ''' razoável, porque a diferença entre "alguém espera resposta" e "isto é
        ''' informação" é maior do que a diferença entre duas semanas e três.
        ''' </summary>
        Public Const PorEsperarResposta As Double = 20.0

        ''' <summary>
        ''' <b>Cinco dias por regra do dono.</b> Menos que o rótulo, e de
        ''' propósito: a regra dele diz <i>sobre o que</i> é a mensagem, não que
        ''' alguém está esperando. Uma reclamação de cliente já respondida não é
        ''' urgente por ser reclamação.
        ''' </summary>
        Public Const PorRegraDoDono As Double = 5.0

        ''' <summary>
        ''' <b>Trinta dias quando o prazo já passou; dez quando ele é esta
        ''' semana.</b> Só entra quando o <b>dono</b> marcou um prazo — nada aqui
        ''' lê data de dentro do corpo do e-mail, porque uma data lida errada de
        ''' um texto viraria urgência inventada.
        ''' </summary>
        Public Const PorPrazoVencido As Double = 30.0
        Public Const PorPrazoPerto As Double = 10.0
        ''' <summary>Quantos dias contam como "esta semana".</summary>
        Public Const DiasDePrazoPerto As Integer = 7

        ''' <summary>
        ''' <b>Dez dias por pessoa próxima.</b> "Próxima" é uma decisão de quem
        ''' chama — este arquivo não sabe o que é um contato. O padrão é
        ''' <c>False</c>, e o padrão é o certo: sem saber, não afirmar.
        ''' </summary>
        Public Const PorPessoaProxima As Double = 10.0

        ''' <summary>
        ''' Pontua uma linha.
        ''' </summary>
        ''' <param name="dias">
        ''' Dias de espera. Vem de fora já contado, porque é o mesmo número que a
        ''' tela mostra — recalculá-lo aqui abriria a porta para a nota e a
        ''' coluna discordarem.
        ''' </param>
        ''' <param name="prazo">
        ''' O prazo que <b>o dono</b> marcou, se marcou. <c>Nothing</c> não vale
        ''' como "sem pressa": vale como "ninguém disse", e não vira parcela
        ''' nenhuma.
        ''' </param>
        Public Shared Function Pontuar(dias As Integer,
                                       rotulo As String,
                                       regrasCasadas As Integer,
                                       Optional pessoaProxima As Boolean = False,
                                       Optional prazo As DateTimeOffset? = Nothing,
                                       Optional agora As DateTimeOffset? = Nothing) As Prioridade

            Dim parcelas As New List(Of Parcela)()

            ' OS DIAS SEMPRE ENTRAM, inclusive zero. Uma parcela ausente e uma
            ' parcela de valor zero contam a mesma historia na soma e historias
            ' diferentes na explicacao -- e a explicacao e o produto aqui.
            Dim deEspera = Math.Max(0, dias)
            parcelas.Add(New Parcela("espera", deEspera * PorDia,
                                     $"{deEspera} dia(s) esperando"))

            If String.Equals(rotulo, "precisa_de_mim", StringComparison.Ordinal) Then
                parcelas.Add(New Parcela("espera_resposta", PorEsperarResposta,
                                         "alguém espera uma resposta sua"))
            End If

            Dim quantasRegras = Math.Max(0, regrasCasadas)
            If quantasRegras > 0 Then
                parcelas.Add(New Parcela("regra_do_dono",
                                         quantasRegras * PorRegraDoDono,
                                         $"casa com {quantasRegras} regra(s) sua(s)"))
            End If

            If pessoaProxima Then
                parcelas.Add(New Parcela("pessoa", PorPessoaProxima,
                                         "é alguém com quem você troca e-mails"))
            End If

            If prazo.HasValue AndAlso agora.HasValue Then
                Dim faltam = (prazo.Value - agora.Value).TotalDays
                If faltam < 0 Then
                    parcelas.Add(New Parcela("prazo", PorPrazoVencido,
                                             "o prazo que você marcou já passou"))
                ElseIf faltam <= DiasDePrazoPerto Then
                    parcelas.Add(New Parcela("prazo", PorPrazoPerto,
                                             "o prazo que você marcou é esta semana"))
                End If
            End If

            Return New Prioridade(parcelas)
        End Function

    End Class

    ''' <summary>
    ''' Uma parcela da nota: o nome interno, quanto vale, e a frase que o dono lê.
    '''
    ''' A frase não é enfeite. Ela é a metade do produto que torna a ordem
    ''' discutível — um nome interno como <c>espera_resposta</c> explica para
    ''' quem escreveu o código, e para mais ninguém.
    ''' </summary>
    Public NotInheritable Class Parcela
        Public ReadOnly Property Nome As String
        Public ReadOnly Property Valor As Double
        Public ReadOnly Property Frase As String

        Friend Sub New(nome As String, valor As Double, frase As String)
            Me.Nome = If(nome, "")
            Me.Valor = valor
            Me.Frase = If(frase, "")
        End Sub
    End Class

    ''' <summary>
    ''' A nota de uma linha, <b>com as contas abertas</b>.
    '''
    ''' <see cref="Total"/> sozinho não é um produto deste arquivo: ele só faz
    ''' sentido acompanhado de <see cref="Parcelas"/>, e quem for mostrá-lo tem
    ''' de conseguir mostrar as duas coisas.
    ''' </summary>
    Public NotInheritable Class Prioridade

        Public ReadOnly Property Parcelas As IReadOnlyList(Of Parcela)

        Public ReadOnly Property Total As Double
            Get
                Return Parcelas.Sum(Function(p) p.Valor)
            End Get
        End Property

        Friend Sub New(parcelas As IReadOnlyList(Of Parcela))
            Me.Parcelas = If(parcelas, Array.Empty(Of Parcela)())
        End Sub

        ''' <summary>
        ''' A conta em português, uma parcela por linha, terminando no total.
        '''
        ''' <b>Todas as parcelas, inclusive as de valor zero.</b> Omitir a de
        ''' zero pareceria mais limpo e esconderia a informação mais útil da
        ''' explicação: <i>esta linha está aqui apesar de estar esperando há zero
        ''' dias</i>.
        ''' </summary>
        Public Function Explicar() As String
            Dim linhas = Parcelas.Select(
                Function(p) $"{p.Frase}: {p.Valor.ToString("0.#", CultureInfo.InvariantCulture)}")
            Return String.Join(Environment.NewLine, linhas) & Environment.NewLine &
                   $"total: {Total.ToString("0.#", CultureInfo.InvariantCulture)}"
        End Function

    End Class

End Namespace

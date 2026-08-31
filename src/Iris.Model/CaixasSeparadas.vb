Imports System.Collections.Generic
Imports System.Linq

Namespace Global.Iris.Model

    ''' <summary>
    ''' <b>A caixa dividida — cada mensagem numa gaveta só.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>UMA MENSAGEM APARECE EM UMA GAVETA, E NÃO EM TRÊS</b>
    '''
    ''' É a decisão que dá forma a tudo aqui. Uma mensagem pode satisfazer duas
    ''' regras do dono <i>e</i> ter um rótulo; mostrá-la nas três gavetas seria
    ''' honesto e inútil — uma caixa dividida em que o mesmo e-mail aparece três
    ''' vezes não está dividida, está triplicada, e o dono passa a ter de lembrar
    ''' se já tratou aquilo em outra gaveta.
    '''
    ''' A ordem de desempate é <b>a do dono primeiro</b>:
    '''
    ''' <list type="number">
    ''' <item>a <b>primeira</b> regra que ela satisfaz, na ordem do arquivo dele
    ''' — ele escreveu aquilo justamente para separar isso;</item>
    ''' <item>senão, o rótulo;</item>
    ''' <item>senão, a gaveta das não classificadas.</item>
    ''' </list>
    '''
    ''' A ordem do arquivo dele vira, assim, uma ordem de <b>prioridade</b>, e
    ''' isso precisa estar escrito na tela: quem escreve regra genérica na
    ''' primeira linha esvazia as de baixo, e vai querer saber por quê.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A GAVETA DAS NÃO CLASSIFICADAS NÃO É UM RESTO</b>
    '''
    ''' Ela é a única que diz a verdade sobre a cobertura. Se a varredura
    ''' classificou 40 de 900, as 860 restantes <b>não</b> são "sem importância":
    ''' são desconhecidas, e escondê-las faria a caixa dividida parecer completa.
    ''' Por isso ela existe mesmo vazia, e por isso vem por último e não some.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ORDEM FIXA, E NÃO POR TAMANHO</b>
    '''
    ''' As gavetas de rótulo vêm sempre na mesma ordem, e é a ordem de <i>quem
    ''' cobra</i>: primeiro o que espera resposta, depois o que espera resposta
    ''' de outro, depois o que não espera nada. Ordenar por tamanho faria a
    ''' caixa mudar de forma a cada varredura, e uma tela que se reorganiza
    ''' sozinha é uma tela em que ninguém acha nada duas vezes.
    ''' </summary>
    Public NotInheritable Class CaixasSeparadas

        ''' <summary>
        ''' A ordem em que as gavetas de rótulo aparecem. Não é a ordem do enum
        ''' por acaso — é a ordem de quem cobra.
        ''' </summary>
        Private Shared ReadOnly Ordem As String() = {
            "precisa_de_mim", "aguardando", "fyi",
            "notificacao", "promocao", "newsletter"}

        ''' <summary>
        ''' O nome que cada gaveta de rótulo mostra. Em português, e dito do
        ''' ponto de vista do dono — <i>"esperam você"</i>, e não
        ''' <i>"precisa_de_mim"</i>, que é o nome do fio.
        ''' </summary>
        Private Shared ReadOnly Nomes As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String)(StringComparer.Ordinal) From {
                {"precisa_de_mim", "Esperam você"},
                {"aguardando", "Você já respondeu"},
                {"fyi", "Só para saber"},
                {"notificacao", "Avisos automáticos"},
                {"promocao", "Promoções"},
                {"newsletter", "Newsletters"}}

        Public Const NomeDasNaoClassificadas As String = "Ainda não classificadas"

        ''' <summary>
        ''' Divide.
        '''
        ''' <paramref name="rotulos"/> e <paramref name="regrasCasadas"/> são
        ''' indexados pela <see cref="ItemKey"/>; mensagem que não aparece em
        ''' nenhum dos dois cai na gaveta das não classificadas.
        '''
        ''' <paramref name="regrasDoDono"/> dá a <b>ordem</b>. Uma regra casada
        ''' que não esteja nessa lista é ignorada: ela é de um arquivo anterior,
        ''' o dono já a apagou, e ressuscitá-la como gaveta seria o programa
        ''' insistindo numa pergunta que ele parou de fazer.
        ''' </summary>
        Public Shared Function Dividir(
                mensagens As IReadOnlyList(Of MensagemNaFila),
                rotulos As IReadOnlyDictionary(Of ItemKey, String),
                regrasCasadas As IReadOnlyDictionary(Of ItemKey, IReadOnlyList(Of String)),
                regrasDoDono As IReadOnlyList(Of String)) As IReadOnlyList(Of Gaveta)

            Dim todas = If(mensagens, CType(Array.Empty(Of MensagemNaFila)(),
                                            IReadOnlyList(Of MensagemNaFila)))
            Dim comRotulo = If(rotulos, New Dictionary(Of ItemKey, String)())
            Dim comRegra = If(regrasCasadas,
                              New Dictionary(Of ItemKey, IReadOnlyList(Of String))())
            Dim doDono = If(regrasDoDono, CType(Array.Empty(Of String)(),
                                                IReadOnlyList(Of String)))

            ' As gavetas nascem TODAS, e vazias. Uma gaveta que so aparece
            ' quando tem conteudo faz a tela mudar de forma a cada varredura --
            ' e faz "nenhuma mensagem espera voce" ficar indistinguivel de
            ' "ninguem perguntou".
            Dim gavetas As New List(Of Gaveta)()
            Dim porNome As New Dictionary(Of String, List(Of MensagemNaFila))(StringComparer.Ordinal)
            Dim nomesNaOrdem As New List(Of String)()

            For Each regra In doDono
                Dim limpa = If(regra, "").Trim()
                If limpa.Length = 0 OrElse porNome.ContainsKey(limpa) Then Continue For
                porNome(limpa) = New List(Of MensagemNaFila)()
                nomesNaOrdem.Add(limpa)
            Next

            Dim primeiraDoDono = nomesNaOrdem.Count

            For Each chaveDoRotulo In Ordem
                porNome(chaveDoRotulo) = New List(Of MensagemNaFila)()
                nomesNaOrdem.Add(chaveDoRotulo)
            Next

            porNome(NomeDasNaoClassificadas) = New List(Of MensagemNaFila)()
            nomesNaOrdem.Add(NomeDasNaoClassificadas)

            For Each m In todas
                If m Is Nothing OrElse m.Chave Is Nothing Then Continue For
                porNome(GavetaDe(m, comRotulo, comRegra, doDono)).Add(m)
            Next

            For i = 0 To nomesNaOrdem.Count - 1
                Dim chaveDaGaveta = nomesNaOrdem(i)
                Dim daRegra = i < primeiraDoDono
                Dim rotuloDela = If(daRegra OrElse
                                    chaveDaGaveta = NomeDasNaoClassificadas,
                                    "", chaveDaGaveta)

                gavetas.Add(New Gaveta(
                    NomeVisivel(chaveDaGaveta, daRegra),
                    rotuloDela,
                    daRegra,
                    porNome(chaveDaGaveta).
                        OrderByDescending(Function(m) m.Quando.GetValueOrDefault()).
                        ThenBy(Function(m) m.Assunto, StringComparer.Ordinal).
                        ToList()))
            Next

            Return gavetas
        End Function

        ''' <summary>
        ''' Em que gaveta esta mensagem cai. <b>Uma só</b> — ver o cabeçalho.
        ''' </summary>
        Private Shared Function GavetaDe(
                m As MensagemNaFila,
                rotulos As IReadOnlyDictionary(Of ItemKey, String),
                regrasCasadas As IReadOnlyDictionary(Of ItemKey, IReadOnlyList(Of String)),
                regrasDoDono As IReadOnlyList(Of String)) As String

            Dim minhas As IReadOnlyList(Of String) = Nothing
            If regrasCasadas.TryGetValue(m.Chave, minhas) AndAlso minhas IsNot Nothing Then
                ' A ORDEM DO ARQUIVO DELE DECIDE, e nao a ordem em que o modelo
                ' devolveu as regras casadas. A segunda nao quer dizer nada; a
                ' primeira ele escreveu.
                For Each regra In regrasDoDono
                    Dim limpa = If(regra, "").Trim()
                    If limpa.Length = 0 Then Continue For
                    If minhas.Any(Function(r) String.Equals(If(r, "").Trim(), limpa,
                                                            StringComparison.Ordinal)) Then
                        Return limpa
                    End If
                Next
            End If

            Dim rotuloDela As String = Nothing
            If rotulos.TryGetValue(m.Chave, rotuloDela) AndAlso
               Ordem.Contains(rotuloDela) Then
                Return rotuloDela
            End If

            ' ROTULO QUE NAO ESTA NA LISTA CAI AQUI, e nao numa gaveta propria.
            ' Ele so pode ter vindo de um banco gravado por uma versao com outro
            ' conjunto de rotulos, e inventar uma gaveta para ele mostraria ao
            ' dono uma categoria que este programa nao sabe explicar.
            Return NomeDasNaoClassificadas
        End Function

        Private Shared Function NomeVisivel(chave As String, daRegra As Boolean) As String
            If daRegra Then Return chave
            Dim nome As String = Nothing
            Return If(Nomes.TryGetValue(chave, nome), nome, chave)
        End Function

    End Class

    ''' <summary>
    ''' Uma gaveta da caixa dividida.
    '''
    ''' <see cref="Vazia"/> é normal e não é escondida: a gaveta existir vazia é
    ''' o que separa <i>"nada aqui"</i> de <i>"esta divisão não existe"</i>.
    ''' </summary>
    Public NotInheritable Class Gaveta

        ''' <summary>O que a tela mostra. Para gaveta do dono, a frase dele.</summary>
        Public ReadOnly Property Nome As String
        ''' <summary>
        ''' O rótulo do fio (<c>precisa_de_mim</c>, …), ou vazio quando a gaveta
        ''' é de uma regra do dono ou a das não classificadas. Serve a quem
        ''' precisa da chave e não do nome — um ícone, um filtro salvo.
        ''' </summary>
        Public ReadOnly Property Rotulo As String
        ''' <summary>Esta gaveta nasceu de uma regra que o dono escreveu?</summary>
        Public ReadOnly Property DoDono As Boolean
        Public ReadOnly Property Mensagens As IReadOnlyList(Of MensagemNaFila)

        Public ReadOnly Property Quantas As Integer
            Get
                Return Mensagens.Count
            End Get
        End Property

        Public ReadOnly Property Vazia As Boolean
            Get
                Return Mensagens.Count = 0
            End Get
        End Property

        Friend Sub New(nome As String, rotulo As String, doDono As Boolean,
                       mensagens As IReadOnlyList(Of MensagemNaFila))
            Me.Nome = If(nome, "")
            Me.Rotulo = If(rotulo, "")
            Me.DoDono = doDono
            Me.Mensagens = If(mensagens, Array.Empty(Of MensagemNaFila)())
        End Sub

    End Class

End Namespace

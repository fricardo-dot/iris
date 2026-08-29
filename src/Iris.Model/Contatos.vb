Imports System.Collections.Generic

Namespace Global.Iris.Model

    ''' <summary>
    ''' Identidade de um contato. Tipo próprio pelo mesmo motivo do
    ''' <see cref="TaskKey"/> e do <see cref="AppointmentKey"/>: o compilador
    ''' impede passar uma mensagem para uma operação de contato.
    ''' </summary>
    Public NotInheritable Class ContactKey
        Implements IEquatable(Of ContactKey)

        Public ReadOnly Property Item As ItemKey

        Public Sub New(item As ItemKey)
            If item Is Nothing Then Throw New ArgumentNullException(NameOf(item))
            Me.Item = item
        End Sub

        Public Overrides Function ToString() As String
            Return "contato " & Item.ToString()
        End Function

        Public Overloads Function Equals(other As ContactKey) As Boolean _
            Implements IEquatable(Of ContactKey).Equals
            If other Is Nothing Then Return False
            Return Equals(Item, other.Item)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, ContactKey))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Item.GetHashCode()
        End Function
    End Class

    ''' <summary>
    ''' <b>O que se quer gravar num contato — e o que ele NÃO tem.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>NÃO HÁ CORPO, NEM NOTA, NEM CARTÃO</b>
    '''
    ''' Um <c>ContactItem</c> do Outlook tem mais de cem propriedades, e várias
    ''' são texto livre — <c>Body</c>, <c>PersonalHomePage</c>, os três
    ''' endereços completos. O Iris não escreve em nenhuma delas.
    '''
    ''' Não é economia de trabalho. É o mesmo desenho das outras fases: o que o
    ''' Iris cria a partir de uma mensagem é <b>só o que a mensagem já dizia</b>
    ''' — quem mandou, e de onde. Um campo de nota convidaria a pôr no catálogo
    ''' de endereços um resumo que o assistente escreveu, e aí um dado gerado
    ''' viraria dado de cadastro, num lugar que outras ferramentas leem como se
    ''' fosse verdade conferida.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E NÃO HÁ NADA QUE ENCAMINHE</b>
    '''
    ''' <c>ContactItem.ForwardAsVcard()</c> devolve um <c>MailItem</c>. Salvar
    ''' contato não manda e-mail, mas encaminhar cartão manda — e é o único
    ''' caminho de envio que existe neste objeto. Não há campo, não há operação,
    ''' e há teste que varre o fonte do escritor.
    ''' </summary>
    Public NotInheritable Class ContactDraft
        Public Property Nome As String = ""
        Public Property Email As String = ""
        Public Property Empresa As String = ""
    End Class

    ''' <summary>
    ''' Um contato lido.
    '''
    ''' <b>Os campos de texto são anuláveis de propósito.</b> Cadeia vazia diria
    ''' "este contato não tem empresa"; <c>Nothing</c> diz "não consegui ler
    ''' este campo". São coisas diferentes, e colapsá-las é a família de
    ''' ausência-virando-fato que esta base corrigiu em cinco lugares — aqui o
    ''' estrago seria afirmar sobre o cadastro de uma pessoa.
    ''' </summary>
    Public NotInheritable Class ContactInfo
        Public Property Key As ItemKey
        Public Property Nome As String
        Public Property Email As String
        Public Property Empresa As String
    End Class

    ''' <summary>
    ''' <b>O que uma leitura de contatos conseguiu — e a ressalva que ela
    ''' obriga a carregar.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>PASTA VAZIA NÃO É CATÁLOGO VAZIO</b>
    '''
    ''' Esta é a coisa inteira desta fase. Numa conta corporativa os contatos
    ''' vivem no <b>GAL</b> — o catálogo do servidor —, e o GAL está fora de
    ''' escopo pela §8. A pasta pessoal de Contatos pode estar vazia com a
    ''' organização inteira endereçável.
    '''
    ''' Uma tela que dissesse "nenhum contato" sobre isso estaria afirmando
    ''' ausência a partir de não ter olhado — e desta vez a afirmação seria
    ''' sobre <i>pessoas</i>. Por isso <see cref="ForaDoAlcance"/> não é um
    ''' detalhe de apresentação: é parte do resultado, e vem marcado do
    ''' escritor, não decidido pela tela.
    ''' </summary>
    Public NotInheritable Class ContactList
        Public Property Items As New List(Of ContactInfo)()

        ''' <summary>
        ''' Quantos itens a leitura recusou. <c>Nothing</c> é "não contei"; zero
        ''' é "contei e não houve".
        ''' </summary>
        Public Property Skipped As Integer?

        Public Property Truncada As Boolean
        Public Property MotivoDoCorte As String = ""

        ''' <summary>
        ''' <b>O que esta leitura não alcança, e por quê.</b>
        '''
        ''' Sempre preenchido, inclusive — sobretudo — quando a leitura deu
        ''' certo e a lista veio vazia. Ver o comentário da classe.
        ''' </summary>
        Public Property ForaDoAlcance As String = ""
    End Class

    ''' <summary>
    ''' <b>O que o leitor e a tela precisam dizer igual.</b>
    '''
    ''' <b>Não se chama <c>Contatos</c>, e o motivo está no CLAUDE.md.</b> A
    ''' tela tem uma propriedade <c>Contatos</c> — a coleção — e em VB um
    ''' membro eclipsa um módulo de mesmo nome ignorando maiúsculas. O erro que
    ''' aparece é "ForaDoAlcance não é membro de ObservableCollection", que não
    ''' diz nada sobre o problema. Colchetes calariam o compilador e deixariam
    ''' a armadilha no lugar para o próximo.
    '''
    ''' Mora no modelo, e não no <c>ContactWriting</c>, porque as duas pontas
    ''' usam: quem lê preenche a ressalva, e quem mostra precisa repeti-la
    ''' quando a leitura falha. Deixá-la só no lado do Outlook obrigaria a
    ''' camada de tela a enxergar a camada do COM — ou, pior, a escrever a
    ''' ressalva de novo com outras palavras, que é como uma ressalva começa a
    ''' divergir de si mesma.
    ''' </summary>
    Public Module RegrasDeContato

        ''' <summary>
        ''' A ressalva que toda leitura de contatos carrega.
        '''
        ''' Constante, e não texto montado na hora, porque precisa sair igual
        ''' em toda leitura — inclusive na que deu certo e veio vazia, que é
        ''' justamente onde a tela seria mais tentada a calar.
        ''' </summary>
        Public Const ForaDoAlcance As String =
            "esta é a pasta pessoal de Contatos. O catálogo da organização " &
            "(GAL) está fora do alcance do Iris, então uma pasta vazia aqui " &
            "não quer dizer que não haja contatos — quer dizer que eles não " &
            "estão aqui."

        ''' <summary>
        ''' <b>Já existe contato com este endereço na lista lida?</b>
        '''
        ''' Devolve o contato encontrado, ou <c>Nothing</c>.
        '''
        ''' <b>"Não encontrei" aqui é fraco de propósito, e quem chama precisa
        ''' saber disso.</b> A busca é sobre os contatos que a leitura
        ''' <i>trouxe</i> — não sobre a pasta inteira quando houve
        ''' truncamento, e nunca sobre o GAL. Um "não existe" dito com esta
        ''' base seria o mesmo defeito que a fase inteira evita, virado do
        ''' avesso.
        '''
        ''' Comparação por endereço, sem diferenciar maiúsculas. Comparar por
        ''' nome acusaria homônimo, que é comum e legítimo.
        ''' </summary>
        Public Function Procurar(lidos As IEnumerable(Of ContactInfo),
                                 email As String) As ContactInfo
            If lidos Is Nothing OrElse String.IsNullOrWhiteSpace(email) Then Return Nothing

            Dim alvo = email.Trim()
            For Each c In lidos
                ' Email Nothing e "nao consegui ler", e nao "nao tem". Nao casa
                ' com nada, e sobretudo nao casa com outro que tambem nao leu.
                If c IsNot Nothing AndAlso c.Email IsNot Nothing AndAlso
                   String.Equals(c.Email.Trim(), alvo, StringComparison.OrdinalIgnoreCase) Then
                    Return c
                End If
            Next
            Return Nothing
        End Function

    End Module

End Namespace

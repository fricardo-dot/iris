Imports System.Collections.Generic
Imports System.Linq

Namespace Global.Iris.Model

    ''' <summary>
    ''' <b>Quem é "eu" numa mensagem.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE NÃO SERVE PERGUNTAR PELA PASTA</b>
    '''
    ''' A tentação é <c>está em Itens Enviados, logo fui eu</c>, e ela quebra em
    ''' quatro situações que existem numa caixa corporativa comum: alias, conta
    ''' adicional, caixa compartilhada, e regra que move mensagem para uma pasta
    ''' que não é a de origem. Pasta é <b>onde</b> a mensagem está; a direção é
    ''' <b>quem</b> a escreveu, e as duas só coincidem por hábito.
    '''
    ''' Aqui a pergunta é sempre a mesma: o remetente está no conjunto explícito
    ''' das minhas identidades? Itens Enviados continua sendo a fonte essencial
    ''' para <i>achar</i> as mensagens; ele apenas não decide mais quem as
    ''' escreveu.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>SEM IDENTIDADE, A RESPOSTA É "NÃO SEI"</b>
    '''
    ''' Conjunto vazio devolve <see cref="Direcao.Desconhecida"/> para tudo —
    ''' nunca <see cref="Direcao.DoOutro"/>. A diferença aparece na fila de
    ''' respostas pendentes: "não sei quem escreveu" vira uma linha que se
    ''' declara incerta, e o palpite viraria uma cobrança contra você por uma
    ''' mensagem que você mesmo mandou.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O ENDEREÇO INTERNO NÃO É SMTP, E ISSO É PROBLEMA CONHECIDO</b>
    '''
    ''' Numa organização Exchange, o remetente de uma mensagem interna chega
    ''' como endereço X.500 — <c>/O=EXCHANGELABS/OU=.../CN=RECIPIENTS/CN=...</c>
    ''' — e não como <c>alguem@empresa.com</c>. Um conjunto só com endereços
    ''' SMTP não reconheceria as suas próprias mensagens internas, que são
    ''' justamente as que enchem a fila.
    '''
    ''' A comparação aqui é sobre a cadeia inteira, normalizada, <b>seja ela
    ''' SMTP ou X.500</b>: quem monta o conjunto pode e deve pôr as duas formas.
    ''' Resolver o X.500 para SMTP é trabalho da borda que fala com o Outlook, e
    ''' quando ela não conseguir, a forma crua ainda casa.
    ''' </summary>
    Public NotInheritable Class MinhasIdentidades

        Private ReadOnly _enderecos As HashSet(Of String)

        Public Sub New(enderecos As IEnumerable(Of String))
            _enderecos = New HashSet(Of String)(StringComparer.Ordinal)
            If enderecos Is Nothing Then Return
            For Each e In enderecos
                Dim normal = Normalizar(e)
                ' TUDO QUE O DONO ESCREVEU ENTRA, e a forma so e exigida do
                ' remetente que chega.
                '
                ' Filtrar aqui parecia simetrico e era pior: o arquivo e editado
                ' a mao, e uma forma legitima que eu nao previ -- "EX:/O=...",
                ' X.400, endereco de provedor que nao e Exchange -- sumia em
                ' SILENCIO. O arquivo continuava la, parecendo certo, com menos
                ' identidades do que ele escreveu.
                '
                ' E nao ha risco no outro sentido: anotacao solta so casaria com
                ' um remetente que se lesse igual, e remetente sem forma de
                ' endereco nunca chega a ser comparado.
                '
                ' Achado por revisao externa em 31/08/2026.
                If normal.Length > 0 Then _enderecos.Add(normal)
            Next
        End Sub

        ''' <summary>Nenhuma identidade declarada — e então nada é "meu".</summary>
        Public ReadOnly Property Vazio As Boolean
            Get
                Return _enderecos.Count = 0
            End Get
        End Property

        ''' <summary>Quantas identidades o conjunto tem, já normalizadas.</summary>
        Public ReadOnly Property Quantas As Integer
            Get
                Return _enderecos.Count
            End Get
        End Property

        ''' <summary>
        ''' <b>De quem é esta mensagem?</b>
        '''
        ''' Remetente ausente devolve <c>Desconhecida</c>, e não <c>DoOutro</c>:
        ''' mensagem sem remetente legível é uma leitura que falhou, e leitura
        ''' que falhou não é evidência de nada.
        ''' </summary>
        Public Function DirecaoDe(remetente As String) As Direcao
            If Vazio Then Return Direcao.Desconhecida

            Dim normal = Normalizar(remetente)
            If normal.Length = 0 Then Return Direcao.Desconhecida

            ' SEM FORMA DE ENDERECO, "NAO SEI" -- e nao "do outro".
            '
            ' O passo anterior tira o nome de exibicao de "Fulano <f@x>". Numa
            ' cadeia como "Diretoria <Regulatorio>" ele produz "regulatorio", que
            ' nao esta no conjunto e viraria DoOutro. E ai a fila cobra do dono
            ' uma resposta por uma leitura que nao deu certo -- justamente o
            ' engano que o piso do conjunto vazio existe para impedir, entrando
            ' pela porta dos fundos.
            '
            ' Achado por revisao externa em 31/08/2026.
            If Not TemFormaDeEndereco(normal) Then Return Direcao.Desconhecida

            Return If(_enderecos.Contains(normal), Direcao.Minha, Direcao.DoOutro)
        End Function

        ''' <summary>
        ''' <b>A forma comparável de um endereço.</b>
        '''
        ''' Tira espaço em volta, tira o nome de exibição de
        ''' <c>Fulano &lt;f@x&gt;</c>, tira os sinais de menor e maior soltos, e
        ''' baixa a caixa pela cultura invariante.
        '''
        ''' <b>Invariante, e não a do usuário.</b> Em turco o "I" minúsculo não é
        ''' "i", e um endereço com I viraria duas cadeias diferentes conforme a
        ''' máquina — o tipo de defeito que não acontece nunca até acontecer numa
        ''' máquina só.
        ''' </summary>
        Friend Shared Function Normalizar(endereco As String) As String
            If String.IsNullOrWhiteSpace(endereco) Then Return ""

            Dim texto = endereco.Trim()

            ' "Fulano <f@x.com>" -> "f@x.com". So quando ha o par completo: um
            ' '<' solto e lixo, e cortar por ele inventaria um endereco.
            Dim abre = texto.LastIndexOf("<"c)
            Dim fecha = texto.LastIndexOf(">"c)
            If abre >= 0 AndAlso fecha > abre Then
                texto = texto.Substring(abre + 1, fecha - abre - 1).Trim()
            End If

            Return texto.ToLowerInvariant()
        End Function

        ''' <summary>
        ''' <b>Isto se parece com um endereço?</b> — duas formas, e só duas.
        '''
        ''' SMTP tem <c>@</c>; X.500 começa com <c>/</c>. Qualquer outra coisa é
        ''' uma leitura que não deu certo, e leitura que não deu certo responde
        ''' "não sei".
        '''
        ''' <b>Não valida endereço</b>, e não deve: rejeitar um endereço
        ''' estranho mas verdadeiro o transformaria em "não sei", e "não sei"
        ''' custa uma linha incerta na fila. Errar para o lado da incerteza é
        ''' barato; errar para o lado da afirmação não é.
        ''' </summary>
        Friend Shared Function TemFormaDeEndereco(normal As String) As Boolean
            If normal.Length = 0 Then Return False
            ' "/o=" e nao so o "/" inicial: o X.500 tambem aparece prefixado,
            ' como "EX:/O=...", e exigir que comece com barra recusaria essa
            ' forma -- transformando a propria mensagem do dono em "nao sei".
            Return normal.Contains("@"c) OrElse
                   normal.StartsWith("/", StringComparison.Ordinal) OrElse
                   normal.Contains("/o=")
        End Function

        ''' <summary>As identidades, para a tela mostrar o que está valendo.</summary>
        Public Function Listar() As IReadOnlyList(Of String)
            Return _enderecos.OrderBy(Function(e) e, StringComparer.Ordinal).ToList()
        End Function

    End Class

    ''' <summary>
    ''' De quem partiu a mensagem, do ponto de vista do dono da caixa.
    ''' </summary>
    Public Enum Direcao
        ''' <summary>
        ''' Não deu para saber. É o <b>zero</b>, então é o que aparece em campo
        ''' esquecido e desserialização incompleta — e a fila trata este caso
        ''' declarando a incerteza, nunca escolhendo um lado.
        ''' </summary>
        Desconhecida = 0
        ''' <summary>Escrita por mim. A espera é do outro.</summary>
        Minha
        ''' <summary>Escrita por outra pessoa. A resposta pode ser minha.</summary>
        DoOutro
    End Enum

End Namespace

Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Model

Namespace Global.Iris.Assist

    ''' <summary>O que se pede à IA. Fechado, e por enquanto pequeno.</summary>
    Public Enum AssistOperation
        ''' <summary>
        ''' Operação não declarada. É o <b>zero</b>, então é o que aparece em
        ''' campo esquecido e desserialização incompleta — e por isso nenhuma
        ''' autorização pode listá-la. A proibição é estrutural, não de fixture.
        ''' </summary>
        Nenhuma = 0
        Resumir
        Redigir
    End Enum

    ''' <summary>
    ''' <b>A cerimônia da §28.3, virada tipo.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ISTO É</b>
    '''
    ''' O registro de que alguém — o usuário, com a política corporativa na
    ''' mão — autorizou conteúdo desta caixa a sair desta máquina. Sem um
    ''' destes, <b>válido</b> e conferido, nada sai.
    '''
    ''' Não é configuração. Configuração se muda sem pensar; isto declara quem
    ''' autorizou, sob qual política, para qual provedor e endpoint, com qual
    ''' retenção aceita, sobre quais pastas e quais rótulos.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>TUDO É LISTA EXPLÍCITA — E NEM TODA LISTA É ACEITA</b>
    '''
    ''' Os <c>LabelReadingKind</c> vêm listados um a um: não existe "aceita os
    ''' seguros" nem "aceita os conclusivos", e um estado que ninguém listou
    ''' nega, inclusive um que ainda não existe.
    '''
    ''' Mas listar não basta para valer. <see cref="LabelPolicy.Elegivel"/>
    ''' recusa estruturalmente os desfechos que descrevem o <b>fracasso de
    ''' ler</b> — <c>Denied</c>, <c>Unreadable</c>, <c>Unknown</c> e os outros.
    ''' Uma ativação que os liste é <b>inválida</b>, e inválida por inteiro:
    ''' ignorar em silêncio a entrada ruim deixaria o erro vivo e mudo.
    ''' </summary>
    Public NotInheritable Class ActivationRecord

        ''' <summary>
        ''' Identidade estável desta autorização, e a versão dela.
        '''
        ''' O diário do 3.3 e a capability do 3.2 precisam registrar <b>sob qual
        ''' autorização</b> algo saiu — "havia autorização" não serve quando a
        ''' pergunta, meses depois, for qual delas.
        ''' </summary>
        Public ReadOnly Property Id As String
        Public ReadOnly Property Versao As Integer

        ''' <summary>Quem autorizou, e sob que política. Texto obrigatório.</summary>
        Public ReadOnly Property Autoridade As String

        ''' <summary>
        ''' <b>A política corporativa aplicável foi verificada?</b>
        '''
        ''' Campo próprio, e não um adjetivo enfiado na <see cref="Autoridade"/>.
        ''' As duas perguntas são diferentes: <i>quem autorizou</i> e <i>sob que
        ''' apuração</i>. Escrever "não verificada" no campo de autoria responde
        ''' a pergunta errada e some da vista de quem procura a resposta certa.
        '''
        ''' <c>False</c> não impede a ativação — é decisão do dono da caixa —,
        ''' mas fica <b>visível na faixa o tempo todo</b>, e não enterrado no
        ''' arquivo.
        ''' </summary>
        Public ReadOnly Property PoliticaCorporativaVerificada As Boolean

        Public ReadOnly Property Quando As DateTimeOffset

        ''' <summary>
        ''' Depois disto, não vale mais. <b>Obrigatório e finito.</b>
        '''
        ''' Já foi opcional, e <c>Nothing</c> queria dizer "sem prazo". Uma
        ''' autorização sem prazo é uma decisão tomada uma vez e obedecida para
        ''' sempre: quem a concedeu esquece, o provedor muda de política, a
        ''' pasta de teste vira pasta de trabalho, e nada obriga ninguém a olhar
        ''' de novo. O prazo é o que força a reencontrar a decisão.
        ''' </summary>
        Public ReadOnly Property Ate As DateTimeOffset?

        Public ReadOnly Property Provedor As String
        ''' <summary>Endpoint fixo. <b>Nunca</b> vem do prompt.</summary>
        Public ReadOnly Property Endpoint As String
        Public ReadOnly Property Modelo As String
        Public ReadOnly Property Regiao As String
        ''' <summary>
        ''' A política de retenção e treinamento que foi aceita, <b>em texto</b>
        ''' — para quem ler o registro depois.
        '''
        ''' Ela <b>descreve</b>; quem <b>impõe</b> é
        ''' <see cref="ExigirRetencaoZero"/> e
        ''' <see cref="ProvedoresPermitidos"/>. Os dois papéis ficam separados
        ''' de propósito: cruzar texto livre com comportamento seria adivinhar o
        ''' que a frase quis dizer.
        ''' </summary>
        Public ReadOnly Property RetencaoAceita As String

        ''' <summary>
        ''' <b>Exigir retenção zero no próprio pedido.</b>
        '''
        ''' Sem isto, a autorização declarava "retenção aceita: nenhuma" e o
        ''' pedido não impunha nada — o registro seria uma frase bonita sem
        ''' efeito, e a frase bonita é pior que o silêncio porque alguém confia
        ''' nela.
        '''
        ''' Com gateway, é isto que vira restrição de roteamento no pedido.
        ''' </summary>
        Public ReadOnly Property ExigirRetencaoZero As Boolean

        ''' <summary>
        ''' Os provedores subjacentes que o pedido autoriza — vazio quer dizer
        ''' <b>qualquer um</b> que satisfaça as outras restrições.
        '''
        ''' Existe porque gateway roteia: o endpoint autorizado é o do gateway, e
        ''' quem de fato processa o conteúdo pode ser outro, com região e
        ''' retenção próprias. Listar aqui é o único jeito de a autorização
        ''' falar sobre quem realmente vê a mensagem.
        ''' </summary>
        Public ReadOnly Property ProvedoresPermitidos As IReadOnlyList(Of String)

        Public ReadOnly Property Operacoes As IReadOnlyList(Of AssistOperation)
        Public ReadOnly Property Pastas As IReadOnlyList(Of FolderKey)
        ''' <summary>GUIDs de rótulo permitidos, canônicos e minúsculos.</summary>
        Public ReadOnly Property Rotulos As IReadOnlyList(Of String)
        ''' <summary>Desfechos de leitura aceitos, listados um a um.</summary>
        Public ReadOnly Property Leituras As IReadOnlyList(Of LabelReadingKind)
        ''' <summary>Valores de <c>ContentBits</c> aceitos, listados um a um.</summary>
        Public ReadOnly Property ContentBits As IReadOnlyList(Of Integer)

        ''' <summary>
        ''' A política declara que registro <b>desligado</b> pode ser ignorado.
        '''
        ''' Sem esta declaração, um registro histórico cujo GUID não está na
        ''' lista <b>nega</b>. Ignorar histórico em silêncio seria decisão de
        ''' política disfarçada de regra técnica: <c>Enabled=False</c> mostra
        ''' que o registro se diz inativo, e não prova que o conteúdo deixou de
        ''' ser sensível, que o cabeçalho está atual, nem que a empresa aceita
        ''' desconsiderar rebaixamento.
        ''' </summary>
        Public ReadOnly Property IgnorarHistorico As Boolean

        Public Sub New(id As String, versao As Integer,
                       autoridade As String, quando As DateTimeOffset,
                       provedor As String, endpoint As String, modelo As String,
                       regiao As String, retencaoAceita As String,
                       operacoes As IEnumerable(Of AssistOperation),
                       pastas As IEnumerable(Of FolderKey),
                       rotulos As IEnumerable(Of String),
                       leituras As IEnumerable(Of LabelReadingKind),
                       contentBits As IEnumerable(Of Integer),
                       Optional ate As DateTimeOffset? = Nothing,
                       Optional ignorarHistorico As Boolean = False,
                       Optional politicaCorporativaVerificada As Boolean = False,
                       Optional exigirRetencaoZero As Boolean = False,
                       Optional provedoresPermitidos As IEnumerable(Of String) = Nothing)
            Me.Id = If(id, "")
            Me.Versao = versao
            Me.Autoridade = If(autoridade, "")
            Me.Quando = quando
            Me.Ate = ate
            Me.Provedor = If(provedor, "")
            Me.Endpoint = If(endpoint, "")
            Me.Modelo = If(modelo, "")
            Me.Regiao = If(regiao, "")
            Me.RetencaoAceita = If(retencaoAceita, "")
            Me.Operacoes = Congelar(operacoes)
            Me.Pastas = Congelar(pastas)
            Me.Rotulos = Congelar(If(rotulos, Enumerable.Empty(Of String)()).
                                  Select(AddressOf Canonizar))
            Me.Leituras = Congelar(leituras)
            Me.ContentBits = Congelar(contentBits)
            Me.IgnorarHistorico = ignorarHistorico
            Me.PoliticaCorporativaVerificada = politicaCorporativaVerificada
            Me.ExigirRetencaoZero = exigirRetencaoZero
            Me.ProvedoresPermitidos = Congelar(provedoresPermitidos)
        End Sub

        Private Shared Function Congelar(Of T)(o As IEnumerable(Of T)) As IReadOnlyList(Of T)
            If o Is Nothing Then Return Array.Empty(Of T)()
            Return o.ToList()
        End Function

        ''' <summary>
        ''' GUID em forma canônica, ou <c>Nothing</c> se não for GUID.
        '''
        ''' Sem canonizar, <c>{BAEA3331-…}</c> na autorização e
        ''' <c>baea3331-…</c> no rótulo seriam strings diferentes, e a
        ''' comparação falharia — negando, mas por confusão de formato em vez
        ''' de por política. E texto que não é GUID nunca vira entrada válida.
        ''' </summary>
        Private Shared Function Canonizar(bruto As String) As String
            Dim g As Guid
            If Not Guid.TryParse(If(bruto, ""), g) Then Return Nothing
            Return g.ToString("D", Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' <b>O que a produção tem hoje: nada.</b>
        '''
        ''' Não é lacuna nem pendência de implementação. É a §28.2: a política
        ''' corporativa aplicável não é inferível desta máquina, e a escolha do
        ''' provedor e da credencial é do usuário. Enquanto isso não acontecer,
        ''' a única política defensável é não mandar nada.
        ''' </summary>
        Public Shared ReadOnly Property DaProducao As ActivationRecord
            Get
                Return Nothing
            End Get
        End Property

        ''' <summary>
        ''' Está formalmente completo? Campo obrigatório em branco invalida —
        ''' um registro pela metade é pior que registro nenhum, porque parece
        ''' autorização.
        ''' </summary>
        Public Function Completo() As Boolean
            Return Id.Trim().Length > 0 AndAlso
                   Versao > 0 AndAlso
                   Autoridade.Trim().Length > 0 AndAlso
                   Provedor.Trim().Length > 0 AndAlso
                   Endpoint.Trim().Length > 0 AndAlso
                   Modelo.Trim().Length > 0 AndAlso
                   Regiao.Trim().Length > 0 AndAlso
                   RetencaoAceita.Trim().Length > 0 AndAlso
                   Operacoes.Count > 0 AndAlso
                   Pastas.Count > 0 AndAlso
                   Ate.HasValue
        End Function

        ''' <summary>
        ''' <b>Coerente por dentro?</b> Separado de <see cref="Completo"/>: um
        ''' registro pode ter todos os campos preenchidos e ainda declarar
        ''' coisas que não podem ser declaradas.
        '''
        ''' O que invalida, e por quê:
        ''' <list type="bullet">
        ''' <item>desfecho não elegível na lista — autorizar sobre "não
        ''' consegui ler" não é autorizar, é decidir sem informação;</item>
        ''' <item><c>AssistOperation.Nenhuma</c> na lista — é o zero do enum, o
        ''' valor de campo esquecido;</item>
        ''' <item>texto que não é GUID na lista de rótulos, que casaria com
        ''' nada ou, pior, com string arbitrária;</item>
        ''' <item>prazo que termina antes de começar.</item>
        ''' </list>
        ''' </summary>
        Public Function Coerente() As Boolean
            ' O PRAZO TEM DE SER UM PRAZO. Igual a Quando nao e prazo: e uma
            ' autorizacao que nasce vencida, e "nasce vencida" e um jeito
            ' silencioso de a IA nunca funcionar sem ninguem entender por que.
            If Not Ate.HasValue OrElse Ate.Value <= Quando Then Return False

            If Operacoes.Contains(AssistOperation.Nenhuma) Then Return False
            If Leituras.Any(Function(k) Not LabelPolicy.Elegivel(k)) Then Return False
            If Rotulos.Any(Function(r) r Is Nothing) Then Return False

            ' Provedor subjacente em branco na lista casaria com nada e pareceria
            ' uma restricao. Restricao que nao restringe e pior que nenhuma.
            If ProvedoresPermitidos.Any(Function(p) String.IsNullOrWhiteSpace(p)) Then Return False

            Return True
        End Function

        ''' <summary>
        ''' O endpoint é HTTPS? Endpoint não-HTTPS manda conteúdo corporativo
        ''' em claro, e nenhuma autorização cobre isso.
        ''' </summary>
        Public Function EndpointSeguro() As Boolean
            Dim u As Uri = Nothing
            If Not Uri.TryCreate(Endpoint, UriKind.Absolute, u) Then Return False
            Return String.Equals(u.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
        End Function

        Public Function Vigente(agora As DateTimeOffset) As Boolean
            If agora < Quando Then Return False
            ' Sem prazo NAO vale mais. Antes era `Not Ate.HasValue OrElse ...`,
            ' que fazia a ausencia de prazo virar vigencia eterna — a falha
            ' aberta de sempre, com a ausencia de dado virando permissao.
            Return Ate.HasValue AndAlso agora <= Ate.Value
        End Function

    End Class

End Namespace

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
        Public ReadOnly Property Quando As DateTimeOffset
        ''' <summary>Depois disto, não vale mais. <c>Nothing</c> = sem prazo.</summary>
        Public ReadOnly Property Ate As DateTimeOffset?

        Public ReadOnly Property Provedor As String
        ''' <summary>Endpoint fixo. <b>Nunca</b> vem do prompt.</summary>
        Public ReadOnly Property Endpoint As String
        Public ReadOnly Property Modelo As String
        Public ReadOnly Property Regiao As String
        ''' <summary>A política de retenção e treinamento que foi aceita.</summary>
        Public ReadOnly Property RetencaoAceita As String

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
                       Optional ignorarHistorico As Boolean = False)
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
                   Pastas.Count > 0
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
            If Ate.HasValue AndAlso Ate.Value < Quando Then Return False
            If Operacoes.Contains(AssistOperation.Nenhuma) Then Return False
            If Leituras.Any(Function(k) Not LabelPolicy.Elegivel(k)) Then Return False
            If Rotulos.Any(Function(r) r Is Nothing) Then Return False
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
            Return Not Ate.HasValue OrElse agora <= Ate.Value
        End Function

    End Class

End Namespace

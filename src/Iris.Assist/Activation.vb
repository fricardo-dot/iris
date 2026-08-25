Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Model

Namespace Global.Iris.Assist

    ''' <summary>O que se pede à IA. Fechado, e por enquanto pequeno.</summary>
    Public Enum AssistOperation
        ''' <summary>Valor zero: operação não declarada. Nunca autorizada.</summary>
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
    ''' destes, válido e conferido, <b>nada sai</b>.
    '''
    ''' Não é configuração. Configuração é coisa que se muda sem pensar; isto
    ''' declara quem autorizou, sob qual política, para qual provedor, com qual
    ''' retenção aceita, sobre quais pastas e quais rótulos.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>TUDO É LISTA EXPLÍCITA, E NADA É IMPLÍCITO</b>
    '''
    ''' Os <c>LabelReadingKind</c> aceitos vêm listados um a um. Não existe
    ''' "aceita os seguros" nem "aceita os conclusivos": a política diz os
    ''' nomes, e um estado que ninguém listou — inclusive um que ainda não
    ''' existe — nega.
    '''
    ''' Foi assim que a Fase 2 tratou as inferências de ambiente, e é a única
    ''' forma que sobrevive a alguém acrescentar um membro no enum sem reler
    ''' esta classe.
    ''' </summary>
    Public NotInheritable Class ActivationRecord

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
        ''' <summary>GUIDs de rótulo permitidos, minúsculos.</summary>
        Public ReadOnly Property Rotulos As IReadOnlyList(Of String)
        ''' <summary>Desfechos de leitura aceitos, listados um a um.</summary>
        Public ReadOnly Property Leituras As IReadOnlyList(Of LabelReadingKind)
        ''' <summary>Valores de <c>ContentBits</c> aceitos, listados um a um.</summary>
        Public ReadOnly Property ContentBits As IReadOnlyList(Of Integer)

        Public Sub New(autoridade As String, quando As DateTimeOffset,
                       provedor As String, endpoint As String, modelo As String,
                       regiao As String, retencaoAceita As String,
                       operacoes As IEnumerable(Of AssistOperation),
                       pastas As IEnumerable(Of FolderKey),
                       rotulos As IEnumerable(Of String),
                       leituras As IEnumerable(Of LabelReadingKind),
                       contentBits As IEnumerable(Of Integer),
                       Optional ate As DateTimeOffset? = Nothing)
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
                                  Select(Function(r) If(r, "").Trim().ToLowerInvariant()))
            Me.Leituras = Congelar(leituras)
            Me.ContentBits = Congelar(contentBits)
        End Sub

        Private Shared Function Congelar(Of T)(o As IEnumerable(Of T)) As IReadOnlyList(Of T)
            If o Is Nothing Then Return Array.Empty(Of T)()
            Return o.ToList()
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
            Return Autoridade.Trim().Length > 0 AndAlso
                   Provedor.Trim().Length > 0 AndAlso
                   Endpoint.Trim().Length > 0 AndAlso
                   Modelo.Trim().Length > 0 AndAlso
                   RetencaoAceita.Trim().Length > 0 AndAlso
                   Operacoes.Count > 0 AndAlso
                   Pastas.Count > 0
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

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq

Namespace Global.Iris.Sync

    Public Enum ProviderKind
        ''' <summary>Primeiro de propósito: não saber é o estado inicial honesto.</summary>
        Desconhecido
        ExchangeCached
        ExchangeOnline
        PstLocal
    End Enum

    ''' <summary>
    ''' O ambiente em que uma medição vale, e ele inclui a JANELA.
    '''
    ''' A §18.4 mudou o que "ambiente" significa neste projeto. "Exchange em
    ''' cache" não é um ambiente: é uma família parametrizada pela janela de
    ''' sincronização, e o parâmetro muda <b>o que existe</b>, não só o que
    ''' custa. O mesmo Iris, no mesmo Outlook, com o cursor num lugar ou
    ''' noutro, enxerga caixas diferentes.
    '''
    ''' Por isso a janela entra na impressão digital. Mudar a janela produz
    ''' um ambiente DIFERENTE, e nenhuma conclusão de ausência tirada no
    ''' anterior sobrevive: itens que "sumiram" quando a janela encolheu não
    ''' foram excluídos.
    ''' </summary>
    Public NotInheritable Class EnvironmentFingerprint

        Public ReadOnly Property Provider As ProviderKind
        Public ReadOnly Property CachedMode As Boolean

        ''' <summary>
        ''' A janela, como TOKEN OPACO lido do perfil — não como número de meses.
        '''
        ''' A §22.3 mediu que o OOM <b>não expõe a janela</b>: enumerando os
        ''' membros de <c>Store</c> não há nenhum <c>Sync*</c> nem
        ''' <c>Window*</c>; só <c>IsCachedExchange</c> e
        ''' <c>ExchangeStoreType</c>. O valor existe apenas no registro do
        ''' perfil, em <c>00036601</c>, num blob cuja codificação eu não
        ''' verifiquei.
        '''
        ''' E não preciso verificar. A impressão digital exige duas
        ''' propriedades — ser <b>estável</b> enquanto o ambiente não muda e
        ''' <b>sensível</b> quando ele muda — e nenhuma das duas exige saber o
        ''' que os bytes significam. Decodificar o blob por palpite para
        ''' escrever "3 meses" num relatório acrescentaria uma afirmação não
        ''' medida a uma classe que existe justamente para impedir isso.
        '''
        ''' <c>Nothing</c> = não foi lido, e isso NÃO é o mesmo que "sem janela".
        ''' </summary>
        Public ReadOnly Property WindowToken As String

        Public ReadOnly Property PolicyVersion As Integer

        Public Sub New(provider As ProviderKind, cachedMode As Boolean,
                       windowToken As String, Optional policyVersion As Integer = 1)
            Me.Provider = provider
            Me.CachedMode = cachedMode
            Me.WindowToken = windowToken
            Me.PolicyVersion = policyVersion
        End Sub

        Public Function Value() As String
            Return String.Join("|", Provider.ToString(),
                               If(CachedMode, "cached", "online"),
                               If(WindowToken, "janela-nao-lida"),
                               PolicyVersion.ToString(CultureInfo.InvariantCulture))
        End Function

        Public Overrides Function ToString() As String
            Return Value()
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Dim o = TryCast(obj, EnvironmentFingerprint)
            Return o IsNot Nothing AndAlso String.Equals(Value(), o.Value(), StringComparison.Ordinal)
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Value().GetHashCode()
        End Function
    End Class

    ''' <summary>
    ''' O que o Iris está AUTORIZADO a concluir neste ambiente.
    '''
    ''' Não é uma lista de recursos: é uma lista de inferências permitidas. A
    ''' diferença importa porque o dano de um ambiente não medido não é
    ''' lentidão — é o Iris afirmar que uma mensagem foi excluída quando ela
    ''' apenas saiu da janela.
    ''' </summary>
    Public NotInheritable Class EnvironmentCapabilities

        ''' <summary>
        ''' Se <c>NaoEncontrado</c> pode virar <c>AusenteDaPasta</c>. Falso num
        ''' ambiente não medido — ali "não encontrei" não distingue excluído de
        ''' fora-da-janela.
        ''' </summary>
        Public ReadOnly Property PodeConcluirAusencia As Boolean

        ''' <summary>
        ''' Se a varredura pode se declarar de cobertura COMPLETA. Falso sem
        ''' medição: a §19.2 mediu pastas cheias reportando zero itens, e uma
        ''' varredura que lê zero e se declara completa apaga a pasta.
        ''' </summary>
        Public ReadOnly Property PodeAfirmarCoberturaCompleta As Boolean

        ''' <summary>
        ''' Se o incremental da Q3 pode substituir a varredura cheia. Falso sem
        ''' medição: incremental sobre universo desconhecido não tem como
        ''' provar que não pulou nada.
        ''' </summary>
        Public ReadOnly Property PodeUsarIncremental As Boolean

        ''' <summary>Por que degradou. Vazio quando não degradou.</summary>
        Public ReadOnly Property Reason As String

        Friend Sub New(ausencia As Boolean, cobertura As Boolean, incremental As Boolean,
                       reason As String)
            PodeConcluirAusencia = ausencia
            PodeAfirmarCoberturaCompleta = cobertura
            PodeUsarIncremental = incremental
            Me.Reason = reason
        End Sub

        Public ReadOnly Property Degradado As Boolean
            Get
                Return Not (PodeConcluirAusencia AndAlso PodeAfirmarCoberturaCompleta AndAlso
                            PodeUsarIncremental)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Uma linha da matriz da Q8: um ambiente que foi MEDIDO, com a seção do
    ''' FASE2.md onde a medição está registrada.
    '''
    ''' A referência não é enfeite. Sem ela, "ambiente suportado" vira uma
    ''' lista que cresce por conveniência — alguém acrescenta uma linha para
    ''' destravar um caso e ninguém consegue perguntar depois "medido onde?".
    ''' </summary>
    Public NotInheritable Class MeasuredEnvironment
        Public ReadOnly Property Fingerprint As EnvironmentFingerprint
        Public ReadOnly Property Evidence As String

        ''' <summary>
        ''' A mensagem mais ANTIGA que o OOM alcançou quando isto foi medido.
        '''
        ''' É <b>alcance medido</b>, não fronteira da janela, e a distinção é
        ''' exatamente o tipo de leitura excessiva que já custou caro aqui. A
        ''' mensagem mais antiga alcançada é onde os dados por acaso acabam:
        ''' pode ser mais nova que o limite (não existe correio mais velho) ou
        ''' mais velha que ele (itens que chegaram por outro caminho).
        '''
        ''' Serve para <b>descrever</b> o que foi visto — nunca para concluir
        ''' "mais antigo que isto está fora da janela".
        ''' </summary>
        Public ReadOnly Property AlcanceMedido As Date?

        Public Sub New(fp As EnvironmentFingerprint, evidence As String,
                       Optional alcanceMedido As Date? = Nothing)
            Fingerprint = fp
            Me.Evidence = evidence
            Me.AlcanceMedido = alcanceMedido
        End Sub
    End Class

    ''' <summary>
    ''' A matriz da Q8 e o que fazer fora dela.
    '''
    ''' <b>A matriz está INCOMPLETA, e de propósito.</b> A §19.3 mediu que
    ''' levantar as outras linhas nesta máquina custa horas e dezenas de GB de
    ''' download — mais espaço livre do que a máquina tem. Um ambiente não
    ''' medido não vira "provavelmente igual": vira degradação explícita.
    '''
    ''' É a resposta que a Q8 admite hoje: <b>declaração de escopo
    ''' suportado</b>, com fallback conservador para todo o resto. Fingir a
    ''' matriz cheia seria pior que declará-la parcial, porque um número
    ''' inventado não avisa quando está errado.
    ''' </summary>
    Public NotInheritable Class EnvironmentPolicy

        Private Shared ReadOnly _matriz As IReadOnlyList(Of MeasuredEnvironment) = Construir()

        Private Shared Function Construir() As IReadOnlyList(Of MeasuredEnvironment)
            Dim m As New List(Of MeasuredEnvironment)()

            ' UNICA linha medida, e ela e a medicao de 2026-08-24, registrada
            ' na §22.3. O token da janela e o valor CRU do perfil, nao um
            ' numero de meses: ver EnvironmentFingerprint.WindowToken.
            '
            ' A tentacao aqui era escrever "1 mes", que e o que o usuario
            ' tinha dito em conversa e o que a §18/§19 mediram em 2026-08-22.
            ' Seria uma afirmacao nao medida DENTRO da classe que existe para
            ' impedir afirmacoes nao medidas: a janela mudou desde entao (o
            ' usuario aumentou o cache), e a medicao de hoje alcanca 1.979
            ' mensagens ate 2024-10-09, contra as 1.004 de dois dias atras.
            m.Add(New MeasuredEnvironment(
                New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "84-09-00-00"),
                "FASE2 §22.3 — caixa corporativa do usuario, medido em 2026-08-24: " &
                "108 pastas, 95 reportando zero, 1.979 mensagens datadas alcancadas",
                New Date(2024, 10, 9)))

            Return m
        End Function

        Public Shared ReadOnly Property Matriz As IReadOnlyList(Of MeasuredEnvironment)
            Get
                Return _matriz
            End Get
        End Property

        Public Shared Function Medido(fp As EnvironmentFingerprint) As MeasuredEnvironment
            If fp Is Nothing Then Return Nothing
            Return _matriz.FirstOrDefault(Function(x) x.Fingerprint.Equals(fp))
        End Function

        ''' <summary>
        ''' O que este ambiente autoriza. FALHA FECHADO: qualquer coisa que não
        ''' esteja na matriz recebe o mínimo.
        ''' </summary>
        Public Shared Function Capacidades(fp As EnvironmentFingerprint) As EnvironmentCapabilities
            If fp Is Nothing Then
                Return Degradar("ambiente nao identificado")
            End If
            If fp.Provider = ProviderKind.Desconhecido Then
                Return Degradar("provider desconhecido")
            End If
            If fp.CachedMode AndAlso fp.WindowToken Is Nothing Then
                ' O caso mais traicoeiro: sabe-se que ha janela e nao se sabe
                ' qual. Pior que nao saber nada, porque parece identificado.
                Return Degradar("modo cached com janela nao lida")
            End If

            Dim linha = Medido(fp)
            If linha Is Nothing Then
                Return Degradar($"ambiente fora da matriz medida: {fp.Value()}")
            End If

            Return New EnvironmentCapabilities(True, True, True, "")
        End Function

        Private Shared Function Degradar(motivo As String) As EnvironmentCapabilities
            Return New EnvironmentCapabilities(False, False, False, motivo)
        End Function

        ''' <summary>
        ''' Se trocar de <paramref name="antes"/> para <paramref name="depois"/>
        ''' invalida as conclusões de ausência já tiradas.
        '''
        ''' Sempre que a impressão digital muda. A tentação é dizer "só encolher
        ''' a janela invalida, aumentar não" — e é falso nos dois sentidos:
        ''' encolher esconde itens que existem, e aumentar revela itens que
        ''' foram concluídos ausentes, o que também torna a conclusão anterior
        ''' errada.
        ''' </summary>
        Public Shared Function ExigeReconciliacao(antes As EnvironmentFingerprint,
                                                  depois As EnvironmentFingerprint) As Boolean
            If antes Is Nothing OrElse depois Is Nothing Then Return True
            Return Not antes.Equals(depois)
        End Function

    End Class

End Namespace

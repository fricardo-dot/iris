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
    ''' As inferências que o ambiente pode ou não autorizar.
    '''
    ''' São enumeradas uma a uma, e não agrupadas num "ambiente suportado",
    ''' porque cada uma exige uma demonstração <b>própria</b>. Medir que o OOM
    ''' alcança 1.979 mensagens não demonstra nenhuma das três: diz quanto se
    ''' enxergou, não que "não encontrei" signifique excluído.
    ''' </summary>
    Public Enum Inference
        ''' <summary>Transformar <c>NaoEncontrado</c> em <c>AusenteDaPasta</c>.</summary>
        ConcluirAusencia
        ''' <summary>Declarar que uma varredura teve cobertura completa.</summary>
        AfirmarCoberturaCompleta
        ''' <summary>Usar o incremental da Q3 no lugar da varredura cheia.</summary>
        UsarIncremental
    End Enum

    ''' <summary>
    ''' Uma inferência autorizada, com a evidência que a autoriza.
    '''
    ''' A evidência é obrigatória por construção. Sem isso a lista de
    ''' autorizações cresce por conveniência — alguém acrescenta uma entrada
    ''' para destravar um caso e ninguém consegue perguntar depois "medido
    ''' onde?".
    ''' </summary>
    Public NotInheritable Class GrantedInference
        Public ReadOnly Property Inference As Inference
        Public ReadOnly Property Evidence As String

        Public Sub New(inference As Inference, evidence As String)
            If String.IsNullOrWhiteSpace(evidence) Then
                Throw New ArgumentException("inferencia autorizada exige evidencia", NameOf(evidence))
            End If
            Me.Inference = inference
            Me.Evidence = evidence
        End Sub
    End Class

    ''' <summary>
    ''' O ambiente em que uma medição vale, e ele inclui a JANELA.
    '''
    ''' A §18.4 mudou o que "ambiente" significa neste projeto. "Exchange em
    ''' cache" não é um ambiente: é uma família parametrizada pela janela de
    ''' sincronização, e o parâmetro muda <b>o que existe</b>, não só o que
    ''' custa. O mesmo Iris, no mesmo Outlook, com o cursor num lugar ou
    ''' noutro, enxerga caixas diferentes.
    '''
    ''' Por isso a janela entra na impressão digital. Mudar a janela produz um
    ''' ambiente DIFERENTE, e nenhuma conclusão de ausência tirada no anterior
    ''' sobrevive: itens que "sumiram" quando a janela encolheu não foram
    ''' excluídos.
    ''' </summary>
    Public NotInheritable Class EnvironmentFingerprint

        Public ReadOnly Property Provider As ProviderKind
        Public ReadOnly Property CachedMode As Boolean

        ''' <summary>
        ''' A janela de sincronização — e hoje é sempre <c>Nothing</c>, porque
        ''' <b>não há de onde lê-la</b>.
        '''
        ''' Isto começou como "token opaco": os bytes crus de <c>00036601</c> no
        ''' registro do perfil, usados como impressão digital sem decodificar,
        ''' com o argumento de que a impressão digital só precisa ser estável e
        ''' sensível, não semântica. O argumento estava certo. A premissa não —
        ''' e a §22.4 mediu isso do jeito mais direto possível.
        '''
        ''' O usuário moveu o cursor da janela três vezes e o registro do perfil
        ''' inteiro — 294 valores, blobs comparados por hash — não mudou em nada:
        '''
        ''' <code>
        '''   1 mês (inicial)               00036601 = 84-09-00-00
        '''   tentativa de 2 semanas                   84-09-00-00
        '''   3 MESES, conferido na tela               84-09-00-00
        '''   de volta a 1 mês                         84-09-00-00
        ''' </code>
        '''
        ''' Antes disso a §22.3 já tinha medido que o OOM não expõe a janela:
        ''' enumerando os membros de <c>Store</c> não há nenhum <c>Sync*</c> nem
        ''' <c>Window*</c>, só <c>IsCachedExchange</c> e
        ''' <c>ExchangeStoreType</c>. Somando as duas medições: <b>a janela não
        ''' é legível nem pelo OOM nem pelo registro do perfil.</b>
        '''
        ''' A consequência é dura, e fica explícita em vez de escondida: em modo
        ''' cached a impressão digital do ambiente é <b>incompleta por
        ''' construção</b>, e por isso
        ''' <see cref="EnvironmentPolicy.ExigeReconciliacao"/> não tem como
        ''' disparar quando a janela muda. O ambiente degrada sempre. É o
        ''' comportamento certo, mas é degradação <b>permanente</b>, não uma
        ''' pendência que alguém destrava depois.
        '''
        ''' A saída provável não é achar a configuração: é medir o <b>efeito</b>
        ''' dela — ver §22.11. O campo continua existindo para quando houver de
        ''' onde ler.
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

        Private ReadOnly _permitidas As HashSet(Of Inference)

        ''' <summary>Por que não autorizou tudo. Vazio quando autorizou.</summary>
        Public ReadOnly Property Reason As String

        Friend Sub New(permitidas As IEnumerable(Of Inference), reason As String)
            _permitidas = New HashSet(Of Inference)(If(permitidas, Enumerable.Empty(Of Inference)()))
            Me.Reason = reason
        End Sub

        Public Function Permite(i As Inference) As Boolean
            Return _permitidas.Contains(i)
        End Function

        Public ReadOnly Property PodeConcluirAusencia As Boolean
            Get
                Return Permite(Inference.ConcluirAusencia)
            End Get
        End Property

        Public ReadOnly Property PodeAfirmarCoberturaCompleta As Boolean
            Get
                Return Permite(Inference.AfirmarCoberturaCompleta)
            End Get
        End Property

        Public ReadOnly Property PodeUsarIncremental As Boolean
            Get
                Return Permite(Inference.UsarIncremental)
            End Get
        End Property

        Public ReadOnly Property Degradado As Boolean
            Get
                Return _permitidas.Count < [Enum].GetValues(GetType(Inference)).Length
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Uma linha da matriz da Q8: um ambiente RECONHECIDO, e o que nele foi
    ''' demonstrado.
    '''
    ''' Reconhecer e autorizar são coisas separadas, e a separação é a correção
    ''' central que esta classe sofreu. A versão anterior tinha uma linha só
    ''' com a evidência do reconhecimento, e daí concedia as três inferências —
    ''' ou seja, "medi quantos itens o OOM alcança" virava "medi que ausência,
    ''' cobertura e incremental funcionam". São coisas diferentes e nenhuma das
    ''' três estava medida.
    ''' </summary>
    Public NotInheritable Class MeasuredEnvironment
        Public ReadOnly Property Fingerprint As EnvironmentFingerprint

        ''' <summary>Onde o ambiente foi RECONHECIDO. Não autoriza nada sozinho.</summary>
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

        ''' <summary>
        ''' Se o token da janela foi validado como impressão digital. Exige
        ''' <b>as duas</b> propriedades, não só a sensibilidade:
        '''
        '''   - SENSIBILIDADE: muda quando o cursor da janela se move;
        '''   - ESTABILIDADE OPERACIONAL: não muda por conta própria ao longo
        '''     de reinícios do Outlook e do computador.
        '''
        ''' A §22.4 mediu apenas estabilidade de leitura em sessão curta, que
        ''' é bem menos que a segunda. Marcar isto como <c>True</c> só depois
        ''' de cobrir reinício é parte do protocolo, não zelo extra: um token
        ''' que muda sozinho ao reiniciar faria o Iris reconciliar a caixa
        ''' inteira toda vez que o usuário abrisse o Outlook.
        '''
        ''' Enquanto for falso, a linha não autoriza inferência nenhuma, mesmo
        ''' que traga evidências. Se o token não mudar quando a janela muda, a
        ''' impressão digital não distingue os dois universos, e toda
        ''' autorização concedida ao antigo continua valendo no novo — o
        ''' mecanismo inteiro vira decoração.
        '''
        ''' <b>O protocolo foi executado em 2026-08-24, e o resultado foi
        ''' negativo</b> (§22.4): o token não é sensível porque não é a janela.
        ''' Este campo continua existindo — e continua exigindo as duas
        ''' propriedades — para quando aparecer um candidato que seja.
        ''' </summary>
        Public ReadOnly Property TokenValidado As Boolean

        Public ReadOnly Property Grants As IReadOnlyList(Of GrantedInference)

        Public Sub New(fp As EnvironmentFingerprint, evidence As String,
                       Optional alcanceMedido As Date? = Nothing,
                       Optional tokenValidado As Boolean = False,
                       Optional grants As IEnumerable(Of GrantedInference) = Nothing)
            Fingerprint = fp
            Me.Evidence = evidence
            Me.AlcanceMedido = alcanceMedido
            Me.TokenValidado = tokenValidado

            Dim lista = If(grants, Enumerable.Empty(Of GrantedInference)()).ToList()

            ' Duplicata nao e detalhe estetico: a autorizacao e deduplicada por
            ' conjunto, e se a razao fosse calculada da CONTAGEM DA LISTA, tres
            ' grants iguais de UsarIncremental produziriam "autorizou tudo" com
            ' uma permissao so — degradado sem motivo escrito, que e o pior
            ' estado possivel: errado e silencioso.
            Dim vistas As New HashSet(Of Inference)()
            For Each g In lista
                If Not [Enum].IsDefined(GetType(Inference), g.Inference) Then
                    Throw New ArgumentException(
                        $"inferencia fora do enum: {CInt(g.Inference)}", NameOf(grants))
                End If
                If Not vistas.Add(g.Inference) Then
                    Throw New ArgumentException(
                        $"inferencia autorizada duas vezes: {g.Inference}", NameOf(grants))
                End If
            Next

            Me.Grants = lista
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
    '''
    ''' <b>Esta política ainda não é consumida por código de produção.</b>
    ''' Ela é o MODELO da autorização, não o cumprimento dela: nada impede
    ''' hoje um chamador de passar <c>FolderCoverage.Completa</c> direto para
    ''' a <see cref="PresencePolicy"/> sem passar por aqui. Ligar as duas
    ''' coisas é do 2.2, e "a matriz concede zero" não é o mesmo que "o
    ''' produto já impõe zero".
    '''
    ''' <b>Hoje a matriz autoriza ZERO inferências.</b> Isso não é um bug nem
    ''' uma pendência escondida: é o estado honesto. O ambiente do usuário está
    ''' reconhecido, a estabilidade do token está medida, e nenhuma das três
    ''' inferências foi demonstrada. A consequência de produto — o Iris não
    ''' pode, hoje, concluir que uma mensagem sumiu de uma pasta — está
    ''' registrada na §22.5, não escondida atrás de um default permissivo.
    ''' </summary>
    Public NotInheritable Class EnvironmentPolicy

        Private Shared ReadOnly _matriz As IReadOnlyList(Of MeasuredEnvironment) = Construir()

        Private Shared Function Construir() As IReadOnlyList(Of MeasuredEnvironment)
            Dim m As New List(Of MeasuredEnvironment)()

            ' UNICA linha, e ela RECONHECE sem AUTORIZAR.
            '
            ' A janela entra como Nothing porque NAO HA DE ONDE LE-LA. A §22.4
            ' mediu: o usuario moveu o cursor tres vezes e o registro do perfil
            ' inteiro nao mudou. O 00036601, que eu usava como token, nao e a
            ' janela.
            '
            ' Duas tentacoes ja foram recusadas aqui, e as duas eram a mesma
            ' coisa com roupas diferentes:
            '   1. escrever "1 mes" — que veio de uma frase do usuario em
            '      conversa, nao de leitura nenhuma;
            '   2. manter "84-09-00-00" — que veio de uma leitura, mas de uma
            '      leitura do valor ERRADO.
            ' A segunda parecia mais rigorosa que a primeira e nao era: medir
            ' com cuidado a coisa errada nao e melhor que nao medir.
            '
            ' Com a janela em Nothing, Capacidades para no degrau
            ' "cached com janela nao legivel" e nunca chega a consultar esta
            ' linha. Ela permanece porque REGISTRA a medicao — o alcance, as
            ' pastas —, nao porque autorize.
            m.Add(New MeasuredEnvironment(
                New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing),
                "FASE2 §22.3 — caixa corporativa do usuario, medido em 2026-08-24: " &
                "108 pastas, 95 reportando zero, 2.044 mensagens datadas alcancadas",
                New Date(2024, 10, 9),
                tokenValidado:=False))

            Return m
        End Function

        Public Shared ReadOnly Property Matriz As IReadOnlyList(Of MeasuredEnvironment)
            Get
                Return _matriz
            End Get
        End Property

        Public Shared Function Medido(fp As EnvironmentFingerprint) As MeasuredEnvironment
            Return Medido(fp, _matriz)
        End Function

        Friend Shared Function Medido(fp As EnvironmentFingerprint,
                                      matriz As IReadOnlyList(Of MeasuredEnvironment)) As MeasuredEnvironment
            If fp Is Nothing Then Return Nothing
            Return matriz.FirstOrDefault(Function(x) x.Fingerprint.Equals(fp))
        End Function

        ''' <summary>
        ''' O que este ambiente autoriza. FALHA FECHADO em três degraus: fora
        ''' da matriz, token não validado, e inferência sem evidência própria.
        ''' </summary>
        Public Shared Function Capacidades(fp As EnvironmentFingerprint) As EnvironmentCapabilities
            Return Capacidades(fp, _matriz)
        End Function

        ''' <summary>
        ''' Sobrecarga com matriz injetada. Existe para o teste poder exercitar
        ''' o caminho de AUTORIZAÇÃO — sem ela, com a matriz de produção
        ''' autorizando zero, uma política que negasse tudo incondicionalmente
        ''' passaria em todos os testes.
        ''' </summary>
        Friend Shared Function Capacidades(fp As EnvironmentFingerprint,
                                           matriz As IReadOnlyList(Of MeasuredEnvironment)) As EnvironmentCapabilities
            If fp Is Nothing Then
                Return Negar("ambiente nao identificado")
            End If
            If fp.Provider = ProviderKind.Desconhecido Then
                Return Negar("provider desconhecido")
            End If
            If fp.CachedMode AndAlso fp.WindowToken Is Nothing Then
                ' Sabe-se que ha janela e nao se sabe qual — pior que nao saber
                ' nada, porque parece identificado.
                '
                ' Hoje este degrau e PERMANENTE em modo cached, e nao uma
                ' pendencia que alguem destrava: a §22.4 mediu que a janela nao
                ' e legivel nem pelo OOM nem pelo registro do perfil. Enquanto
                ' isso valer, todo ambiente cached degrada aqui e nenhuma linha
                ' da matriz chega a ser consultada.
                Return Negar("modo cached com janela nao legivel (§22.4)")
            End If

            Dim linha = Medido(fp, matriz)
            If linha Is Nothing Then
                Return Negar($"ambiente fora da matriz medida: {fp.Value()}")
            End If

            If Not linha.TokenValidado Then
                ' Sem sensibilidade verificada, a impressao digital nao
                ' distingue dois universos, e autorizar aqui vazaria a
                ' autorizacao para o universo seguinte sem ninguem notar.
                Return Negar("token da janela nao validado: sensibilidade nao medida (§22.4)")
            End If

            If linha.Grants.Count = 0 Then
                Return Negar("ambiente reconhecido, nenhuma inferencia demonstrada")
            End If

            ' A razao vem do conjunto DISTINTO, nunca da contagem da lista.
            Dim permitidas = linha.Grants.Select(Function(g) g.Inference).Distinct().ToList()
            Return New EnvironmentCapabilities(permitidas,
                If(permitidas.Count < [Enum].GetValues(GetType(Inference)).Length,
                   "autorizacao parcial: so o que foi demonstrado", ""))
        End Function

        Private Shared Function Negar(motivo As String) As EnvironmentCapabilities
            Return New EnvironmentCapabilities(Nothing, motivo)
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

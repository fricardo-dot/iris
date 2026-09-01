Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Cache
Imports Iris.Model

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>Os rótulos de todas as pastas, endereçados por <see cref="ItemKey"/>.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ESTA CLASSE PRECISA EXISTIR</b>
    '''
    ''' O cache guarda rótulo por <i>(pasta, encarnação, geração)</i>. A caixa
    ''' dividida e a fila trabalham por <see cref="ItemKey"/>, que é
    ''' <c>EntryID + StoreID</c> — e <b>não carrega a pasta</b>.
    '''
    ''' A ponte entre as duas coisas estava faltando, e a tentação era escrevê-la
    ''' num <c>Select</c> qualquer no meio do caminho. Ela está aqui, sozinha,
    ''' porque a conversão <b>perde informação</b> e a perda precisa de nome.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DUAS PASTAS PODEM DISCORDAR — E AÍ NINGUÉM DECIDE</b>
    '''
    ''' A mesma mensagem pode estar associada a duas pastas, cada uma com a sua
    ''' geração e a sua observação. A pasta A diz <c>precisa_de_mim</c>; a pasta
    ''' B diz <c>fyi</c>. As duas viram o mesmo <see cref="ItemKey"/>.
    '''
    ''' Jogar as duas num dicionário faria a última vencer — e "a última" é a
    ''' ordem em que o acervo enumerou as pastas, que não quer dizer nada. A
    ''' mensagem apareceria numa gaveta que metade da evidência contradiz, e o
    ''' dono não teria como saber que houve desacordo.
    '''
    ''' Então <b>discordância tira a mensagem do mapa</b> e é contada em
    ''' <see cref="LeituraDeRotulos.Divergentes"/>. Ela cai na gaveta das não
    ''' classificadas, que é onde vai tudo que este programa não sabe explicar —
    ''' e a contagem existe para o desacordo não sumir junto.
    '''
    ''' <b>Concordância não é desacordo.</b> Duas pastas dizendo <c>fyi</c> são
    ''' uma informação só, repetida; ela passa.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>UMA LEITURA POR VEZ</b>
    '''
    ''' Vai ao banco por pasta, pela mesma conexão do resto do cache. Quem chama
    ''' serializa — é o mesmo contrato que o acervo já declara.
    ''' </summary>
    Public NotInheritable Class RotulosDoAcervo

        Private ReadOnly _acervo As AcervoDeTodasAsPastas
        Private ReadOnly _cache As RotulosNoCache

        Public Sub New(acervo As AcervoDeTodasAsPastas, cache As RotulosNoCache)
            If acervo Is Nothing Then Throw New ArgumentNullException(NameOf(acervo))
            If cache Is Nothing Then Throw New ArgumentNullException(NameOf(cache))
            _acervo = acervo
            _cache = cache
        End Sub

        ''' <summary>
        ''' Lê tudo o que está publicado, em todas as pastas.
        ''' </summary>
        Public Function Ler() As LeituraDeRotulos
            Dim vistos As New Dictionary(Of ItemKey, RotuloObservado)()
            Dim brigados As New HashSet(Of ItemKey)()
            Dim pastas = 0

            For Each pasta In _acervo.Pastas
                ' PASTA SEM GERACAO PUBLICADA NAO E PASTA VAZIA -- e a mesma
                ' regra da busca e da fila. Publicados ja devolveria vazio, e
                ' pular aqui evita contar a pasta como lida.
                If Not pasta.Manifesto.GenerationKey.HasValue Then Continue For
                pastas += 1

                For Each par In _cache.Publicados(pasta.Chave)
                    Dim chave As New ItemKey(par.Key, pasta.Store)
                    If chave.IsEmpty Then Continue For

                    Dim jaVisto As RotuloObservado = Nothing
                    If Not vistos.TryGetValue(chave, jaVisto) Then
                        vistos(chave) = par.Value
                        Continue For
                    End If

                    If Concordam(jaVisto, par.Value) Then Continue For
                    brigados.Add(chave)
                Next
            Next

            For Each chave In brigados
                vistos.Remove(chave)
            Next

            Dim rotulos As New Dictionary(Of ItemKey, String)()
            Dim casadas As New Dictionary(Of ItemKey, IReadOnlyList(Of String))()

            For Each par In vistos
                rotulos(par.Key) = par.Value.Rotulo
                ' NULO NAO ENTRA NO MAPA, e vazio entra. E o mesmo contrato do
                ' cache, atravessando mais uma camada: ausencia da chave e
                ' "ninguem respondeu sobre regras"; lista vazia e "respondeu, e
                ' nenhuma casou".
                If par.Value.RegrasCasadas IsNot Nothing Then
                    casadas(par.Key) = par.Value.RegrasCasadas
                End If
            Next

            Return New LeituraDeRotulos(rotulos, casadas, brigados.Count, pastas)
        End Function

        ''' <summary>
        ''' Duas observações da mesma mensagem dizem a mesma coisa?
        '''
        ''' <b>Só o rótulo e as regras contam.</b> Confiança, ativação e instante
        ''' diferem naturalmente entre duas pastas varridas em momentos
        ''' diferentes, e tratar isso como desacordo apagaria a classificação de
        ''' toda mensagem que estivesse em duas pastas — trocando um problema
        ''' raro por um comum.
        ''' </summary>
        Private Shared Function Concordam(a As RotuloObservado, b As RotuloObservado) As Boolean
            If Not String.Equals(a.Rotulo, b.Rotulo, StringComparison.Ordinal) Then Return False

            Dim semA = a.RegrasCasadas Is Nothing
            Dim semB = b.RegrasCasadas Is Nothing
            If semA <> semB Then Return False
            If semA Then Return True

            Return a.RegrasCasadas.Count = b.RegrasCasadas.Count AndAlso
                   a.RegrasCasadas.OrderBy(Function(x) x, StringComparer.Ordinal).
                     SequenceEqual(
                       b.RegrasCasadas.OrderBy(Function(x) x, StringComparer.Ordinal),
                       StringComparer.Ordinal)
        End Function

    End Class

    ''' <summary>
    ''' O que a leitura achou — <b>com as perdas contadas</b>.
    '''
    ''' <see cref="Divergentes"/> não é erro, é desacordo entre pastas. Zero é o
    ''' caso normal; qualquer outra coisa quer dizer que aquelas mensagens
    ''' apareceram como não classificadas <i>apesar</i> de terem rótulo.
    ''' </summary>
    Public NotInheritable Class LeituraDeRotulos
        Public ReadOnly Property Rotulos As IReadOnlyDictionary(Of ItemKey, String)
        ''' <summary>
        ''' Ausência da chave é "ninguém respondeu sobre regras"; lista vazia é
        ''' "respondeu, e nenhuma casou". A distinção vem do cache e sobrevive
        ''' até aqui.
        ''' </summary>
        Public ReadOnly Property RegrasCasadas As IReadOnlyDictionary(Of ItemKey, IReadOnlyList(Of String))
        ''' <summary>Mensagens cujo rótulo duas pastas contam diferente.</summary>
        Public ReadOnly Property Divergentes As Integer
        ''' <summary>Quantas pastas tinham geração publicada para ler.</summary>
        Public ReadOnly Property PastasLidas As Integer

        Friend Sub New(rotulos As IReadOnlyDictionary(Of ItemKey, String),
                       regrasCasadas As IReadOnlyDictionary(Of ItemKey, IReadOnlyList(Of String)),
                       divergentes As Integer, pastasLidas As Integer)
            Me.Rotulos = rotulos
            Me.RegrasCasadas = regrasCasadas
            Me.Divergentes = divergentes
            Me.PastasLidas = pastasLidas
        End Sub

        ''' <summary>
        ''' Nada lido. <b>Não é "nada classificado"</b>: é "não deu para ler", e
        ''' quem a devolve tem de dizer isso na tela.
        ''' </summary>
        Public Shared Function Vazia() As LeituraDeRotulos
            Return New LeituraDeRotulos(
                New Dictionary(Of ItemKey, String)(),
                New Dictionary(Of ItemKey, IReadOnlyList(Of String))(), 0, 0)
        End Function
    End Class

End Namespace

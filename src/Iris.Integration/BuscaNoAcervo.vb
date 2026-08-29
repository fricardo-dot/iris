Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Text
Imports Iris.Cache
Imports Iris.Sync
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>Busca textual sobre o acervo — a entrega que o ESCOPO dizia existir.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO NASCE EM 28/08/2026, E NÃO NA FASE 2</b>
    '''
    ''' O <c>ESCOPO.md</c> listava "busca textual" como entregue pela Fase 2
    ''' desde 25/08. Não estava: não havia esquema de busca, nem serviço, nem
    ''' tela. O <see cref="ManifestReader"/> lê o manifesto de uma pasta e
    ''' devolve as linhas publicadas; ele nunca procurou nada.
    '''
    ''' Foi descoberto ao planejar a Fase 4, por revisão externa, e a lacuna é
    ''' pré-condição do resto: sem busca textual não há linha de base contra a
    ''' qual comparar busca semântica, e sem linha de base a Fase 4 não tem
    ''' como ser avaliada.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ELA PROCURA, E O QUE ELA NÃO PODE PROCURAR</b>
    '''
    ''' Assunto e nome do remetente. <b>Não</b> corpo, <b>não</b> anexo — e
    ''' isso não é limitação de implementação: é a regra D1, que proíbe corpo e
    ''' anexo no cache. Quem procurar por uma palavra que só existe no corpo
    ''' não vai achar, e a ressalva do resultado tem de dizer isso.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>EM MEMÓRIA, E O NÚMERO QUE JUSTIFICA</b>
    '''
    ''' O casamento acontece em memória, sobre as linhas do manifesto, e não
    ''' por <c>LIKE</c> no SQLite. Dois motivos, e nenhum é preguiça:
    '''
    '''   • <c>LIKE</c> do SQLite ignora maiúsculas <b>só em ASCII</b>. Numa
    '''     caixa em português, "Regulatório" e "regulatorio" seriam palavras
    '''     diferentes, e uma busca que não acha o que o usuário vê na tela é
    '''     pior que busca nenhuma.
    '''   • O acervo medido em 28/08/2026 tem <b>1.123 linhas</b>. Percorrer
    '''     mil registros de metadado é irrelevante ao lado da leitura do
    '''     banco que já acontece de qualquer jeito.
    '''
    ''' Isto <b>deixa</b> de valer quando o acervo crescer uma ou duas ordens
    ''' de grandeza — e ele foi desenhado para acumular. O ponto de virada não
    ''' é opinião: é quando a busca deixar de ser instantânea. Trocar por FTS5
    ''' ou por coluna normalizada é migração de esquema, e migração é decisão
    ''' de tamanho, não conserto.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ZERO RESULTADO NÃO É "NÃO EXISTE"</b>
    '''
    ''' A §23 proíbe concluir ausência, e a proibição vale aqui mais que em
    ''' qualquer outro lugar: busca é justamente onde o usuário interpreta
    ''' silêncio como resposta.
    '''
    ''' Por isso <see cref="ResultadoDaBusca"/> carrega, <b>no mesmo objeto</b>
    ''' que os achados, quais pastas foram consultadas, com que cobertura, de
    ''' que geração, e quais pastas conhecidas <b>não têm acervo nenhum</b>.
    ''' Zero achados sobre uma caixa cujo acervo é parcial não é informação
    ''' sobre a caixa; é informação sobre o acervo.
    ''' </summary>
    ''' <summary>
    ''' <b>Quanto se sabe sobre este achado.</b>
    '''
    ''' Existe porque a busca ganhou um segundo passe, tolerante a erro de
    ''' digitação e a flexão — e um achado por aproximação é um palpite bom,
    ''' não um achado. Misturar os dois na mesma lista seria dizer "a busca
    ''' achou" quando o certo é "a busca achou algo parecido".
    ''' </summary>
    Public Enum GrauDoAchado
        ''' <summary>Não casou de jeito nenhum.</summary>
        Nenhum = 0

        ''' <summary>Todas as palavras estão ali, como foram digitadas.</summary>
        Exato = 1

        ''' <summary>
        ''' Casou por radical ou por uma letra de diferença. Provavelmente é
        ''' isto — e "provavelmente" é uma palavra que a tela tem de dizer.
        ''' </summary>
        Aproximado = 2
    End Enum

    Public NotInheritable Class BuscaNoAcervo

        Private ReadOnly _acervo As AcervoDeTodasAsPastas
        Private ReadOnly _dreno As PublicationDrain

        ''' <summary>
        ''' A busca recebe o acervo <b>já drenado</b>, e não o banco.
        '''
        ''' Foi o conserto de 28/08/2026, à tarde: antes ela abria o
        ''' <c>ManifestReader</c> por conta própria a cada pergunta, o que é o
        ''' contorno que a §26.2 proíbe. Agora a única fonte é o consumidor que
        ''' o dreno alimenta. Com a entrega travada antes de qualquer uma das
        ''' duas receber, as duas ficam no retrato anterior — e a busca diz
        ''' isso. Uma falha <i>entre</i> as duas entregas deixa o painel à
        ''' frente; ver o <see cref="ConsumidorComposto"/>.
        ''' </summary>
        Public Sub New(acervo As AcervoDeTodasAsPastas, dreno As PublicationDrain)
            If acervo Is Nothing Then Throw New ArgumentNullException(NameOf(acervo))
            _acervo = acervo
            _dreno = dreno
        End Sub

        ''' <summary>
        ''' Procura em todas as pastas que têm geração publicada.
        '''
        ''' Termo vazio devolve resultado vazio <b>com</b> as pastas
        ''' consultadas — e não uma lista de tudo. Busca sem termo que devolve
        ''' o acervo inteiro parece funcionalidade e é acidente.
        ''' </summary>
        Public Function Procurar(termo As String) As ResultadoDaBusca
            Dim t = New TermoDeBusca(termo)
            Dim consultadas As New List(Of PastaConsultada)()
            Dim semAcervo As New List(Of PastaConsultada)()
            Dim achados As New List(Of AchadoDaBusca)()

            For Each pasta In _acervo.Pastas
                Dim manifesto = pasta.Manifesto

                Dim descrita As New PastaConsultada(pasta.Chave, pasta.Nome,
                                                    manifesto.GenerationKey,
                                                    manifesto.Cobertura,
                                                    manifesto.PublishedAt,
                                                    manifesto.Ressalva,
                                                    manifesto.Items.Count)

                ' PASTA SEM GERAÇÃO PUBLICADA NÃO É PASTA VAZIA.
                '
                ' Misturá-la com as consultadas faria o resultado dizer
                ' "procurei aqui e não achei" sobre um lugar onde ninguém
                ' procurou.
                '
                ' E note o que isto NÃO diz. Até 28/08 eu chamava estas pastas
                ' de "nunca varridas", e é mais do que se sabe: sem geração
                ' publicada cabe também a tentativa que foi rejeitada pela S6,
                ' a que foi cancelada, e a que falhou. O que o cache afirma é
                ' que não há acervo publicado — e é só isso que o texto pode
                ' dizer.
                If manifesto.GenerationKey Is Nothing Then
                    semAcervo.Add(descrita)
                    Continue For
                End If

                consultadas.Add(descrita)
                If t.Vazio Then Continue For

                For Each item In manifesto.Items
                    Dim grau = t.Grau(item)
                    If grau <> GrauDoAchado.Nenhum Then
                        achados.Add(New AchadoDaBusca(pasta.Chave, pasta.Nome, item, grau))
                    End If
                Next
            Next

            ' O DRENO: A BUSCA PASSA POR ELE, E TAMBEM O CONSULTA.
            '
            ' ESTE COMENTARIO DESCREVIA CODIGO QUE NAO EXISTE MAIS, e ficou
            ' assim por uma passada inteira. Ele dizia que a busca "le o
            ' ManifestReader de cada pasta e depois so CONSULTA a fila" -- um
            ' desvio declarado da §26.2. O desvio saiu no mesmo dia: a busca le
            ' o AcervoDeTodasAsPastas, que e um IPublicationConsumer ligado ao
            ' dreno e so muda quando o dreno entrega. Eu tirei o desvio e
            ' esqueci a confissao dele aqui; a revisao externa pegou.
            '
            ' O que sobrou e legitimo, e continua necessario: alem de LER pelo
            ' consumidor, a busca CONSULTA a fila para montar a ressalva. Sem
            ' isso o dreno travado sumiria atras de uma lista que parece
            ' completa -- e quem procura concluiria ausencia do silencio.
            Dim pendentes As Integer
            Dim travado As Long?
            Try
                ' SEM DRENO NAO E FILA LIMPA, E SIM FILA NAO OBSERVADA.
                '
                ' Isto devolvia 0 -- a mesma resposta de "olhei e nao ha nada
                ' pendente". O unico chamador de producao passa um dreno real,
                ' entao a tela nunca chegou nesse estado; mas a classe
                ' aceita Nothing, e "aceita e mente" e pior que "nao aceita".
                ' O -1 e o mesmo caminho do banco travado, e cai na frase que
                ' ja existe: "nao consegui conferir".
                pendentes = If(_dreno Is Nothing, -1, _dreno.Pendentes().Count)
                travado = If(_dreno Is Nothing, CType(Nothing, Long?), _dreno.TravadoEm())
            Catch
                ' Banco travado nao pode derrubar a busca: o que ela ja leu
                ' continua valendo. O que nao vale e AFIRMAR que a fila esta
                ' limpa sem ter conseguido olhar.
                pendentes = -1
            End Try

            Return New ResultadoDaBusca(t, achados, consultadas, semAcervo, pendentes, travado)
        End Function

    End Class

    ''' <summary>
    ''' O termo, já normalizado e partido em palavras.
    '''
    ''' <b>Normalização:</b> minúsculas por cultura invariante e remoção de
    ''' diacríticos por decomposição Unicode. "REGULATÓRIO", "regulatorio" e
    ''' "Regulatório" viram a mesma coisa — que é o que qualquer pessoa espera
    ''' de uma busca, e o que o <c>LIKE</c> do SQLite não faz.
    '''
    ''' <b>Conjunção, não disjunção:</b> duas palavras exigem as duas. Busca
    ''' por "amostras aquaba" que devolvesse tudo o que tem "amostras" seria
    ''' ruído com cara de resultado.
    '''
    ''' <b>Subcadeia, não palavra inteira:</b> "regulat" acha "Regulatório".
    ''' Numa caixa em português, com flexão e composição, exigir palavra
    ''' inteira faria o usuário adivinhar a forma exata.
    ''' </summary>
    Public NotInheritable Class TermoDeBusca

        Public ReadOnly Property Original As String
        Public ReadOnly Property Palavras As IReadOnlyList(Of String)

        Public Sub New(termo As String)
            Original = If(termo, "")
            Palavras = Normalizar(Original).
                       Split({" "c, ChrW(9), ChrW(10), ChrW(13)}, StringSplitOptions.RemoveEmptyEntries).
                       ToList()
        End Sub

        Public ReadOnly Property Vazio As Boolean
            Get
                Return Palavras.Count = 0
            End Get
        End Property

        ''' <summary>
        ''' Minúsculas invariantes e sem diacrítico.
        '''
        ''' <c>ToLowerInvariant</c> e não <c>ToLower</c>: a cultura do host não
        ''' pode decidir se uma busca acha. Já houve um teste nesta suíte que
        ''' media a cultura da máquina em vez do código.
        ''' </summary>
        Public Shared Function Normalizar(s As String) As String
            If String.IsNullOrEmpty(s) Then Return ""
            Dim decomposto = s.Normalize(NormalizationForm.FormD)
            Dim sb As New StringBuilder(decomposto.Length)
            For Each c In decomposto
                If CharUnicodeInfo.GetUnicodeCategory(c) <> UnicodeCategory.NonSpacingMark Then
                    sb.Append(c)
                End If
            Next
            Return sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant()
        End Function

        ''' <summary>
        ''' Casa contra assunto <b>e</b> remetente, juntos.
        '''
        ''' Juntos e não separados: procurar "kate regulatorio" tem de achar
        ''' uma mensagem de "Regulatório - Kate", e exigir que cada palavra
        ''' caia no mesmo campo faria isso falhar por um motivo que o usuário
        ''' não tem como adivinhar.
        ''' </summary>
        Public Function Casa(item As ManifestItem) As Boolean
            Return Grau(item) = GrauDoAchado.Exato
        End Function

        ''' <summary>
        ''' <b>Exato, aproximado, ou nada.</b>
        '''
        ''' ------------------------------------------------------------
        ''' <b>POR QUE HÁ UM SEGUNDO PASSE, E O NÚMERO QUE O JUSTIFICA</b>
        '''
        ''' Medido em 29/08/2026 (<c>tools/medir-busca.py</c>, 300 mensagens
        ''' do acervo real, consultas derivadas do próprio assunto de cada
        ''' uma). Tudo o que esta busca <i>promete</i> vale 100%: palavra
        ''' exata, sem acento, caixa trocada, fora de ordem, pedaço de
        ''' palavra, assunto junto com remetente.
        '''
        ''' O que ela não prometia:
        '''
        ''' <list type="bullet">
        ''' <item><b>erro de digitação — 0,4%</b> de 242 consultas;</item>
        ''' <item><b>flexão de número — 0%</b> de 64.</item>
        ''' </list>
        '''
        ''' Numa caixa em português, "reuniões" não acha "reunião" porque o
        ''' singular não é subcadeia do plural. E uma letra trocada zera a
        ''' busca. Eram as duas falhas mecânicas que sobravam, e nenhuma
        ''' delas precisa de significado para ser consertada.
        '''
        ''' ------------------------------------------------------------
        ''' <b>ADITIVO, E NUNCA SUBSTITUTIVO</b>
        '''
        ''' O primeiro passe é <b>exatamente</b> o que existia. Nada que
        ''' achava antes deixa de achar: o segundo passe só roda quando o
        ''' primeiro falha. Foi assim de propósito — os seis casos de 100%
        ''' são a linha de base, e mexer neles para consertar os outros dois
        ''' seria trocar defeito por defeito.
        '''
        ''' ------------------------------------------------------------
        ''' <b>E O APROXIMADO SE DECLARA</b>
        '''
        ''' O grau sai no resultado, e a tela separa. Um achado por
        ''' aproximação é um palpite bom, e palpite misturado com certeza é
        ''' a mesma família de defeito que este projeto passou uma série de
        ''' revisões corrigindo — aqui ela apareceria como "a busca achou",
        ''' quando o certo é "a busca achou algo parecido".
        '''
        ''' O custo tem número também: o segundo passe é mais frouxo, e
        ''' frouxidão é ruído. A medição de 29/08 já mostrava 67,7% das
        ''' buscas por pedaço de palavra casando com mais de dez mensagens.
        ''' Separar os graus é o que impede o ruído novo de contaminar o
        ''' resultado bom.
        ''' </summary>
        Public Function Grau(item As ManifestItem) As GrauDoAchado
            If item Is Nothing OrElse Vazio Then Return GrauDoAchado.Nenhum

            Dim alvo = Normalizar($"{item.Subject} {item.SenderName}")
            If Palavras.All(Function(p) alvo.Contains(p)) Then
                Return GrauDoAchado.Exato
            End If

            ''' PARTE POR TUDO QUE NÃO É LETRA OU DÍGITO, e não só por espaço.
            '''
            ''' Assunto de e-mail vem cheio de pontuação colada: "[Contrato]",
            ''' "contrato/fornecedor", "RES:contrato". Partindo só por espaço, o
            ''' token fica "[contrato]" e a distância até "contrado" passa de
            ''' um — o segundo passe falharia por um motivo que nada tem a ver
            ''' com o que ele quer medir.
            '''
            ''' O primeiro passe não precisa disto: ele casa por subcadeia, e
            ''' subcadeia atravessa pontuação sozinha.
            Dim doAlvo = EmPalavras(alvo)
            If doAlvo.Count = 0 Then Return GrauDoAchado.Nenhum

            If Palavras.All(Function(p) doAlvo.Any(Function(a) Parecidas(p, a))) Then
                Return GrauDoAchado.Aproximado
            End If

            Return GrauDoAchado.Nenhum
        End Function

        ''' <summary>
        ''' <b>Parte por tudo o que não é letra nem dígito.</b>
        '''
        ''' Era uma lista de pontuação ASCII, e o comentário dizia "tudo o que
        ''' não é letra ou dígito" — a revisão externa mediu a diferença e ela é
        ''' real: aspas curvas, travessão, reticências e espaço não separável
        ''' ficavam colados na palavra, e <c>"Contrato"</c> entre aspas tipográficas
        ''' não casava com <c>contrado</c>. Assunto de e-mail vem cheio deles.
        '''
        ''' <c>Char.IsLetterOrDigit</c> em vez de lista: a lista nunca termina, e
        ''' uma lista incompleta com comentário completo é pior que as duas
        ''' coisas separadas.
        ''' </summary>
        Friend Shared Function EmPalavras(alvo As String) As List(Of String)
            Dim saida As New List(Of String)()
            If String.IsNullOrEmpty(alvo) Then Return saida

            Dim atual As New StringBuilder()
            For Each c In alvo
                If Char.IsLetterOrDigit(c) Then
                    atual.Append(c)
                ElseIf atual.Length > 0 Then
                    saida.Add(atual.ToString())
                    atual.Clear()
                End If
            Next
            If atual.Length > 0 Then saida.Add(atual.ToString())
            Return saida
        End Function

        ''' <summary>
        ''' <b>A palavra em PONTOS DE CÓDIGO, e não em unidades UTF-16.</b>
        '''
        ''' Esta é a correção de uma divergência real, e não teórica. O harness
        ''' de medição é Python, que conta pontos de código; o VB conta
        ''' <c>Char</c>, que são unidades UTF-16. Um emoji ocupa <b>dois</b>
        ''' <c>Char</c> e <b>um</b> ponto de código — então
        ''' <c>contra😀to</c> contra <c>Contrato</c> dava <i>Aproximado</i> num
        ''' lado e <i>Nenhum</i> no outro.
        '''
        ''' A revisão externa achou isso <b>fora</b> da tabela de casos
        ''' compartilhada, que é exatamente onde a tabela não protege. O caso
        ''' entrou nela junto com esta correção.
        ''' </summary>
        Friend Shared Function Pontos(p As String) As Integer()
            If String.IsNullOrEmpty(p) Then Return Array.Empty(Of Integer)()

            Dim saida As New List(Of Integer)()
            Dim i = 0
            While i < p.Length
                If Char.IsHighSurrogate(p(i)) AndAlso i + 1 < p.Length AndAlso
                   Char.IsLowSurrogate(p(i + 1)) Then
                    saida.Add(Char.ConvertToUtf32(p(i), p(i + 1)))
                    i += 2
                Else
                    saida.Add(AscW(p(i)))
                    i += 1
                End If
            End While
            Return saida.ToArray()
        End Function

        ''' <summary>
        ''' Duas palavras são "a mesma coisa mal digitada ou flexionada".
        '''
        ''' A ordem importa: radical antes de distância. Radical é barato e
        ''' explica o caso comum; a distância é o resto.
        ''' </summary>
        Friend Shared Function Parecidas(consulta As String, doAlvo As String) As Boolean
            If String.IsNullOrEmpty(consulta) OrElse String.IsNullOrEmpty(doAlvo) Then Return False
            If doAlvo.Contains(consulta) Then Return True

            ''' Radical dos dois lados, e não só de um: quem digita o plural
            ''' procurando o singular e quem faz o contrário têm o mesmo
            ''' direito de achar.
            Dim rc = Radical(consulta)
            Dim ra = Radical(doAlvo)
            If rc.Length >= 4 AndAlso ra.Contains(rc) Then Return True

            ''' UMA LETRA, E SÓ EM PALAVRA LONGA. Distância 1 sobre palavra
            ''' de quatro letras casa metade do dicionário -- e "aproximado"
            ''' que casa com tudo não é aproximado, é lixo.
            '''
            ''' Contado em PONTOS DE CÓDIGO: <c>Length</c> conta unidades UTF-16,
            ''' e um emoji valeria por dois. Ver <see cref="Pontos"/>.
            Dim pc = Pontos(consulta)
            Dim pa = Pontos(doAlvo)
            If pc.Length >= 5 AndAlso Math.Abs(pc.Length - pa.Length) <= 1 Then
                Return DistanciaAte1(consulta, doAlvo)
            End If

            Return False
        End Function

        ''' <summary>
        ''' <b>Radical pobre, de propósito.</b>
        '''
        ''' Não é um stemmer de português — é a lista curta de terminações em
        ''' que o singular <b>não é subcadeia</b> do plural, que são
        ''' exatamente as que a busca por subcadeia já não resolvia. Onde o
        ''' singular É subcadeia ("contratos" contém "contrato"), o primeiro
        ''' passe já achava e não há o que consertar.
        '''
        ''' Só a partir de cinco letras. "mais" viraria "mal" com a mesma
        ''' regra que conserta "contratuais", e um radical que inventa
        ''' palavra curta gera ruído em cima de ruído.
        ''' </summary>
        ''' <b>E o piso hoje NAO e observavel pelo <c>Grau</c>.</b> A guarda
        ''' <c>rc.Length >= 4</c> do <see cref="Parecidas"/> ja rejeita tudo o
        ''' que sairia daqui com menos de cinco letras, entao tirar o piso nao
        ''' muda nenhum resultado de busca -- foi medido, desfazendo-o, e a
        ''' tabela compartilhada continuou passando.
        '''
        ''' Ele fica por dois motivos: e observavel por quem chama
        ''' <c>Radical</c> direto (e ha teste), e a guarda do <c>Parecidas</c>
        ''' pode mudar. Guarda redundante e barata; guarda redundante que
        ''' ninguem sabe que e redundante e armadilha.
        Friend Shared Function Radical(p As String) As String
            If String.IsNullOrEmpty(p) OrElse p.Length < 5 Then Return If(p, "")

            If p.EndsWith("oes") Then Return p.Substring(0, p.Length - 3) & "ao"
            If p.EndsWith("aes") Then Return p.Substring(0, p.Length - 3) & "ao"
            If p.EndsWith("ais") Then Return p.Substring(0, p.Length - 3) & "al"
            If p.EndsWith("eis") Then Return p.Substring(0, p.Length - 3) & "el"
            If p.EndsWith("ois") Then Return p.Substring(0, p.Length - 3) & "ol"
            If p.EndsWith("uis") Then Return p.Substring(0, p.Length - 3) & "ul"
            If p.EndsWith("ns") Then Return p.Substring(0, p.Length - 2) & "m"
            If p.EndsWith("es") Then Return p.Substring(0, p.Length - 2)
            If p.EndsWith("s") Then Return p.Substring(0, p.Length - 1)
            Return p
        End Function

        ''' <summary>
        ''' Distância de edição <b>até 1</b>, e nunca o valor exato.
        '''
        ''' Não é economia: é a pergunta certa. Ninguém aqui quer saber
        ''' "quão diferentes", quer saber "é a mesma palavra mal digitada".
        ''' Parar no primeiro segundo erro é O(n) em vez da matriz inteira.
        ''' </summary>
        Friend Shared Function DistanciaAte1(a As String, b As String) As Boolean
            If a = b Then Return True

            ''' PONTOS DE CÓDIGO, e não Char. Ver <see cref="Pontos"/>: contar
            ''' em UTF-16 fazia um emoji valer por dois e divergir do harness.
            Dim x = Pontos(a)
            Dim y = Pontos(b)

            Dim ia = 0, ib = 0, erros = 0
            While ia < x.Length AndAlso ib < y.Length
                If x(ia) = y(ib) Then
                    ia += 1 : ib += 1
                    Continue While
                End If

                erros += 1
                If erros > 1 Then Return False

                If x.Length = y.Length Then
                    ia += 1 : ib += 1        ' substituicao
                ElseIf x.Length > y.Length Then
                    ia += 1                  ' sobra em a
                Else
                    ib += 1                  ' sobra em b
                End If
            End While

            ''' O que sobrou no fim conta: "abc" e "abcd" saem do laço com
            ''' zero erros e uma letra pendente.
            Return erros + (x.Length - ia) + (y.Length - ib) <= 1
        End Function

    End Class

    ''' <summary>Uma linha achada, e em que pasta.</summary>
    Public NotInheritable Class AchadoDaBusca
        Public ReadOnly Property FolderKey As Long
        Public ReadOnly Property NomeDaPasta As String
        Public ReadOnly Property Item As ManifestItem

        ''' <summary>
        ''' Exato ou aproximado. Viaja <b>com</b> o achado, e não numa lista
        ''' paralela: lista paralela é onde um resultado perde a ressalva no
        ''' caminho até a tela.
        ''' </summary>
        Public ReadOnly Property Grau As GrauDoAchado

        Friend Sub New(folderKey As Long, nome As String, item As ManifestItem,
                       grau As GrauDoAchado)
            Me.FolderKey = folderKey
            NomeDaPasta = nome
            Me.Item = item
            Me.Grau = grau
        End Sub
    End Class

    ''' <summary>Onde se procurou, e com que alcance.</summary>
    Public NotInheritable Class PastaConsultada
        Public ReadOnly Property FolderKey As Long
        Public ReadOnly Property Nome As String
        Public ReadOnly Property GenerationKey As Long?
        Public ReadOnly Property Cobertura As FolderCoverage
        Public ReadOnly Property PublishedAt As String
        Public ReadOnly Property Ressalva As String
        Public ReadOnly Property Itens As Integer

        Friend Sub New(folderKey As Long, nome As String, geracao As Long?,
                       cobertura As FolderCoverage, publicadaEm As String,
                       ressalva As String, itens As Integer)
            Me.FolderKey = folderKey
            Me.Nome = nome
            GenerationKey = geracao
            Me.Cobertura = cobertura
            PublishedAt = publicadaEm
            Me.Ressalva = ressalva
            Me.Itens = itens
        End Sub
    End Class

    ''' <summary>
    ''' O resultado — <b>e a qualificação dele, no mesmo objeto</b>.
    '''
    ''' A revisão externa de 28/08 foi explícita: a ressalva de cobertura tem
    ''' de ser <b>estrutural</b>, e não um texto que uma tela futura possa
    ''' esquecer de mostrar. Aqui ela é: não dá para pegar
    ''' <see cref="Achados"/> sem ter <see cref="Consultadas"/> e
    ''' <see cref="SemAcervo"/> na mão.
    '''
    ''' Isso não é <i>enforcement</i> — nada impede o chamador de ignorar os
    ''' dois. É o mesmo limite que o <see cref="FolderManifest"/> já
    ''' reconhece sobre si mesmo, e ele está dito lá com todas as letras.
    ''' </summary>
    Public NotInheritable Class ResultadoDaBusca

        Public ReadOnly Property Termo As TermoDeBusca
        Public ReadOnly Property Achados As IReadOnlyList(Of AchadoDaBusca)

        ''' <summary>Pastas com acervo publicado — onde se procurou de fato.</summary>
        Public ReadOnly Property Consultadas As IReadOnlyList(Of PastaConsultada)

        ''' <summary>
        ''' Pastas conhecidas que <b>nunca foram varridas</b>.
        '''
        ''' Separadas das consultadas de propósito. Uma pasta sem geração
        ''' publicada não é uma pasta onde não se achou nada: é uma pasta onde
        ''' ninguém procurou, e o usuário precisa saber a diferença antes de
        ''' concluir qualquer coisa do silêncio.
        ''' </summary>
        Public ReadOnly Property SemAcervo As IReadOnlyList(Of PastaConsultada)

        ''' <summary>
        ''' Quantas publicacoes esperam entrega. <b>-1</b> quer dizer que nao
        ''' deu para olhar — e nao zero. Zero afirma fila limpa.
        ''' </summary>
        Public ReadOnly Property PublicacoesPendentes As Integer

        ''' <summary>A geracao em que o dreno emperrou, se emperrou.</summary>
        Public ReadOnly Property DrenoTravadoEm As Long?

        Friend Sub New(termo As TermoDeBusca, achados As IEnumerable(Of AchadoDaBusca),
                       consultadas As IEnumerable(Of PastaConsultada),
                       semAcervo As IEnumerable(Of PastaConsultada),
                       Optional pendentes As Integer = 0,
                       Optional travadoEm As Long? = Nothing)
            Me.Termo = termo
            Me.Achados = If(achados, Enumerable.Empty(Of AchadoDaBusca)()).ToList()
            Me.Consultadas = If(consultadas, Enumerable.Empty(Of PastaConsultada)()).ToList()
            Me.SemAcervo = If(semAcervo, Enumerable.Empty(Of PastaConsultada)()).ToList()
            PublicacoesPendentes = pendentes
            DrenoTravadoEm = travadoEm
        End Sub

        Public ReadOnly Property TotalNoAcervo As Integer
            Get
                Return Consultadas.Sum(Function(p) p.Itens)
            End Get
        End Property

        ''' <summary>
        ''' Alguma pasta consultada tem cobertura menor que completa.
        '''
        ''' Em Exchange em cache isto é <b>sempre</b> verdade hoje (§23), e o
        ''' fato de ser sempre verdade não é motivo para parar de dizer.
        ''' </summary>
        Public ReadOnly Property AlgumaParcial As Boolean
            Get
                Return Consultadas.Any(Function(p) p.Cobertura <> FolderCoverage.Completa)
            End Get
        End Property

        ''' <summary>
        ''' Como a busca <b>deve</b> ser qualificada na tela.
        '''
        ''' Nunca diz "não existe". Zero achados sobre acervo parcial é
        ''' informação sobre o acervo, e não sobre a caixa — e a frase muda
        ''' conforme o que de fato se sabe.
        ''' </summary>
        Public ReadOnly Property Ressalva As String
            Get
                If Termo.Vazio Then
                    Return "Digite alguma coisa para procurar no acervo."
                End If

                Dim partes As New List(Of String)()

                If Consultadas.Count = 0 Then
                    ' "FOI VARRIDA" E MAIS DO QUE SE SABE, e o comentario do
                    ' laco la em cima ja dizia isso desde 28/08: sem geracao
                    ' publicada cabe tambem a varredura rejeitada pela S6, a
                    ' cancelada e a que falhou. Eu corrigi o comentario e nao
                    ' corrigi a FRASE -- e o teste cobrava a frase errada, o
                    ' que e pior: um teste prendendo o comportamento que a
                    ' revisao tinha mandado tirar.
                    partes.Add("Nenhuma pasta tem acervo publicado, então não há acervo " &
                               "onde procurar. Isso não quer dizer que ninguém tentou varrer.")
                Else
                    Dim onde = $"Procurei em {Consultadas.Count} pasta(s) varrida(s), " &
                               $"sobre {TotalNoAcervo} mensagem(ns) guardada(s)."
                    partes.Add(onde)

                    If Achados.Count = 0 Then
                        partes.Add("Nada no acervo observado casa com esse termo. " &
                                   "Isso não quer dizer que não exista na caixa.")
                    End If

                    ' DUAS CAUSAS DIFERENTES, DUAS FRASES DIFERENTES.
                    '
                    ' Até 28/08 isto tratava cobertura Desconhecida como
                    ' parcial e depois afirmava a CAUSA da parcial — "o Outlook
                    ' não expõe tudo". Cobertura desconhecida não tem causa
                    ' conhecida; dizer a causa da outra é inventar diagnóstico.
                    ' Count(Of T) do LINQ, e nao a propriedade Count da lista:
                    ' em VB, `Consultadas.Count(...)` le a PROPRIEDADE e tenta
                    ' indexa-la. Enumerable.Count explicito resolve.
                    Dim parciais = Consultadas.Where(
                        Function(p) p.Cobertura = FolderCoverage.Parcial).Count()
                    Dim ignotas = Consultadas.Where(
                        Function(p) p.Cobertura = FolderCoverage.Desconhecida).Count()

                    If parciais > 0 Then
                        partes.Add("O acervo é parcial: o Outlook não expõe tudo o que existe " &
                                   "no servidor, e o Iris não conclui ausência.")
                    End If
                    If ignotas > 0 Then
                        partes.Add($"Em {ignotas} pasta(s) não dá para dizer o quanto o Iris " &
                                   "enxerga.")
                    End If
                End If

                ' A LIMITAÇÃO QUE MAIS SURPREENDE QUEM USA.
                '
                ' A busca não alcança o corpo, e quem procura uma palavra que
                ' só existe no corpo não vai achar. Dizer isso sempre é o que
                ' impede o usuário de concluir que a mensagem não existe.
                partes.Add("A busca alcança assunto e remetente. O corpo da mensagem não " &
                           "é guardado no cache, então não é procurável.")

                If SemAcervo.Count > 0 Then
                    partes.Add($"{SemAcervo.Count} pasta(s) conhecida(s) não têm acervo publicado " &
                               "e ficaram de fora: " &
                               String.Join(", ", SemAcervo.Select(Function(p) p.Nome)) & ".")
                End If

                ' A RESSALVA DE CADA PASTA NÃO PODE SUMIR.
                '
                ' O PastaConsultada carrega a ressalva do manifesto — que
                ' inclui a CONTRAÇÃO de alcance, quando existe — e até 28/08
                ' este resumo nunca a lia. Uma pasta cujo alcance ENCOLHEU
                ' desde a última varredura é a informação mais acionável que
                ' existe aqui, e ela estava sendo calculada e jogada fora.
                Dim encolheram = Consultadas.
                    Where(Function(p) Not String.IsNullOrWhiteSpace(p.Ressalva) AndAlso
                                      p.Ressalva.Contains("encolheu")).
                    Select(Function(p) p.Nome).ToList()
                If encolheram.Count > 0 Then
                    partes.Add("O alcance do Iris ENCOLHEU em: " &
                               String.Join(", ", encolheram) &
                               ". O que ele guardou antes pode não estar mais lá.")
                End If

                ' O DRENO. Publicação que existe e não foi entregue quer dizer
                ' que o acervo mostrado está atrás do que já foi varrido — e
                ' quem procura precisa saber disso antes de concluir do
                ' silêncio.
                '
                ' ------------------------------------------------------------
                ' ESTA FRASE JA ESTEVE ERRADA TRES VEZES, E NAO NO MESMO EIXO.
                '
                ' (1) Dizia que as publicacoes "nao foram entregues ao acervo".
                ' Falso: a publicacao JA materializa o acervo. O que fica
                ' pendente e a entrega ao CONSUMIDOR.
                '
                ' (2) Virou "a busca ja as enxerga; o painel pode estar
                ' atrasado". Era verdade enquanto a busca contornava o dreno, e
                ' deixou de ser no MESMO dia em que o contorno saiu.
                '
                ' (3) Virou "o retrato anterior -- na busca E no painel". Essa
                ' errou por GENERALIZAR: e verdade quando ninguem drenou, e
                ' falsa na ENTREGA PARCIAL. O ConsumidorComposto entrega ao
                ' painel primeiro e a busca depois, sem transacao. Se a segunda
                ' falha, a geracao continua pendente -- e o painel JA a recebeu.
                ' Nesse estado, tanto o "nem a busca nem o painel" do ramo
                ' travado quanto o "na busca e no painel" deste ramo afirmam
                ' sobre o painel uma coisa que este objeto nao pode saber.
                '
                ' (4) Virou "esta busca nao esta enxergando" / "o que a busca
                ' mostra e o retrato anterior". Eu tinha tirado a afirmacao
                ' categorica sobre o PAINEL e mantido uma sobre a BUSCA -- e ela
                ' e falsa pela OUTRA divida, a de o consumidor ignorar qual
                ' geracao chegou. Com 10 e 11 pendentes e o manifesto ja
                ' apontando para 11, entregar a 10 faz esta busca recarregar a
                ' 11; se a entrega da 11 falhar, a busca esta enxergando
                ' exatamente a geracao que a ressalva jura que ela nao ve.
                '
                ' O PADRAO DAS QUATRO VOLTAS: toda vez eu afirmei o ESTADO de
                ' alguem -- ora do painel, ora da busca. E este objeto nao sabe
                ' o estado de ninguem; ele sabe o estado da FILA.
                '
                ' Entao esta versao afirma so isso, e no modo certo: havendo
                ' entrega pendente, NADA na tela pode ser TRATADO COMO o retrato
                ' da ultima varredura. Nao diz que esta atras, nao diz que esta
                ' a frente, e continua sendo exatamente o que quem procura
                ' precisa saber antes de concluir do silencio.
                '
                ' (5) E A ABERTURA AINDA ERA CATEGORICA. Ela dizia "ainda nao
                ' foram entregues", e Pendentes() nao significa isso: significa
                ' drained_at IS NULL. A entrega e AO MENOS UMA VEZ, e ha uma
                ' janela -- coberta pelo DrenoAposCrashTests, que diz com todas
                ' as letras "o disco diz que a UI NAO recebeu, e ela recebeu".
                ' Entrega NAO CONFIRMADA e o que a fila sabe, e e o que ela diz.
                If DrenoTravadoEm.HasValue Then
                    partes.Add($"A entrega da geração {DrenoTravadoEm.Value} está travada. " &
                               "Enquanto ela não completar, nada aqui pode ser tratado como o " &
                               "retrato da última varredura — nem esta busca, nem o painel do " &
                               "acervo.")
                ElseIf PublicacoesPendentes < 0 Then
                    partes.Add("Não consegui conferir se há varredura esperando entrega.")
                ElseIf PublicacoesPendentes > 0 Then
                    partes.Add($"{PublicacoesPendentes} varredura(s) publicada(s) com entrega " &
                               "não confirmada. Nada aqui pode ser tratado como o retrato da última " &
                               "varredura — nem esta busca, nem o painel do acervo, e os dois " &
                               "podem estar em pontos diferentes.")
                End If

                Return String.Join(" ", partes)
            End Get
        End Property

    End Class

End Namespace

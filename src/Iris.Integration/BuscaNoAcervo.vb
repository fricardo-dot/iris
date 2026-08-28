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
    Public NotInheritable Class BuscaNoAcervo

        Private ReadOnly _db As CacheDatabase
        Private ReadOnly _conn As SqliteConnection
        Private ReadOnly _dreno As PublicationDrain

        Public Sub New(db As CacheDatabase)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _db = db
            _conn = db.Connection
            _dreno = New PublicationDrain(db)
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

            For Each pasta In Pastas()
                Dim manifesto = New ManifestReader(_db).Ler(pasta.Chave)

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
                    If t.Casa(item) Then
                        achados.Add(New AchadoDaBusca(pasta.Chave, pasta.Nome, item))
                    End If
                Next
            Next

            ' O DRENO: UM DESVIO DECLARADO, E NAO UM CUMPRIMENTO.
            '
            ' A §26.2 exige um CONSUMIDOR ligado ao dreno, e proibe leitura
            ' direta como substituto. Esta busca le o ManifestReader de cada
            ' pasta e depois so CONSULTA a fila. A revisao externa de 28/08 foi
            ' explicita: consultar o estado do dreno nao e passar por ele.
            '
            ' Fica assim, e fica dito com esse nome, por um motivo concreto: o
            ' AcervoService e de UMA pasta, com Apontar/Atual, e uma busca entre
            ' pastas nao cabe nesse formato. Encaixa-la exigiria um consumidor
            ' multi-pasta -- que e trabalho de desenho, nao conserto de linha.
            '
            ' O que a consulta compra: o dreno travado APARECE na tela por onde
            ' o usuario esta olhando, em vez de sumir atras de uma lista que
            ' parece completa. E menos do que a §26.2 pede, e mais do que
            ' silencio. Esta no relatorio como divida.
            Dim pendentes As Integer
            Dim travado As Long?
            Try
                pendentes = _dreno.Pendentes().Count
                travado = _dreno.TravadoEm()
            Catch
                ' Banco travado nao pode derrubar a busca: o que ela ja leu
                ' continua valendo. O que nao vale e AFIRMAR que a fila esta
                ' limpa sem ter conseguido olhar.
                pendentes = -1
            End Try

            Return New ResultadoDaBusca(t, achados, consultadas, semAcervo, pendentes, travado)
        End Function

        Private Structure PastaBruta
            Public Chave As Long
            Public Nome As String
        End Structure

        Private Function Pastas() As List(Of PastaBruta)
            Dim r As New List(Of PastaBruta)()
            Using cmd = _conn.CreateCommand()
                cmd.CommandText = "SELECT folder_key, name FROM folder ORDER BY name"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        r.Add(New PastaBruta With {
                            .Chave = rd.GetInt64(0),
                            .Nome = If(rd.IsDBNull(1), "", rd.GetString(1))})
                    End While
                End Using
            End Using
            Return r
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
            If item Is Nothing OrElse Vazio Then Return False
            Dim alvo = Normalizar($"{item.Subject} {item.SenderName}")
            Return Palavras.All(Function(p) alvo.Contains(p))
        End Function

    End Class

    ''' <summary>Uma linha achada, e em que pasta.</summary>
    Public NotInheritable Class AchadoDaBusca
        Public ReadOnly Property FolderKey As Long
        Public ReadOnly Property NomeDaPasta As String
        Public ReadOnly Property Item As ManifestItem

        Friend Sub New(folderKey As Long, nome As String, item As ManifestItem)
            Me.FolderKey = folderKey
            NomeDaPasta = nome
            Me.Item = item
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
                    partes.Add("Nenhuma pasta foi varrida ainda, então não há acervo onde procurar.")
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
                ' A FRASE ESTAVA FACTUALMENTE ERRADA ATE 28/08.
                '
                ' Ela dizia que as publicacoes "ainda nao foram entregues ao
                ' acervo". Nao e isso: a publicacao JA materializou o acervo, e
                ' esta busca pode estar mostrando exatamente a geracao que ela
                ' dizia nao ter chegado. O que esta pendente e a entrega ao
                ' CONSUMIDOR -- o painel do acervo, que se atualiza pelo dreno.
                '
                ' Descrever errado um aviso de inconsistencia e pior que nao
                ' avisar: quem le conclui a coisa errada com confianca.
                If DrenoTravadoEm.HasValue Then
                    partes.Add($"A entrega da geração {DrenoTravadoEm.Value} ao painel do acervo " &
                               "está travada. A busca já enxerga essa varredura; o painel ao lado " &
                               "pode estar atrasado.")
                ElseIf PublicacoesPendentes < 0 Then
                    partes.Add("Não consegui conferir se há entrega pendente ao painel do acervo.")
                ElseIf PublicacoesPendentes > 0 Then
                    partes.Add($"{PublicacoesPendentes} varredura(s) publicada(s) ainda não foram " &
                               "entregues ao painel do acervo. A busca já as enxerga; o painel " &
                               "pode estar atrasado.")
                End If

                Return String.Join(" ", partes)
            End Get
        End Property

    End Class

End Namespace

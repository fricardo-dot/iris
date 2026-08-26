Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.Json
Imports Iris.Assist
Imports Iris.Model

Namespace Global.Iris.Integration

    ''' <summary>Por que a ativação não foi carregada. Fechado, e sem texto livre.</summary>
    Public Enum ActivationLoadFailure
        ''' <summary>Zero: nada aconteceu. Nunca significa "carregou".</summary>
        Nenhuma = 0
        ''' <summary>Não há arquivo de ativação. O caso normal: a IA nasce desligada.</summary>
        Ausente
        ''' <summary>
        ''' O arquivo existe e <b>não passou na conferência de plataforma</b> —
        ''' dono errado, ou permissão de escrita para quem não devia.
        ''' </summary>
        PermissaoRuim
        ''' <summary>
        ''' O caminho existe e <b>não é um arquivo comum</b> — diretório, link,
        ''' ponto de nova análise. Um link faz o conteúdo vir de outro lugar que
        ''' ninguém conferiu.
        ''' </summary>
        NaoEhArquivoComum
        ''' <summary>Passa do teto. Ativação é dezenas de linhas, não megabytes.</summary>
        GrandeDemais
        ''' <summary>Não deu para ler o arquivo.</summary>
        NaoLeu
        ''' <summary>JSON malformado, comentado, com vírgula sobrando, ou fundo demais.</summary>
        JsonInvalido
        ''' <summary>Campo que o Iris não conhece. Pode ser erro de digitação num campo que importa.</summary>
        CampoDesconhecido
        ''' <summary>O mesmo campo duas vezes. Qual das duas valeria?</summary>
        CampoDuplicado
        ''' <summary>Campo obrigatório ausente.</summary>
        CampoFaltando
        ''' <summary>Campo presente com o tipo errado.</summary>
        TipoErrado
        ''' <summary>Valor que não dá para interpretar: enum desconhecido, data não canônica, GUID que não é GUID.</summary>
        ValorInvalido
        ''' <summary>Formalmente incompleta — <see cref="ActivationRecord.Completo"/>.</summary>
        Incompleta
        ''' <summary>Internamente incoerente — <see cref="ActivationRecord.Coerente"/>.</summary>
        Incoerente
        ''' <summary>O endpoint não é HTTPS.</summary>
        EndpointInseguro
        ''' <summary>O prazo passa do máximo que o Iris aceita.</summary>
        PrazoLongoDemais
    End Enum

    ''' <summary>
    ''' O resultado de tentar carregar — <b>o registro ou o motivo</b>, nunca os
    ''' dois nem nenhum.
    '''
    ''' Existe porque <c>ActivationRecord.DaProducao</c> era uma propriedade que
    ''' devolvia <c>Nothing</c>, e <c>Nothing</c> não sabe dizer <i>por quê</i>.
    ''' O usuário que escreveu o arquivo com um campo errado veria a IA
    ''' desligada com a mesma frase de quem nunca escreveu arquivo nenhum, e
    ''' passaria a tarde procurando.
    ''' </summary>
    Public NotInheritable Class ActivationLoadResult

        Public ReadOnly Property Record As ActivationRecord
        Public ReadOnly Property Falha As ActivationLoadFailure
        ''' <summary>
        ''' Qual campo, quando isso ajuda. <b>Nunca o valor</b>: o arquivo tem
        ''' endpoint e nome de pasta, e mensagem de erro vaza para log e para a
        ''' tela.
        ''' </summary>
        Public ReadOnly Property Campo As String

        Private Sub New(r As ActivationRecord, f As ActivationLoadFailure, campo As String)
            Record = r
            Falha = f
            Me.Campo = If(campo, "")
        End Sub

        Friend Shared Function Ok(r As ActivationRecord) As ActivationLoadResult
            Return New ActivationLoadResult(r, ActivationLoadFailure.Nenhuma, Nothing)
        End Function

        Friend Shared Function Nao(f As ActivationLoadFailure,
                                   Optional campo As String = Nothing) As ActivationLoadResult
            Return New ActivationLoadResult(Nothing, f, campo)
        End Function

        Public ReadOnly Property Carregou As Boolean
            Get
                Return Record IsNot Nothing AndAlso Falha = ActivationLoadFailure.Nenhuma
            End Get
        End Property

    End Class

    ''' <summary>
    ''' <b>A cerimônia da §28.3, lida de um arquivo.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE UM ARQUIVO, E NÃO UMA TELA</b>
    '''
    ''' A ativação é o único ponto do Iris em que o usuário assume, por escrito,
    ''' que conteúdo da caixa dele pode sair da máquina. Uma caixa de diálogo com
    ''' um botão "Ativar" transformaria isso num clique entre outros. Um arquivo
    ''' que ele escreve, relê e guarda é um ato deliberado, e continua legível
    ''' meses depois, quando a pergunta for "sob que autorização isto saiu?".
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O PARSE É ESTRITO, E A RECUSA É POR INTEIRO</b>
    '''
    ''' Campo desconhecido, campo repetido, tipo errado, enum que não existe,
    ''' data em formato não canônico: <b>nada é carregado</b>. A tentação de
    ''' ignorar o que não se entende e aproveitar o resto é exatamente como se
    ''' constrói uma autorização que ninguém escreveu — <c>pastas</c> digitado
    ''' <c>pasta</c> viraria lista vazia, e lista vazia num campo obrigatório é
    ''' a diferença entre negar e nem chegar a perguntar.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTE CARREGADOR NÃO CONFERE</b>
    '''
    ''' <b>Dono e permissões do arquivo.</b> Conferir isso exige as APIs de ACL
    ''' do Windows, que não existem no alvo <c>net10.0</c> deste assembly sem
    ''' arrastar uma dependência de plataforma para dentro de uma camada que hoje
    ''' não tem nenhuma.
    '''
    ''' O que se confere no lugar: que é <b>arquivo comum</b>, que <b>não é
    ''' link</b>, e que cabe no teto. O caminho padrão fica sob
    ''' <c>%LOCALAPPDATA%</c>, que o Windows já protege por usuário; numa máquina
    ''' onde isso não vale, um arquivo de ativação é o menor dos problemas.
    ''' Fica <b>declarado</b> em vez de omitido.
    ''' </summary>
    Public NotInheritable Class ActivationLoader

        ''' <summary>Teto do arquivo. Ativação é dezenas de linhas.</summary>
        Public Const TetoDeBytes As Integer = 64 * 1024

        ''' <summary>
        ''' O prazo máximo que o Iris aceita numa ativação.
        '''
        ''' Não é desconfiança do usuário: é que a decisão envelhece. O provedor
        ''' muda de política, a pasta de teste vira pasta de trabalho, e a pessoa
        ''' que autorizou esquece que autorizou. Noventa dias força a
        ''' reencontrar a decisão enquanto ela ainda é lembrada.
        ''' </summary>
        Public Shared ReadOnly PrazoMaximo As TimeSpan = TimeSpan.FromDays(90)

        Private Shared ReadOnly Conhecidos As New HashSet(Of String)(StringComparer.Ordinal) From {
            "id", "versao", "autoridade", "politicaCorporativaVerificada",
            "quando", "ate",
            "provedor", "endpoint", "modelo", "regiao",
            "retencaoAceita", "exigirRetencaoZero", "provedoresPermitidos",
            "operacoes", "pastas", "rotulos", "leituras", "contentBits",
            "ignorarHistorico"}

        Private Shared ReadOnly Obrigatorios As String() = {
            "id", "versao", "autoridade", "politicaCorporativaVerificada",
            "quando", "ate", "provedor", "endpoint", "modelo", "regiao",
            "retencaoAceita", "exigirRetencaoZero", "provedoresPermitidos",
            "operacoes", "pastas", "leituras", "contentBits"}

        ''' <summary><c>%LOCALAPPDATA%\Iris\ativacao.json</c>.</summary>
        Public Shared Function CaminhoPadrao() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "ativacao.json")
        End Function

        ''' <param name="verificador">
        ''' A conferência que <b>depende da plataforma</b> — dono do arquivo e
        ''' quem pode escrever nele. Recebe o <c>FileStream</c> já aberto, e não
        ''' o caminho: conferir por caminho e ler depois deixaria a janela em que
        ''' alguém troca o arquivo entre uma coisa e outra.
        '''
        ''' <c>Nothing</c> quer dizer <b>sem conferência</b>, e é o padrão porque
        ''' este assembly não tem — nem quer ter — dependência de plataforma. Quem
        ''' monta em produção passa a de verdade.
        ''' </param>
        Public Shared Function Carregar(agora As DateTimeOffset,
                                        Optional verificador As Func(Of FileStream, Boolean) = Nothing) _
                                        As ActivationLoadResult
            Return Carregar(CaminhoPadrao(), agora, verificador)
        End Function

        Public Shared Function Carregar(caminho As String, agora As DateTimeOffset,
                                        Optional verificador As Func(Of FileStream, Boolean) = Nothing) _
                                        As ActivationLoadResult
            Dim bruto = LerArquivo(caminho, verificador)
            If bruto.Falha <> ActivationLoadFailure.Nenhuma Then
                Return ActivationLoadResult.Nao(bruto.Falha)
            End If

            Dim doc As JsonDocument = Nothing
            Try
                doc = JsonDocument.Parse(bruto.Bytes, New JsonDocumentOptions With {
                    .AllowTrailingCommas = False,
                    .CommentHandling = JsonCommentHandling.Disallow,
                    .MaxDepth = 8})
            Catch ex As JsonException
                Return ActivationLoadResult.Nao(ActivationLoadFailure.JsonInvalido)
            End Try

            Try
                Return Interpretar(doc.RootElement, agora)
            Finally
                doc.Dispose()
            End Try
        End Function

        ' ==============================================================

        Private Structure Arquivo
            Public Bytes As Byte()
            Public Falha As ActivationLoadFailure
        End Structure

        ''' <summary>
        ''' Abre <b>uma vez</b> e faz tudo sobre o arquivo aberto.
        '''
        ''' A versão anterior media com <c>FileInfo</c> e depois chamava
        ''' <c>ReadAllBytes</c> pelo <b>caminho</b> — duas aberturas, e entre
        ''' elas a janela para trocar o arquivo por um link, ou por um maior. O
        ''' que foi conferido não era necessariamente o que foi lido.
        '''
        ''' Agora: abre, e daí em diante tamanho, permissão e conteúdo saem todos
        ''' do mesmo <c>FileStream</c>.
        ''' </summary>
        Private Shared Function LerArquivo(caminho As String,
                                           verificador As Func(Of FileStream, Boolean)) As Arquivo
            Dim r As Arquivo
            Try
                Dim fi As New FileInfo(caminho)
                If Not fi.Exists Then
                    r.Falha = ActivationLoadFailure.Ausente
                    Return r
                End If

                ' LINK NAO, e a conferencia e ANTES de abrir porque abrir um
                ' link ja teria seguido o link. Sobra uma janela minuscula, e e
                ' o verificador — que le pelo HANDLE — quem a fecha.
                If fi.LinkTarget IsNot Nothing OrElse
                   (fi.Attributes And FileAttributes.ReparsePoint) <> 0 OrElse
                   (fi.Attributes And FileAttributes.Directory) <> 0 Then
                    r.Falha = ActivationLoadFailure.NaoEhArquivoComum
                    Return r
                End If

                Using fs As New FileStream(caminho, FileMode.Open, FileAccess.Read,
                                           FileShare.Read)
                    If fs.Length > TetoDeBytes Then
                        r.Falha = ActivationLoadFailure.GrandeDemais
                        Return r
                    End If

                    If verificador IsNot Nothing AndAlso Not verificador(fs) Then
                        r.Falha = ActivationLoadFailure.PermissaoRuim
                        Return r
                    End If

                    Dim bytes(CInt(fs.Length) - 1) As Byte
                    fs.ReadExactly(bytes)
                    r.Bytes = bytes
                    Return r
                End Using

            Catch ex As Exception
                r.Falha = ActivationLoadFailure.NaoLeu
                Return r
            End Try
        End Function

        Private Shared Function Interpretar(raiz As JsonElement,
                                            agora As DateTimeOffset) As ActivationLoadResult

            If raiz.ValueKind <> JsonValueKind.Object Then
                Return ActivationLoadResult.Nao(ActivationLoadFailure.JsonInvalido)
            End If

            ' CAMPO DESCONHECIDO E CAMPO REPETIDO param aqui, antes de qualquer
            ' leitura. O JsonDocument aceita duplicata em silencio e entrega a
            ' ultima; quem escreveu duas nao sabe qual vale.
            Dim vistos As New HashSet(Of String)(StringComparer.Ordinal)
            For Each p In raiz.EnumerateObject()
                If Not Conhecidos.Contains(p.Name) Then
                    Return ActivationLoadResult.Nao(ActivationLoadFailure.CampoDesconhecido, p.Name)
                End If
                If Not vistos.Add(p.Name) Then
                    Return ActivationLoadResult.Nao(ActivationLoadFailure.CampoDuplicado, p.Name)
                End If
            Next

            For Each nome In Obrigatorios
                If Not vistos.Contains(nome) Then
                    Return ActivationLoadResult.Nao(ActivationLoadFailure.CampoFaltando, nome)
                End If
            Next

            Dim falha As ActivationLoadFailure
            Dim campo As String = Nothing

            Dim id = Texto(raiz, "id", falha, campo)
            Dim versao = Inteiro(raiz, "versao", falha, campo)
            Dim autoridade = Texto(raiz, "autoridade", falha, campo)
            Dim verificada = Booleano(raiz, "politicaCorporativaVerificada", falha, campo)
            Dim quando = Instante(raiz, "quando", falha, campo)
            Dim ate = Instante(raiz, "ate", falha, campo)
            Dim provedor = Texto(raiz, "provedor", falha, campo)
            Dim endpoint = Texto(raiz, "endpoint", falha, campo)
            Dim modelo = Texto(raiz, "modelo", falha, campo)
            Dim regiao = Texto(raiz, "regiao", falha, campo)
            Dim retencao = Texto(raiz, "retencaoAceita", falha, campo)
            Dim zero = Booleano(raiz, "exigirRetencaoZero", falha, campo)
            Dim historico = If(vistos.Contains("ignorarHistorico"),
                               Booleano(raiz, "ignorarHistorico", falha, campo), False)

            Dim slugs = Textos(raiz, "provedoresPermitidos", falha, campo, obrigatorio:=True)
            Dim rotulos = Textos(raiz, "rotulos", falha, campo, obrigatorio:=False)
            Dim operacoes = Enums(Of AssistOperation)(raiz, "operacoes", falha, campo)
            Dim leituras = Enums(Of LabelReadingKind)(raiz, "leituras", falha, campo)
            Dim bits = Inteiros(raiz, "contentBits", falha, campo)
            Dim pastas = LerPastas(raiz, falha, campo)

            If falha <> ActivationLoadFailure.Nenhuma Then
                Return ActivationLoadResult.Nao(falha, campo)
            End If

            Dim r As New ActivationRecord(
                id, versao, autoridade, quando,
                provedor, endpoint, modelo, regiao, retencao,
                operacoes, pastas, rotulos, leituras, bits,
                ate:=ate, ignorarHistorico:=historico,
                politicaCorporativaVerificada:=verificada,
                exigirRetencaoZero:=zero,
                provedoresPermitidos:=slugs)

            ' A ORDEM AQUI IMPORTA: completo, coerente, seguro, e so entao o
            ' prazo. Um registro incompleto tem campos em branco, e conferir
            ' prazo sobre campo em branco daria o motivo errado.
            If Not r.Completo() Then Return ActivationLoadResult.Nao(ActivationLoadFailure.Incompleta)
            If Not r.Coerente() Then Return ActivationLoadResult.Nao(ActivationLoadFailure.Incoerente)
            If Not r.EndpointSeguro() Then
                Return ActivationLoadResult.Nao(ActivationLoadFailure.EndpointInseguro, "endpoint")
            End If
            If r.Ate.Value - r.Quando > PrazoMaximo Then
                Return ActivationLoadResult.Nao(ActivationLoadFailure.PrazoLongoDemais, "ate")
            End If

            Return ActivationLoadResult.Ok(r)
        End Function

        ' ==============================================================
        ' Os leitores. Cada um marca a falha e segue; quem chama confere no fim.

        Private Shared Function Pega(raiz As JsonElement, nome As String,
                                     esperado As JsonValueKind,
                                     ByRef falha As ActivationLoadFailure,
                                     ByRef campo As String,
                                     ByRef achou As JsonElement) As Boolean
            If falha <> ActivationLoadFailure.Nenhuma Then Return False
            If Not raiz.TryGetProperty(nome, achou) Then Return False
            If achou.ValueKind <> esperado Then
                falha = ActivationLoadFailure.TipoErrado : campo = nome
                Return False
            End If
            Return True
        End Function

        Private Shared Function Texto(raiz As JsonElement, nome As String,
                                      ByRef falha As ActivationLoadFailure,
                                      ByRef campo As String) As String
            Dim e As JsonElement = Nothing
            If Not Pega(raiz, nome, JsonValueKind.String, falha, campo, e) Then Return ""
            Return e.GetString()
        End Function

        Private Shared Function Booleano(raiz As JsonElement, nome As String,
                                         ByRef falha As ActivationLoadFailure,
                                         ByRef campo As String) As Boolean
            If falha <> ActivationLoadFailure.Nenhuma Then Return False
            Dim e As JsonElement = Nothing
            If Not raiz.TryGetProperty(nome, e) Then Return False
            ' Booleano NAO aceita 0/1 nem "true": o campo diz sim ou nao, e
            ' aceitar sinonimos e onde "0" de alguem vira True de outro.
            If e.ValueKind = JsonValueKind.True Then Return True
            If e.ValueKind = JsonValueKind.False Then Return False
            falha = ActivationLoadFailure.TipoErrado : campo = nome
            Return False
        End Function

        Private Shared Function Inteiro(raiz As JsonElement, nome As String,
                                        ByRef falha As ActivationLoadFailure,
                                        ByRef campo As String) As Integer
            Dim e As JsonElement = Nothing
            If Not Pega(raiz, nome, JsonValueKind.Number, falha, campo, e) Then Return 0
            Dim v As Integer
            If Not e.TryGetInt32(v) Then
                falha = ActivationLoadFailure.ValorInvalido : campo = nome
                Return 0
            End If
            Return v
        End Function

        ''' <summary>
        ''' Data <b>canônica</b>: ISO-8601 com deslocamento explícito.
        '''
        ''' Sem deslocamento, <c>2026-08-25T12:00:00</c> vira hora local de quem
        ''' lê — e uma autorização que vence em horários diferentes conforme a
        ''' máquina não é uma autorização.
        ''' </summary>
        Private Shared Function Instante(raiz As JsonElement, nome As String,
                                         ByRef falha As ActivationLoadFailure,
                                         ByRef campo As String) As DateTimeOffset
            Dim e As JsonElement = Nothing
            If Not Pega(raiz, nome, JsonValueKind.String, falha, campo, e) Then
                Return DateTimeOffset.MinValue
            End If

            Dim v As DateTimeOffset
            If Not DateTimeOffset.TryParseExact(
                e.GetString(),
                {"yyyy-MM-ddTHH:mm:sszzz", "yyyy-MM-ddTHH:mm:ssZ",
                 "yyyy-MM-ddTHH:mm:ss.FFFFFFFzzz", "yyyy-MM-ddTHH:mm:ss.FFFFFFFZ"},
                Globalization.CultureInfo.InvariantCulture,
                Globalization.DateTimeStyles.AssumeUniversal Or
                Globalization.DateTimeStyles.AdjustToUniversal, v) Then
                falha = ActivationLoadFailure.ValorInvalido : campo = nome
                Return DateTimeOffset.MinValue
            End If
            Return v
        End Function

        Private Shared Function Textos(raiz As JsonElement, nome As String,
                                       ByRef falha As ActivationLoadFailure,
                                       ByRef campo As String,
                                       obrigatorio As Boolean) As List(Of String)
            Dim saida As New List(Of String)()
            If falha <> ActivationLoadFailure.Nenhuma Then Return saida
            Dim e As JsonElement = Nothing
            If Not raiz.TryGetProperty(nome, e) Then
                If obrigatorio Then falha = ActivationLoadFailure.CampoFaltando : campo = nome
                Return saida
            End If
            If e.ValueKind <> JsonValueKind.Array Then
                falha = ActivationLoadFailure.TipoErrado : campo = nome
                Return saida
            End If
            For Each item In e.EnumerateArray()
                If item.ValueKind <> JsonValueKind.String Then
                    falha = ActivationLoadFailure.TipoErrado : campo = nome
                    Return saida
                End If
                saida.Add(item.GetString())
            Next
            Return saida
        End Function

        Private Shared Function Inteiros(raiz As JsonElement, nome As String,
                                         ByRef falha As ActivationLoadFailure,
                                         ByRef campo As String) As List(Of Integer)
            Dim saida As New List(Of Integer)()
            Dim e As JsonElement = Nothing
            If Not Pega(raiz, nome, JsonValueKind.Array, falha, campo, e) Then Return saida
            For Each item In e.EnumerateArray()
                Dim v As Integer
                If item.ValueKind <> JsonValueKind.Number OrElse Not item.TryGetInt32(v) Then
                    falha = ActivationLoadFailure.ValorInvalido : campo = nome
                    Return saida
                End If
                saida.Add(v)
            Next
            Return saida
        End Function

        ''' <summary>
        ''' Enum <b>por nome</b>, e nunca por número.
        '''
        ''' Número num arquivo escrito à mão é um convite ao engano: <c>3</c> não
        ''' diz nada a quem relê, e um membro inserido no meio do enum muda o que
        ''' o arquivo antigo autoriza sem ninguém tocar nele.
        ''' </summary>
        Private Shared Function Enums(Of T As Structure)(raiz As JsonElement, nome As String,
                                                         ByRef falha As ActivationLoadFailure,
                                                         ByRef campo As String) As List(Of T)
            Dim saida As New List(Of T)()
            Dim e As JsonElement = Nothing
            If Not Pega(raiz, nome, JsonValueKind.Array, falha, campo, e) Then Return saida

            For Each item In e.EnumerateArray()
                If item.ValueKind <> JsonValueKind.String Then
                    falha = ActivationLoadFailure.TipoErrado : campo = nome
                    Return saida
                End If

                Dim bruto = If(item.GetString(), "")
                Dim v As T

                ' NUMERO NAO E NOME, e o TryParse do .NET nao concorda: ele
                ' aceita "1", devolve True, e o IsDefined APROVA o resultado —
                ' foi assim que o primeiro teste deste arquivo passou verde
                ' devendo falhar. Um nome de membro comeca com letra ou
                ' sublinhado; "1", "-1" e "+1" nao.
                If bruto.Length = 0 OrElse
                   Not (Char.IsLetter(bruto(0)) OrElse bruto(0) = "_"c) Then
                    falha = ActivationLoadFailure.ValorInvalido : campo = nome
                    Return saida
                End If

                ' ignoreCase:=False de proposito. E o IsDefined depois do
                ' TryParse porque o enum do .NET NAO e fechado: uma lista com
                ' virgula ("Resumir, Redigir") tambem passa pelo TryParse e
                ' produz um valor combinado que nao e membro nenhum.
                If Not [Enum].TryParse(Of T)(bruto, False, v) OrElse
                   Not [Enum].IsDefined(GetType(T), v) Then
                    falha = ActivationLoadFailure.ValorInvalido : campo = nome
                    Return saida
                End If
                saida.Add(v)
            Next
            Return saida
        End Function

        ''' <summary>
        ''' As pastas, cada uma com <c>storeId</c> e <c>entryId</c>.
        '''
        ''' Objeto e não string: pasta é um par, e achatar num texto só faria o
        ''' usuário inventar um separador que um dia aparece dentro de um id.
        ''' </summary>
        Private Shared Function LerPastas(raiz As JsonElement,
                                       ByRef falha As ActivationLoadFailure,
                                       ByRef campo As String) As List(Of FolderKey)
            Dim saida As New List(Of FolderKey)()
            Dim e As JsonElement = Nothing
            If Not Pega(raiz, "pastas", JsonValueKind.Array, falha, campo, e) Then Return saida

            For Each item In e.EnumerateArray()
                If item.ValueKind <> JsonValueKind.Object Then
                    falha = ActivationLoadFailure.TipoErrado : campo = "pastas"
                    Return saida
                End If

                Dim quantas = item.EnumerateObject().Count()
                Dim store As JsonElement = Nothing, entry As JsonElement = Nothing
                If quantas <> 2 OrElse
                   Not item.TryGetProperty("storeId", store) OrElse
                   Not item.TryGetProperty("entryId", entry) OrElse
                   store.ValueKind <> JsonValueKind.String OrElse
                   entry.ValueKind <> JsonValueKind.String Then
                    falha = ActivationLoadFailure.TipoErrado : campo = "pastas"
                    Return saida
                End If

                Dim idDoStore = store.GetString(), idDaPasta = entry.GetString()
                If String.IsNullOrWhiteSpace(idDoStore) OrElse
                   String.IsNullOrWhiteSpace(idDaPasta) Then
                    falha = ActivationLoadFailure.ValorInvalido : campo = "pastas"
                    Return saida
                End If
                ' A ORDEM E (entryId, storeId), e nao a alfabetica que a
                ' intuicao sugere. Trocar os dois compila, roda, e faz a
                ' autorizacao valer para uma pasta que nao existe.
                saida.Add(New FolderKey(entryId:=idDaPasta, storeId:=idDoStore))
            Next
            Return saida
        End Function

    End Class

End Namespace

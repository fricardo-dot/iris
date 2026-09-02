Imports System.Collections.Generic
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json

Namespace Global.Iris.Update

    ''' <summary>
    ''' <b>O que uma versão publicada declara — e a assinatura que prova quem a
    ''' publicou.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ASSINATURA, E NÃO LOGIN</b>
    '''
    ''' Login autentica <i>quem baixa</i>. O que precisa ser garantido aqui é o
    ''' contrário: que o pacote veio de quem diz ter vindo e não foi trocado no
    ''' caminho. Isso é assinatura.
    '''
    ''' E é sério de um jeito que uma atualização de aplicativo comum não é: o
    ''' Iris lê o e-mail do dono. Um atualizador é um <b>canal de execução de
    ''' código</b> apontado para dentro desse programa. Ele merece o mesmo rigor
    ''' que o portão de divulgação tem para o que sai.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DOIS ARQUIVOS, E NÃO UM CAMPO DENTRO DO JSON</b>
    '''
    ''' A assinatura é <b>destacada</b>: <c>iris.json</c> e <c>iris.json.sig</c>.
    ''' A alternativa — um campo <c>assinatura</c> dentro do próprio JSON — obriga
    ''' a assinar "o documento sem aquele campo", e aí a verificação depende de
    ''' reconstruir exatamente os mesmos bytes que o assinador viu: ordem de
    ''' chaves, espaços, escapes, terminador de linha. É uma família inteira de
    ''' defeitos sutis, e cada um deles é uma assinatura que confere sobre bytes
    ''' que não são os que serão lidos.
    '''
    ''' Assinando o arquivo <b>inteiro</b>, byte a byte, não há canonicalização e
    ''' não há como divergir.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE A ASSINATURA NÃO COMPRA</b>
    '''
    ''' Ela prova <i>origem</i>, e não <i>qualidade</i>: uma versão ruim assinada
    ''' pelo dono é uma versão ruim que o Iris vai aceitar. E não substitui o
    ''' certificado de código do Windows — o SmartScreen continua avisando na
    ''' primeira execução de um executável sem certificado de uma autoridade, que
    ''' é outra coisa e custa dinheiro.
    ''' </summary>
    Public NotInheritable Class ManifestoDeVersao

        ''' <summary>A versão publicada, como <c>1.2.3</c>.</summary>
        Public ReadOnly Property Versao As Version

        ''' <summary>Quando ela foi publicada. Só informativo, para a tela.</summary>
        Public ReadOnly Property Publicada As DateTimeOffset

        ''' <summary>O que mudou, em português, para o dono decidir.</summary>
        Public ReadOnly Property Notas As String

        ''' <summary>De onde baixar o pacote. Sempre <c>https</c>.</summary>
        Public ReadOnly Property Endereco As String

        ''' <summary>
        ''' O SHA-256 do pacote, em hexadecimal.
        '''
        ''' <b>Vem de dentro do manifesto assinado</b>, e é isso que o torna útil:
        ''' um hash publicado ao lado do arquivo prova apenas que os dois vieram
        ''' juntos. Este aqui prova que o pacote é o que o dono <i>descreveu no
        ''' documento que assinou</i> — que é o mais forte que se consegue sem
        ''' assinar o executável em si.
        ''' </summary>
        Public ReadOnly Property Sha256 As String

        ''' <summary>O tamanho declarado, em bytes. Ver <see cref="TamanhoMaximo"/>.</summary>
        Public ReadOnly Property Bytes As Long

        Private Sub New(versao As Version, publicada As DateTimeOffset, notas As String,
                        endereco As String, sha256 As String, bytes As Long)
            Me.Versao = versao
            Me.Publicada = publicada
            Me.Notas = If(notas, "")
            Me.Endereco = endereco
            Me.Sha256 = sha256
            Me.Bytes = bytes
        End Sub

        ''' <summary>
        ''' Teto do manifesto, em bytes. Ele tem algumas centenas; qualquer coisa
        ''' maior não é um manifesto, e não vale materializar para descobrir.
        ''' </summary>
        Public Const ManifestoMaximo As Integer = 64 * 1024

        ''' <summary>
        ''' Teto do pacote. Um autocontido de WPF fica na casa das dezenas de
        ''' megabytes; trezentos é folga larga e ainda impede que um endereço
        ''' trocado entregue um download infinito.
        ''' </summary>
        Public Const TamanhoMaximo As Long = 300L * 1024L * 1024L

        ''' <summary>
        ''' <b>Confere a assinatura e só então interpreta o conteúdo.</b>
        '''
        ''' A ordem é o ponto. Interpretar primeiro e conferir depois faria o
        ''' <c>JsonDocument</c> processar bytes de origem desconhecida — que é
        ''' pequeno risco, e é risco desnecessário quando a inversão é de graça.
        '''
        ''' <c>Nothing</c> em qualquer recusa, com o motivo em
        ''' <paramref name="motivo"/>. Não há exceção saindo daqui: a chamada
        ''' acontece num caminho de rede, e desfecho é melhor que exceção.
        ''' </summary>
        Public Shared Function Ler(bytesDoManifesto As Byte(),
                                   assinatura As Byte(),
                                   chavePublica As Byte(),
                                   ByRef motivo As String) As ManifestoDeVersao
            motivo = ""

            If bytesDoManifesto Is Nothing OrElse bytesDoManifesto.Length = 0 Then
                motivo = "o manifesto veio vazio"
                Return Nothing
            End If
            If bytesDoManifesto.Length > ManifestoMaximo Then
                motivo = "o manifesto é grande demais para ser um manifesto"
                Return Nothing
            End If
            If assinatura Is Nothing OrElse assinatura.Length = 0 Then
                motivo = "a versão veio sem assinatura"
                Return Nothing
            End If

            If Not Confere(bytesDoManifesto, assinatura, chavePublica) Then
                motivo = "a assinatura não confere — este arquivo não foi publicado por você"
                Return Nothing
            End If

            Return Interpretar(bytesDoManifesto, motivo)
        End Function

        ''' <summary>
        ''' ECDSA P-256 sobre SHA-256, com a chave pública embutida no programa.
        '''
        ''' <b>Qualquer falha é "não confere"</b>, e não uma exceção: chave
        ''' malformada, assinatura de outro tamanho, curva diferente — nenhuma
        ''' delas é motivo para o Iris parar, e todas são o mesmo desfecho para
        ''' quem pergunta.
        ''' </summary>
        ''' <summary>
        ''' O OID da curva P-256 — <c>prime256v1</c>, <c>secp256r1</c>. É o nome
        ''' dela, e não uma pista sobre ela: <c>KeySize = 256</c> sozinho aceita
        ''' brainpoolP256r1 e secp256k1, que o Windows conhece.
        ''' </summary>
        Private Const OidDaP256 As String = "1.2.840.10045.3.1.7"

        Private Shared Function Confere(dados As Byte(), assinatura As Byte(),
                                        chavePublica As Byte()) As Boolean
            If chavePublica Is Nothing OrElse chavePublica.Length = 0 Then Return False

            Try
                ' "verificador", e nao "ecdsa": o local eclipsaria o TIPO ECDsa,
                ' porque VB e insensivel a maiusculas -- e o erro sai como "tipo
                ' nao pode ser inferido a partir de expressao contendo ecdsa".
                ' CLAUDE.md, primeira secao.
                Using verificador = ECDsa.Create()
                    Dim lidos = 0
                    verificador.ImportSubjectPublicKeyInfo(chavePublica, lidos)

                    ' A SPKI TEM DE SER CONSUMIDA INTEIRA. ImportSubjectPublicKeyInfo
                    ' diz quantos bytes leu e ignora o resto; uma chave valida com
                    ' lixo colado atras seria aceita, e o lixo seria conteudo que
                    ' ninguem olhou dentro do executavel.
                    If lidos <> chavePublica.Length Then Return False

                    ' E TEM DE SER P-256, PELO OID.
                    '
                    ' KeySize = 256 sozinho nao basta: brainpoolP256r1 e
                    ' secp256k1 tambem tem 256 bits, e o Windows conhece as duas.
                    '
                    ' As duas pontas erravam, e de jeitos diferentes: a ferramenta
                    ' de publicacao conferia o COMPRIMENTO da SPKI -- heuristica --
                    ' e o cliente so o KeySize. As duas passaram a conferir o OID
                    ' ao mesmo tempo, e o teste com brainpoolP256r1 e que separou
                    ' os dois casos: ele so ficava verde depois desta linha.
                    If Not EhP256(verificador) Then Return False

                    Return verificador.VerifyData(dados, assinatura, HashAlgorithmName.SHA256)
                End Using
            Catch
                Return False
            End Try
        End Function

        ' "conteudo", e nao "bytes": o parametro eclipsaria a propriedade
        ' Bytes -- VB e insensivel a maiusculas -- e uma leitura nao qualificada
        ' de Bytes acrescentada aqui pegaria o vetor em vez do tamanho.
        ''' <summary>
        ''' A curva pelo nome, e não pelo tamanho. <c>False</c> para curva
        ''' explícita — uma chave que não diga qual curva é não é uma chave que
        ''' se possa afirmar ser P-256.
        ''' </summary>
        Private Shared Function EhP256(quem As ECDsa) As Boolean
            Try
                Dim qual = quem.ExportParameters(False).Curve
                Return qual.IsNamed AndAlso
                       String.Equals(qual.Oid?.Value, OidDaP256, StringComparison.Ordinal)
            Catch
                Return False
            End Try
        End Function

        Private Shared Function Interpretar(conteudo As Byte(),
                                            ByRef motivo As String) As ManifestoDeVersao
            Try
                Using doc = JsonDocument.Parse(conteudo)
                    Dim raiz = doc.RootElement
                    If raiz.ValueKind <> JsonValueKind.Object Then
                        motivo = "o manifesto não é um objeto"
                        Return Nothing
                    End If

                    ' NOMES QUE NAO ECLIPSAM AS PROPRIEDADES. "versao" e
                    ' "endereco" sombreariam Versao e Endereco, ignorando maiusculas,
                    ' e uma referencia nao qualificada acrescentada depois pegaria o
                    ' local. CLAUDE.md, secao 1.
                    Dim numero = Texto(raiz, "versao")
                    Dim quando = Texto(raiz, "publicada")
                    Dim ondeBaixar = Texto(raiz, "endereco")
                    Dim sha = Texto(raiz, "sha256")

                    Dim aVersao As Version = Nothing
                    If Not Version.TryParse(numero, aVersao) Then
                        motivo = "o manifesto não traz uma versão legível"
                        Return Nothing
                    End If

                    ' EXATAMENTE TRES COMPONENTES, e nao "o que o TryParse aceitar".
                    '
                    ' Ele aceita "1.2" e "1.2.3.4". O primeiro faz ToString(3) LANCAR
                    ' -- numa tela, longe daqui. O segundo e pior por ser silencioso:
                    ' 1.2.3.4 e maior que 1.2.3 na comparacao, mas as duas aparecem
                    ' como "1.2.3" na tela e produzem o mesmo nome de arquivo. O dono
                    ' veria o Iris oferecendo a versao que ele ja tem.
                    If aVersao.Build < 0 OrElse aVersao.Revision >= 0 Then
                        motivo = "a versão do manifesto não tem três números"
                        Return Nothing
                    End If

                    ' HTTPS OU NADA, e mesmo vindo de dentro da assinatura.
                    '
                    ' Quem assina é o dono, e um endereco http:// assinado por ele
                    ' seria um engano dele -- e um engano que entrega o pacote a
                    ' quem estiver no caminho. A conferencia custa uma linha.
                    If Not ondeBaixar.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                        motivo = "o endereço do pacote não é https"
                        Return Nothing
                    End If

                    If sha.Length <> 64 OrElse
                       Not sha.All(Function(c) Uri.IsHexDigit(c)) Then
                        motivo = "o manifesto não traz um SHA-256 legível"
                        Return Nothing
                    End If

                    Dim tamanho As Long = 0
                    Dim campoBytes As JsonElement = Nothing
                    If raiz.TryGetProperty("bytes", campoBytes) Then
                        campoBytes.TryGetInt64(tamanho)
                    End If
                    If tamanho <= 0 OrElse tamanho > TamanhoMaximo Then
                        motivo = "o tamanho declarado do pacote não é plausível"
                        Return Nothing
                    End If

                    Dim publicada As DateTimeOffset
                    DateTimeOffset.TryParse(quando, Globalization.CultureInfo.InvariantCulture,
                                            Globalization.DateTimeStyles.RoundtripKind,
                                            publicada)

                    Return New ManifestoDeVersao(aVersao, publicada, Texto(raiz, "notas"),
                                                 ondeBaixar, sha.ToLowerInvariant(), tamanho)
                End Using
            Catch
                motivo = "o manifesto não é JSON legível"
                Return Nothing
            End Try
        End Function

        Private Shared Function Texto(raiz As JsonElement, nome As String) As String
            Dim campo As JsonElement = Nothing
            If Not raiz.TryGetProperty(nome, campo) Then Return ""
            If campo.ValueKind <> JsonValueKind.String Then Return ""
            Return If(campo.GetString(), "")
        End Function

    End Class

End Namespace

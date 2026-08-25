Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Runtime.InteropServices
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' Leitura do rótulo de sensibilidade do Purview (<c>MSIP_Labels</c>).
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ESTE MÓDULO NÃO DECIDE</b>
    '''
    ''' Nada. Ele lê e classifica a leitura. Se o conteúdo pode sair da
    ''' máquina é decisão da política da Fase 3, e ela não autoriza — ver a
    ''' §28.2 do FASE3.md. Misturar leitura com decisão foi o erro que a
    ''' §29.2 descreve.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A PROPRIEDADE E DE ONDE ELA VEM</b>
    '''
    ''' O DASL aponta para uma named property do namespace de <b>cabeçalhos
    ''' de internet</b>. Isso importa para o portão e está registrado como a
    ''' P16: cabeçalho recebido pode ter origem fora do mecanismo
    ''' corporativo, então ler o valor com perfeição não prova que ninguém
    ''' consegue apresentar uma classificação baixa falsa.
    '''
    ''' Daí a assimetria que o resto da fase respeita: sinal restritivo serve
    ''' para <b>negar</b>; ausência ou rótulo baixo não servem, sozinhos,
    ''' para <b>permitir</b>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ARMADILHAS QUE ESTE CÓDIGO EXISTE PARA NÃO REPETIR</b>
    '''
    ''' • Nem toda <c>COMException</c> é ausência. <c>MAPI_E_NOT_FOUND</c>
    '''   (<c>0x8004010F</c>) é; acesso negado, ocupado e item não baixado
    '''   não são, e tratá-los como ausência abriria o portão exatamente
    '''   quando ele mais precisa estar fechado.
    ''' • O valor pode trazer <b>vários</b> GUIDs e registros históricos,
    '''   inclusive <c>Enabled=False</c>. Não é um par nome/valor.
    ''' • <c>Name</c> não é identidade estável; a identidade é o GUID.
    ''' • Valor malformado vira <see cref="LabelReadingKind.Malformed"/>,
    '''   <b>nunca</b> "sem rótulo".
    ''' • A propriedade pode existir com string <b>vazia</b>.
    ''' • Nunca <c>SetProperty</c>, nem para "testar round-trip".
    ''' • R7: <c>mail.PropertyAccessor.GetProperty(...)</c> encadeado deixa
    '''   RCW sem dono. Cada objeto COM é adquirido, usado e liberado.
    ''' </summary>
    Friend Module SensitivityLabels

        ''' <summary>
        ''' O DASL do <c>MSIP_Labels</c>. O GUID é o do conjunto de
        ''' propriedades nomeadas de cabeçalhos de internet.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>A CAIXA ALTA IMPORTA, E CUSTOU UMA MEDIÇÃO INTEIRA</b>
        '''
        ''' A primeira rodada usava <c>msip_labels</c> em minúsculas e devolveu
        ''' <c>Blank</c> para <b>120 de 120</b> itens — o que se leria como
        ''' "ninguém nesta caixa tem rótulo". O controle negativo mostrou por
        ''' quê, no mesmo item e pelo mesmo caminho:
        '''
        '''   <c>MSIP_Labels</c>  → lança <c>MAPI_E_NOT_FOUND</c>
        '''   <c>msip_labels</c>  → devolve <c>String</c> de 0 caracteres
        '''   nome inventado      → lança <c>MAPI_E_NOT_FOUND</c>
        '''
        ''' Nome de named property é <b>sensível a maiúsculas</b>, e a versão em
        ''' minúsculas é <b>outra propriedade</b>, que existe vazia neste store.
        ''' Eu estava lendo a propriedade errada, e ela respondia "vazio" com a
        ''' mesma cara de "sem rótulo".
        '''
        ''' Com a caixa alta certa, os mesmos 100 itens deram 43 <c>Absent</c> e
        ''' 57 <c>Blank</c> — uma distribuição, não um valor único.
        '''
        ''' E o nome inventado ter lançado é o que prova que <c>GetProperty</c>
        ''' <b>não cria</b> mapeamento de named property: o <c>msip_labels</c>
        ''' minúsculo já existia neste store, não foi eu que criei.
        ''' </summary>
        Friend Const DaslMsipLabels As String =
            "http://schemas.microsoft.com/mapi/string/" &
            "{00020386-0000-0000-C000-000000000046}/MSIP_Labels"

        ''' <summary><c>PR_CHANGE_KEY</c>, evidência de versão (P10).</summary>
        Private Const DaslChangeKey As String =
            "http://schemas.microsoft.com/mapi/proptag/0x65E20102"

        Private Const MapiNotFound As Integer = &H8004010F
        Private Const InvalidArg As Integer = &H80070057

        ' ==============================================================

        ''' <summary>
        ''' Lê o rótulo de cada item pedido. Um item que falha <b>não</b>
        ''' derruba os outros: cada linha carrega o próprio desfecho, porque
        ''' um lote que falha inteiro por causa de um item esconde qual era.
        ''' </summary>
        Public Function ReadLabels(ns As OL.NameSpace,
                                   items As IReadOnlyList(Of ItemKey)) _
                                   As OperationResult(Of IReadOnlyList(Of LabelReading))

            Dim saida As New List(Of LabelReading)()
            For Each item In items
                saida.Add(LerUm(ns, item))
            Next
            Return OperationResult(Of IReadOnlyList(Of LabelReading)).Ok(saida)
        End Function

        ''' <summary>
        ''' O controle negativo da P7, no MESMO item e pelo MESMO caminho.
        '''
        ''' Lê três DASLs: o canônico, a variante em minúsculas, e um nome que
        ''' <b>comprovadamente não existe</b>. O que o terceiro devolver é o que
        ''' "ausente" parece nesta conta — e sem isso não dá para saber se
        ''' <c>Blank</c> quer dizer "vazio" ou "não existe".
        '''
        ''' Conjunto <b>fixo</b>. Named property tem mapeamento próprio no
        ''' store, e gerar candidatos em massa mexe nesse mapeamento.
        ''' </summary>
        Public Function ProbeSemantics(ns As OL.NameSpace, item As ItemKey) _
                                       As OperationResult(Of NamedPropertyProbe)
            Dim obj As Object = Nothing
            Dim acessor As OL.PropertyAccessor = Nothing
            Try
                Try
                    obj = ns.GetItemFromID(item.EntryId, item.StoreId)
                Catch ex As COMException
                    Return OperationResult(Of NamedPropertyProbe).Fail(ErrorKind.NotFound, "item")
                End Try
                If obj Is Nothing Then
                    Return OperationResult(Of NamedPropertyProbe).Fail(ErrorKind.NotFound, "item")
                End If

                acessor = DirectCast(CallByName(obj, "PropertyAccessor", CallType.Get),
                                     OL.PropertyAccessor)
                Dim tentativas As New List(Of NamedPropertyProbe.Tentativa)()
                For Each candidato In Candidatos()
                    tentativas.Add(Tentar(acessor, candidato.Rotulo, candidato.Dasl))
                Next
                Return OperationResult(Of NamedPropertyProbe).Ok(
                    New NamedPropertyProbe(item, tentativas))
            Finally
                ComHelpers.Release(acessor)
                ComHelpers.Release(obj)
            End Try
        End Function

        Private Function Candidatos() As IEnumerable(Of (Rotulo As String, Dasl As String))
            Const raiz = "http://schemas.microsoft.com/mapi/string/" &
                         "{00020386-0000-0000-C000-000000000046}/"
            Return {
                ("canonico  (MSIP_Labels)", raiz & "MSIP_Labels"),
                ("minusculo (msip_labels)", raiz & "msip_labels"),
                ("CONTROLE  (nao existe) ", raiz & "x-iris-controle-nao-existe")
            }
        End Function

        Private Function Tentar(acessor As OL.PropertyAccessor, rotulo As String,
                                dasl As String) As NamedPropertyProbe.Tentativa
            Try
                Dim v = acessor.GetProperty(dasl)
                Dim texto = TryCast(v, String)
                Return New NamedPropertyProbe.Tentativa(
                    rotulo, dasl, lancou:=False, hresult:=Nothing, excecao:=Nothing,
                    tipoDoValor:=If(v Is Nothing, Nothing, v.GetType().Name),
                    comprimento:=If(texto Is Nothing, CType(Nothing, Integer?), texto.Length))
            Catch ex As Exception
                Return New NamedPropertyProbe.Tentativa(
                    rotulo, dasl, lancou:=True, hresult:=ex.HResult,
                    excecao:=ex.GetType().Name, tipoDoValor:=Nothing, comprimento:=Nothing)
            End Try
        End Function

        Private Function LerUm(ns As OL.NameSpace, item As ItemKey) As LabelReading
            Dim obj As Object = Nothing
            Dim acessor As OL.PropertyAccessor = Nothing
            Try
                Try
                    obj = ns.GetItemFromID(item.EntryId, item.StoreId)
                Catch ex As COMException
                    Return Falha(item, LabelReadStage.Item, ex.HResult)
                End Try
                If obj Is Nothing Then
                    Return New LabelReading(item, LabelReadingKind.Unreadable, LabelReadStage.Item)
                End If

                Dim versao = LerVersao(obj)

                Try
                    acessor = DirectCast(CallByName(obj, "PropertyAccessor", CallType.Get), OL.PropertyAccessor)
                Catch ex As COMException
                    Return Falha(item, LabelReadStage.Accessor, ex.HResult, versao)
                Catch ex As MissingMemberException
                    ' Item sem PropertyAccessor (classe inesperada). Nao e
                    ' ausencia de rotulo: e ausencia de caminho para saber.
                    Return New LabelReading(item, LabelReadingKind.Unreadable,
                                            LabelReadStage.Accessor, version:=versao)
                End Try
                If acessor Is Nothing Then
                    Return New LabelReading(item, LabelReadingKind.Unreadable,
                                            LabelReadStage.Accessor, version:=versao)
                End If

                Dim cru As Object
                Try
                    cru = acessor.GetProperty(DaslMsipLabels)
                Catch ex As COMException
                    Return Falha(item, LabelReadStage.Value, ex.HResult, versao)
                End Try

                Return Interpretar(item, cru, versao)
            Finally
                ComHelpers.Release(acessor)
                ComHelpers.Release(obj)
            End Try
        End Function

        ' ==============================================================

        ''' <summary>
        ''' Traduz HRESULT em desfecho, e é <b>aqui</b> que a distinção de que a
        ''' fase inteira depende acontece.
        '''
        ''' <b>Só <c>MAPI_E_NOT_FOUND</c> vira ausência</b>, e só quando a etapa
        ''' foi ler o valor. A primeira versão aceitava <c>E_INVALIDARG</c>
        ''' junto — e a medição do 3.0, no mesmo dia, mostrou que esse HRESULT é
        ''' o que a <c>Table</c> devolve para "não aceito este DASL". Ou seja: a
        ''' minha própria medição provava que ele significa "operação recusada",
        ''' e eu o estava lendo como "o item não tem rótulo". Falha aberta.
        '''
        ''' Agora ele é <see cref="LabelReadingKind.Unsupported"/>.
        ''' </summary>
        Private Function Falha(item As ItemKey, etapa As LabelReadStage, hr As Integer,
                               Optional versao As LabelVersionEvidence = Nothing) As LabelReading
            Dim tipo As LabelReadingKind
            Select Case hr
                Case MapiNotFound
                    ' Ausencia so vale se a etapa foi LER O VALOR. Nao achar o
                    ' ITEM com o mesmo HRESULT nao diz nada sobre rotulo.
                    tipo = If(etapa = LabelReadStage.Value,
                              LabelReadingKind.Absent, LabelReadingKind.Unreadable)
                Case InvalidArg
                    ' "Nao aceito esta operacao", e NAO "o item nao tem".
                    tipo = LabelReadingKind.Unsupported
                Case Else
                    tipo = ClassificarPorPolitica(hr)
            End Select
            Return New LabelReading(item, tipo, etapa, hr, version:=versao)
        End Function

        ''' <summary>
        ''' O que não é ausência nem transiente passa pelo MESMO classificador
        ''' que o resto do broker usa. Uma segunda tabela de HRESULTs viveria
        ''' aqui divergindo em silêncio — a lição do <c>ErrorPolicy</c> da
        ''' Fase 2, que já esteve em três cópias.
        ''' </summary>
        Private Function ClassificarPorPolitica(hr As Integer) As LabelReadingKind
            Select Case OutlookFailurePolicy.ClassifyFailure(hr, isMutation:=False,
                                                             mutationAttemptStarted:=False)
                Case ErrorKind.Denied
                    Return LabelReadingKind.Denied
                Case ErrorKind.Busy, ErrorKind.NotConnected
                    Return LabelReadingKind.Transient
                Case Else
                    Return LabelReadingKind.Unreadable
            End Select
        End Function


        ''' <summary>
        ''' Lê o valor em registros, e devolve o que se sabe deles.
        '''
        ''' O valor é uma lista de <c>chave=valor</c> separados por <c>;</c>, e a
        ''' chave tem a forma <c>MSIP_Label_&lt;guid&gt;_&lt;campo&gt;</c>. Um item
        ''' pode carregar <b>vários</b> registros, inclusive de rótulo já
        ''' removido (<c>Enabled=False</c>).
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O QUE A PRIMEIRA VERSÃO DEIXAVA PASSAR</b>
        '''
        ''' Ela ignorava em silêncio todo par que não reconhecia. Com isso:
        '''
        '''   • um valor <b>meio corrompido</b> — um GUID bom e um inválido —
        '''     saía <c>Present</c>, como se estivesse inteiro;
        '''   • o <b>mesmo</b> GUID com <c>Enabled=True</c> e <c>Enabled=False</c>
        '''     saía <c>Present</c>, porque o conjunto ordenado terminava com um
        '''     ativo só. O comentário prometia detectar conflito e o código só
        '''     detectava <i>mais de um GUID</i>.
        '''
        ''' Agora <b>tudo o que não está na gramática conhecida contamina</b>, e
        ''' campo contraditório para o mesmo GUID é conflito.
        '''
        ''' A regra intermediária — contaminar só quando o fragmento mencionasse
        ''' <c>MSIP_Label</c> — também caiu, e pelo mesmo motivo: se outro
        ''' cabeçalho foi colado <i>dentro</i> do valor, o valor está malformado.
        ''' Numa barreira de divulgação, "não entendi este pedaço" nunca pode
        ''' virar "o resto vale".
        '''
        ''' O que continua válido é <b>campo desconhecido dentro de um registro
        ''' bem formado</b>: ele pertence inequivocamente àquele GUID, entra na
        ''' lista de campos observados, e não inventa semântica nenhuma.
        ''' </summary>
        Private Function Interpretar(item As ItemKey, cru As Object,
                                     versao As LabelVersionEvidence) As LabelReading
            If cru Is Nothing Then
                ' Propriedade resolvida e valor nulo: NAO e o mesmo que
                ' MAPI_E_NOT_FOUND, e nao da para afirmar ausencia.
                Return New LabelReading(item, LabelReadingKind.Malformed,
                                        LabelReadStage.Value, version:=versao)
            End If

            Dim texto = TryCast(cru, String)
            If texto Is Nothing Then
                Return New LabelReading(item, LabelReadingKind.Malformed,
                                        LabelReadStage.Parse, version:=versao)
            End If

            If texto.Trim().Length = 0 Then
                Return New LabelReading(item, LabelReadingKind.Blank, LabelReadStage.Value,
                                        rawLength:=texto.Length, version:=versao)
            End If

            Dim lido = Analisar(texto)
            Return New LabelReading(item, lido.Tipo, LabelReadStage.Parse,
                                    labelIds:=lido.Ativos, rawLength:=texto.Length,
                                    version:=versao, campos:=lido.Campos)
        End Function

        ''' <summary>
        ''' O parser. Separado de <c>Interpretar</c> para ser testável sem COM —
        ''' os casos que interessam (corrompido pela metade, conflito no mesmo
        ''' GUID, só histórico) não aparecem nesta caixa, e um parser que só a
        ''' caixa real exercita nunca vê o caso difícil.
        ''' </summary>
        Friend Function Analisar(texto As String) _
                                 As (Tipo As LabelReadingKind, Ativos As IReadOnlyList(Of String),
                                     Campos As IReadOnlyList(Of String))
            Const prefixo = "MSIP_Label_"
            Dim vazio = CType(Array.Empty(Of String)(), IReadOnlyList(Of String))

            Dim ativos As New SortedSet(Of String)(StringComparer.Ordinal)
            Dim inativos As New SortedSet(Of String)(StringComparer.Ordinal)
            Dim campos As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim algumReconhecido = False
            Dim contaminado = False

            For Each bruto In texto.Split(";"c)
                ' Fragmento vazio e so pontuacao: ";;" e ";" no fim.
                If bruto.Trim().Length = 0 Then Continue For

                ' TUDO o que nao esta na gramatica conhecida contamina.
                '
                ' A versao anterior so contaminava se o fragmento MENCIONASSE
                ' MSIP_Label, com o argumento de que lixo sem o prefixo "pode
                ' ser outro cabecalho colado". O argumento se derruba sozinho:
                ' se outro cabecalho foi colado DENTRO do valor de MSIP_Labels,
                ' o valor esta malformado. Numa barreira de divulgacao,
                ' "nao entendi este pedaco" nunca pode virar "o resto vale".
                Dim corteIgual = bruto.IndexOf("="c)
                If corteIgual <= 0 Then
                    contaminado = True
                    Continue For
                End If

                ' Corta no PRIMEIRO "=": o valor pode conter "=" (um Name
                ' escolhido pela empresa, por exemplo) sem que isso seja erro.
                Dim chave = bruto.Substring(0, corteIgual).Trim()
                Dim valor = bruto.Substring(corteIgual + 1).Trim()
                If Not chave.StartsWith(prefixo, StringComparison.OrdinalIgnoreCase) Then
                    contaminado = True
                    Continue For
                End If

                Dim resto = chave.Substring(prefixo.Length)
                Dim corte = resto.LastIndexOf("_"c)
                If corte <= 0 Then
                    contaminado = True
                    Continue For
                End If

                Dim id As Guid
                If Not Guid.TryParse(resto.Substring(0, corte), id) Then
                    ' GUID invalido num registro que se diz de rotulo. Isto e
                    ' exatamente o caso que a versao anterior engolia.
                    contaminado = True
                    Continue For
                End If

                algumReconhecido = True
                Dim campo = resto.Substring(corte + 1)
                campos.Add(campo)
                If Not String.Equals(campo, "Enabled",
                                     StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim g = id.ToString("D", CultureInfo.InvariantCulture)
                If String.Equals(valor, "True", StringComparison.OrdinalIgnoreCase) Then
                    ativos.Add(g)
                ElseIf String.Equals(valor, "False", StringComparison.OrdinalIgnoreCase) Then
                    inativos.Add(g)
                Else
                    ' Enabled com valor que nao e nem True nem False.
                    contaminado = True
                End If
            Next

            ' Contaminacao vence tudo: valor que eu nao entendo inteiro e o caso
            ' em que eu NAO posso decidir.
            If contaminado OrElse Not algumReconhecido Then
                Return (LabelReadingKind.Malformed, vazio, campos.ToList())
            End If

            ' O MESMO GUID dizendo as duas coisas. A versao anterior devolvia
            ' Present aqui, porque o conjunto de ativos terminava com um so.
            If ativos.Overlaps(inativos) Then
                Return (LabelReadingKind.Conflicting, ativos.ToList(), campos.ToList())
            End If

            If ativos.Count > 1 Then
                Return (LabelReadingKind.Conflicting, ativos.ToList(), campos.ToList())
            End If
            If ativos.Count = 1 Then
                Return (LabelReadingKind.Present, ativos.ToList(), campos.ToList())
            End If

            ' Forma boa, historico presente, nenhum ativo: rotulo removido.
            If inativos.Count > 0 Then
                Return (LabelReadingKind.HistoricalOnly, vazio, campos.ToList())
            End If

            ' Reconheceu registro de rotulo mas nenhum campo Enabled.
            Return (LabelReadingKind.Malformed, vazio, campos.ToList())
        End Function

        ''' <summary>
        ''' Evidência de versão, best-effort. Falhar aqui não pode derrubar a
        ''' leitura do rótulo — é informação a mais, não requisito.
        ''' </summary>
        Private Function LerVersao(obj As Object) As LabelVersionEvidence
            Dim entryId = TextoDe(Function() CStr(CallByName(obj, "EntryID", CallType.Get)))
            Dim modificado As DateTimeOffset? = Nothing
            Try
                Dim d = CDate(CallByName(obj, "LastModificationTime", CallType.Get))
                modificado = New DateTimeOffset(d)
            Catch
            End Try
            Return New LabelVersionEvidence(entryId, modificado, ChangeKeyDe(obj))
        End Function

        Private Function ChangeKeyDe(obj As Object) As String
            Dim acessor As OL.PropertyAccessor = Nothing
            Try
                acessor = DirectCast(CallByName(obj, "PropertyAccessor", CallType.Get), OL.PropertyAccessor)
                If acessor Is Nothing Then Return Nothing
                Dim bytes = TryCast(acessor.GetProperty(DaslChangeKey), Byte())
                If bytes Is Nothing OrElse bytes.Length = 0 Then Return Nothing
                Return Convert.ToHexString(bytes)
            Catch
                Return Nothing
            Finally
                ComHelpers.Release(acessor)
            End Try
        End Function

        Private Function TextoDe(f As Func(Of String)) As String
            Try
                Return If(f(), "")
            Catch
                Return ""
            End Try
        End Function

    End Module

End Namespace

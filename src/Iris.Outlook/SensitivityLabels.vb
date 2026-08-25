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
        ''' Traduz HRESULT em desfecho, e é <b>aqui</b> que a distinção que a
        ''' fase inteira depende acontece.
        '''
        ''' <c>MAPI_E_NOT_FOUND</c> e <c>E_INVALIDARG</c> significam "esta
        ''' propriedade não existe neste item". Tudo o mais é outra coisa, e
        ''' outra coisa nunca vira ausência.
        ''' </summary>
        Private Function Falha(item As ItemKey, etapa As LabelReadStage, hr As Integer,
                               Optional versao As LabelVersionEvidence = Nothing) As LabelReading
            Dim tipo As LabelReadingKind
            Select Case hr
                Case MapiNotFound, InvalidArg
                    ' Ausencia so vale se a etapa foi LER O VALOR. Nao achar o
                    ' ITEM com o mesmo HRESULT nao diz nada sobre rotulo.
                    tipo = If(etapa = LabelReadStage.Value,
                              LabelReadingKind.Absent, LabelReadingKind.Unreadable)
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
        ''' Interpreta o valor cru. Só string interessa; qualquer outro tipo é
        ''' formato que eu não conheço, e formato desconhecido é
        ''' <see cref="LabelReadingKind.Malformed"/> — nunca ausência.
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

            Dim ativos = GuidsAtivos(texto)
            If ativos Is Nothing Then
                Return New LabelReading(item, LabelReadingKind.Malformed, LabelReadStage.Parse,
                                        rawLength:=texto.Length, version:=versao)
            End If

            Dim tipo = If(ativos.Count > 1, LabelReadingKind.Conflicting, LabelReadingKind.Present)
            If ativos.Count = 0 Then
                ' Ha valor, e nenhum rotulo ATIVO nele. Historico de rotulo
                ' removido cai aqui. Nao e ausencia da propriedade, e nao e
                ' um rotulo: e um valor que eu nao sei traduzir em decisao.
                tipo = LabelReadingKind.Malformed
            End If

            Return New LabelReading(item, tipo, LabelReadStage.Parse,
                                    labelIds:=ativos, rawLength:=texto.Length, version:=versao)
        End Function

        ''' <summary>
        ''' Os GUIDs dos rótulos <b>ativos</b>.
        '''
        ''' O valor é uma lista de registros <c>chave=valor</c> separados por
        ''' <c>;</c>, e um item pode carregar <b>mais de um</b> registro —
        ''' inclusive de rótulo já removido, marcado <c>Enabled=False</c>.
        ''' Modelar como par único perderia exatamente o caso interessante.
        '''
        ''' Devolve <c>Nothing</c> quando o texto não tem forma reconhecível —
        ''' que é diferente de ter forma e nenhum ativo.
        ''' </summary>
        Private Function GuidsAtivos(texto As String) As IReadOnlyList(Of String)
            Dim pares = texto.Split(";"c).
                        Select(Function(p) p.Split("="c)).
                        Where(Function(p) p.Length = 2).
                        Select(Function(p) (Chave:=p(0).Trim(), Valor:=p(1).Trim())).
                        ToList()
            If pares.Count = 0 Then Return Nothing

            ' Chave e da forma "MSIP_Label_<guid>_<campo>". O <campo> Enabled
            ' diz se aquele rotulo vale; ausencia de Enabled explicito NAO
            ' vira ativo por conta propria.
            Dim ativos As New SortedSet(Of String)(StringComparer.Ordinal)
            Dim algumReconhecido = False

            For Each par In pares
                Dim k = par.Chave
                If Not k.StartsWith("MSIP_Label_", StringComparison.OrdinalIgnoreCase) Then Continue For
                Dim resto = k.Substring("MSIP_Label_".Length)
                Dim corte = resto.LastIndexOf("_"c)
                If corte <= 0 Then Continue For

                Dim bruto = resto.Substring(0, corte)
                Dim campo = resto.Substring(corte + 1)
                Dim id As Guid
                If Not Guid.TryParse(bruto, id) Then Continue For

                algumReconhecido = True
                If String.Equals(campo, "Enabled", StringComparison.OrdinalIgnoreCase) AndAlso
                   String.Equals(par.Valor, "True", StringComparison.OrdinalIgnoreCase) Then
                    ativos.Add(id.ToString("D", CultureInfo.InvariantCulture))
                End If
            Next

            If Not algumReconhecido Then Return Nothing
            Return ativos.ToList()
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

Imports System.Collections.Generic
Imports System.Linq

Namespace Global.Iris.Core

    ''' <summary>
    ''' O schema do cache, descrito como DADO.
    '''
    ''' Não é string de DDL: é um modelo que o gerador de DDL e os testes de
    ''' invariante leem. Se fosse DDL solta, os testes teriam de reparsear e
    ''' os dois lados poderiam divergir — que é exatamente o erro que a Q1
    ''' cobrou quando o teste sintético verificava um algoritmo diferente do
    ''' que rodava contra o Outlook.
    '''
    ''' As proibições que este modelo existe para tornar verificáveis estão
    ''' em FASE2.md §14 (I1–I8) e §17 (S1–S7). Cada uma tem um teste que
    ''' FALHA quando ela é violada — sem isso são comentário, e este projeto
    ''' já viu a regra do <c>Permission</c> ficar escrita por três marcos
    ''' enquanto o gate falhava aberto.
    ''' </summary>
    Public NotInheritable Class SchemaColumn
        Public ReadOnly Property Name As String
        Public ReadOnly Property Kind As String
        Public ReadOnly Property IsPrimaryKey As Boolean
        Public ReadOnly Property IsRequired As Boolean
        ''' <summary>Tabela referenciada, ou Nothing.</summary>
        Public ReadOnly Property References As String
        ''' <summary>
        ''' Marca que este valor vem do PROVIDER (Outlook), e não do Iris.
        ''' É o que o I1 usa para proibir que ele seja chave primária.
        ''' </summary>
        Public ReadOnly Property IsProviderValue As Boolean

        Public Sub New(name As String, kind As String,
                       Optional isPrimaryKey As Boolean = False,
                       Optional isRequired As Boolean = False,
                       Optional references As String = Nothing,
                       Optional isProviderValue As Boolean = False)
            Me.Name = name
            Me.Kind = kind
            Me.IsPrimaryKey = isPrimaryKey
            Me.IsRequired = isRequired
            Me.References = references
            Me.IsProviderValue = isProviderValue
        End Sub
    End Class

    Public NotInheritable Class SchemaTable
        Public ReadOnly Property Name As String
        Public ReadOnly Property Columns As IReadOnlyList(Of SchemaColumn)
        ''' <summary>Conjuntos de colunas com índice ÚNICO.</summary>
        Public ReadOnly Property UniqueIndexes As IReadOnlyList(Of String())

        Public Sub New(name As String, columns As IEnumerable(Of SchemaColumn),
                       Optional uniqueIndexes As IEnumerable(Of String()) = Nothing)
            Me.Name = name
            Me.Columns = columns.ToList()
            Me.UniqueIndexes = If(uniqueIndexes, Enumerable.Empty(Of String())()).ToList()
        End Sub

        Public Function Column(nome As String) As SchemaColumn
            Return Columns.FirstOrDefault(
                Function(c) String.Equals(c.Name, nome, StringComparison.OrdinalIgnoreCase))
        End Function
    End Class

    Public NotInheritable Class CacheSchema
        Public ReadOnly Property Tables As IReadOnlyList(Of SchemaTable)

        Public Sub New(tables As IEnumerable(Of SchemaTable))
            Me.Tables = tables.ToList()
        End Sub

        Public Function Table(nome As String) As SchemaTable
            Return Tables.FirstOrDefault(
                Function(t) String.Equals(t.Name, nome, StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>
        ''' O schema pretendido pelo 2.1.
        '''
        ''' A forma vem inteira da Fase 2. Cada separação abaixo existe
        ''' porque uma medição obrigou:
        '''
        '''   • item x incarnation — a §11.1 mediu que EntryID e RecordKey
        '''     MUDAM num Move. Identidade e localizador não podem ser a
        '''     mesma linha.
        '''   • correlation_edge — a §11.2 mediu que nenhuma propriedade é
        '''     única E estável. Toda união é asserção, e asserção que não
        '''     dá para desfazer é dívida permanente.
        '''   • association com estado — a §16.3 mediu que ausência na pasta
        '''     não distingue movido de excluído.
        '''   • generation com cobertura — a §19.2 achou pastas cheias que o
        '''     OOM reporta com ZERO itens.
        ''' </summary>
        ''' <summary>Um indice unico sobre as colunas dadas.</summary>
        Private Shared Function Unico(ParamArray colunas As String()) As String()
            Return colunas
        End Function

        Private Shared Function Col(nome As String, tipo As String,
                                    Optional pk As Boolean = False,
                                    Optional obrigatoria As Boolean = False,
                                    Optional refs As String = Nothing,
                                    Optional doProvider As Boolean = False) As SchemaColumn
            Return New SchemaColumn(nome, tipo, pk, obrigatoria, refs, doProvider)
        End Function

        ''' <summary>
        ''' O schema pretendido pelo 2.1.
        '''
        ''' A forma vem inteira da Fase 2. Cada separacao abaixo existe porque
        ''' uma medicao obrigou, e a referencia esta junto de cada uma.
        '''
        ''' NOTA DE SINTAXE: isto e montado com List e Add, e nao com um
        ''' literal aninhado. O literal nao aceita linha em branco nem linha
        ''' so de comentario no meio — a continuacao implicita do VB quebra, e
        ''' o erro aparece dezenas de linhas adiante.
        ''' </summary>
        Public Shared Function Intended() As CacheSchema
            Dim t As New List(Of SchemaTable)()

            t.Add(New SchemaTable("store", {
                Col("store_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("provider_store_id", "TEXT", obrigatoria:=True, doProvider:=True),
                Col("display_name", "TEXT")
            }, {Unico("provider_store_id")}))

            ' §19.2: a pasta pode estar PARCIALMENTE no cache. Sem 'coverage',
            ' Count == 0 vira 'pasta vazia' — e na caixa medida dezenas de
            ' pastas cheias reportam zero.
            t.Add(New SchemaTable("folder", {
                Col("folder_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("store_key", "INTEGER", obrigatoria:=True, refs:="store"),
                Col("provider_entry_id", "TEXT", obrigatoria:=True, doProvider:=True),
                Col("name", "TEXT"),
                Col("coverage", "TEXT", obrigatoria:=True)
            }))

            ' A identidade logica. Nao carrega nada do provider: a §11.1 mediu
            ' que EntryID e RecordKey MUDAM num Move.
            t.Add(New SchemaTable("item", {
                Col("item_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("created_at", "TEXT", obrigatoria:=True)
            }))

            ' A encarnacao: onde os localizadores do provider vivem, e eles
            ' podem trocar sem que o item mude.
            t.Add(New SchemaTable("incarnation", {
                Col("incarnation_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("item_key", "INTEGER", obrigatoria:=True, refs:="item"),
                Col("folder_key", "INTEGER", obrigatoria:=True, refs:="folder"),
                Col("provider_entry_id", "TEXT", obrigatoria:=True, doProvider:=True),
                Col("provider_record_key", "TEXT", doProvider:=True),
                Col("search_key", "TEXT", doProvider:=True),
                Col("internet_message_id", "TEXT", doProvider:=True),
                Col("first_seen_generation", "INTEGER", refs:="generation"),
                Col("last_seen_generation", "INTEGER", refs:="generation")
            }))

            ' §16.3 e §3.4: ausencia e ESTADO, nunca DELETE da linha.
            ' §17 S2: e so geracao completa e valida pode marca-la.
            t.Add(New SchemaTable("association", {
                Col("association_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("item_key", "INTEGER", obrigatoria:=True, refs:="item"),
                Col("folder_key", "INTEGER", obrigatoria:=True, refs:="folder"),
                Col("presence", "TEXT", obrigatoria:=True),
                Col("absent_by_generation", "INTEGER", refs:="generation")
            }))

            ' §14 I2: pende do ITEM, nunca da encarnacao. Um Move nao pode
            ' apagar o trabalho do usuario.
            t.Add(New SchemaTable("user_state", {
                Col("user_state_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("item_key", "INTEGER", obrigatoria:=True, refs:="item"),
                Col("triaged", "INTEGER"),
                Col("ai_summary_ref", "TEXT")
            }, {Unico("item_key")}))

            ' §11.2: nenhuma propriedade e unica E estavel, entao toda uniao e
            ' ASSERCAO. Por isso: procedencia, confianca, instante,
            ' coexistencia confirmada (I8) e retratacao (I3).
            t.Add(New SchemaTable("correlation_edge", {
                Col("edge_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("from_item_key", "INTEGER", obrigatoria:=True, refs:="item"),
                Col("to_item_key", "INTEGER", obrigatoria:=True, refs:="item"),
                Col("kind", "TEXT", obrigatoria:=True),
                Col("evidence", "TEXT", obrigatoria:=True),
                Col("confidence", "REAL", obrigatoria:=True),
                Col("observed_at", "TEXT", obrigatoria:=True),
                Col("generation_key", "INTEGER", refs:="generation"),
                Col("coexistence_checked", "INTEGER", obrigatoria:=True),
                Col("retracted_at", "TEXT")
            }))

            ' §17 S4: geracao e POR PASTA — nao existe geracao da caixa.
            ' §14 I7: o universo, incluindo o cutoff, faz parte da identidade.
            ' §17 S6: as contagens que precisam concordar.
            t.Add(New SchemaTable("generation", {
                Col("generation_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("folder_key", "INTEGER", obrigatoria:=True, refs:="folder"),
                Col("coverage_kind", "TEXT", obrigatoria:=True),
                Col("universe_fingerprint", "TEXT", obrigatoria:=True),
                Col("retention_cutoff", "TEXT"),
                Col("rows_read", "INTEGER", obrigatoria:=True),
                Col("count_before", "INTEGER", obrigatoria:=True),
                Col("count_after", "INTEGER", obrigatoria:=True),
                Col("distinct_keys", "INTEGER", obrigatoria:=True),
                Col("is_valid", "INTEGER", obrigatoria:=True),
                Col("started_at", "TEXT", obrigatoria:=True),
                Col("committed_at", "TEXT")
            }))

            Return New CacheSchema(t)
        End Function
    End Class

End Namespace

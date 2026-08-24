Imports System.Collections.Generic
Imports System.Linq

Namespace Global.Iris.Core

    Public Enum DeleteAction
        NoAction
        Restrict
        Cascade
        SetNull
    End Enum

    Public NotInheritable Class SchemaColumn
        Public ReadOnly Property Name As String
        Public ReadOnly Property Kind As String
        Public ReadOnly Property IsPrimaryKey As Boolean
        Public ReadOnly Property IsRequired As Boolean
        ''' <summary>Tabela referenciada, ou Nothing.</summary>
        Public ReadOnly Property References As String
        ''' <summary>
        ''' Valor vindo do PROVIDER (Outlook), não do Iris. É o que o I1 usa
        ''' para proibir que ele seja chave primária.
        ''' </summary>
        Public ReadOnly Property IsProviderValue As Boolean
        ''' <summary>Restrição de valor, ou Nothing.</summary>
        Public ReadOnly Property Check As String
        Public ReadOnly Property OnDelete As DeleteAction

        Public Sub New(name As String, kind As String,
                       Optional isPrimaryKey As Boolean = False,
                       Optional isRequired As Boolean = False,
                       Optional references As String = Nothing,
                       Optional isProviderValue As Boolean = False,
                       Optional check As String = Nothing,
                       Optional onDelete As DeleteAction = DeleteAction.NoAction)
            Me.Name = name
            Me.Kind = kind
            Me.IsPrimaryKey = isPrimaryKey
            Me.IsRequired = isRequired
            Me.References = references
            Me.IsProviderValue = isProviderValue
            Me.Check = check
            Me.OnDelete = onDelete
        End Sub
    End Class

    Public NotInheritable Class SchemaTable
        Public ReadOnly Property Name As String
        Public ReadOnly Property Columns As IReadOnlyList(Of SchemaColumn)
        Public ReadOnly Property UniqueIndexes As IReadOnlyList(Of String())
        ''' <summary>Índices não únicos — para consulta, não para invariante.</summary>
        Public ReadOnly Property Indexes As IReadOnlyList(Of String())
        ''' <summary>Restrições de tabela, além das de coluna.</summary>
        Public ReadOnly Property Checks As IReadOnlyList(Of String)

        Public Sub New(name As String, columns As IEnumerable(Of SchemaColumn),
                       Optional uniqueIndexes As IEnumerable(Of String()) = Nothing,
                       Optional indexes As IEnumerable(Of String()) = Nothing,
                       Optional checks As IEnumerable(Of String) = Nothing)
            Me.Name = name
            Me.Columns = columns.ToList()
            Me.UniqueIndexes = If(uniqueIndexes, Enumerable.Empty(Of String())()).ToList()
            Me.Indexes = If(indexes, Enumerable.Empty(Of String())()).ToList()
            Me.Checks = If(checks, Enumerable.Empty(Of String)()).ToList()
        End Sub

        Public Function Column(nome As String) As SchemaColumn
            Return Columns.FirstOrDefault(
                Function(c) String.Equals(c.Name, nome, StringComparison.OrdinalIgnoreCase))
        End Function

        Public Function TemUnico(ParamArray colunas As String()) As Boolean
            Return UniqueIndexes.Any(
                Function(ix) ix.Length = colunas.Length AndAlso
                             ix.SequenceEqual(colunas, StringComparer.OrdinalIgnoreCase))
        End Function
    End Class

    ''' <summary>
    ''' O schema do cache, descrito como DADO.
    '''
    ''' Não é string de DDL: é um modelo que o gerador de DDL e os testes de
    ''' invariante leem. Se fosse DDL solta, os testes teriam de reparsear e
    ''' os dois lados poderiam divergir — o erro que a Q1 cobrou quando o
    ''' teste sintético verificava um algoritmo diferente do que rodava
    ''' contra o Outlook.
    '''
    ''' As proibições estão em FASE2.md §14 (I1–I8) e §17 (S1–S7). Cada uma
    ''' tem teste que FALHA quando é violada — sem isso são comentário, e
    ''' este projeto já viu a regra do <c>Permission</c> ficar escrita por
    ''' três marcos enquanto o gate falhava aberto.
    ''' </summary>
    Public NotInheritable Class CacheSchema
        Public ReadOnly Property Tables As IReadOnlyList(Of SchemaTable)

        Public Sub New(tables As IEnumerable(Of SchemaTable))
            Me.Tables = tables.ToList()
        End Sub

        Public Function Table(nome As String) As SchemaTable
            Return Tables.FirstOrDefault(
                Function(t) String.Equals(t.Name, nome, StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>Um índice único sobre as colunas dadas.</summary>
        Private Shared Function Unico(ParamArray colunas As String()) As String()
            Return colunas
        End Function

        Private Shared Function Col(nome As String, tipo As String,
                                    Optional pk As Boolean = False,
                                    Optional obrigatoria As Boolean = False,
                                    Optional refs As String = Nothing,
                                    Optional doProvider As Boolean = False,
                                    Optional check As String = Nothing,
                                    Optional aoExcluir As DeleteAction = DeleteAction.NoAction) As SchemaColumn
            Return New SchemaColumn(nome, tipo, pk, obrigatoria, refs, doProvider, check, aoExcluir)
        End Function

        ''' <summary>
        ''' O schema pretendido pelo 2.1.
        '''
        ''' A forma vem inteira da Fase 2. Cada separação existe porque uma
        ''' medição obrigou, e a referência está junto de cada uma.
        '''
        ''' NOTA DE SINTAXE: montado com List e Add, não com literal
        ''' aninhado. O literal não aceita linha em branco nem linha só de
        ''' comentário no meio — a continuação implícita do VB quebra, e o
        ''' erro aparece dezenas de linhas adiante.
        ''' </summary>
        Public Shared Function Intended() As CacheSchema
            Dim t As New List(Of SchemaTable)()

            ' AMBIENTE — a allowlist do D2 é DADO, não constante no código.
            ' A §19.3 mediu que abrir a janela de cache custa horas e dezenas
            ' de GB, então a matriz de providers não é levantável a custo
            ' baixo e "ambiente não medido" tem de RECUSAR operar.
            t.Add(New SchemaTable("environment_profile", {
                Col("environment_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("fingerprint", "TEXT", obrigatoria:=True),
                Col("provider", "TEXT", obrigatoria:=True),
                Col("cached_mode", "INTEGER", obrigatoria:=True),
                Col("sync_window", "TEXT"),
                Col("policy_version", "INTEGER", obrigatoria:=True),
                Col("allowed", "INTEGER", obrigatoria:=True, check:="allowed IN (0,1)")
            }, {Unico("fingerprint")}))

            t.Add(New SchemaTable("store", {
                Col("store_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("provider_store_id", "TEXT", obrigatoria:=True, doProvider:=True),
                Col("display_name", "TEXT")
            }, {Unico("provider_store_id")}))

            ' published_generation_key + reconcile_epoch são o CAS que impede
            ' geração velha sobrescrever nova.
            t.Add(New SchemaTable("folder", {
                Col("folder_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("store_key", "INTEGER", obrigatoria:=True, refs:="store", aoExcluir:=DeleteAction.Restrict),
                Col("provider_entry_id", "TEXT", obrigatoria:=True, doProvider:=True),
                Col("name", "TEXT"),
                Col("published_generation_key", "INTEGER", refs:="generation"),
                Col("reconcile_epoch", "INTEGER", obrigatoria:=True, check:="reconcile_epoch >= 0"),
                Col("stability", "TEXT", obrigatoria:=True,
                    check:="stability IN ('estavel','instavel')")
            }, {Unico("store_key", "provider_entry_id")}))

            ' COBERTURA versionada, porque MUDA com sessão, janela e
            ' sincronização. A §19.2 mediu pastas cheias reportando ZERO
            ' itens; a conclusão só vale no universo em que foi tirada.
            t.Add(New SchemaTable("coverage_observation", {
                Col("coverage_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("folder_key", "INTEGER", obrigatoria:=True, refs:="folder", aoExcluir:=DeleteAction.Cascade),
                Col("environment_key", "INTEGER", obrigatoria:=True, refs:="environment_profile"),
                Col("universe_fingerprint", "TEXT", obrigatoria:=True),
                Col("coverage", "TEXT", obrigatoria:=True,
                    check:="coverage IN ('completa','parcial','desconhecida')"),
                Col("source", "TEXT", obrigatoria:=True),
                Col("observed_at", "TEXT", obrigatoria:=True),
                Col("superseded_by", "INTEGER", refs:="coverage_observation")
            }, Nothing, {Unico("folder_key", "observed_at")}))

            t.Add(New SchemaTable("item", {
                Col("item_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("created_at", "TEXT", obrigatoria:=True)
            }))

            ' Idempotência: o localizador é único DENTRO da pasta.
            t.Add(New SchemaTable("incarnation", {
                Col("incarnation_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("item_key", "INTEGER", obrigatoria:=True, refs:="item", aoExcluir:=DeleteAction.Restrict),
                Col("folder_key", "INTEGER", obrigatoria:=True, refs:="folder", aoExcluir:=DeleteAction.Restrict),
                Col("provider_entry_id", "TEXT", obrigatoria:=True, doProvider:=True),
                Col("provider_record_key", "TEXT", doProvider:=True),
                Col("search_key", "TEXT", doProvider:=True),
                Col("internet_message_id", "TEXT", doProvider:=True),
                Col("first_seen_generation", "INTEGER", refs:="generation"),
                Col("last_seen_generation", "INTEGER", refs:="generation")
            }, {Unico("folder_key", "provider_entry_id")},
               {Unico("search_key"), Unico("internet_message_id")}))

            ' METADADO — o que o 2.1 promete guardar, e faltava no modelo.
            ' Medido (§18.3): 349 bytes por mensagem, 11,7 MB para a caixa
            ' inteira. SEM corpo e SEM anexo (D1).
            t.Add(New SchemaTable("metadata_observation", {
                Col("metadata_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("incarnation_key", "INTEGER", obrigatoria:=True, refs:="incarnation", aoExcluir:=DeleteAction.Cascade),
                Col("generation_key", "INTEGER", refs:="generation"),
                Col("subject", "TEXT"),
                Col("sender_name", "TEXT"),
                Col("received_at", "TEXT"),
                Col("last_modified_at", "TEXT"),
                Col("size_bytes", "INTEGER"),
                Col("has_attachments", "INTEGER", check:="has_attachments IN (0,1)"),
                Col("is_unread", "INTEGER", check:="is_unread IN (0,1)"),
                Col("message_class", "TEXT")
            }, Nothing, {Unico("incarnation_key")}))

            t.Add(New SchemaTable("association", {
                Col("association_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("item_key", "INTEGER", obrigatoria:=True, refs:="item", aoExcluir:=DeleteAction.Restrict),
                Col("folder_key", "INTEGER", obrigatoria:=True, refs:="folder", aoExcluir:=DeleteAction.Restrict),
                Col("presence", "TEXT", obrigatoria:=True,
                    check:="presence IN ('nao_verificado','presente','suspeito','ausente_da_pasta')"),
                Col("observability", "TEXT", obrigatoria:=True,
                    check:="observability IN ('observavel','nao_observavel_no_universo','desconhecida')"),
                Col("concluded_in_universe", "TEXT"),
                Col("version", "INTEGER", obrigatoria:=True, check:="version >= 0"),
                Col("generation_key", "INTEGER", refs:="generation"),
                Col("absent_by_generation", "INTEGER", refs:="generation"),
                Col("absent_by_coverage", "INTEGER", refs:="coverage_observation")
            }, {Unico("item_key", "folder_key")}))

            ' §14 I2: pende do ITEM, nunca da encarnação.
            t.Add(New SchemaTable("user_state", {
                Col("user_state_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("item_key", "INTEGER", obrigatoria:=True, refs:="item", aoExcluir:=DeleteAction.Restrict),
                Col("triaged", "INTEGER"),
                Col("ai_summary_ref", "TEXT")
            }, {Unico("item_key")}))

            ' EVIDÊNCIA DE COEXISTÊNCIA — I8, e não é booleano. Um booleano
            ' afirma "checado" sem dizer onde nem quando, e a §16.1 mediu que
            ' duas linhas no cache NÃO provam coexistência.
            t.Add(New SchemaTable("coexistence_evidence", {
                Col("evidence_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("left_incarnation", "INTEGER", obrigatoria:=True, refs:="incarnation"),
                Col("right_incarnation", "INTEGER", obrigatoria:=True, refs:="incarnation"),
                Col("generation_key", "INTEGER", obrigatoria:=True, refs:="generation"),
                Col("universe_fingerprint", "TEXT", obrigatoria:=True),
                Col("observed_at", "TEXT", obrigatoria:=True)
            }))

            ' §11.2: nenhuma propriedade é única E estável, então toda união
            ' é ASSERÇÃO — com procedência, e retratável.
            t.Add(New SchemaTable("correlation_edge", {
                Col("edge_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("from_item_key", "INTEGER", obrigatoria:=True, refs:="item", aoExcluir:=DeleteAction.Restrict),
                Col("to_item_key", "INTEGER", obrigatoria:=True, refs:="item", aoExcluir:=DeleteAction.Restrict),
                Col("kind", "TEXT", obrigatoria:=True,
                    check:="kind IN ('mesmo_item','copia_de','variante_de_conflito_de','duplicata_suspeita')"),
                Col("evidence", "TEXT", obrigatoria:=True),
                Col("confidence", "REAL", obrigatoria:=True,
                    check:="confidence >= 0 AND confidence <= 1"),
                Col("observed_at", "TEXT", obrigatoria:=True),
                Col("generation_key", "INTEGER", refs:="generation"),
                Col("coexistence_evidence_key", "INTEGER", refs:="coexistence_evidence"),
                Col("retracted_at", "TEXT")
            }, Nothing, {Unico("from_item_key"), Unico("to_item_key")},
               {"from_item_key <> to_item_key"}))

            ' TENTATIVA — mutável, e é ONDE o checkpoint vive.
            t.Add(New SchemaTable("scan_attempt", {
                Col("attempt_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("folder_key", "INTEGER", obrigatoria:=True, refs:="folder", aoExcluir:=DeleteAction.Cascade),
                Col("environment_key", "INTEGER", obrigatoria:=True, refs:="environment_profile"),
                Col("universe_fingerprint", "TEXT", obrigatoria:=True),
                Col("retention_cutoff", "TEXT"),
                Col("algorithm_version", "INTEGER", obrigatoria:=True),
                Col("reconcile_epoch", "INTEGER", obrigatoria:=True),
                Col("attempt_number", "INTEGER", obrigatoria:=True, check:="attempt_number >= 1"),
                Col("stage", "TEXT", obrigatoria:=True,
                    check:="stage IN ('aberta','contagem_inicial','varrendo','contagem_final','publicada','descartada')"),
                Col("count_before", "INTEGER"),
                Col("count_after", "INTEGER"),
                Col("rows_read", "INTEGER", obrigatoria:=True, check:="rows_read >= 0"),
                Col("cursor", "TEXT"),
                Col("started_at", "TEXT", obrigatoria:=True),
                Col("ended_at", "TEXT"),
                Col("rejection", "TEXT")
            }, Nothing, {Unico("folder_key", "stage")}))

            ' As linhas ainda NÃO publicadas. É aqui que vive o conjunto de
            ' chaves vistas — nunca serializado dentro do cursor.
            t.Add(New SchemaTable("scan_stage", {
                Col("stage_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("attempt_key", "INTEGER", obrigatoria:=True, refs:="scan_attempt", aoExcluir:=DeleteAction.Cascade),
                Col("provider_entry_id", "TEXT", obrigatoria:=True, doProvider:=True),
                Col("page_number", "INTEGER", obrigatoria:=True, check:="page_number >= 1"),
                Col("cursor_after", "TEXT")
            }, {Unico("attempt_key", "provider_entry_id")}))

            ' O resultado PUBLICADO. Imutável.
            t.Add(New SchemaTable("generation", {
                Col("generation_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("folder_key", "INTEGER", obrigatoria:=True, refs:="folder", aoExcluir:=DeleteAction.Restrict),
                Col("attempt_key", "INTEGER", obrigatoria:=True, refs:="scan_attempt"),
                Col("coverage_kind", "TEXT", obrigatoria:=True,
                    check:="coverage_kind IN ('completa','incremental')"),
                Col("coverage_key", "INTEGER", refs:="coverage_observation"),
                Col("universe_fingerprint", "TEXT", obrigatoria:=True),
                Col("retention_cutoff", "TEXT"),
                Col("rows_read", "INTEGER", obrigatoria:=True),
                Col("count_before", "INTEGER", obrigatoria:=True),
                Col("count_after", "INTEGER", obrigatoria:=True),
                Col("distinct_keys", "INTEGER", obrigatoria:=True),
                Col("reconcile_epoch", "INTEGER", obrigatoria:=True),
                Col("published_at", "TEXT", obrigatoria:=True)
            }, {Unico("attempt_key")}))

            ' O que precisa ser reprocessado depois de crash. Entra no MESMO
            ' commit da publicação: se saísse antes, um crash entre os dois
            ' reprocessaria uma publicação que nunca houve.
            t.Add(New SchemaTable("publication_log", {
                Col("log_key", "INTEGER", pk:=True, obrigatoria:=True),
                Col("generation_key", "INTEGER", obrigatoria:=True, refs:="generation", aoExcluir:=DeleteAction.Cascade),
                Col("emitted_at", "TEXT", obrigatoria:=True),
                Col("drained_at", "TEXT")
            }, {Unico("generation_key")}))

            Return New CacheSchema(t)
        End Function

    End Class

End Namespace

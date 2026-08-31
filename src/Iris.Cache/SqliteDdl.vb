Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Iris.Core

Namespace Global.Iris.Cache

    ''' <summary>
    ''' Gera o DDL a partir do <see cref="CacheSchema"/>.
    '''
    ''' O DDL não é escrito à mão em lugar nenhum. Se fosse, o modelo que os
    ''' testes de invariante leem e o banco que o programa abre seriam duas
    ''' coisas que podem divergir — e a Q1 já cobrou esse desenho quando o
    ''' teste sintético verificava um algoritmo diferente do que rodava
    ''' contra o Outlook.
    ''' </summary>
    Public NotInheritable Class SqliteDdl

        ''' <summary>
        ''' Versão do schema gravada em <c>PRAGMA user_version</c>. Abrir um
        ''' banco de versão diferente FALHA FECHADO.
        ''' </summary>
        ' 2: disclosure_log ganhou http_status.
        ' 3: generation ganhou discarded.
        ' 4: metadata_observation e scan_stage ganharam conversation_id,
        '    conversation_index e sender_address -- a Fase 3 pergunta 'quem
        '    falou por ultimo nesta conversa', e ate aqui o cache nao sabia
        '    responder nem o que e uma conversa nem quem e um remetente.
        ' 5: label_observation -- o rotulo da IA, por encarnacao e geracao.
        Public Const SchemaVersion As Integer = 5

        ''' <summary>
        ''' <b>Os passos de migração conhecidos, e só eles.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE ISTO NÃO CONTRADIZ "NÃO MIGRAR"</b>
        '''
        ''' <see cref="CacheDatabase.Open"/> recusava toda versão divergente,
        ''' com a razão escrita: <i>migrar sem saber de onde para onde é pior
        ''' que recusar</i>. A razão continua valendo — e ela fala de migração
        ''' <b>cega</b>. Um passo listado aqui sabe exatamente de onde para
        ''' onde, e o que não está listado <b>continua falhando fechado</b>.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE O CACHE PASSOU A MERECER MIGRAÇÃO</b>
        '''
        ''' Enquanto o arquivo guardava só metadado do Outlook, apagá-lo não
        ''' custava nada: ele se reconstrói. Depois que o <b>diário do egress</b>
        ''' passou a morar dentro dele, apagar virou destruir o registro do que
        ''' saiu desta máquina — que não se reconstrói de lugar nenhum.
        '''
        ''' E a saída que eu tinha indicado não existia: mandar rodar o harness
        ''' para ler o diário velho <b>antes</b> de apagar não funciona, porque
        ''' o harness também abre pelo <c>CacheDatabase.Open</c> da versão nova
        ''' e leva a mesma recusa. A instrução de preservação era circular.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O QUE UM PASSO PODE FAZER</b>
        '''
        ''' Só o que for <b>aditivo</b>: acrescentar coluna nula, acrescentar
        ''' índice, acrescentar tabela. Nenhum passo aqui pode apagar, renomear
        ''' ou reinterpretar dado já gravado — para isso a recusa continua sendo
        ''' a resposta certa, porque aí sim ninguém sabe de onde para onde.
        '''
        ''' E o resultado <b>não é aceito na confiança</b>: depois de migrar, o
        ''' <see cref="SchemaIntrospector"/> compara o arquivo real com o
        ''' modelo, como em qualquer abertura. Migração que produza a forma
        ''' errada é pega ali.
        ''' </summary>
        Public Shared ReadOnly Property Migracoes _
                                 As IReadOnlyDictionary(Of Integer, IReadOnlyList(Of String))
            Get
                Return _migracoes
            End Get
        End Property

        ' 1 -> 2 e 2 -> 3: aditivas e nulas, entao nenhuma linha ja gravada
        ' muda de sentido. O CHECK acompanha a coluna, para o banco migrado
        ' ficar com a mesma guarda do banco criado do zero.
        '
        ' E NULO NAO E ZERO. Numa geracao anterior a coluna, nulo quer dizer
        ' "esta varredura nao contou o que descartou". Zero seria a afirmacao
        ' de que nada foi descartado, e nenhuma geracao velha tem como
        ' sustenta-la -- inclusive a que varreu a Caixa de Entrada e jogou
        ' fora a contagem de 12.
        '
        ' (Comentario AQUI, e nao dentro do inicializador: em VB a continuacao
        ' implicita de { } nao aceita uma linha so de comentario, e o erro sai
        ' na linha ANTERIOR.)
        ' FECHADA DE VERDADE, e nao so no tipo declarado.
        '
        ' Guardar um Dictionary numa variavel tipada IReadOnlyDictionary nao
        ' fecha nada: DirectCast de volta para Dictionary devolve a colecao
        ' VIVA, e um Clear() ali desliga a migracao inteira. Array exposto como
        ' IReadOnlyList tem o mesmo buraco.
        '
        ' Este projeto ja levou exatamente este golpe uma vez, em
        ' ActivationRecord.Congelar, que devolvia ToList() tipado como
        ' IReadOnlyList -- e um TryCast reabria a lista de operacoes
        ' autorizadas. ReadOnlyDictionary e Array.AsReadOnly nao tem volta.
        ' A MIGRACAO 4 CRIA UMA TABELA, e repete a DDL em vez de chamar o
        ' gerador. O gerador produz a forma de HOJE: uma migracao que o
        ' chamasse passaria a criar outra coisa no dia em que a tabela
        ' mudasse, e bancos migrados em datas diferentes ficariam diferentes
        ' entre si.
        '
        ' A MIGRACAO 3 NAO CRIA INDICE, e e de proposito. O CREATE INDEX de
        ' conversation_id vale para banco NOVO; num banco migrado a coluna
        ' nasce toda nula, e indice sobre coluna nula so ocupa espaco ate a
        ' primeira varredura -- que quem migra vai fazer de qualquer jeito,
        ' porque ate ela a conversa e desconhecida.
        Private Shared ReadOnly _migracoes As IReadOnlyDictionary(Of Integer, IReadOnlyList(Of String)) =
            New ObjectModel.ReadOnlyDictionary(Of Integer, IReadOnlyList(Of String))(
                New Dictionary(Of Integer, IReadOnlyList(Of String)) From {
                    {1, Array.AsReadOnly(New String() {
                        "ALTER TABLE disclosure_log ADD COLUMN http_status INTEGER " &
                        "CHECK (http_status IS NULL OR " &
                        "(http_status >= 100 AND http_status <= 599))"
                    })},
                    {2, Array.AsReadOnly(New String() {
                        "ALTER TABLE generation ADD COLUMN discarded INTEGER " &
                        "CHECK (discarded IS NULL OR discarded >= 0)"
                    })},
                    {4, Array.AsReadOnly(New String() {
                        "CREATE TABLE label_observation (" &
                        "  label_key INTEGER PRIMARY KEY, " &
                        "  incarnation_key INTEGER NOT NULL REFERENCES incarnation(incarnation_key) ON DELETE CASCADE, " &
                        "  generation_key INTEGER NOT NULL REFERENCES generation(generation_key), " &
                        "  label TEXT NOT NULL CHECK (label IN ('precisa_de_mim','aguardando','fyi','notificacao','promocao','newsletter')), " &
                        "  confidence REAL NOT NULL CHECK (confidence >= 0 AND confidence <= 1), " &
                        "  activation_id TEXT NOT NULL, " &
                        "  observed_at TEXT NOT NULL)",
                        "CREATE UNIQUE INDEX ux_label_observation_1 ON label_observation " &
                        "  (incarnation_key, generation_key)"
                    })},
                    {3, Array.AsReadOnly(New String() {
                        "ALTER TABLE metadata_observation ADD COLUMN conversation_id TEXT",
                        "ALTER TABLE metadata_observation ADD COLUMN conversation_index TEXT",
                        "ALTER TABLE metadata_observation ADD COLUMN sender_address TEXT",
                        "ALTER TABLE scan_stage ADD COLUMN conversation_id TEXT",
                        "ALTER TABLE scan_stage ADD COLUMN conversation_index TEXT",
                        "ALTER TABLE scan_stage ADD COLUMN sender_address TEXT"
                    })}
                })

        Public Shared Function Generate(schema As CacheSchema) As IReadOnlyList(Of String)
            Dim comandos As New List(Of String)()

            ' FK precisa ser LIGADA por conexão — no SQLite ela é opcional e
            ' vem desligada. Um schema com FK declarada e desligada dá a
            ' impressão de integridade sem tê-la.
            comandos.Add("PRAGMA foreign_keys = ON")
            ' WAL: crash no meio de uma transação não corrompe o banco.
            comandos.Add("PRAGMA journal_mode = WAL")

            For Each t In OrdenarPorDependencia(schema)
                comandos.Add(CriarTabela(t, schema))
            Next

            For Each t In schema.Tables
                Dim n = 0
                For Each ix In t.UniqueIndexes
                    n += 1
                    comandos.Add($"CREATE UNIQUE INDEX ux_{t.Name}_{n} ON {t.Name} ({String.Join(", ", ix)})")
                Next
                n = 0
                For Each ix In t.Indexes
                    n += 1
                    comandos.Add($"CREATE INDEX ix_{t.Name}_{n} ON {t.Name} ({String.Join(", ", ix)})")
                Next
            Next

            comandos.Add($"PRAGMA user_version = {SchemaVersion}")
            Return comandos
        End Function

        ''' <summary>
        ''' Tabelas na ordem em que dá para criar sem FK apontando para o
        ''' vazio. Há ciclo real no modelo — <c>folder</c> aponta para
        ''' <c>generation</c> e vice-versa —, e nesse caso a ordem não
        ''' resolve: o SQLite aceita FK para tabela ainda inexistente desde
        ''' que ela exista na hora do INSERT.
        ''' </summary>
        Private Shared Function OrdenarPorDependencia(schema As CacheSchema) As IReadOnlyList(Of SchemaTable)
            Dim restantes = schema.Tables.ToList()
            Dim ordenadas As New List(Of SchemaTable)()
            Dim jaCriadas As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            While restantes.Count > 0
                Dim prontas = restantes.Where(
                    Function(t) t.Columns.All(
                        Function(c) c.References Is Nothing OrElse
                                    jaCriadas.Contains(c.References) OrElse
                                    String.Equals(c.References, t.Name, StringComparison.OrdinalIgnoreCase))).ToList()

                If prontas.Count = 0 Then
                    ' Ciclo. Cria o resto na ordem em que veio: o SQLite nao
                    ' valida o alvo da FK no CREATE.
                    ordenadas.AddRange(restantes)
                    Exit While
                End If

                For Each t In prontas
                    ordenadas.Add(t)
                    jaCriadas.Add(t.Name)
                    restantes.Remove(t)
                Next
            End While

            Return ordenadas
        End Function

        Private Shared Function CriarTabela(t As SchemaTable, schema As CacheSchema) As String
            Dim sb As New StringBuilder()
            sb.Append($"CREATE TABLE {t.Name} (")
            Dim partes As New List(Of String)()

            For Each c In t.Columns
                Dim p As New StringBuilder()
                p.Append($"{c.Name} {c.Kind}")
                If c.IsPrimaryKey Then p.Append(" PRIMARY KEY")
                If c.IsRequired AndAlso Not c.IsPrimaryKey Then p.Append(" NOT NULL")
                If c.Check IsNot Nothing Then p.Append($" CHECK ({c.Check})")
                partes.Add(p.ToString())
            Next

            For Each c In t.Columns
                If c.References Is Nothing Then Continue For
                Dim acao = AcaoSql(c.OnDelete)
                Dim alvo = ChavePrimariaDe(schema, c.References)
                partes.Add($"FOREIGN KEY ({c.Name}) REFERENCES {c.References} ({alvo}){acao}")
            Next

            For Each chk In t.Checks
                partes.Add($"CHECK ({chk})")
            Next

            sb.Append(String.Join(", ", partes))
            sb.Append(")")
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' A chave primária REAL da tabela referenciada, lida do modelo.
        '''
        ''' Já foi por convenção — "a PK de <c>x</c> chama-se <c>x_key</c>" —,
        ''' e a convenção estava errada em <c>environment_profile</c>, cuja PK
        ''' é <c>environment_key</c>. O SQLite aceita o CREATE TABLE apontando
        ''' para coluna inexistente e só reclama no primeiro INSERT, com
        ''' "foreign key mismatch", muito longe de onde o erro foi escrito.
        ''' </summary>
        Friend Shared Function ChavePrimariaDe(schema As CacheSchema, tabela As String) As String
            Dim alvo = schema.Tables.FirstOrDefault(
                Function(x) String.Equals(x.Name, tabela, StringComparison.OrdinalIgnoreCase))
            If alvo Is Nothing Then
                Throw New InvalidOperationException($"FK aponta para tabela inexistente: {tabela}")
            End If
            Dim pk = alvo.Columns.FirstOrDefault(Function(c) c.IsPrimaryKey)
            If pk Is Nothing Then
                Throw New InvalidOperationException($"tabela {tabela} nao tem chave primaria")
            End If
            Return pk.Name
        End Function

        Private Shared Function AcaoSql(a As DeleteAction) As String
            Select Case a
                Case DeleteAction.Cascade : Return " ON DELETE CASCADE"
                Case DeleteAction.Restrict : Return " ON DELETE RESTRICT"
                Case DeleteAction.SetNull : Return " ON DELETE SET NULL"
                Case Else : Return ""
            End Select
        End Function

    End Class

End Namespace

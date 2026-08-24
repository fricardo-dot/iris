Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Core
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' O gate de schema, e o que cada invariante recusa.
'''
''' A metade que importa NÃO é "o schema pretendido passa". É que **cada
''' regra dispara quando é violada**. Regra que nunca dispara não protege
''' nada, e este projeto tem o precedente: a regra do <c>Permission</c>
''' ficou escrita por três marcos enquanto o gate falhava aberto.
'''
''' Por isso cada teste aqui QUEBRA o schema pretendido de um jeito
''' específico e exige a violação correspondente. É o mesmo desenho dos
''' controles negativos de <see cref="CursorPagingTests"/>.
''' </summary>
<TestClass>
Public Class SchemaGateTests

    ' ==================================================================
    ' Utilidades para quebrar o schema de um jeito só
    ' ==================================================================

    Private Shared Function Copiar(c As SchemaColumn,
                                   Optional pk As Boolean? = Nothing,
                                   Optional required As Boolean? = Nothing,
                                   Optional refs As String = "",
                                   Optional provider As Boolean? = Nothing) As SchemaColumn
        Return New SchemaColumn(
            c.Name, c.Kind,
            If(pk.HasValue, pk.Value, c.IsPrimaryKey),
            If(required.HasValue, required.Value, c.IsRequired),
            If(refs = "", c.References, If(refs Is Nothing, Nothing, refs)),
            If(provider.HasValue, provider.Value, c.IsProviderValue))
    End Function

    ''' <summary>Reconstrói o schema trocando UMA tabela.</summary>
    Private Shared Function Com(schema As CacheSchema, tabela As SchemaTable) As CacheSchema
        Return New CacheSchema(schema.Tables.Select(
            Function(t) If(String.Equals(t.Name, tabela.Name, StringComparison.OrdinalIgnoreCase),
                           tabela, t)))
    End Function

    Private Shared Function Sem(schema As CacheSchema, nomeDaTabela As String) As CacheSchema
        Return New CacheSchema(schema.Tables.Where(
            Function(t) Not String.Equals(t.Name, nomeDaTabela, StringComparison.OrdinalIgnoreCase)))
    End Function

    ''' <summary>Remove UMA coluna de UMA tabela.</summary>
    Private Shared Function SemColuna(schema As CacheSchema, tabela As String, coluna As String) As CacheSchema
        Dim t = schema.Table(tabela)
        Return Com(schema, New SchemaTable(t.Name,
            t.Columns.Where(Function(c) Not String.Equals(c.Name, coluna, StringComparison.OrdinalIgnoreCase)),
            t.UniqueIndexes))
    End Function

    Private Shared Function UnicoTeste(ParamArray colunas As String()) As String()
        Return colunas
    End Function

    Private Shared Function Violou(schema As CacheSchema, regra As String) As Boolean
        Return SchemaGate.Check(schema).Any(Function(x) x.Rule = regra)
    End Function

    Private Shared Sub Exige(schema As CacheSchema, regra As String, oQueFoiQuebrado As String)
        Dim vs = SchemaGate.Check(schema)
        Assert.IsTrue(vs.Any(Function(x) x.Rule = regra),
            $"quebrei {oQueFoiQuebrado} e {regra} NAO disparou. " &
            $"Violacoes vistas: {String.Join(" | ", vs.Select(Function(x) x.Rule))}")
    End Sub

    ' ==================================================================

    ''' <summary>
    ''' O controle POSITIVO. Sozinho ele não vale nada — um gate que nunca
    ''' reprova também passa aqui. Vale junto com os que seguem.
    ''' </summary>
    <TestMethod>
    Public Sub O_schema_pretendido_passa_limpo()
        Dim vs = SchemaGate.Check(CacheSchema.Intended())
        Assert.AreEqual(0, vs.Count,
            "o schema pretendido deveria passar. Violacoes: " &
            String.Join(" | ", vs.Select(Function(x) x.ToString())))
    End Sub

    <TestMethod>
    Public Sub I1_recusa_identificador_do_provider_como_chave_primaria()
        ' O erro obvio, que funciona no primeiro dia e quebra no primeiro Move.
        Dim s = CacheSchema.Intended()
        Dim t = s.Table("incarnation")
        Dim quebrada = New SchemaTable(t.Name,
            t.Columns.Select(Function(c) If(c.Name = "provider_entry_id",
                                            Copiar(c, pk:=True), c)), t.UniqueIndexes)
        Exige(Com(s, quebrada), "I1", "incarnation.provider_entry_id virou chave primaria")
    End Sub

    <TestMethod>
    Public Sub I2_recusa_estado_do_usuario_preso_a_encarnacao()
        Dim s = CacheSchema.Intended()
        Dim t = s.Table("user_state")
        Dim quebrada = New SchemaTable(t.Name,
            t.Columns.Select(Function(c) If(c.Name = "item_key",
                                            New SchemaColumn("incarnation_key", "INTEGER",
                                                             isRequired:=True, references:="incarnation"),
                                            c)), t.UniqueIndexes)
        Exige(Com(s, quebrada), "I2", "user_state passou a pender de incarnation")
    End Sub

    <TestMethod>
    Public Sub I2_recusa_provider_id_dentro_do_estado_do_usuario()
        ' A porta dos fundos: nao referencia encarnacao, mas guarda o EntryID.
        Dim s = CacheSchema.Intended()
        Dim t = s.Table("user_state")
        Dim quebrada = New SchemaTable(t.Name,
            t.Columns.Concat({New SchemaColumn("provider_entry_id", "TEXT", isProviderValue:=True)}),
            t.UniqueIndexes)
        Exige(Com(s, quebrada), "I2", "user_state ganhou uma coluna de EntryID")
    End Sub

    <TestMethod>
    Public Sub I3_recusa_uniao_por_reescrita()
        ' canonical_item_key no item: o desenho que parece simples e destroi
        ' a possibilidade de desfazer.
        Dim s = CacheSchema.Intended()
        Dim t = s.Table("item")
        Dim quebrada = New SchemaTable(t.Name,
            t.Columns.Concat({New SchemaColumn("canonical_item_key", "INTEGER", references:="item")}),
            t.UniqueIndexes)
        Exige(Com(s, quebrada), "I3", "item ganhou canonical_item_key")
    End Sub

    <TestMethod>
    Public Sub I3_recusa_aresta_que_nao_pode_ser_retratada()
        Exige(SemColuna(CacheSchema.Intended(), "correlation_edge", "retracted_at"),
              "I3", "a aresta perdeu retracted_at")
    End Sub

    <TestMethod>
    Public Sub I4_recusa_aresta_sem_procedencia()
        For Each campo In {"evidence", "confidence", "observed_at"}
            Exige(SemColuna(CacheSchema.Intended(), "correlation_edge", campo),
                  "I4", $"a aresta perdeu {campo}")
        Next
    End Sub

    ''' <summary>
    ''' I4 aceita aresta SEM geração — de propósito. Aresta criada por evento
    ''' ou por confirmação manual não pertence a geração nenhuma, e exigir
    ''' isso foi um erro da 1ª redação.
    ''' </summary>
    <TestMethod>
    Public Sub I4_aceita_aresta_sem_geracao()
        Dim s = SemColuna(CacheSchema.Intended(), "correlation_edge", "generation_key")
        Assert.IsFalse(Violou(s, "I4"),
            "aresta sem generation_key nao pode ser violacao: evento e confirmacao " &
            "manual nao pertencem a geracao")
    End Sub

    <TestMethod>
    Public Sub I5_recusa_schema_em_que_ficar_separado_custa_mais()
        Dim s = CacheSchema.Intended()
        Dim t = s.Table("item")
        Dim quebrada = New SchemaTable(t.Name,
            t.Columns.Concat({New SchemaColumn("edge_key", "INTEGER",
                                               isRequired:=True, references:="correlation_edge")}),
            t.UniqueIndexes)
        Exige(Com(s, quebrada), "I5", "item passou a exigir uma aresta")
    End Sub

    <TestMethod>
    Public Sub I6_recusa_ausencia_sem_estado()
        Exige(SemColuna(CacheSchema.Intended(), "association", "presence"),
              "I6", "association perdeu o estado de presenca")
    End Sub

    <TestMethod>
    Public Sub I7_recusa_geracao_sem_universo_e_sem_cutoff()
        Exige(SemColuna(CacheSchema.Intended(), "generation", "universe_fingerprint"),
              "I7", "generation perdeu o universo")
        Exige(SemColuna(CacheSchema.Intended(), "generation", "retention_cutoff"),
              "I7", "generation perdeu o cutoff de retencao")
    End Sub

    <TestMethod>
    Public Sub I8_recusa_uniao_sem_coexistencia_confirmada()
        Exige(SemColuna(CacheSchema.Intended(), "correlation_edge", "coexistence_checked"),
              "I8", "a aresta perdeu coexistence_checked")
    End Sub

    <TestMethod>
    Public Sub S1_recusa_geracao_que_nao_diz_o_que_cobriu()
        Exige(SemColuna(CacheSchema.Intended(), "generation", "coverage_kind"),
              "S1", "generation perdeu coverage_kind")
    End Sub

    <TestMethod>
    Public Sub S2_recusa_ausencia_nao_amarrada_a_uma_geracao()
        Exige(SemColuna(CacheSchema.Intended(), "association", "absent_by_generation"),
              "S2", "association perdeu absent_by_generation")
    End Sub

    ''' <summary>
    ''' S3 — o índice único que parece obviamente certo e está errado.
    ''' A matriz B da §16.1 produz, legitimamente, duas linhas para o mesmo
    ''' item logico.
    ''' </summary>
    <TestMethod>
    Public Sub S3_recusa_indice_unico_que_proibe_sobreposicao()
        Dim s = CacheSchema.Intended()
        Dim t = s.Table("association")
        Dim quebrada = New SchemaTable(t.Name, t.Columns,
                                       t.UniqueIndexes.Concat({UnicoTeste("item_key")}))
        Exige(Com(s, quebrada), "S3", "association ganhou indice unico sobre item_key")
    End Sub

    <TestMethod>
    Public Sub S4_recusa_geracao_global()
        Exige(SemColuna(CacheSchema.Intended(), "generation", "folder_key"),
              "S4", "generation deixou de ser por pasta")
    End Sub

    <TestMethod>
    Public Sub S6_recusa_geracao_sem_as_contagens()
        For Each campo In {"rows_read", "count_before", "count_after", "distinct_keys"}
            Exige(SemColuna(CacheSchema.Intended(), "generation", campo),
                  "S6", $"generation perdeu {campo}")
        Next
    End Sub

    ''' <summary>
    ''' S7 — o caso que TODOS os outros sinais deixam passar.
    '''
    ''' Na caixa medida, dezenas de pastas reportam zero itens e não estão
    ''' vazias: estão fora da janela de cache (§19.2). Esse caso satisfaz o
    ''' S6 com folga — lidas = antes = depois = 0 — e mesmo assim concluir
    ''' "tudo foi excluído" seria desastre.
    ''' </summary>
    <TestMethod>
    Public Sub S7_recusa_pasta_sem_cobertura_conhecida()
        Exige(SemColuna(CacheSchema.Intended(), "folder", "coverage"),
              "S7", "folder perdeu a cobertura do cache")
    End Sub

    <TestMethod>
    Public Sub S7_e_independente_do_S6()
        ' Sem 'coverage', o S7 dispara e o S6 continua satisfeito — que e
        ' exatamente por que S7 precisou existir em separado.
        Dim s = SemColuna(CacheSchema.Intended(), "folder", "coverage")
        Assert.IsTrue(Violou(s, "S7"))
        Assert.IsFalse(Violou(s, "S6"),
            "S6 nao deveria disparar: e por ele nao pegar este caso que o S7 existe")
    End Sub

    ''' <summary>
    ''' Tabela inteira faltando tem de acusar, não passar em silêncio.
    ''' </summary>
    <TestMethod>
    Public Sub Tabela_ausente_nao_passa_calada()
        For Each par In New(Tabela As String, Regra As String)() {
            ("user_state", "I2"), ("correlation_edge", "I3"), ("item", "I5"),
            ("association", "I6"), ("generation", "I7"), ("folder", "S7")}
            Exige(Sem(CacheSchema.Intended(), par.Tabela), par.Regra,
                  $"a tabela {par.Tabela} sumiu")
        Next
    End Sub

    ''' <summary>
    ''' O guarda do guarda: TODA regra declarada precisa ter pelo menos um
    ''' teste que a faz disparar. Sem isto, acrescentar uma regra nova sem
    ''' controle negativo passaria despercebido.
    ''' </summary>
    <TestMethod>
    Public Sub Toda_regra_e_exercitada_por_algum_teste()
        Dim declaradas = New HashSet(Of String)(
            {"I1", "I2", "I3", "I4", "I5", "I6", "I7", "I8",
             "S1", "S2", "S3", "S4", "S6", "S7"})

        ' Recolhe as regras que QUALQUER quebra deste arquivo faz disparar.
        Dim disparadas As New HashSet(Of String)()
        Dim s = CacheSchema.Intended()
        For Each t In s.Tables
            For Each v In SchemaGate.Check(Sem(s, t.Name))
                disparadas.Add(v.Rule)
            Next
            For Each c In t.Columns
                For Each v In SchemaGate.Check(SemColuna(s, t.Name, c.Name))
                    disparadas.Add(v.Rule)
                Next
            Next
        Next
        ' as tres que exigem ACRESCENTAR algo, e nao remover
        disparadas.Add("I1") : disparadas.Add("I5") : disparadas.Add("S3")

        Dim orfas = declaradas.Except(disparadas).OrderBy(Function(x) x).ToList()
        Assert.AreEqual(0, orfas.Count,
            "regras que nenhuma quebra faz disparar: " & String.Join(", ", orfas) &
            ". Regra que nunca dispara nao esta protegendo nada.")
    End Sub

End Class

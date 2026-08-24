Imports System.Collections.Generic
Imports System.Linq

Namespace Global.Iris.Core

    Public NotInheritable Class GateViolation
        Public ReadOnly Property Rule As String
        Public ReadOnly Property Detail As String
        Public Sub New(rule As String, detail As String)
            Me.Rule = rule
            Me.Detail = detail
        End Sub
        Public Overrides Function ToString() As String
            Return $"{Rule}: {Detail}"
        End Function
    End Class

    ''' <summary>
    ''' Os invariantes de FASE2.md §14 (I1–I8) e §17 (S1–S7), como código que
    ''' recusa um schema errado.
    '''
    ''' Cada regra existe porque uma medição obrigou, e a referência está no
    ''' comentário — não por burocracia, mas porque daqui a seis meses a
    ''' pergunta vai ser "por que isso é proibido?" e a resposta tem de estar
    ''' junto da proibição.
    '''
    ''' O teste correspondente alimenta um schema DELIBERADAMENTE errado e
    ''' exige que a regra dispare. Regra que nunca dispara não está
    ''' protegendo nada — foi assim que o gate do <c>Permission</c> ficou
    ''' três marcos falhando aberto com a regra escrita ao lado.
    ''' </summary>
    Public NotInheritable Class SchemaGate

        Public Shared Function Check(schema As CacheSchema) As IReadOnlyList(Of GateViolation)
            Dim v As New List(Of GateViolation)()
            If schema Is Nothing Then
                v.Add(New GateViolation("gate", "schema nulo"))
                Return v
            End If

            VerificarI1(schema, v)
            VerificarI2(schema, v)
            VerificarI3eI4eI8(schema, v)
            VerificarI5(schema, v)
            VerificarI6(schema, v)
            VerificarI7(schema, v)
            VerificarS1eS6(schema, v)
            VerificarS2(schema, v)
            VerificarS3(schema, v)
            VerificarS4(schema, v)
            VerificarS7(schema, v)
            Return v
        End Function

        ''' <summary>
        ''' I1 — nenhum identificador do provider é chave primária.
        ''' Medido: §11.1, o EntryID muda num Move.
        ''' </summary>
        Private Shared Sub VerificarI1(s As CacheSchema, v As List(Of GateViolation))
            For Each t In s.Tables
                For Each c In t.Columns
                    If c.IsProviderValue AndAlso c.IsPrimaryKey Then
                        v.Add(New GateViolation("I1",
                            $"{t.Name}.{c.Name} é do provider e é chave primária. " &
                            "O EntryID muda num Move (§11.1)."))
                    End If
                Next
            Next
        End Sub

        ''' <summary>
        ''' I2 — estado do usuário pende do ITEM, nunca da encarnação.
        ''' Senão arrastar uma mensagem entre pastas apaga o trabalho dele.
        ''' </summary>
        Private Shared Sub VerificarI2(s As CacheSchema, v As List(Of GateViolation))
            Dim t = s.Table("user_state")
            If t Is Nothing Then
                v.Add(New GateViolation("I2", "não existe tabela de estado do usuário"))
                Return
            End If
            If t.Columns.Any(Function(c) c.References = "incarnation") Then
                v.Add(New GateViolation("I2",
                    "user_state referencia incarnation. Um Move apagaria o " &
                    "estado, em silêncio (§11.1)."))
            End If
            If Not t.Columns.Any(Function(c) c.References = "item" AndAlso c.IsRequired) Then
                v.Add(New GateViolation("I2", "user_state não pende de item de forma obrigatória"))
            End If
            ' Guardar provider id aqui reintroduz o problema pela porta dos fundos.
            For Each c In t.Columns
                If c.IsProviderValue Then
                    v.Add(New GateViolation("I2",
                        $"user_state.{c.Name} é valor do provider; o estado ficaria " &
                        "preso a uma encarnação."))
                End If
            Next
        End Sub

        ''' <summary>
        ''' I3 — unir é ARESTA, nunca reescrita. I4 — procedência, confiança e
        ''' instante. I8 — coexistência CONFIRMADA.
        ''' Medido: §11.2, nenhuma propriedade é única E estável.
        ''' </summary>
        Private Shared Sub VerificarI3eI4eI8(s As CacheSchema, v As List(Of GateViolation))
            Dim t = s.Table("correlation_edge")
            If t Is Nothing Then
                v.Add(New GateViolation("I3", "não existe tabela de arestas de correlação"))
                Return
            End If

            ' I3: a aresta tem de ser removível.
            If t.Column("retracted_at") Is Nothing Then
                v.Add(New GateViolation("I3",
                    "a aresta não tem como ser retratada. Uma união errada " &
                    "viraria permanente, e a §11.2 garante que haverá uniões erradas."))
            End If

            ' I3: coluna de canonicalização no ITEM é reescrita disfarçada.
            Dim item = s.Table("item")
            If item IsNot Nothing Then
                For Each c In item.Columns
                    If c.Name.IndexOf("canonical", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       c.References = "item" Then
                        v.Add(New GateViolation("I3",
                            $"item.{c.Name} aponta para outro item. Isso é união " &
                            "por reescrita: depois de gravada, não há o que desfazer."))
                    End If
                Next
            End If

            ' I4
            For Each obrigatoria In {"evidence", "confidence", "observed_at"}
                Dim c = t.Column(obrigatoria)
                If c Is Nothing OrElse Not c.IsRequired Then
                    v.Add(New GateViolation("I4",
                        $"correlation_edge.{obrigatoria} ausente ou opcional. Sem " &
                        "procedência não dá para reavaliar quando a regra melhorar (§10.6)."))
                End If
            Next

            ' I8
            Dim coex = t.Column("coexistence_checked")
            If coex Is Nothing OrElse Not coex.IsRequired Then
                v.Add(New GateViolation("I8",
                    "correlation_edge não registra se a coexistência foi CONFIRMADA. " &
                    "Duas linhas no cache não provam coexistência: o corte fraturado " &
                    "da §16.1 produz exatamente isso para um Move legítimo."))
            End If
        End Sub

        ''' <summary>
        ''' I5 — "não unido" é o estado barato e o padrão.
        ''' Violação: item que EXIGE aresta para existir.
        ''' </summary>
        Private Shared Sub VerificarI5(s As CacheSchema, v As List(Of GateViolation))
            Dim item = s.Table("item")
            If item Is Nothing Then
                v.Add(New GateViolation("I5", "não existe tabela de item"))
                Return
            End If
            For Each c In item.Columns
                If c.IsRequired AndAlso c.References = "correlation_edge" Then
                    v.Add(New GateViolation("I5",
                        $"item.{c.Name} exige aresta de correlação. Deixar separado " &
                        "tem de ser o caminho BARATO, senão a pressão empurra para unir."))
                End If
            Next
        End Sub

        ''' <summary>
        ''' I6 — ausência é estado da ASSOCIAÇÃO, com valor transitório.
        ''' Medido: §16.3, ausência na pasta não distingue movido de excluído.
        ''' </summary>
        Private Shared Sub VerificarI6(s As CacheSchema, v As List(Of GateViolation))
            Dim t = s.Table("association")
            If t Is Nothing Then
                v.Add(New GateViolation("I6", "não existe tabela de associação item–pasta"))
                Return
            End If
            Dim p = t.Column("presence")
            If p Is Nothing OrElse Not p.IsRequired Then
                v.Add(New GateViolation("I6",
                    "association não tem estado de presença obrigatório. Sem ele, " &
                    "'ausente' vira apagar a linha — e apagar não tem estado transitório."))
            End If
        End Sub

        ''' <summary>
        ''' I7 — o universo faz parte da identidade da geração.
        ''' Inclui o cutoff de retenção, que a §18 mostrou existir de fato.
        ''' </summary>
        Private Shared Sub VerificarI7(s As CacheSchema, v As List(Of GateViolation))
            Dim g = s.Table("generation")
            If g Is Nothing Then
                v.Add(New GateViolation("I7", "não existe tabela de geração"))
                Return
            End If
            Dim u = g.Column("universe_fingerprint")
            If u Is Nothing OrElse Not u.IsRequired Then
                v.Add(New GateViolation("I7",
                    "generation não carrega o universo. Comparar gerações de " &
                    "universos diferentes conclui que milhares de itens sumiram."))
            End If
            If g.Column("retention_cutoff") Is Nothing Then
                v.Add(New GateViolation("I7",
                    "generation não registra o cutoff de retenção. A §18 mediu que " &
                    "a janela existe e não foi o Iris que a escolheu."))
            End If
        End Sub

        ''' <summary>
        ''' S1 — o tipo de cobertura é persistido. S6 — as três contagens.
        ''' </summary>
        Private Shared Sub VerificarS1eS6(s As CacheSchema, v As List(Of GateViolation))
            Dim g = s.Table("generation")
            If g Is Nothing Then Return
            Dim k = g.Column("coverage_kind")
            If k Is Nothing OrElse Not k.IsRequired Then
                v.Add(New GateViolation("S1",
                    "generation não diz se foi completa ou incremental. A §16.2 mediu " &
                    "que o incremental não descobre Move-in."))
            End If
            For Each nome In {"rows_read", "count_before", "count_after", "distinct_keys"}
                Dim c = g.Column(nome)
                If c Is Nothing OrElse Not c.IsRequired Then
                    v.Add(New GateViolation("S6",
                        $"generation.{nome} ausente. É um dos números que precisam " &
                        "concordar para a geração ser aceitável (§17.1)."))
                End If
            Next
        End Sub

        ''' <summary>
        ''' S2 — só geração completa e válida marca ausência.
        ''' </summary>
        Private Shared Sub VerificarS2(s As CacheSchema, v As List(Of GateViolation))
            Dim a = s.Table("association")
            If a Is Nothing Then Return
            Dim c = a.Column("absent_by_generation")
            If c Is Nothing OrElse c.References <> "generation" Then
                v.Add(New GateViolation("S2",
                    "association não amarra a ausência a uma geração. Sem isso não " &
                    "dá para exigir que ela tenha sido completa e válida."))
            End If
        End Sub

        ''' <summary>
        ''' S3 — o schema aceita sobreposição transitória.
        ''' Medido: §16.1 matriz B, a encarnação velha e a nova coexistem no
        ''' cache vindas de gerações de instantes diferentes.
        ''' </summary>
        Private Shared Sub VerificarS3(s As CacheSchema, v As List(Of GateViolation))
            For Each t In {s.Table("association"), s.Table("incarnation")}
                If t Is Nothing Then Continue For
                For Each idx In t.UniqueIndexes
                    If idx.Length = 1 AndAlso
                       String.Equals(idx(0), "item_key", StringComparison.OrdinalIgnoreCase) Then
                        v.Add(New GateViolation("S3",
                            $"{t.Name} tem índice único sobre item_key: força 'um item " &
                            "está em exatamente uma pasta'. O corte fraturado da §16.1 " &
                            "produz legitimamente duas linhas."))
                    End If
                Next
            Next
        End Sub

        ''' <summary>
        ''' S4 — observações de pastas diferentes não formam instantâneo
        ''' global. Geração é POR PASTA.
        ''' </summary>
        Private Shared Sub VerificarS4(s As CacheSchema, v As List(Of GateViolation))
            Dim g = s.Table("generation")
            If g Is Nothing Then Return
            Dim f = g.Column("folder_key")
            If f Is Nothing OrElse Not f.IsRequired OrElse f.References <> "folder" Then
                v.Add(New GateViolation("S4",
                    "generation não é por pasta. Uma geração 'da caixa' afirmaria " &
                    "simultaneidade entre pastas, que a §16.1 mediu não existir."))
            End If
        End Sub

        ''' <summary>
        ''' S7 — nenhuma conclusão sobre ausência sem saber se a pasta está
        ''' INTEIRA no cache.
        '''
        ''' Medido: §19.2. Dezenas de pastas do usuário reportam ZERO itens e
        ''' não estão vazias — estão fora da janela. E o caso passa em TODOS
        ''' os sinais do S6, porque lidas = antes = depois = 0.
        ''' </summary>
        Private Shared Sub VerificarS7(s As CacheSchema, v As List(Of GateViolation))
            Dim f = s.Table("folder")
            If f Is Nothing Then
                v.Add(New GateViolation("S7", "não existe tabela de pasta"))
                Return
            End If
            Dim c = f.Column("coverage")
            If c Is Nothing OrElse Not c.IsRequired Then
                v.Add(New GateViolation("S7",
                    "folder não registra a cobertura do cache. Sem isso, " &
                    "'Count == 0' é lido como 'pasta vazia' — e na caixa medida " &
                    "dezenas de pastas cheias reportam zero (§19.2)."))
            End If
        End Sub

    End Class

End Namespace

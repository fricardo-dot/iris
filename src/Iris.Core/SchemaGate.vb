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
            VerificarS1b(schema, v)
            VerificarS5(schema, v)
            VerificarD1(schema, v)
            VerificarD2(schema, v)
            VerificarIdempotencia(schema, v)
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

            ' I8 — e NAO como booleano.
            '
            ' Um booleano afirma "checado" sem dizer ONDE, QUANDO nem em qual
            ' universo. E a §16.1 mediu que duas linhas no cache nao provam
            ' coexistencia: o corte fraturado produz exatamente isso para um
            ' Move legitimo. Por isso e REFERENCIA a evidencia.
            Dim coex = t.Column("coexistence_evidence_key")
            If coex Is Nothing OrElse coex.References <> "coexistence_evidence" Then
                v.Add(New GateViolation("I8",
                    "correlation_edge não aponta para evidência de coexistência. " &
                    "Booleano não serve: ele afirma 'checado' sem dizer onde nem quando."))
            End If
            If s.Table("coexistence_evidence") Is Nothing Then
                v.Add(New GateViolation("I8", "não existe tabela de evidência de coexistência"))
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

            ' A cobertura e VERSIONADA, nao uma coluna estatica na pasta:
            ' ela muda com sessao, janela e sincronizacao, e a conclusao so
            ' vale no universo em que foi tirada.
            Dim cob = s.Table("coverage_observation")
            If cob Is Nothing Then
                v.Add(New GateViolation("S7",
                    "não existe coverage_observation. Sem cobertura conhecida, " &
                    "'Count == 0' é lido como 'pasta vazia' — e na caixa medida " &
                    "dezenas de pastas cheias reportam zero (§19.2)."))
                Return
            End If
            For Each obrigatoria In {"folder_key", "coverage", "universe_fingerprint", "observed_at"}
                Dim c = cob.Column(obrigatoria)
                If c Is Nothing OrElse Not c.IsRequired Then
                    v.Add(New GateViolation("S7",
                        $"coverage_observation.{obrigatoria} ausente ou opcional"))
                End If
            Next

            ' E a ausencia tem de dizer QUAL cobertura a autorizou.
            Dim assoc = s.Table("association")
            If assoc IsNot Nothing Then
                Dim ac = assoc.Column("absent_by_coverage")
                If ac Is Nothing OrElse ac.References <> "coverage_observation" Then
                    v.Add(New GateViolation("S7",
                        "association não amarra a ausência a uma cobertura. Sem isso " &
                        "não dá para exigir que a pasta estivesse INTEIRA no cache."))
                End If
            End If
        End Sub

        ''' <summary>
        ''' S1 (parte b) — a TENTATIVA é separada da GERAÇÃO.
        '''
        ''' Checkpoint é estado de trabalho INCOMPLETO. Se ele morasse na
        ''' geração, ou a geração existiria antes de ser válida, ou não
        ''' haveria onde retomar. Foi o bloqueador da 1ª versão do plano.
        ''' </summary>
        Private Shared Sub VerificarS1b(s As CacheSchema, v As List(Of GateViolation))
            Dim tentativa = s.Table("scan_attempt")
            Dim geracao = s.Table("generation")
            If tentativa Is Nothing Then
                v.Add(New GateViolation("S1b", "nao existe scan_attempt: o checkpoint nao tem onde morar"))
                Return
            End If
            If geracao Is Nothing Then Return

            ' Checkpoint (cursor) NAO pode viver na geracao publicada.
            If geracao.Column("cursor") IsNot Nothing Then
                v.Add(New GateViolation("S1b",
                    "generation.cursor: checkpoint numa tabela imutavel. " &
                    "Ou a geracao existe antes de ser valida, ou nao ha onde retomar."))
            End If
            If tentativa.Column("cursor") Is Nothing Then
                v.Add(New GateViolation("S1b", "scan_attempt nao tem cursor: nao da para retomar"))
            End If
            ' O staging precisa existir e ser POR TENTATIVA.
            Dim staging = s.Table("scan_stage")
            If staging Is Nothing Then
                v.Add(New GateViolation("S1b",
                    "nao existe scan_stage. As chaves vistas iriam para dentro do cursor, " &
                    "que e exatamente o desenho que a revisao recusou."))
            ElseIf staging.Column("attempt_key") Is Nothing Then
                v.Add(New GateViolation("S1b", "scan_stage nao pertence a uma tentativa"))
            End If
            ' publication_log, para reprocessar depois de crash.
            If s.Table("publication_log") Is Nothing Then
                v.Add(New GateViolation("S1b",
                    "nao existe publication_log: um crash entre o commit e a publicacao " &
                    "perderia o evento sem deixar rastro"))
            End If
        End Sub

        ''' <summary>
        ''' S5 — retenção expurga dado derivado, nunca estado do usuário.
        ''' Aqui: apagar item ou pasta NÃO pode cascatear para user_state.
        ''' </summary>
        Private Shared Sub VerificarS5(s As CacheSchema, v As List(Of GateViolation))
            Dim us = s.Table("user_state")
            If us Is Nothing Then Return
            For Each c In us.Columns
                If c.OnDelete = DeleteAction.Cascade Then
                    v.Add(New GateViolation("S5",
                        $"user_state.{c.Name} cascateia. Apagar um item apagaria o " &
                        "trabalho do usuario junto."))
                End If
            Next
        End Sub

        ''' <summary>
        ''' D1 — o cache guarda METADADO, e nao conteudo.
        '''
        ''' Medido: metadado da caixa inteira = 11,7 MB; com texto = 372 MB.
        ''' E texto significa correspondencia EM CLARO no disco antes de
        ''' criptografia e retencao estarem decididas.
        ''' </summary>
        Private Shared Sub VerificarD1(s As CacheSchema, v As List(Of GateViolation))
            ' O nome sozinho nao basta: "has_attachments INTEGER" e uma FLAG,
            ' e "body TEXT" e conteudo. A primeira versao desta regra deu
            ' falso positivo em has_attachments — e regra que grita no
            ' schema certo e regra que sera desligada.
            '
            ' Conteudo mora em TEXT ou BLOB. Flag e contagem moram em
            ' INTEGER. So os dois primeiros sao suspeitos.
            Dim proibidos = {"body", "html", "attachment", "anexo", "corpo", "content_bytes"}
            For Each t In s.Tables
                For Each c In t.Columns
                    Dim tipo = c.Kind.ToUpperInvariant()
                    If tipo <> "TEXT" AndAlso tipo <> "BLOB" Then Continue For
                    Dim nome = c.Name.ToLowerInvariant()
                    For Each p In proibidos
                        If nome.Contains(p) Then
                            v.Add(New GateViolation("D1",
                                $"{t.Name}.{c.Name} e {tipo} e parece guardar conteudo. " &
                                "O 2.1 e metadado: 372 MB e correspondencia em claro no " &
                                "disco exigem a decisao de criptografia antes."))
                        End If
                    Next
                Next
            Next
            If s.Table("metadata_observation") Is Nothing Then
                v.Add(New GateViolation("D1",
                    "nao existe metadata_observation: o cache nao guarda o metadado " &
                    "que o 2.1 promete"))
            End If
        End Sub

        ''' <summary>
        ''' D2 — ambiente nao medido recusa operar, e a allowlist e DADO.
        ''' </summary>
        Private Shared Sub VerificarD2(s As CacheSchema, v As List(Of GateViolation))
            Dim amb = s.Table("environment_profile")
            If amb Is Nothing Then
                v.Add(New GateViolation("D2",
                    "nao existe environment_profile: a allowlist viraria constante no " &
                    "codigo, e ambiente nao medido operaria por omissao"))
                Return
            End If
            If amb.Column("allowed") Is Nothing OrElse Not amb.Column("allowed").IsRequired Then
                v.Add(New GateViolation("D2", "environment_profile nao diz se o ambiente e permitido"))
            End If
            If amb.Column("fingerprint") Is Nothing Then
                v.Add(New GateViolation("D2", "environment_profile sem fingerprint"))
            End If
        End Sub

        ''' <summary>
        ''' Idempotencia — reimportar nao pode duplicar.
        '''
        ''' Sem isto, uma retomada ou uma segunda importacao criam pasta,
        ''' encarnacao e associacao repetidas, e o S6 passa a rejeitar toda
        ''' varredura por chave duplicada. O sintoma aparece longe da causa.
        ''' </summary>
        Private Shared Sub VerificarIdempotencia(s As CacheSchema, v As List(Of GateViolation))
            Dim casos = {
                ("folder", New String() {"store_key", "provider_entry_id"}),
                ("incarnation", New String() {"folder_key", "provider_entry_id"}),
                ("association", New String() {"item_key", "folder_key"}),
                ("user_state", New String() {"item_key"})
            }
            For Each caso In casos
                Dim t = s.Table(caso.Item1)
                If t Is Nothing Then Continue For
                If Not t.TemUnico(caso.Item2) Then
                    v.Add(New GateViolation("IDEM",
                        $"{caso.Item1} sem unico sobre ({String.Join(", ", caso.Item2)}): " &
                        "reimportar duplicaria"))
                End If
            Next
        End Sub

    End Class

End Namespace

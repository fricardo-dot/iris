Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Assist
Imports Iris.Integration
Imports Iris.Model

Namespace Global.Iris.CrashHarness

    ''' <summary>
    ''' Um processo que morre de verdade no meio de uma gravação.
    '''
    ''' Existe porque injetar exceção e matar o processo provam coisas
    ''' DIFERENTES, e chamar as duas de "teste de crash" seria o mesmo erro
    ''' que a §16.5 cometeu com o Restrict — declarar coberto o que não foi
    ''' exercido:
    '''
    '''   - exceção injetada prova ATOMICIDADE: a transação desfaz o que
    '''     escreveu. Mas o processo continua vivo, o SQLite roda o rollback
    '''     e fecha a conexão com ordem. Nada disso acontece num crash.
    '''   - TerminateProcess prova DURABILIDADE: ninguém desfaz nada, ninguém
    '''     fecha nada, o arquivo fica como estava no disco. Quem recupera é
    '''     o WAL, na próxima abertura.
    '''
    ''' O que este harness NÃO prova: perda de energia. TerminateProcess mata
    ''' o processo, mas o Windows continua vivo e o que já foi entregue ao
    ''' sistema de arquivos continua lá. Faltar luz pode perder escrita que o
    ''' SO ainda não descarregou. Provar isso exigiria injeção de falha no
    ''' sistema de arquivos, e não está feito.
    ''' </summary>
    Module Program

        Public Function Main(args As String()) As Integer
            ' Modo DIARIO: grava a intencao, comeca o voo, e MORRE. Existe
            ' porque fechar um Using nao e morrer - o SQLite fecha com ordem,
            ' e a prova de durabilidade do diario precisa de um processo que
            ' nao fecha nada. Ver o XML doc deste modulo.
            '
            ' A conferencia vem ANTES da aridade do modo antigo: com tres
            ' argumentos ela nunca seria alcancada.
            If args.Length = 3 AndAlso String.Equals(args(1), "diario", StringComparison.Ordinal) Then
                Return Diario(args(0), args(2))
            End If

            If args.Length < 4 Then
                Console.Error.WriteLine("uso: <db> <folderKey> <ponto> <kill|throw> [attemptKeyParaRetomar]")
                Console.Error.WriteLine("     <db> diario <apos-intencao|em-voo|nenhum>")
                Return 2
            End If

            Dim caminho = args(0)

            Dim folderKey = Long.Parse(args(1))
            Dim ponto = args(2)
            Dim modo = args(3)
            Dim retomar As Long = If(args.Length > 4, Long.Parse(args(4)), 0L)

            ' Defeito deliberado, ligado pelo teste que serve de controle
            ' negativo. Ver CacheWriterDefects.
            If args.Length > 5 AndAlso args(5) = "checkpoint-antes" Then
                CacheWriterDefects.CheckpointAntesDasLinhas = True
            End If

            ' arg 6: "drenar" faz o harness seguir ate o consumidor.
            Dim drenar = args.Length > 6 AndAlso args(6) = "drenar" 

            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(caminho, CacheSchema.Intended(), falha)
                If db Is Nothing Then
                    Console.Error.WriteLine($"abrir falhou: {falha}")
                    Return 3
                End If

                Dim w As New CacheWriter(db)

                If Not String.Equals(ponto, "nenhum", StringComparison.Ordinal) Then
                    CrashInjection.Armar(ponto, AddressOf Morrer)
                    _modo = modo
                End If

                Dim tentativa As Long
                Dim primeiraPagina As Integer = 1
                If retomar > 0 Then
                    tentativa = retomar
                    ' Retomar de onde o cursor parou. O cursor E o checkpoint:
                    ' se ele diz "cursor-2", as paginas 1 e 2 estao gravadas.
                    Dim c = w.CursorDaTentativa(tentativa)
                    If c IsNot Nothing AndAlso c.StartsWith("cursor-") Then
                        primeiraPagina = Integer.Parse(c.Substring("cursor-".Length)) + 1
                    End If
                Else
                    tentativa = w.AbrirTentativa(folderKey, 1, "universo-teste", EpocaDe(db, folderKey), 1)
                    Console.Out.WriteLine($"tentativa={tentativa}")
                    Console.Out.Flush()
                End If

                For p = primeiraPagina To TotalDePaginas
                    w.GravarPagina(tentativa, folderKey, p, LinhasDaPagina(p), $"cursor-{p}")
                Next

                Dim g As Long = 0
                Dim r = w.Publicar(tentativa, folderKey, "completa",
                                   TotalDePaginas * LinhasPorPagina,
                                   TotalDePaginas * LinhasPorPagina, g)
                Console.Out.WriteLine($"resultado={r} geracao={g}")

                ' --- o DRENO, quando pedido ---
                '
                ' A condicao do 2.4 (§26.2) exige matar o processo entre
                ' Receber e MarcarDrenada e provar que a reabertura converge.
                ' O consumidor usado aqui e o AcervoService de verdade, o
                ' mesmo que o ViewModel usa - senao a prova seria sobre uma
                ' imitacao, que e o erro que a Q1 cobrou.
                If drenar Then
                    Dim acervo As New AcervoService(db, folderKey)
                    Dim dreno As New PublicationDrain(db)
                    Dim n = dreno.Drenar(acervo)
                    Console.Out.WriteLine(
                        $"drenadas={n} recebidas={acervo.Recebidas} " &
                        $"itens={acervo.Atual.Items.Count} cobertura={acervo.Atual.Cobertura}")
                End If
                Return 0
            End Using
        End Function

        ''' <summary>
        ''' O diário do egress, morrendo no ponto pedido.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE UM PROCESSO DE VERDADE</b>
        '''
        ''' O teste do diário fecha o <c>Using</c> e reabre o banco. Isso prova
        ''' que a reconciliação lê o que ficou escrito — e <b>não</b> prova que
        ''' ficou escrito: fechar o <c>Using</c> dá ao SQLite a chance de
        ''' descarregar tudo com ordem, que é exatamente o que um crash não dá.
        '''
        ''' Aqui o processo é morto com <c>TerminateProcess</c> depois de gravar
        ''' o passo. Se a intenção e o "em voo" não estiverem <b>durados</b> no
        ''' disco, a reabertura não acha nada — e o diário estaria afirmando, por
        ''' omissão, que nada saiu.
        ''' </summary>
        Private Function Diario(caminho As String, ponto As String) As Integer
            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(caminho, CacheSchema.Intended(), falha)
                If db Is Nothing Then
                    Console.Error.WriteLine($"abrir falhou: {falha}")
                    Return 3
                End If

                Dim j As New SqliteDisclosureJournal(db)
                Dim c = CapabilityDeTeste()
                Console.Out.WriteLine($"requestId={c.RequestId}")
                Console.Out.Flush()

                If Not j.Intencao(c, DateTimeOffset.UtcNow) Then
                    Console.Error.WriteLine("a intencao nao pegou")
                    Return 4
                End If
                If String.Equals(ponto, "apos-intencao", StringComparison.Ordinal) Then Morrer()

                ' So toca na "rede" depois de o voo estar DURADO. E a regra que
                ' o Iniciando devolver Boolean existe para impor: um passo que
                ' nao pegou e seguido de um envio produz egress sem registro.
                If Not j.Iniciando(c.RequestId, DateTimeOffset.UtcNow) Then
                    Console.Error.WriteLine("o inicio do voo nao pegou")
                    Return 5
                End If
                If String.Equals(ponto, "em-voo", StringComparison.Ordinal) Then Morrer()

                j.Concluir(c.RequestId, DateTimeOffset.UtcNow)
                Return 0
            End Using
        End Function

        ''' <summary>
        ''' Uma capability de verdade, do portão de verdade.
        '''
        ''' Fabricar uma aqui provaria o diário contra um objeto que a produção
        ''' nunca produz — o erro que a Q1 cobrou.
        ''' </summary>
        Private Function CapabilityDeTeste() As DisclosureCapability
            Dim agora = DateTimeOffset.UtcNow
            Dim pasta As New FolderKey("store-1", "pasta-1")
            Dim destino As New AssistDestination("provedor-de-teste",
                                                 "https://exemplo.invalido/v1",
                                                 "modelo-de-teste")
            Dim ativacao As New ActivationRecord(
                "ativacao-harness", 1, "harness", agora.AddDays(-1),
                "provedor-de-teste", "https://exemplo.invalido/v1", "modelo-de-teste",
                "local", "sem retencao", {AssistOperation.Resumir}, {pasta},
                Array.Empty(Of String)(), {LabelReadingKind.Absent}, {0})

            Dim mensagens As New List(Of MessageClassification)()
            Dim partes As New List(Of MessagePart)()
            For n = 1 To 2
                Dim chave As New ItemKey($"E-{n}", "store-1")
                Dim leitura As New LabelReading(chave, LabelReadingKind.Absent,
                                                LabelReadStage.Parse,
                                                version:=New LabelVersionEvidence(
                                                    $"E-{n}", agora, $"CK-{n}"))
                mensagens.Add(New MessageClassification(chave, pasta, leitura, temAnexo:=False))
                partes.Add(ContentPipeline.Preparar(
                    New MessageSnapshot(chave, $"CK-{n}", $"assunto {n}", "de@x.invalido",
                                        {"para@x.invalido"}, "corpo", False, True, temAnexo:=False)).Parte)
            Next

            Dim voo As New PreflightRequest(AssistOperation.Resumir, pasta, destino)
            Dim d = New DisclosurePolicy(ativacao).
                    Decidir(New DisclosureRequest(voo, mensagens), agora)
            Dim env = New EnvelopeBuilder().Montar(AssistOperation.Resumir, "resuma", partes)
            Return New CapabilityLedger().Emitir(d, env.Envelope, agora)
        End Function

        Public Const TotalDePaginas As Integer = 3
        Public Const LinhasPorPagina As Integer = 2

        ''' <summary>
        ''' As linhas são deterministicas de proposito: o teste precisa saber
        ''' exatamente o que deveria estar no disco depois do crash.
        ''' </summary>
        Public Function LinhasDaPagina(p As Integer) As IReadOnlyList(Of StagedRow)
            Dim r As New List(Of StagedRow)()
            For i = 1 To LinhasPorPagina
                r.Add(New StagedRow With {
                    .ProviderEntryId = $"E-{p}-{i}",
                    .Subject = $"Assunto {p}.{i}",
                    .SenderName = "Fulano",
                    .ReceivedAt = New DateTime(2026, 1, p, 12, i, 0, DateTimeKind.Utc).ToString("o"),
                    .SizeBytes = 1000L + p * 10 + i,
                    .HasAttachments = (i Mod 2 = 0),
                    .IsUnread = True,
                    .MessageClass = "IPM.Note"})
            Next
            Return r
        End Function

        Private _modo As String

        Private Sub Morrer()
            If String.Equals(_modo, "throw", StringComparison.Ordinal) Then
                Throw New InvalidOperationException("crash injetado")
            End If
            ' Kill() = TerminateProcess: sem finalizadores, sem rollback, sem
            ' fechar conexao. E o mais perto de arrancar o processo que da
            ' para fazer de dentro dele.
            '
            ' NAO usar Environment.Exit aqui: ele roda finalizadores e
            ' handlers de saida, ou seja, deixa o processo ARRUMAR A CASA
            ' antes de morrer. Um teste que "mata" assim exercita o caminho
            ' de encerramento limpo e chama isso de crash.
            Process.GetCurrentProcess().Kill()
        End Sub

        Private Function EpocaDe(db As CacheDatabase, folderKey As Long) As Long
            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText = $"SELECT reconcile_epoch FROM folder WHERE folder_key = {folderKey}"
                Return Convert.ToInt64(cmd.ExecuteScalar())
            End Using
        End Function

    End Module

End Namespace

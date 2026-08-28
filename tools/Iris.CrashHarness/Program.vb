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

            ' Modo ATIVACAO: le a cerimonia pelo caminho de PRODUCAO e conta o
            ' que aconteceu. Existe porque "salvei o arquivo, deu certo?" nao
            ' tem resposta na interface antes de abrir o Outlook — e porque
            ' conferir com um roteiro paralelo provaria o roteiro, e nao o
            ' carregador que o Iris usa de verdade.
            If args.Length >= 1 AndAlso String.Equals(args(0), "ativacao",
                                                      StringComparison.Ordinal) Then
                Return Ativacao(If(args.Length > 1, args(1), Nothing))
            End If

            ' Modo BUSCAR: exercita a busca do acervo contra um cache DE
            ' VERDADE, e mede quanto ela demora.
            '
            ' Existe por dois motivos. O primeiro e o de sempre neste projeto:
            ' teste com dado sintetico prova o algoritmo contra o vocabulario
            ' que eu mesmo inventei -- sem acento, sem prefixo de conversa, com
            ' um remetente so. Mil linhas de uma caixa real dizem outra coisa.
            '
            ' O segundo e o numero. A busca casa EM MEMORIA, e isso so se
            ' sustenta enquanto for instantaneo; o gatilho para trocar por FTS5
            ' esta escrito no BuscaNoAcervo. Sem uma forma de medir, o gatilho
            ' seria uma frase.
            '
            ' SO LEITURA. Abre o cache, procura, imprime. Nao grava nada.
            If args.Length >= 1 AndAlso String.Equals(args(0), "buscar",
                                                      StringComparison.Ordinal) Then
                Return Buscar(args)
            End If

            ' Modo AMBIENTE: a cerimonia do ambiente.
            '
            ' Nao e roteiro PowerShell -- eu escrevi um e ele nao roda: o
            ' PowerShell 5.1 vive no .NET Framework e nao carrega o
            ' Microsoft.Data.Sqlite do .NET 10. E melhor assim: aqui a leitura
            ' passa pelo MESMO CacheDatabase.Open do Iris, entao a cerimonia
            ' confere o banco que o programa abre, e nao um que so ela abre.
            If args.Length >= 1 AndAlso String.Equals(args(0), "ambiente",
                                                      StringComparison.Ordinal) Then
                Return Ambiente(args)
            End If

            ' Modo HISTORICO: le o diario de divulgacao do cache de producao.
            ' Existe para conferir o canario -- uma linha, um RequestId, um
            ' desfecho -- sem abrir o banco na mao.
            If args.Length >= 1 AndAlso String.Equals(args(0), "historico",
                                                      StringComparison.Ordinal) Then
                Return Historico(If(args.Length > 1, args(1), Nothing))
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
        ''' <summary>
        ''' Le a ativacao e diz o que o Iris vai fazer com ela.
        '''
        ''' Nao imprime identificador de pasta nem endpoint por extenso: a saida
        ''' costuma ir para um chat ou um bilhete, e o arquivo descreve a caixa
        ''' corporativa de alguem.
        ''' </summary>
        ''' <summary>
        ''' <b>A cerimônia do ambiente: mostra o que o Iris mediu, e grava a
        ''' decisão de quem responde pela caixa.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE ISTO É UMA CERIMÔNIA E NÃO UMA DETECÇÃO</b>
        '''
        ''' O gate D2 diz que a allowlist do ambiente é <b>dado</b>, e não
        ''' constante no código, justamente para que "ambiente não medido"
        ''' possa recusar operar. Se o Iris marcasse <c>allowed = 1</c> no que
        ''' ele mesmo detectou, estaria medindo e aprovando a própria medição, e
        ''' o D2 viraria decoração.
        '''
        ''' Então o programa grava o perfil com <c>allowed = 0</c> e para. Aqui
        ''' é onde o zero vira um.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O QUE ESTA AUTORIZAÇÃO NÃO DESTRAVA</b>
        '''
        ''' Em Exchange em cache a janela de sincronização <b>não é legível</b>
        ''' — a §22.3 mediu que <c>Store</c> não a expõe e a §22.4 mediu que o
        ''' registro do perfil também não. O Iris nunca vai saber que ela mudou,
        ''' então continua sem poder afirmar cobertura completa nem concluir
        ''' ausência. Autorizar destrava <b>varrer e contar</b>. A degradação é
        ''' permanente, e de propósito.
        ''' </summary>
        ''' <summary>
        ''' Onde o cache do usuario mora. Um lugar so decide isso: dois modos
        ''' com o caminho montado na mao divergiriam no dia em que ele mudasse,
        ''' e um deles passaria a conferir um arquivo que nao e o do programa.
        ''' </summary>
        Private Function CacheDoUsuario() As String
            Return IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "cache.db")
        End Function

        ''' <summary>
        ''' Procura no acervo e conta quanto tempo levou.
        '''
        ''' <c>Iris.CrashHarness buscar &lt;termo&gt; [caminho-do-cache]</c>
        ''' </summary>
        Private Function Buscar(args As String()) As Integer
            If args.Length < 2 Then
                Console.Error.WriteLine("uso: buscar <termo> [caminho-do-cache]")
                Return 2
            End If

            Dim termo = args(1)
            Dim alvo = If(args.Length > 2, args(2), CacheDoUsuario())
            Console.WriteLine($"cache:  {alvo}")
            Console.WriteLine($"termo:  {termo}")
            Console.WriteLine()

            If Not IO.File.Exists(alvo) Then
                Console.WriteLine("Nao existe cache. Ele nasce na primeira abertura do Iris.")
                Return 1
            End If

            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(alvo, CacheSchema.Intended(), falha)
                If db Is Nothing Then
                    Console.WriteLine($"o cache nao abriu: {falha}")
                    Return 1
                End If

                Dim relogio = Diagnostics.Stopwatch.StartNew()
                Dim r = New Iris.Integration.BuscaNoAcervo(db).Procurar(termo)
                relogio.Stop()

                Console.WriteLine($"achados:  {r.Achados.Count}")
                Console.WriteLine($"acervo:   {r.TotalNoAcervo} mensagens em {r.Consultadas.Count} pasta(s)")
                Console.WriteLine($"tempo:    {relogio.Elapsed.TotalMilliseconds:N1} ms")
                Console.WriteLine()

                For Each a In r.Achados.Take(10)
                    Dim assunto = If(a.Item.Subject, "")
                    If assunto.Length > 64 Then assunto = assunto.Substring(0, 61) & "..."
                    Console.WriteLine($"  [{a.NomeDaPasta}] {assunto}  <{a.Item.SenderName}>")
                Next
                If r.Achados.Count > 10 Then
                    Console.WriteLine($"  ... e mais {r.Achados.Count - 10}")
                End If

                Console.WriteLine()
                Console.WriteLine("RESSALVA:")
                Console.WriteLine("  " & r.Ressalva)
                Return 0
            End Using
        End Function

        Private Function Ambiente(args As String()) As Integer
            Dim autorizar As Long = 0
            Dim revogar = False
            Dim caminho As String = Nothing

            For i = 1 To args.Length - 1
                Select Case args(i)
                    Case "--autorizar"
                        If i + 1 < args.Length Then
                            i += 1
                            If Not Long.TryParse(args(i), autorizar) Then
                                Console.Error.WriteLine("chave invalida: " & args(i))
                                Return 2
                            End If
                        End If
                    Case "--revogar"
                        revogar = True
                    Case Else
                        caminho = args(i)
                End Select
            Next

            Dim alvo = If(String.IsNullOrWhiteSpace(caminho), CacheDoUsuario(), caminho)
            Console.WriteLine($"cache:  {alvo}")
            Console.WriteLine()

            If Not IO.File.Exists(alvo) Then
                Console.WriteLine("Nao existe cache. Ele nasce na primeira abertura do Iris.")
                Return 1
            End If

            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(alvo, CacheSchema.Intended(), falha)
                If db Is Nothing Then
                    Console.WriteLine($"o cache nao abriu: {falha}")
                    Return 1
                End If

                If autorizar > 0 Then Return Decidir(db, autorizar, Not revogar)
                Return Listar(db)
            End Using
        End Function

        Private Function Listar(db As CacheDatabase) As Integer
            Dim quantos = 0
            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText =
                    "SELECT environment_key, fingerprint, provider, cached_mode, " &
                    "       sync_window, allowed " &
                    "FROM environment_profile ORDER BY environment_key"
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        quantos += 1
                        Console.WriteLine(New String("-"c, 66))
                        Console.WriteLine($"  chave ....... {r.GetInt64(0)}")
                        Console.WriteLine($"  ambiente .... {r.GetString(1)}")
                        Console.WriteLine($"  provedor .... {r.GetString(2)}")
                        Console.WriteLine($"  cache ....... {If(r.GetInt32(3) = 1, "sim", "nao")}")
                        Console.WriteLine($"  janela ...... " &
                            If(r.IsDBNull(4), "NAO LEGIVEL (medido: 22.3 e 22.4)", r.GetString(4)))
                        If r.GetInt32(5) = 1 Then
                            Console.WriteLine("  AUTORIZADO por voce.")
                        Else
                            Console.WriteLine("  nao autorizado -- o Iris NAO varre neste ambiente.")
                            Console.WriteLine($"  Para autorizar:  -- ambiente --autorizar {r.GetInt64(0)}")
                        End If
                    End While
                End Using
            End Using
            Console.WriteLine(New String("-"c, 66))

            If quantos = 0 Then
                Console.WriteLine()
                Console.WriteLine("Nenhum ambiente medido ainda.")
                Console.WriteLine("Abra o Iris com o Outlook aberto: ele mede na primeira leitura.")
            End If

            Console.WriteLine()
            Console.WriteLine("O QUE VOCE AUTORIZA: que o Iris trate o que consegue ler nesta")
            Console.WriteLine("caixa como base para contar itens e varrer pastas.")
            Console.WriteLine()
            Console.WriteLine("O QUE NAO DESTRAVA: em Exchange em cache a janela de sincronizacao")
            Console.WriteLine("nao e legivel, entao o Iris nunca sabe que ela mudou. Ele continua")
            Console.WriteLine("sem afirmar cobertura completa e sem concluir ausencia.")
            Return 0
        End Function

        Private Function Decidir(db As CacheDatabase, chave As Long,
                                 permitir As Boolean) As Integer
            Dim impressao As String = Nothing
            Dim antes = -1

            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText =
                    "SELECT fingerprint, allowed FROM environment_profile " &
                    "WHERE environment_key = $k"
                cmd.Parameters.AddWithValue("$k", chave)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        impressao = r.GetString(0)
                        antes = r.GetInt32(1)
                    End If
                End Using
            End Using

            If impressao Is Nothing Then
                Console.WriteLine($"Nao existe perfil com a chave {chave}.")
                Console.WriteLine("Rode sem --autorizar para ver a lista.")
                Return 1
            End If

            Dim novo = If(permitir, 1, 0)
            If antes = novo Then
                Console.WriteLine($"Nada a fazer: o perfil {chave} ja esta " &
                                  If(permitir, "AUTORIZADO", "nao autorizado") & ".")
                Return 0
            End If

            Console.WriteLine($"Perfil ...... {chave}")
            Console.WriteLine($"Ambiente .... {impressao}")
            Console.WriteLine($"Passa de .... {If(antes = 1, "AUTORIZADO", "nao autorizado")} " &
                              $"para {If(permitir, "AUTORIZADO", "nao autorizado")}")
            Console.WriteLine()

            Using cmd = db.Connection.CreateCommand()
                cmd.CommandText =
                    "UPDATE environment_profile SET allowed = $a WHERE environment_key = $k"
                cmd.Parameters.AddWithValue("$a", novo)
                cmd.Parameters.AddWithValue("$k", chave)
                If cmd.ExecuteNonQuery() <> 1 Then
                    Console.WriteLine("a gravacao nao pegou.")
                    Return 1
                End If
            End Using

            Console.WriteLine("GRAVADO.")
            If permitir Then
                Console.WriteLine()
                Console.WriteLine("Lembre do que isto NAO destrava: a janela de sincronizacao")
                Console.WriteLine("continua ilegivel, entao nada de cobertura completa nem de")
                Console.WriteLine("concluir ausencia. A degradacao e permanente, e de proposito.")
            End If
            Return 0
        End Function

        ''' <summary>
        ''' O diário de divulgação, do mais recente para o mais antigo.
        '''
        ''' <b>Não mostra conteúdo</b>, porque não há: o diário registra o que
        ''' saiu, quando e sob qual autorização — nunca o texto. Se um dia
        ''' aparecer texto aqui, é defeito, e é grave.
        ''' </summary>
        Private Function Historico(caminho As String) As Integer
            Dim alvo = If(String.IsNullOrWhiteSpace(caminho), CacheDoUsuario(), caminho)

            Console.WriteLine($"cache:  {alvo}")
            Console.WriteLine($"existe: {IO.File.Exists(alvo)}")
            Console.WriteLine()

            If Not IO.File.Exists(alvo) Then
                Console.WriteLine("Sem cache nao ha diario — e sem diario a IA fica")
                Console.WriteLine("desligada, de proposito. O cache nasce na primeira")
                Console.WriteLine("abertura do Iris.")
                Return 1
            End If

            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(alvo, CacheSchema.Intended(), falha)
                If db Is Nothing Then
                    Console.WriteLine($"o cache nao abriu: {falha}")
                    Return 1
                End If

                Dim linhas = New SqliteDisclosureJournal(db).Ler(20)
                If linhas.Count = 0 Then
                    Console.WriteLine("Diario VAZIO: nada saiu desta maquina.")
                    Return 0
                End If

                Console.WriteLine($"{linhas.Count} linha(s), da mais recente:")
                Console.WriteLine()
                For Each e In linhas
                    Console.WriteLine($"  {e.Estagio,-12} req={e.RequestId}")
                    Console.WriteLine($"    ativacao ... {e.AtivacaoId} v{e.AtivacaoVersao}")
                    Console.WriteLine($"    operacao ... {e.Operacao}   mensagens: {e.Mensagens}")
                    Console.WriteLine($"    provedor ... {e.Provedor} / {e.Modelo}")
                    Console.WriteLine($"    bytes ...... {e.Bytes}   hash: {e.Hash}")
                    Console.WriteLine($"    nota ....... {e.Nota}" &
                                      If(e.CodigoHttp.HasValue,
                                         $"   HTTP {e.CodigoHttp.Value}", ""))
                    Console.WriteLine()
                Next
            End Using
            Return 0
        End Function

        Private Function Ativacao(caminho As String) As Integer
            Dim alvo = If(String.IsNullOrWhiteSpace(caminho),
                          ActivationLoader.CaminhoPadrao(), caminho)
            Dim agora = DateTimeOffset.Now

            Console.WriteLine($"arquivo: {alvo}")
            Console.WriteLine($"existe:  {IO.File.Exists(alvo)}")
            Console.WriteLine()

            ' ISTO NAO CONFERE PERMISSAO, e dizer isso alto importa.
            '
            ' A conferencia de dono e ACL mora no Iris.App, que e WPF e alvo
            ' -windows; este harness e console e net10.0, e referencia-lo
            ' arrastaria a interface inteira para dentro de uma ferramenta de
            ' diagnostico. Sem o aviso, "CARREGOU" seria lido como "esta tudo
            ' certo" — e o que ele prova e so conteudo e politica.
            Console.WriteLine("AVISO: este modo NAO confere dono, ACL, nem a pasta.")
            Console.WriteLine("       Para isso: tools\conferir-permissao.ps1")
            Console.WriteLine()

            Dim r = ActivationLoader.Carregar(alvo, agora)

            If Not r.Carregou Then
                Console.WriteLine($"NAO CARREGOU: {r.Falha}" &
                                  If(r.Campo.Length > 0, $"  (campo ""{r.Campo}"")", ""))
                Return 1
            End If

            Dim a = r.Record
            Console.WriteLine("CARREGOU.")
            Console.WriteLine($"  id .................. {a.Id} (versao {a.Versao})")
            Console.WriteLine($"  autoridade .......... {a.Autoridade}")
            Console.WriteLine($"  politica verificada . {a.PoliticaCorporativaVerificada}")
            Console.WriteLine($"  modelo .............. {a.Modelo}")
            Console.WriteLine($"  retencao zero ....... {a.ExigirRetencaoZero}")
            Console.WriteLine($"  provedores .......... {String.Join(", ", a.ProvedoresPermitidos)}")
            Console.WriteLine($"  operacoes ........... {String.Join(", ", a.Operacoes)}")
            Console.WriteLine($"  pastas .............. {a.Pastas.Count}")
            Console.WriteLine($"  leituras aceitas .... {String.Join(", ", a.Leituras)}")
            Console.WriteLine($"  vence em ............ {a.Ate.Value.ToLocalTime():dd/MM/yyyy HH:mm}")
            Console.WriteLine($"  VIGENTE agora ....... {a.Vigente(agora)}")
            Console.WriteLine()

            ' O PORTAO, com a mesma politica que a producao usa. Carregar nao e
            ' o mesmo que autorizar: um registro valido ainda pode ser negado.
            Dim politica As New DisclosurePolicy(a)
            Dim destino As New AssistDestination(a.Provedor, a.Endpoint, a.Modelo)
            For Each op In {AssistOperation.Resumir, AssistOperation.Redigir}
                Dim d = politica.Preflight(New PreflightRequest(op, a.Pastas(0), destino), agora)
                Console.WriteLine($"  preflight {op,-8} ... {If(d.Permitido, "PASSA", "NEGA: " & d.Explicacao)}")
            Next
            Return 0
        End Function

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

                ' Nothing, e nao 200: ESTE CAMINHO NAO FALA HTTP.
                '
                ' Ele exercita o protocolo do diario contra um crash, e nao ha
                ' provedor nenhum do outro lado. Gravar 200 aqui poria no
                ' registro a evidencia de uma resposta externa que nunca houve
                ' -- num arquivo cuja unica serventia e dizer o que de fato
                ' aconteceu. Nothing e a verdade: nao houve resposta.
                j.Concluir(c.RequestId, DateTimeOffset.UtcNow, Nothing)
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
                Array.Empty(Of String)(), {LabelReadingKind.Absent}, {0}, ate:=agora.AddDays(30), provedoresPermitidos:={"provedor-subjacente"})

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
            Return New CapabilityLedger().Emitir(d, env.Envelope, env.Envelope.Bytes(), agora)
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

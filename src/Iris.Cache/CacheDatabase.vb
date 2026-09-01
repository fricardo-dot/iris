Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Core
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Cache

    Public NotInheritable Class OpenFailure
        Public ReadOnly Property Reason As String
        Public ReadOnly Property Detail As String
        Friend Sub New(reason As String, detail As String)
            Me.Reason = reason
            Me.Detail = detail
        End Sub
        Public Overrides Function ToString() As String
            Return $"{Reason}: {Detail}"
        End Function
    End Class

    ''' <summary>
    ''' Abre o cache, e RECUSA abrir um banco que não corresponda ao modelo.
    '''
    ''' A diferença entre isto e o <see cref="SchemaGate"/> importa: o gate
    ''' valida o modelo que ESTÁ NO CÓDIGO. Ele prova que o schema que eu
    ''' escrevi respeita os invariantes — não que o arquivo no disco é aquele
    ''' schema. Um banco criado por uma versão antiga, ou editado à mão,
    ''' passa no gate e não corresponde a nada.
    '''
    ''' Por isso a abertura faz as duas coisas:
    '''   1. o gate, sobre o modelo;
    '''   2. a INTROSPECÇÃO, comparando o arquivo real com o modelo.
    '''
    ''' É o mesmo padrão que a Q1 cobrou com a coluna <c>Permission</c> e a
    ''' §16.5 com o <c>Restrict</c>: "não lançou" não é "funciona".
    ''' </summary>
    Public NotInheritable Class CacheDatabase
        Implements IDisposable

        Private ReadOnly _conn As SqliteConnection
        Private _disposed As Boolean

        ''' <summary>
        ''' <b>A trava desta conexão — e ela existe porque a conexão é uma só.</b>
        '''
        ''' <c>SqliteConnection</c> não tem contrato de uso simultâneo. O WAL
        ''' coordena <i>conexões diferentes</i>; ele não torna uma conexão
        ''' reentrante. Um leitor aberto enquanto outro caminho abre
        ''' <c>BEGIN IMMEDIATE</c> é erro em tempo de execução, e o erro sai no
        ''' caminho azarado — não no que causou.
        '''
        ''' O projeto vinha dizendo "quem chama serializa", e isso bastou enquanto
        ''' só a interface lia. Deixou de bastar quando a classificação passou a
        ''' gravar em lote a partir de outra thread, e a tela a ler rótulos por
        ''' linha desenhada: três caminhos, três travas diferentes, nenhuma delas
        ''' protegendo o recurso que os três disputam. Achado por revisão externa
        ''' em 01/09/2026.
        '''
        ''' <b>É reentrante</b> (é um <c>Monitor</c>), então uma leitura composta —
        ''' o leitor de rótulos chamando o do cache — não trava a si mesma.
        '''
        ''' <b>E ela protege a conexão, não a corretude do que se lê.</b> Duas
        ''' consultas seguidas sob a mesma trava não formam um retrato atômico —
        ''' para isso é preciso uma transação, e é o que a gravação de rótulos faz.
        ''' </summary>
        Public ReadOnly Property Trava As Object
            Get
                Return _trava
            End Get
        End Property
        Private ReadOnly _trava As New Object()

        Public ReadOnly Property Connection As SqliteConnection
            Get
                Return _conn
            End Get
        End Property

        Private Sub New(conn As SqliteConnection)
            _conn = conn
        End Sub

        ''' <summary>
        ''' Abre ou cria. Devolve <c>Nothing</c> em <paramref name="falha"/>
        ''' quando deu certo.
        ''' </summary>
        Public Shared Function Open(caminho As String, schema As CacheSchema,
                                    ByRef falha As OpenFailure) As CacheDatabase
            falha = Nothing

            ' 1. O gate, sobre o MODELO. Se o proprio modelo viola um
            '    invariante, nem faz sentido criar banco.
            Dim violacoes = SchemaGate.Check(schema)
            If violacoes.Count > 0 Then
                falha = New OpenFailure("gate",
                    String.Join(" | ", violacoes.Select(Function(x) x.ToString())))
                Return Nothing
            End If

            Dim conn As SqliteConnection = Nothing
            Try
                conn = New SqliteConnection($"Data Source={caminho}")
                conn.Open()

                ' FK vem DESLIGADA no SQLite. Sem isto, todas as FKs do
                ' schema seriam decorativas.
                Executar(conn, "PRAGMA foreign_keys = ON")

                ' 0. O ARQUIVO ESTA INTEIRO?
                '
                ' O cabecalho e o sqlite_schema podem estar legiveis com uma
                ' pagina de dados corrompida no meio -- e ai a introspecao passa,
                ' o banco abre, e a corrupcao aparece semanas depois num SELECT
                ' qualquer, como "database disk image is malformed". Aqui mora o
                ' diario do egress, que nao se reconstroi de lugar nenhum: e
                ' melhor recusar cedo, com o arquivo intacto para quem for
                ' socorrer.
                '
                ' quick_check e nao integrity_check: o segundo le o banco inteiro
                ' e conferiria tambem as FKs, o que num arquivo grande custa
                ' segundos a cada abertura. O primeiro pega corrupcao de pagina,
                ' que e o que este passo existe para pegar.
                ' Achado por revisao externa em 31/08/2026.
                Dim inteiro = Integro(conn)
                If inteiro IsNot Nothing Then
                    falha = New OpenFailure("integridade", inteiro)
                    conn.Dispose()
                    Return Nothing
                End If

                Dim versao = LerVersao(conn)
                Dim vazio = ContarTabelas(conn) = 0

                ' ARQUIVO VAZIO MARCADO COM VERSAO FUTURA NAO E ARQUIVO NOVO.
                '
                ' O ramo do vazio vinha primeiro e chamava Criar, que grava
                ' user_version = 6 por cima do 7 -- rebaixando em silencio a marca
                ' deixada por uma versao do programa que sabia coisas que esta nao
                ' sabe. Um banco recem-criado tem versao 0; qualquer outra coisa
                ' num arquivo sem tabelas e um arquivo que alguem preparou.
                If vazio AndAlso versao > SqliteDdl.SchemaVersion Then
                    falha = New OpenFailure("versao",
                        $"arquivo sem tabelas marcado como versao {versao}, mais nova " &
                        $"que a esperada {SqliteDdl.SchemaVersion}")
                    conn.Dispose()
                    Return Nothing
                End If

                If vazio Then
                    Criar(conn, schema)
                ElseIf versao <> SqliteDdl.SchemaVersion Then
                    ' Versao divergente so passa por um caminho CONHECIDO.
                    '
                    ' Migrar sem saber de onde para onde continua sendo pior
                    ' que recusar -- e por isso o que nao esta em
                    ' SqliteDdl.Migracoes continua sendo recusado, com a mesma
                    ' mensagem de antes. O que mudou foi que apagar o arquivo
                    ' deixou de ser barato: o diario do egress mora aqui, e ele
                    ' nao se reconstroi do Outlook como o resto.
                    If Not Migrar(conn, schema, falha) Then
                        conn.Dispose()
                        Return Nothing
                    End If
                End If

                ' 2. INTROSPECCAO: o arquivo REAL corresponde ao modelo?
                Dim diferencas = SchemaIntrospector.Comparar(conn, schema)
                If diferencas.Count > 0 Then
                    falha = New OpenFailure("divergencia", String.Join(" | ", diferencas))
                    conn.Dispose()
                    Return Nothing
                End If

                ' 3. E as FKs estao mesmo LIGADAS?
                If Not ForeignKeysLigadas(conn) Then
                    falha = New OpenFailure("fk", "foreign_keys nao ficou ligada")
                    conn.Dispose()
                    Return Nothing
                End If

                Return New CacheDatabase(conn)

            Catch ex As Exception
                If conn IsNot Nothing Then conn.Dispose()
                falha = New OpenFailure("excecao", ex.Message)
                Return Nothing
            End Try
        End Function

        Private Shared Sub Criar(conn As SqliteConnection, schema As CacheSchema)
            Using tx = conn.BeginTransaction()
                For Each cmd In SqliteDdl.Generate(schema)
                    ' PRAGMA nao vai dentro de transacao em alguns casos;
                    ' journal_mode em particular.
                    If cmd.StartsWith("PRAGMA journal_mode", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                    Executar(conn, cmd, tx)
                Next
                tx.Commit()
            End Using
            Executar(conn, "PRAGMA journal_mode = WAL")
        End Sub

        Private Shared Sub Executar(conn As SqliteConnection, sql As String,
                                    Optional tx As SqliteTransaction = Nothing)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = sql
                If tx IsNot Nothing Then cmd.Transaction = tx
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' Sobe de <paramref name="de"/> até <see cref="SqliteDdl.SchemaVersion"/>
        ''' usando <b>só</b> passos listados. Um degrau faltando recusa tudo.
        ''' </summary>
        ''' <remarks>
        ''' ------------------------------------------------------------------
        ''' <b>UMA TRANSAÇÃO PARA A ESCADA INTEIRA, E ELA SÓ FECHA SE O RESULTADO
        ''' ESTIVER CERTO</b>
        '''
        ''' A primeira versão fazia um <c>COMMIT</c> por degrau e deixava a
        ''' conferência para a introspecção que roda <b>depois</b>, em
        ''' <see cref="Open"/>. Isso escolhia o degrau pelo número declarado no
        ''' <c>user_version</c> e não pela forma real do arquivo: um banco
        ''' marcado como 1 mas com outra coluna faltando levava o
        ''' <c>ALTER TABLE</c>, era promovido a 2, e só então era recusado — e
        ''' na abertura seguinte já não havia caminho nenhum, porque a versão de
        ''' origem tinha sido apagada pela própria tentativa.
        '''
        ''' Agora a escada inteira e a <b>verificação do resultado</b> moram na
        ''' mesma transação. Forma errada não é diagnosticada por antecipação:
        ''' ela é descoberta pela mesma introspecção de sempre, e o
        ''' <c>ROLLBACK</c> devolve o arquivo <b>intacto</b>, na versão em que
        ''' ele estava. Uma migração que não produz o schema esperado não deixa
        ''' rastro nenhum.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>O LOCK VEM ANTES DA DECISÃO</b>
        '''
        ''' <c>BEGIN IMMEDIATE</c>, e a versão é <b>relida dentro</b> dele. Duas
        ''' instâncias do Iris podiam ler "versão 1" ao mesmo tempo; a primeira
        ''' migrava, e a segunda tentava acrescentar de novo uma coluna que já
        ''' existia e recusava abrir um banco que estava <b>correto</b>.
        '''
        ''' Relendo sob o lock, quem chegou depois vê a versão nova e não tem o
        ''' que fazer.
        ''' </remarks>
        Private Shared Function Migrar(conn As SqliteConnection, schema As CacheSchema,
                                       ByRef falha As OpenFailure) As Boolean
            Try
                ' IMMEDIATE, e nao o DEFERRED que e o padrao: a trava de
                ' escrita e tomada AGORA, e nao na primeira escrita.
                '
                ' Um "SELECT 1" nao serve de substituto -- ele pega trava de
                ' LEITURA, e duas instancias leem a mesma versao felizes. A
                ' decisao de qual degrau subir tem de ser tomada com a escrita
                ' ja garantida, ou a primeira instancia migra e a segunda
                ' tenta acrescentar de novo uma coluna que ja existe, e recusa
                ' abrir um banco que esta CORRETO.
                Using tx = conn.BeginTransaction(deferred:=False)
                    Dim atual = LerVersao(conn, tx)

                    ' Para tras nao ha caminho: um banco mais NOVO que o
                    ' programa foi escrito por uma versao que sabia coisas que
                    ' esta nao sabe.
                    If atual > SqliteDdl.SchemaVersion Then
                        falha = New OpenFailure("versao",
                            $"banco na versao {atual}, mais nova que a esperada " &
                            $"{SqliteDdl.SchemaVersion}")
                        Return False
                    End If

                    While atual < SqliteDdl.SchemaVersion
                        Dim passo As IReadOnlyList(Of String) = Nothing
                        If Not SqliteDdl.Migracoes.TryGetValue(atual, passo) Then
                            falha = New OpenFailure("versao",
                                $"banco na versao {atual}, esperado " &
                                $"{SqliteDdl.SchemaVersion}, e nao ha migracao conhecida")
                            Return False
                        End If

                        For Each cmd In passo
                            Executar(conn, cmd, tx)
                        Next
                        Executar(conn, $"PRAGMA user_version = {atual + 1}", tx)
                        atual += 1
                    End While

                    ' A CONFERENCIA ANTES DO COMMIT.
                    '
                    ' A introspecao le pela MESMA conexao, entao enxerga o que
                    ' a transacao ainda nao publicou. Se a forma nao bate, o
                    ' Using desfaz tudo -- inclusive o user_version -- e o
                    ' arquivo do usuario continua exatamente como estava.
                    '
                    ' E "a forma" e o que o SchemaIntrospector compara: tabela,
                    ' coluna, tipo, nulidade, chave primaria, indice unico e
                    ' FK. Ele NAO compara CHECK, trigger, view, default nem
                    ' collation -- entao um banco divergente so numa dessas
                    ' passa aqui. Nao e buraco desta migracao: o mesmo arquivo
                    ' ja seria aceito se estivesse marcado como versao 2. E
                    ' dito por escrito para ninguem ler "produziu o schema
                    ' esperado" como equivalencia de DDL.
                    Dim diferencas = SchemaIntrospector.Comparar(conn, schema)
                    If diferencas.Count > 0 Then
                        falha = New OpenFailure("migracao",
                            "a migracao nao produziu o schema esperado, e foi desfeita: " &
                            String.Join(" | ", diferencas))
                        Return False
                    End If

                    tx.Commit()
                End Using
            Catch ex As Exception
                falha = New OpenFailure("migracao",
                    $"a migracao falhou e foi desfeita: {ex.Message}")
                Return False
            End Try

            Return True
        End Function

        Private Shared Function LerVersao(conn As SqliteConnection,
                                          Optional tx As SqliteTransaction = Nothing) As Integer
            Using cmd = conn.CreateCommand()
                cmd.Transaction = tx
                cmd.CommandText = "PRAGMA user_version"
                Return Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Function

        Private Shared Function ContarTabelas(conn As SqliteConnection) As Integer
            Using cmd = conn.CreateCommand()
                cmd.CommandText =
                    "SELECT COUNT(*) FROM sqlite_schema WHERE type='table' AND name NOT LIKE 'sqlite_%'"
                Return Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Function

        ''' <summary>
        ''' <c>Nothing</c> quando o arquivo está íntegro; a queixa do SQLite
        ''' quando não está.
        '''
        ''' <b>Arquivo novo passa de graça</b>: <c>quick_check</c> num banco vazio
        ''' devolve <c>ok</c> sem ler nada.
        ''' </summary>
        Private Shared Function Integro(conn As SqliteConnection) As String
            Try
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "PRAGMA quick_check(1)"
                    Dim v = cmd.ExecuteScalar()
                    Dim r = If(v Is Nothing OrElse v Is DBNull.Value, "", CStr(v))
                    If String.Equals(r, "ok", StringComparison.OrdinalIgnoreCase) Then Return Nothing
                    Return $"quick_check: {r}"
                End Using
            Catch ex As Exception
                ' O proprio quick_check estourar ja e a resposta.
                Return $"quick_check nao completou: {ex.Message}"
            End Try
        End Function

        Private Shared Function ForeignKeysLigadas(conn As SqliteConnection) As Boolean
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "PRAGMA foreign_keys"
                Return Convert.ToInt32(cmd.ExecuteScalar()) = 1
            End Using
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            _conn?.Dispose()
        End Sub

    End Class

End Namespace

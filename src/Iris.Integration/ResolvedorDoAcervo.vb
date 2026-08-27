Imports Iris.Cache
Imports Iris.Model
Imports Iris.Sync
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Integration

    ''' <summary>
    ''' O que o cache sabe sobre um ambiente: a chave, e se ele foi
    ''' <b>autorizado</b>.
    ''' </summary>
    Public NotInheritable Class PerfilDoAmbiente

        Public ReadOnly Property Chave As Long
        ''' <summary>A impressão digital, como ela está gravada.</summary>
        Public ReadOnly Property Fingerprint As String
        ''' <summary>
        ''' <b>Alguém autorizou este ambiente?</b>
        '''
        ''' Nasce <c>False</c>. Quem muda para verdadeiro é a cerimônia —
        ''' <c>tools/autorizar-ambiente.ps1</c> —, e não o programa.
        ''' </summary>
        Public ReadOnly Property Permitido As Boolean
        ''' <summary><c>True</c> quando esta execução criou a linha.</summary>
        Public ReadOnly Property Novo As Boolean

        Friend Sub New(chave As Long, fingerprint As String,
                       permitido As Boolean, novo As Boolean)
            Me.Chave = chave
            Me.Fingerprint = fingerprint
            Me.Permitido = permitido
            Me.Novo = novo
        End Sub
    End Class

    ''' <summary>
    ''' <b>Traduz identidade do Outlook em chave do cache.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O BURACO QUE ISTO FECHA</b>
    '''
    ''' A Fase 2 entregou o <c>SweepRunner</c>, o <c>OutlookSweepSource</c> e o
    ''' <c>SqliteSweepSink</c> — e nada que os ligasse. O sink pede
    ''' <c>folderKey</c> e <c>environmentKey</c> como <c>Long</c>; o Outlook
    ''' fala em <c>(StoreId, EntryId)</c>, que são strings. Nenhum código de
    ''' produção fazia a travessia, então <b>o aplicativo nunca varreu</b>: os
    ''' testes semeavam <c>1, 1, 1</c> na mão e seguiam.
    '''
    ''' Enquanto a ponte não existia, a faixa do acervo mostrava a pasta de
    ''' chave 1 — que era a que o teste tinha criado, e não a que o usuário
    ''' estava olhando.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>IDEMPOTENTE, E POR ISSO SEM "CRIAR OU ATUALIZAR"</b>
    '''
    ''' Resolver a mesma pasta duas vezes devolve a mesma chave e <b>não</b>
    ''' toca em mais nada da linha. Em particular não mexe em
    ''' <c>reconcile_epoch</c>, <c>published_generation_key</c> nem
    ''' <c>stability</c>: esses três são estado de sincronização, e uma
    ''' tradução de identidade que os reescrevesse apagaria o trabalho da
    ''' varredura anterior toda vez que alguém clicasse na pasta.
    '''
    ''' O nome é a única exceção, e é enfeite: ele muda quando o usuário
    ''' renomeia a pasta no Outlook, e nada depende dele.
    ''' </summary>
    Public NotInheritable Class ResolvedorDoAcervo

        Private ReadOnly _db As CacheDatabase

        Public Sub New(db As CacheDatabase)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _db = db
        End Sub

        ' ==============================================================

        ''' <summary>
        ''' A chave da pasta, criando <c>store</c> e <c>folder</c> se preciso.
        ''' </summary>
        ''' <remarks>
        ''' Tudo numa transação: sem ela, morrer entre o <c>store</c> e o
        ''' <c>folder</c> deixaria um store órfão, e a próxima tentativa
        ''' esbarraria no índice único de <c>provider_store_id</c> sem ter
        ''' pasta nenhuma para mostrar.
        ''' </remarks>
        Public Function Pasta(storeId As String, entryId As String,
                              nome As String) As Long
            If String.IsNullOrWhiteSpace(storeId) Then
                Throw New ArgumentException("store sem identificador", NameOf(storeId))
            End If
            If String.IsNullOrWhiteSpace(entryId) Then
                Throw New ArgumentException("pasta sem identificador", NameOf(entryId))
            End If

            Using tx = _db.Connection.BeginTransaction(deferred:=False)
                Dim storeKey = Achar(tx, "SELECT store_key FROM store WHERE provider_store_id = $v",
                                     ("$v", storeId))
                If Not storeKey.HasValue Then
                    storeKey = Inserir(tx,
                        "INSERT INTO store (provider_store_id, display_name) VALUES ($v, $n); " &
                        "SELECT last_insert_rowid()",
                        ("$v", storeId), ("$n", CObj(DBNull.Value)))
                End If

                Dim folderKey = Achar(tx,
                    "SELECT folder_key FROM folder WHERE store_key = $s AND provider_entry_id = $e",
                    ("$s", storeKey.Value), ("$e", entryId))

                If folderKey.HasValue Then
                    ' SO o nome. Ver o doc da classe: epoca, geracao publicada e
                    ' estabilidade sao estado de sincronizacao, e reescreve-los
                    ' aqui apagaria a varredura anterior a cada clique.
                    Executar(tx, "UPDATE folder SET name = $n WHERE folder_key = $k",
                             ("$n", Texto(nome)), ("$k", folderKey.Value))
                Else
                    ' 'estavel' e o padrao ate uma varredura dizer o contrario;
                    ' reconcile_epoch 0 e "nunca reconciliada".
                    folderKey = Inserir(tx,
                        "INSERT INTO folder (store_key, provider_entry_id, name, " &
                        "  reconcile_epoch, stability) VALUES ($s, $e, $n, 0, 'estavel'); " &
                        "SELECT last_insert_rowid()",
                        ("$s", storeKey.Value), ("$e", entryId), ("$n", Texto(nome)))
                End If

                tx.Commit()
                Return folderKey.Value
            End Using
        End Function

        ' ==============================================================

        ''' <summary>
        ''' O perfil do ambiente medido. <b>Nasce não autorizado.</b>
        ''' </summary>
        ''' <remarks>
        ''' ------------------------------------------------------------------
        ''' <b>O PROGRAMA NÃO SE AUTORIZA</b>
        '''
        ''' O gate D2 diz que a allowlist do ambiente é <b>dado</b>, e não
        ''' constante no código, justamente para "ambiente não medido" poder
        ''' recusar operar. Se este método gravasse <c>allowed = 1</c> para o
        ''' que ele mesmo detectou, o D2 viraria decoração: o Iris estaria
        ''' medindo e aprovando a própria medição.
        '''
        ''' Então ele grava <c>allowed = 0</c> e para. Quem vira para 1 é a
        ''' cerimônia — <c>tools/autorizar-ambiente.ps1</c> —, que mostra o que
        ''' foi detectado e pede a decisão de quem responde pela caixa. É o
        ''' mesmo desenho da ativação da IA, pelo mesmo motivo.
        '''
        ''' <b>Reencontrar um perfil nunca rebaixa.</b> Ler a linha existente e
        ''' devolver o <c>allowed</c> dela é o que faz a autorização durar; se
        ''' este método reescrevesse 0 a cada abertura, a cerimônia valeria até
        ''' o próximo start do programa.
        ''' </remarks>
        Public Function Ambiente(f As EnvironmentFingerprint) As PerfilDoAmbiente
            If f Is Nothing Then Throw New ArgumentNullException(NameOf(f))
            Dim valor = f.Value()

            Using tx = _db.Connection.BeginTransaction(deferred:=False)
                Using cmd = _db.Connection.CreateCommand()
                    cmd.Transaction = tx
                    cmd.CommandText =
                        "SELECT environment_key, allowed FROM environment_profile " &
                        "WHERE fingerprint = $f"
                    cmd.Parameters.AddWithValue("$f", valor)
                    Using r = cmd.ExecuteReader()
                        If r.Read() Then
                            Dim achado As New PerfilDoAmbiente(
                                r.GetInt64(0), valor, r.GetInt32(1) = 1, False)
                            r.Close()
                            tx.Commit()
                            Return achado
                        End If
                    End Using
                End Using

                Dim chave = Inserir(tx,
                    "INSERT INTO environment_profile (fingerprint, provider, cached_mode, " &
                    "  sync_window, policy_version, allowed) " &
                    "VALUES ($f, $p, $c, $w, $v, 0); SELECT last_insert_rowid()",
                    ("$f", valor),
                    ("$p", f.Provider.ToString()),
                    ("$c", If(f.CachedMode, 1, 0)),
                    ("$w", Texto(f.WindowToken)),
                    ("$v", f.PolicyVersion))

                tx.Commit()
                Return New PerfilDoAmbiente(chave.Value, valor, False, True)
            End Using
        End Function

        ' ==============================================================

        Private Function Achar(tx As SqliteTransaction, sql As String,
                               ParamArray p As (Nome As String, Valor As Object)()) As Long?
            Using cmd = _db.Connection.CreateCommand()
                cmd.Transaction = tx
                cmd.CommandText = sql
                For Each par In p
                    cmd.Parameters.AddWithValue(par.Nome, par.Valor)
                Next
                Dim v = cmd.ExecuteScalar()
                If v Is Nothing OrElse v Is DBNull.Value Then Return Nothing
                Return Convert.ToInt64(v)
            End Using
        End Function

        Private Function Inserir(tx As SqliteTransaction, sql As String,
                                 ParamArray p As (Nome As String, Valor As Object)()) As Long?
            Return Achar(tx, sql, p)
        End Function

        Private Sub Executar(tx As SqliteTransaction, sql As String,
                             ParamArray p As (Nome As String, Valor As Object)())
            Using cmd = _db.Connection.CreateCommand()
                cmd.Transaction = tx
                cmd.CommandText = sql
                For Each par In p
                    cmd.Parameters.AddWithValue(par.Nome, par.Valor)
                Next
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>Texto vazio vira NULO: coluna opcional não guarda "".</summary>
        Private Shared Function Texto(v As String) As Object
            If String.IsNullOrWhiteSpace(v) Then Return DBNull.Value
            Return v
        End Function

    End Class

    ''' <summary>
    ''' <b>O ambiente que o Outlook está reportando agora.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' Puro e sem COM: recebe o <see cref="StoreInfo"/> que o broker já
    ''' devolve e traduz. A §22.3 mediu que <c>Store</c> não expõe janela de
    ''' sincronização — só <c>IsCachedExchange</c> e <c>ExchangeStoreType</c> —
    ''' e a §22.4 mediu que o registro do perfil também não. Por isso o
    ''' <c>WindowToken</c> sai <c>Nothing</c>, e sai <b>declaradamente</b>.
    ''' </summary>
    Public Module AmbienteMedido

        Public Function De(store As StoreInfo) As EnvironmentFingerprint
            If store Is Nothing Then
                Return New EnvironmentFingerprint(ProviderKind.Desconhecido, False, Nothing)
            End If
            Return New EnvironmentFingerprint(Especie(store), store.IsCachedExchange, Nothing)
        End Function

        ''' <summary>
        ''' <c>ExchangeStoreType</c> é o número que o OOM devolve, como texto.
        ''' Valor que não se reconhece vira <see cref="ProviderKind.Desconhecido"/>
        ''' — e desconhecido recusa, que é o lado seguro.
        ''' </summary>
        Public Function Especie(store As StoreInfo) As ProviderKind
            If store Is Nothing Then Return ProviderKind.Desconhecido

            Select Case Trim(If(store.ExchangeStoreType, ""))
                Case "0", "1", "2", "3"
                    ' 0 PrimaryExchangeMailbox, 1 DelegateMailbox,
                    ' 2 PublicFolder, 3 ExchangeMailbox.
                    Return If(store.IsCachedExchange,
                              ProviderKind.ExchangeCached, ProviderKind.ExchangeOnline)
                Case "4"
                    ' 4 = NotExchange: PST ou outro provedor local.
                    Return ProviderKind.PstLocal
                Case Else
                    Return ProviderKind.Desconhecido
            End Select
        End Function

    End Module

End Namespace

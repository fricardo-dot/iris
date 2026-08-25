Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Cache
Imports Iris.Sync
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Integration

    ''' <summary>Um item do manifesto: metadado, e o que se sabe sobre a presença dele.</summary>
    Public NotInheritable Class ManifestItem
        Public ReadOnly Property ProviderEntryId As String
        Public ReadOnly Property Subject As String
        Public ReadOnly Property SenderName As String
        Public ReadOnly Property ReceivedAt As String
        Public ReadOnly Property IsUnread As Boolean?
        Public ReadOnly Property Presence As PresenceState

        Friend Sub New(id As String, subject As String, sender As String, received As String,
                       unread As Boolean?, presence As PresenceState)
            ProviderEntryId = id
            Me.Subject = subject
            SenderName = sender
            ReceivedAt = received
            IsUnread = unread
            Me.Presence = presence
        End Sub
    End Class

    ''' <summary>
    ''' O que o Iris sabe sobre uma pasta — <b>e o quanto disso ele alcançou</b>.
    '''
    ''' A cobertura vem no MESMO objeto que os itens, então quem pega a lista
    ''' já tem a ressalva na mão. É acoplamento útil, e é o que dá para fazer
    ''' desta camada — <b>não é enforcement de apresentação</b>: nada impede o
    ''' chamador de ler só <c>Items</c> e ignorar o resto. Impedir de verdade
    ''' exigiria um modelo de apresentação já qualificado, ou um componente
    ''' visual que recuse renderizar conteúdo parcial sem a ressalva.
    '''
    ''' O motivo é concreto. Em modo cached o Iris publica cobertura
    ''' <c>Parcial</c>, e o manifesto é um <b>acervo</b>, não o estado corrente
    ''' da caixa: pode faltar mensagem que existe no servidor, e pode conter
    ''' mensagem que o usuário já apagou — essa aparece como
    ''' <see cref="PresenceState.Suspeito"/>. Uma UI que exiba isso como se
    ''' fosse a verdade corrente engana, e engana em silêncio.
    ''' </summary>
    Public NotInheritable Class FolderManifest
        Public ReadOnly Property FolderKey As Long
        Public ReadOnly Property GenerationKey As Long?
        Public ReadOnly Property Cobertura As FolderCoverage
        Public ReadOnly Property PublishedAt As String
        Public ReadOnly Property Items As IReadOnlyList(Of ManifestItem)

        ''' <summary>
        ''' Se o alcance do Iris ENCOLHEU desde a geração anterior.
        '''
        ''' Está aqui porque o <see cref="ContractionDetector"/> precisava de um
        ''' consumidor. Sem ele, o detector calculava um veredito que ninguém
        ''' lia — e a §25 afirmava que "encolheu → as conclusões anteriores
        ''' caem", quando o que acontecia era "o detector consegue devolver um
        ''' diagnóstico, se alguém o chamar".
        '''
        ''' O efeito real é este: a contração entra na <see cref="Ressalva"/>, e
        ''' a ressalva é o que a UI mostra. Não é invalidação de conclusão —
        ''' <b>não há conclusão a invalidar</b>, porque em cached a cobertura já
        ''' é sempre parcial e ausência já é proibida (§23). É aviso, e o aviso
        ''' é a única consequência que o escopo aceito comporta.
        ''' </summary>
        Public ReadOnly Property Contracao As ContractionReport

        Friend Sub New(folderKey As Long, generationKey As Long?, cobertura As FolderCoverage,
                       publishedAt As String, items As IEnumerable(Of ManifestItem),
                       Optional contracao As ContractionReport = Nothing)
            Me.FolderKey = folderKey
            Me.GenerationKey = generationKey
            Me.Cobertura = cobertura
            Me.PublishedAt = publishedAt
            Me.Items = If(items, Enumerable.Empty(Of ManifestItem)()).ToList()
            Me.Contracao = contracao
        End Sub

        ''' <summary>
        ''' Se este manifesto pode ser exibido como o estado corrente da pasta.
        '''
        ''' Falso sempre que a cobertura não for <c>Completa</c> — e hoje ela
        ''' nunca é, em modo cached (§23).
        ''' </summary>
        Public ReadOnly Property EhEstadoCorrente As Boolean
            Get
                Return Cobertura = FolderCoverage.Completa
            End Get
        End Property

        ''' <summary>
        ''' Como a UI deve qualificar o que está mostrando. Texto e não
        ''' booleano porque "por que não é o estado corrente" é a informação
        ''' que o usuário precisa, e um booleano a joga fora.
        ''' </summary>
        Public ReadOnly Property Ressalva As String
            Get
                If GenerationKey Is Nothing Then
                    Return "Esta pasta ainda não foi varrida."
                End If
                Dim base As String
                Select Case Cobertura
                    Case FolderCoverage.Completa
                        base = Nothing
                    Case FolderCoverage.Parcial
                        base = "Acervo parcial: o Outlook não expõe tudo o que existe no servidor. " &
                               "Pode faltar mensagem, e mensagem marcada como suspeita pode já ter sido apagada."
                    Case Else
                        base = "Alcance desconhecido: não dá para dizer o quanto desta pasta o Iris enxerga."
                End Select

                ' A contração vem PRIMEIRO quando existe: é a informação mais
                ' recente e a mais acionável — alguma coisa mudou agora.
                Dim aviso = Contracao?.Aviso
                If aviso Is Nothing Then Return base
                Return If(base Is Nothing, aviso, aviso & " " & base)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Lê o manifesto publicado de uma pasta.
    '''
    ''' Devolve as associações que já pertencem a alguma geração publicada.
    '''
    ''' Isso <b>não</b> é o mesmo que "o retrato exato da última geração": uma
    ''' associação carrega a geração em que foi vista por último, e o filtro é
    ''' <c>generation_key IS NOT NULL</c>. O que a fronteira garante é que
    ''' linhas de uma tentativa que nunca publicou não aparecem — e desde que a
    ''' encenação parou de tocar o acervo, elas nem chegam a existir fora de
    ''' <c>scan_stage</c>.
    '''
    ''' <b>As duas consultas não estão no mesmo snapshot.</b> A cabeça pode
    ''' mudar entre a primeira e a segunda, e o manifesto sairia com a cobertura
    ''' de uma geração e os itens de outra. Não é grave hoje — publicação é
    ''' rara e a UI relê — mas é dívida escrita, não descuido.
    ''' </summary>
    Public NotInheritable Class ManifestReader

        Private ReadOnly _conn As SqliteConnection
        Private ReadOnly _detector As ContractionDetector

        Public Sub New(db As CacheDatabase)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _conn = db.Connection
            _detector = New ContractionDetector(db)
        End Sub

        Public Function Ler(folderKey As Long) As FolderManifest
            Dim geracao As Long? = Nothing
            Dim cobertura = FolderCoverage.Desconhecida
            Dim publicadaEm As String = Nothing

            Using cmd = _conn.CreateCommand()
                cmd.CommandText =
                    "SELECT g.generation_key, g.published_at, c.coverage " &
                    "FROM folder f " &
                    "JOIN generation g ON g.generation_key = f.published_generation_key " &
                    "LEFT JOIN coverage_observation c ON c.coverage_key = g.coverage_key " &
                    "WHERE f.folder_key = $f"
                cmd.Parameters.AddWithValue("$f", folderKey)
                Using rd = cmd.ExecuteReader()
                    If rd.Read() Then
                        geracao = rd.GetInt64(0)
                        publicadaEm = rd.GetString(1)
                        cobertura = DaCobertura(If(rd.IsDBNull(2), Nothing, rd.GetString(2)))
                    End If
                End Using
            End Using

            If geracao Is Nothing Then
                Return New FolderManifest(folderKey, Nothing, FolderCoverage.Desconhecida,
                                          Nothing, Nothing)
            End If

            ' O detector tem um consumidor: e este.
            Dim contracao = _detector.Comparar(folderKey)

            Dim itens As New List(Of ManifestItem)()
            Using cmd = _conn.CreateCommand()
                ' So associacoes cuja geracao ja foi PUBLICADA. Linhas
                ' encenadas de uma tentativa em curso tem generation_key nulo
                ' ate a publicacao, entao nao entram aqui.
                cmd.CommandText =
                    "SELECT i.provider_entry_id, m.subject, m.sender_name, m.received_at, " &
                    "       m.is_unread, a.presence " &
                    "FROM association a " &
                    "JOIN incarnation i ON i.item_key = a.item_key AND i.folder_key = a.folder_key " &
                    "LEFT JOIN metadata_observation m ON m.incarnation_key = i.incarnation_key " &
                    "WHERE a.folder_key = $f AND a.generation_key IS NOT NULL " &
                    "ORDER BY m.received_at DESC"
                cmd.Parameters.AddWithValue("$f", folderKey)
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        itens.Add(New ManifestItem(
                            rd.GetString(0),
                            If(rd.IsDBNull(1), Nothing, rd.GetString(1)),
                            If(rd.IsDBNull(2), Nothing, rd.GetString(2)),
                            If(rd.IsDBNull(3), Nothing, rd.GetString(3)),
                            If(rd.IsDBNull(4), CType(Nothing, Boolean?), rd.GetInt32(4) = 1),
                            DaPresenca(If(rd.IsDBNull(5), Nothing, rd.GetString(5)))))
                    End While
                End Using
            End Using

            Return New FolderManifest(folderKey, geracao, cobertura, publicadaEm, itens, contracao)
        End Function

        Private Shared Function DaCobertura(s As String) As FolderCoverage
            Select Case s
                Case "completa" : Return FolderCoverage.Completa
                Case "parcial" : Return FolderCoverage.Parcial
                Case Else : Return FolderCoverage.Desconhecida
            End Select
        End Function

        Private Shared Function DaPresenca(s As String) As PresenceState
            Select Case s
                Case "presente" : Return PresenceState.Presente
                Case "suspeito" : Return PresenceState.Suspeito
                Case "ausente_da_pasta" : Return PresenceState.AusenteDaPasta
                Case Else : Return PresenceState.NaoVerificado
            End Select
        End Function

    End Class

End Namespace

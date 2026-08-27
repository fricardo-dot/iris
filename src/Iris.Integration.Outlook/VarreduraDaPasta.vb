Imports System.Threading
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Model
Imports Iris.Sync

Namespace Global.Iris.Integration.Outlook

    ''' <summary>Por que uma varredura não aconteceu.</summary>
    Public Enum RecusaDeVarredura
        Nenhuma = 0
        ''' <summary>Ninguém escolheu pasta.</summary>
        SemPasta
        ''' <summary>O store da pasta não está entre os que o Outlook reportou.</summary>
        StoreDesconhecido
        ''' <summary>
        ''' O ambiente foi medido e <b>ninguém o autorizou</b>. Não é erro:
        ''' é a cerimônia ainda não feita.
        ''' </summary>
        AmbienteNaoAutorizado
        ''' <summary>O cache não estava disponível.</summary>
        SemCache
        ''' <summary>Explodiu no meio. A varredura descarta a tentativa.</summary>
        Falhou
    End Enum

    ''' <summary>O que a varredura fez, ou por que não fez.</summary>
    Public NotInheritable Class ResultadoDaVarredura

        Public ReadOnly Property Recusa As RecusaDeVarredura
        ''' <summary><c>Nothing</c> quando houve recusa antes de varrer.</summary>
        Public ReadOnly Property Varredura As SweepResult
        ''' <summary>A impressão digital do ambiente, quando houve uma.</summary>
        Public ReadOnly Property Ambiente As String
        ''' <summary>A chave do perfil, para a cerimônia poder ser apontada.</summary>
        Public ReadOnly Property ChaveDoAmbiente As Long
        ''' <summary>A pasta varrida, no cache. Zero quando não se chegou lá.</summary>
        Public ReadOnly Property Pasta As Long

        Friend Sub New(recusa As RecusaDeVarredura, varredura As SweepResult,
                       ambiente As String, chaveDoAmbiente As Long, pasta As Long)
            Me.Recusa = recusa
            Me.Varredura = varredura
            Me.Ambiente = If(ambiente, "")
            Me.ChaveDoAmbiente = chaveDoAmbiente
            Me.Pasta = pasta
        End Sub

        Public ReadOnly Property Ok As Boolean
            Get
                Return Recusa = RecusaDeVarredura.Nenhuma AndAlso Varredura IsNot Nothing
            End Get
        End Property
    End Class

    ''' <summary>
    ''' <b>Varre uma pasta do Outlook para dentro do cache.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>É ISTO QUE FALTAVA PARA O APLICATIVO VARRER</b>
    '''
    ''' As três peças existiam desde a Fase 2 — <see cref="OutlookSweepSource"/>,
    ''' <c>SweepRunner</c> e <c>SqliteSweepSink</c> — e nada as ligava. Só os
    ''' testes montavam a cadeia, semeando as chaves na mão. Em produção
    ''' ninguém chamava o <c>SweepRunner</c>, então <b>o cache só tinha o que
    ''' uma importação manual tivesse posto nele</b>, e a faixa do acervo
    ''' mostrava a pasta de chave 1 para sempre.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A ORDEM É A GARANTIA, E ELA COMEÇA NO AMBIENTE</b>
    '''
    ''' <list type="number">
    ''' <item>mede o ambiente pelo <c>StoreInfo</c> que o broker já devolve;</item>
    ''' <item>resolve o perfil no cache — que <b>nasce não autorizado</b>;</item>
    ''' <item><b>para</b> se ninguém autorizou. Este é o gate D2 em ação: sem
    ''' ele, "a allowlist é dado" seria uma frase sem consequência;</item>
    ''' <item>resolve a pasta, criando <c>store</c> e <c>folder</c> se preciso;</item>
    ''' <item>varre, com as capacidades que a política dá <b>àquele</b>
    ''' ambiente.</item>
    ''' </list>
    '''
    ''' O gate de abertura do próprio <c>SweepRunner</c> é estreito de
    ''' propósito — só recusa ambiente não identificado — e ambiente limitado
    ''' varre e publica parcial. Quem decide se vale varrer <b>aqui</b> é a
    ''' autorização, que é outra pergunta: a primeira é "o Iris sabe onde
    ''' está?", a segunda é "alguém deixou?".
    ''' </summary>
    Public NotInheritable Class VarreduraDaPasta

        ''' <summary>
        ''' Medido na §D5: com lote de 100 na Caixa de Entrada a latência
        ''' máxima por lote passou de 100 ms em execução repetida. 50 é o
        ''' tamanho que coube no orçamento.
        ''' </summary>
        Public Const TamanhoDoLote As Integer = 50

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _db As CacheDatabase
        Private ReadOnly _resolvedor As ResolvedorDoAcervo

        Public Sub New(broker As IOutlookBroker, db As CacheDatabase)
            If broker Is Nothing Then Throw New ArgumentNullException(NameOf(broker))
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _broker = broker
            _db = db
            _resolvedor = New ResolvedorDoAcervo(db)
        End Sub

        ''' <summary>
        ''' Varre <paramref name="pasta"/>, se o ambiente estiver autorizado.
        ''' </summary>
        ''' <param name="store">
        ''' O store a que a pasta pertence, como o broker o reportou. É de onde
        ''' sai a medição do ambiente — e por isso vem de fora: quem já
        ''' enumerou os stores não vai enumerá-los de novo por pasta.
        ''' </param>
        Public Function Executar(pasta As FolderKey, nome As String,
                                 store As StoreInfo,
                                 ct As CancellationToken) As ResultadoDaVarredura
            If pasta Is Nothing OrElse String.IsNullOrWhiteSpace(pasta.EntryId) Then
                Return Recusar(RecusaDeVarredura.SemPasta)
            End If
            If store Is Nothing OrElse String.IsNullOrWhiteSpace(store.StoreId) Then
                Return Recusar(RecusaDeVarredura.StoreDesconhecido)
            End If

            Dim impressao = AmbienteMedido.De(store)
            Dim perfil = _resolvedor.Ambiente(impressao)

            ' O GATE D2. Sem esta linha, "a allowlist do ambiente e dado" seria
            ' uma frase sem consequencia nenhuma: o perfil seria gravado, e
            ' varreria do mesmo jeito.
            If Not perfil.Permitido Then
                Return New ResultadoDaVarredura(RecusaDeVarredura.AmbienteNaoAutorizado,
                                                Nothing, perfil.Fingerprint, perfil.Chave, 0)
            End If

            Dim chaveDaPasta = _resolvedor.Pasta(pasta.StoreId, pasta.EntryId, nome)

            ' O universo: o que esta varredura afirma ter percorrido. A
            ' impressao digital do ambiente entra nele porque a mesma pasta,
            ' no mesmo Outlook, com outra janela, e outro universo -- e uma
            ' tentativa aberta num universo nao pode publicar noutro.
            Dim universo As New SweepUniverse(pasta.StoreId, pasta.EntryId, "todos",
                                              Nothing, 1, impressao.Value())

            Dim epoca = EpocaDe(chaveDaPasta)
            Dim tentativa = ProximaTentativa(chaveDaPasta)

            Try
                Dim fonte As New OutlookSweepSource(_broker, pasta, universo, epoca)
                Dim destino As New SqliteSweepSink(_db, chaveDaPasta, perfil.Chave)
                Dim corredor As New SweepRunner(fonte, destino, TamanhoDoLote)

                Dim r = corredor.Executar(universo, epoca, tentativa,
                                          EnvironmentPolicy.Capacidades(impressao), ct)

                Return New ResultadoDaVarredura(RecusaDeVarredura.Nenhuma, r,
                                                perfil.Fingerprint, perfil.Chave, chaveDaPasta)
            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                ' O SweepRunner ja descarta a tentativa em qualquer excecao.
                ' Aqui so se traduz para a tela, sem o texto da excecao: ele
                ' pode carregar assunto de mensagem, e a faixa nao e lugar
                ' para conteudo vazar por mensagem de erro.
                Return New ResultadoDaVarredura(RecusaDeVarredura.Falhou, Nothing,
                                                perfil.Fingerprint, perfil.Chave, chaveDaPasta)
            End Try
        End Function

        ' ==============================================================

        Private Shared Function Recusar(r As RecusaDeVarredura) As ResultadoDaVarredura
            Return New ResultadoDaVarredura(r, Nothing, Nothing, 0, 0)
        End Function

        ''' <summary>
        ''' A época de reconciliação da pasta. É o CAS que impede uma geração
        ''' velha sobrescrever uma nova.
        ''' </summary>
        Private Function EpocaDe(chaveDaPasta As Long) As Long
            Using cmd = _db.Connection.CreateCommand()
                cmd.CommandText = "SELECT reconcile_epoch FROM folder WHERE folder_key = $k"
                cmd.Parameters.AddWithValue("$k", chaveDaPasta)
                Dim v = cmd.ExecuteScalar()
                If v Is Nothing OrElse v Is DBNull.Value Then Return 0
                Return Convert.ToInt64(v)
            End Using
        End Function

        ''' <summary>
        ''' O número desta tentativa. Contar as que existem e somar um: o
        ''' número entra na tentativa gravada, e repetir um já usado tornaria
        ''' duas varreduras diferentes indistinguíveis no histórico.
        ''' </summary>
        Private Function ProximaTentativa(chaveDaPasta As Long) As Integer
            Using cmd = _db.Connection.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM scan_attempt WHERE folder_key = $k"
                cmd.Parameters.AddWithValue("$k", chaveDaPasta)
                Return Convert.ToInt32(cmd.ExecuteScalar()) + 1
            End Using
        End Function

    End Class

End Namespace

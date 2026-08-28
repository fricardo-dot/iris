Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Cache
Imports Microsoft.Data.Sqlite

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>O acervo de TODAS as pastas, alimentado pelo dreno.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO NASCE, E O QUE ELE CONSERTA</b>
    '''
    ''' A §26.2 diz que nenhuma UI exibe dado do cache sem passar pelo
    ''' <see cref="PublicationDrain"/>, e a <c>ArchitectureTests</c> proíbe a
    ''' apresentação de instanciar <see cref="ManifestReader"/> direto —
    ''' porque o contorno tentador é ler o manifesto e nunca descobrir que a
    ''' entrega parou.
    '''
    ''' A busca do acervo, entregue em 28/08/2026, fazia exatamente o
    ''' contorno: lia o <c>ManifestReader</c> de cada pasta e depois só
    ''' <em>consultava</em> a fila do dreno. A revisão externa foi explícita —
    ''' <b>consultar o estado do dreno não é passar por ele</b>.
    '''
    ''' O motivo de eu ter feito assim era real: o
    ''' <see cref="AcervoService"/> é de <b>uma</b> pasta, com
    ''' <c>Apontar</c>/<c>Atual</c>, e uma busca entre pastas não cabe nesse
    ''' formato. A resposta certa não era contornar; era este consumidor.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE MUDA NA PRÁTICA</b>
    '''
    ''' Antes, a busca lia o banco a cada pergunta e podia mostrar uma geração
    ''' que o painel do acervo ainda não tinha recebido — dois lugares da
    ''' mesma janela discordando sobre o que existe.
    '''
    ''' Agora ela lê deste cache em memória, que só muda quando o dreno
    ''' entrega. Se a entrega travar, a busca <b>congela junto com o painel</b>
    ''' — e diz isso. Ficar para trás junto é honesto; ficar na frente em
    ''' silêncio não era.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A PRIMEIRA CARGA É DIRETA, E PRECISA SER</b>
    '''
    ''' Na abertura não há entrega pendente: tudo o que foi publicado em
    ''' sessões anteriores já está drenado, e esperar por um dreno que não vem
    ''' deixaria a busca vazia para sempre.
    '''
    ''' Isso não é o contorno de novo. O contorno era <b>ignorar</b> o dreno
    ''' em regime; aqui ele é a única fonte de mudança depois da abertura, e o
    ''' <see cref="Recarregado"/> conta quantas vezes o estado mudou — para
    ''' que "a busca está velha" seja uma pergunta com resposta.
    ''' </summary>
    Public NotInheritable Class AcervoDeTodasAsPastas
        Implements IPublicationConsumer

        Private ReadOnly _db As CacheDatabase
        Private ReadOnly _conn As SqliteConnection
        Private ReadOnly _trava As New Object()

        Private _pastas As IReadOnlyList(Of PastaNoAcervo) = Array.Empty(Of PastaNoAcervo)()
        Private _recarregado As Integer

        Public Sub New(db As CacheDatabase)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _db = db
            _conn = db.Connection
            Recarregar()
        End Sub

        ''' <summary>
        ''' Quantas vezes o retrato foi refeito. A primeira carga conta como
        ''' uma — não existe estado "nunca carregado" para alguém confundir
        ''' com "não há nada".
        ''' </summary>
        Public ReadOnly Property Recarregado As Integer
            Get
                SyncLock _trava
                    Return _recarregado
                End SyncLock
            End Get
        End Property

        ''' <summary>O retrato corrente: cada pasta com o manifesto dela.</summary>
        Public ReadOnly Property Pastas As IReadOnlyList(Of PastaNoAcervo)
            Get
                SyncLock _trava
                    Return _pastas
                End SyncLock
            End Get
        End Property

        ''' <summary>
        ''' O dreno entregou uma geração. Reler tudo, e não só a pasta dela.
        '''
        ''' Reler o acervo inteiro a cada entrega parece desperdício e não é:
        ''' as gerações chegam raramente — uma varredura por clique — e um
        ''' retrato parcial abriria a porta para as pastas discordarem entre
        ''' si sobre de quando é o retrato.
        ''' </summary>
        Public Sub Receber(geracao As Long) Implements IPublicationConsumer.Receber
            Recarregar()
        End Sub

        ''' <summary>
        ''' Relê o retrato inteiro. Substitui, nunca acumula — a mesma
        ''' idempotência do <see cref="AcervoService"/>.
        ''' </summary>
        Public Sub Recarregar()
            Dim leitor As New ManifestReader(_db)
            Dim novo As New List(Of PastaNoAcervo)()

            For Each p In PastasBrutas()
                novo.Add(New PastaNoAcervo(p.Chave, p.Nome, leitor.Ler(p.Chave)))
            Next

            SyncLock _trava
                _pastas = novo
                _recarregado += 1
            End SyncLock
        End Sub

        Private Structure Bruta
            Public Chave As Long
            Public Nome As String
        End Structure

        Private Function PastasBrutas() As List(Of Bruta)
            Dim r As New List(Of Bruta)()
            Using cmd = _conn.CreateCommand()
                cmd.CommandText = "SELECT folder_key, name FROM folder ORDER BY name"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        r.Add(New Bruta With {
                            .Chave = rd.GetInt64(0),
                            .Nome = If(rd.IsDBNull(1), "", rd.GetString(1))})
                    End While
                End Using
            End Using
            Return r
        End Function

    End Class

    ''' <summary>Uma pasta e o manifesto dela, como o dreno os entregou.</summary>
    Public NotInheritable Class PastaNoAcervo
        Public ReadOnly Property Chave As Long
        Public ReadOnly Property Nome As String
        Public ReadOnly Property Manifesto As FolderManifest

        Friend Sub New(chave As Long, nome As String, manifesto As FolderManifest)
            Me.Chave = chave
            Me.Nome = If(nome, "")
            Me.Manifesto = manifesto
        End Sub
    End Class

    ''' <summary>
    ''' <b>Um consumidor que repassa a dois.</b>
    '''
    ''' O <see cref="PublicationDrain.Drenar"/> entrega a <b>um</b> consumidor,
    ''' e a partir de 28/08/2026 há dois interessados: o painel do acervo, que
    ''' mostra uma pasta, e o acervo de todas as pastas, que a busca lê.
    '''
    ''' Poderia haver dois drenos, e seria pior: cada um marcaria a geração
    ''' como entregue por conta própria, e o segundo nunca veria o que o
    ''' primeiro já drenou. <b>Um dreno, um consumidor, e o fan-out aqui.</b>
    '''
    ''' Se <b>qualquer</b> um dos dois falhar, a exceção sobe — e a cabeça da
    ''' fila trava, que é o comportamento que o dreno já tem de propósito.
    ''' Engolir a falha de um para agradar o outro faria a geração ser marcada
    ''' como entregue a quem não a recebeu.
    ''' </summary>
    Public NotInheritable Class ConsumidorComposto
        Implements IPublicationConsumer

        Private ReadOnly _partes As IPublicationConsumer()

        Public Sub New(ParamArray partes As IPublicationConsumer())
            _partes = If(partes, Array.Empty(Of IPublicationConsumer)()).
                      Where(Function(p) p IsNot Nothing).ToArray()
        End Sub

        Public Sub Receber(geracao As Long) Implements IPublicationConsumer.Receber
            For Each p In _partes
                p.Receber(geracao)
            Next
        End Sub

    End Class

End Namespace

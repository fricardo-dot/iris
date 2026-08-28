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
    ''' <b>O CONSTRUTOR NÃO LÊ NADA, E ISSO É O CONSERTO</b>
    '''
    ''' A primeira versão desta classe lia o manifesto no construtor, com a
    ''' justificativa de que <i>"na abertura não há entrega pendente"</i>.
    ''' <b>É falso</b>, e a revisão externa de 28/08/2026 nomeou o caso: uma
    ''' queda entre publicar e marcar drenada deixa publicação pendente
    ''' <b>persistida</b>. Na abertura seguinte, o construtor leria o manifesto
    ''' novo antes de o dreno entregar — que é o contorno de novo, reduzido à
    ''' abertura.
    '''
    ''' Agora ele nasce <b>vazio</b>. Quem o enche é <see cref="Receber"/>, ou
    ''' um <see cref="Recarregar"/> que o dono chame <b>depois</b> de drenar.
    ''' A ordem é do dono, e está escrita no <c>AcervoViewModel</c>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ELE AINDA NÃO GARANTE, E É PRECISO DIZER</b>
    '''
    ''' <see cref="Receber"/> ignora <i>qual</i> geração chegou e relê o
    ''' manifesto corrente. Com <b>duas</b> gerações pendentes — 10 e 11 — e o
    ''' manifesto já apontando para 11, a entrega da 10 torna a 11 visível
    ''' antes de a 11 ser entregue.
    '''
    ''' Na prática a janela é curta: o <c>Drenar</c> percorre a fila inteira em
    ''' ordem numa chamada só, então a inconsistência não sobrevive ao laço —
    ''' a menos que um consumidor falhe no meio. Mas <b>curta não é
    ''' inexistente</b>, e a propriedade "não vê o que o dreno não entregou"
    ''' vale para uma geração pendente, e não para uma fila.
    '''
    ''' Consertar isso pede o manifesto <b>de uma geração específica</b>, e o
    ''' <see cref="ManifestReader"/> lê a publicada da pasta. É trabalho de
    ''' desenho, e está no relatório.
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
            ' NAO LE AQUI. Ver o cabecalho: ler no construtor e ler na frente
            ' do dreno quando ha publicacao pendente de uma queda anterior.
        End Sub

        ''' <summary>
        ''' Quantas vezes o retrato foi refeito. <b>Zero quer dizer que ninguém
        ''' carregou ainda</b> — e isso é diferente de "não há nada guardado".
        ''' Quem lê e vê zero está lendo antes do dreno.
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
        ''' as gerações chegam raramente — uma varredura por clique.
        '''
        ''' <b>Mas ele não produz um retrato de um instante</b>, e eu escrevi
        ''' que produzia. As leituras de manifesto não estão numa transação
        ''' única, então duas pastas podem refletir instantes diferentes. Reler
        ''' tudo <i>reduz</i> a divergência; não a elimina. E o custo de reler
        ''' todas as pastas nunca foi medido — com poucas pastas é irrelevante,
        ''' e "poucas" não é um número que alguém tenha conferido.
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
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ISTO NÃO É: ENTREGA ATÔMICA</b>
    '''
    ''' As partes recebem <b>em sequência</b>. Se a primeira concluir e a
    ''' segunda falhar, o painel muda e a busca fica para trás — e eu escrevi,
    ''' na primeira versão, que as duas <i>"congelam juntas"</i>. Não congelam.
    '''
    ''' O que salva é a semântica do dreno: a geração continua pendente e será
    ''' repetida, e as duas partes são idempotentes. Então a divergência é
    ''' <b>temporária</b>, e não perda silenciosa. Chamar isso de simultâneo
    ''' era mais forte que o mecanismo.
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

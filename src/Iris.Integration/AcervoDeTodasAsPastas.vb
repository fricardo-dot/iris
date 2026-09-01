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
    ''' entrega. Se a entrega travar antes de <b>qualquer</b> uma das duas
    ''' receber, as duas ficam no retrato anterior — e a busca diz isso. Ficar
    ''' para trás junto é honesto; ficar na frente em silêncio não era.
    '''
    ''' <b>"Junto" tem um limite, e ele está no <see cref="ConsumidorComposto"/>:</b>
    ''' as entregas são sequenciais, então uma falha <i>entre</i> elas deixa o
    ''' painel à frente da busca.
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
            ' A LEITURA INTEIRA SOB TRAVA, e nao so a troca do retrato.
            '
            ' Duas recargas ao mesmo tempo -- o dreno entregando enquanto alguem
            ' recarrega a mao -- liam a MESMA SqliteConnection em paralelo, que e
            ' o que esta classe pede a quem a usa que nao faca. E, mesmo sem
            ' estourar, a recarga velha podia terminar depois da nova e substituir
            ' o retrato por um mais antigo -- com _recarregado subindo, o que faz
            ' a regressao parecer progresso.
            '
            ' A TRAVA E A DA CONEXAO, e nao uma so desta classe.
            '
            ' Uma trava propria serializava duas recargas entre si e nao servia de
            ' nada contra os outros caminhos que tocam a MESMA conexao -- a leitura
            ' de rotulos e a gravacao em lote. Tres travas diferentes protegendo o
            ' mesmo recurso e o mesmo que nenhuma. Achado por revisao externa em
            ' 01/09/2026.
            '
            ' Ela nao e a mesma que serve o getter de Pastas: manter a leitura do
            ' banco dentro daquela faria a tela esperar a varredura terminar para
            ' conseguir desenhar a lista.
            SyncLock _db.Trava
                Reler()
            End SyncLock
        End Sub

        Private Sub Reler()
            Dim leitor As New ManifestReader(_db)
            Dim novo As New List(Of PastaNoAcervo)()

            For Each p In PastasBrutas()
                novo.Add(New PastaNoAcervo(p.Chave, p.Nome, leitor.Ler(p.Chave), p.Store))
            Next

            SyncLock _trava
                _pastas = novo
                _recarregado += 1
            End SyncLock
        End Sub

        Private Structure Bruta
            Public Chave As Long
            Public Nome As String
            Public Store As String
        End Structure

        Private Function PastasBrutas() As List(Of Bruta)
            Dim r As New List(Of Bruta)()
            Using cmd = _conn.CreateCommand()
                ' O STORE VEM JUNTO. Sem ele, quem monta um ItemKey a partir do
                ' acervo tem o EntryID e nao tem onde abri-lo -- e ItemKey sem
                ' store nao identifica mensagem nenhuma fora de uma caixa so.
                cmd.CommandText =
                    "SELECT f.folder_key, f.name, s.provider_store_id " &
                    "FROM folder f JOIN store s ON s.store_key = f.store_key " &
                    "ORDER BY f.name"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        r.Add(New Bruta With {
                            .Chave = rd.GetInt64(0),
                            .Nome = If(rd.IsDBNull(1), "", rd.GetString(1)),
                            .Store = If(rd.IsDBNull(2), "", rd.GetString(2))})
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

        ''' <summary>
        ''' O <c>StoreID</c> do provedor. Existe porque <see cref="ItemKey"/>
        ''' precisa dele: EntryID sozinho não identifica mensagem fora de uma
        ''' caixa só, e a tela abre pela chave inteira.
        ''' </summary>
        Public ReadOnly Property Store As String

        Friend Sub New(chave As Long, nome As String, manifesto As FolderManifest,
                       Optional store As String = Nothing)
            Me.Chave = chave
            Me.Nome = If(nome, "")
            Me.Manifesto = manifesto
            Me.Store = If(store, "")
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
    ''' O que o dreno garante é que a geração <b>continua pendente</b> e será
    ''' tentada de novo, e que as duas partes são idempotentes. Isso é
    ''' possibilidade de convergência, e <b>não</b> convergência: uma falha
    ''' persistente mantém a divergência enquanto durar.
    '''
    ''' Eu escrevi "temporária" na correção anterior, e a revisão externa
    ''' apontou que continua sendo mais forte que o mecanismo. O que se pode
    ''' afirmar é: <b>nada se perde em silêncio</b> — a pendência fica no
    ''' banco, e a busca a anuncia.
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

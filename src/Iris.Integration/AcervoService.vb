Imports System.Threading
Imports Iris.Cache

Namespace Global.Iris.Integration

    ''' <summary>
    ''' O acervo de uma pasta, como a aplicação o vê — e o consumidor que o
    ''' mantém em dia.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO EXISTE, E NÃO SÓ UM CONSUMIDOR NA UI</b>
    '''
    ''' A condição do 2.4 (§26.2) pede um teste que mate o processo entre
    ''' <c>Receber</c> e <c>MarcarDrenada</c> e prove que, na reabertura, o que
    ''' a UI mostra converge sem perda.
    '''
    ''' Se o consumidor vivesse dentro do ViewModel, a prova teria de rodar
    ''' contra WPF — ou contra uma imitação dele, e aí provaria a imitação. É o
    ''' erro que a Q1 cobrou quando o teste sintético verificava um algoritmo
    ''' diferente do que rodava contra o Outlook.
    '''
    ''' Então a lógica de convergência mora aqui, fora da UI: o ViewModel
    ''' observa este serviço, e o harness de crash usa <b>este mesmo</b>
    ''' serviço. A prova é sobre o código que roda.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>IDEMPOTÊNCIA</b>
    '''
    ''' A entrega é <i>ao menos uma vez</i> (§24.3), então este consumidor
    ''' recebe a mesma geração mais de uma vez sempre que o processo morre no
    ''' meio do dreno. Ele é idempotente pela forma mais simples que existe:
    ''' <b>não acumula nada</b>. Receber uma geração significa reler o
    ''' manifesto inteiro e substituir o que estava.
    '''
    ''' Recontar, somar ou anexar exigiria deduplicação, e deduplicação exige
    ''' lembrar o que já veio — que é estado novo, com o seu próprio problema de
    ''' durabilidade. Reler é mais caro e não tem esse problema.
    ''' </summary>
    Public NotInheritable Class AcervoService
        Implements IPublicationConsumer

        Private ReadOnly _db As CacheDatabase
        ''' <summary>
        ''' A pasta que este serviço está mostrando. <b>Muda</b>: era
        ''' <c>ReadOnly</c>, e por isso o acervo ficava preso à pasta escolhida
        ''' na construção — que em produção era a constante 1, e não a que o
        ''' usuário estava olhando.
        '''
        ''' Guardada sob a mesma trava do manifesto: quem lê o manifesto tem de
        ''' ver a pasta a que ele pertence, e não a de meio segundo atrás.
        ''' </summary>
        Private _folderKey As Long
        Private ReadOnly _trava As New Object()
        Private _atual As FolderManifest
        Private _recebidas As Integer

        ''' <summary>Disparado quando o acervo muda. A UI se pendura aqui.</summary>
        Public Event Mudou As EventHandler

        Public Sub New(db As CacheDatabase, folderKey As Long)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _db = db
            _folderKey = folderKey
            Recarregar()
        End Sub

        ''' <summary>O manifesto corrente. Nunca <c>Nothing</c>.</summary>
        Public ReadOnly Property Atual As FolderManifest
            Get
                SyncLock _trava
                    Return _atual
                End SyncLock
            End Get
        End Property

        ''' <summary>
        ''' Quantas entregas este consumidor recebeu. Inclui as repetidas —
        ''' e é justamente isso que o teste de crash observa.
        ''' </summary>
        Public ReadOnly Property Recebidas As Integer
            Get
                Return Volatile.Read(_recebidas)
            End Get
        End Property

        Public Sub Receber(geracao As Long) Implements IPublicationConsumer.Receber
            Interlocked.Increment(_recebidas)
            Recarregar()
            RaiseEvent Mudou(Me, EventArgs.Empty)
        End Sub

        ''' <summary>
        ''' <b>Passa a mostrar OUTRA pasta.</b>
        '''
        ''' ------------------------------------------------------------------
        ''' Chamado quando o usuário troca de pasta na árvore. Sem isto o acervo
        ''' ficava na pasta de chave 1 — a que uma importação manual tinha
        ''' criado — enquanto a lista ao lado mostrava outra coisa. Dois painéis
        ''' falando de pastas diferentes, sem nada dizendo isso.
        '''
        ''' Trocar para a mesma pasta não é trabalho perdido: relê. O manifesto
        ''' pode ter mudado por uma varredura que rodou no meio.
        ''' </summary>
        Public Sub Apontar(folderKey As Long)
            SyncLock _trava
                _folderKey = folderKey
            End SyncLock
            Recarregar()
            RaiseEvent Mudou(Me, EventArgs.Empty)
        End Sub

        ''' <summary>A pasta que este serviço está mostrando.</summary>
        Public ReadOnly Property Alvo As Long
            Get
                SyncLock _trava
                    Return _folderKey
                End SyncLock
            End Get
        End Property

        ''' <summary>
        ''' Relê o manifesto inteiro. É a idempotência: substituir, nunca
        ''' acumular.
        ''' </summary>
        Public Sub Recarregar()
            SyncLock _db.Trava
                ' A chave e lida SOB A TRAVA e usada fora: ler direto no argumento
                ' deixaria a leitura do manifesto correr com uma chave que pode
                ' mudar no meio, e o manifesto resultante seria atribuido como se
                ' fosse da pasta nova.
                Dim qual As Long
                SyncLock _trava
                    qual = _folderKey
            End SyncLock

            Dim novo = New ManifestReader(_db).Ler(qual)

            SyncLock _trava
                ' Se a pasta mudou enquanto lia, este manifesto e de uma pasta
                ' que ja nao e a mostrada. Descartar e o certo: quem trocou
                ' chamou Recarregar de novo.
                If qual <> _folderKey Then Return
                _atual = novo
            End SyncLock
            End SyncLock
        End Sub

    End Class

End Namespace

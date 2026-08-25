Imports System.Collections.Generic
Imports Iris.Cache

Namespace Global.Iris.Integration

    ''' <summary>
    ''' Quem consome uma geração publicada. A UI vai implementar isto — hoje
    ''' só o teste implementa, e é honesto dizer que o mecanismo existe e o
    ''' consumidor da UI ainda não.
    '''
    ''' <b>Precisa ser idempotente</b>, e não é zelo: a entrega é
    ''' <i>ao menos uma vez</i>. Ver <see cref="PublicationDrain"/>.
    ''' </summary>
    Public Interface IPublicationConsumer
        Sub Receber(geracao As Long)
    End Interface

    ''' <summary>
    ''' Drena as gerações publicadas que a UI ainda não consumiu.
    '''
    ''' É a metade que faltava do critério 9. O 2.1 provou que a dívida
    ''' <b>persiste</b> e é consultável depois de o processo morrer; o que não
    ''' existia era ninguém para consumi-la. Sem consumidor, "publicar para a
    ''' UI" estava entregue pela metade — e o relatório dizia isso.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ENTREGA AO MENOS UMA VEZ, e a ordem das duas linhas é a escolha</b>
    '''
    ''' Consumir vem antes de marcar drenada. Morrer entre as duas repete a
    ''' entrega na próxima abertura.
    '''
    ''' A ordem inversa — marcar e depois consumir — daria <i>no máximo uma
    ''' vez</i>: morrer no meio perderia a geração para sempre, e a UI ficaria
    ''' mostrando estado velho sem nada no disco indicando que faltou algo.
    '''
    ''' Entre repetir e perder, repetir é recuperável. Mas isso <b>transfere
    ''' uma obrigação</b>: o consumidor tem de ser idempotente. É contrato
    ''' dele, está escrito na interface, e não há como este arquivo garantir.
    ''' </summary>
    Public NotInheritable Class PublicationDrain

        Private ReadOnly _writer As CacheWriter

        Public Sub New(db As CacheDatabase)
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            _writer = New CacheWriter(db)
        End Sub

        Public Function Pendentes() As IReadOnlyList(Of Long)
            Return _writer.PublicacoesPendentes()
        End Function

        ''' <summary>
        ''' Entrega as pendentes ao consumidor, na ordem em que foram emitidas,
        ''' e devolve quantas foram drenadas.
        '''
        ''' Consumidor que lança INTERROMPE o dreno e deixa o resto pendente —
        ''' inclusive a que falhou. Engolir a exceção e seguir marcaria como
        ''' drenada uma geração que a UI não recebeu, que é perder em silêncio
        ''' o que este desenho inteiro existe para não perder.
        ''' </summary>
        Public Function Drenar(consumidor As IPublicationConsumer) As Integer
            If consumidor Is Nothing Then Throw New ArgumentNullException(NameOf(consumidor))

            Dim n = 0
            For Each g In _writer.PublicacoesPendentes()
                Try
                    consumidor.Receber(g)
                Catch ex As Exception
                    ' A falha fica REGISTRADA antes de propagar. Sem isso, um
                    ' consumidor que falha sempre trava a fila em silencio: as
                    ' geracoes seguintes nunca chegam e nada no banco diz por
                    ' que. Ver TravadoEm.
                    _writer.RegistrarFalhaNaEntrega(g, $"{ex.GetType().Name}: {ex.Message}")
                    Throw
                End Try
                ' A janela que torna a entrega AO MENOS UMA VEZ. Morrer aqui
                ' deixa a UI ja tendo agido e o disco ainda dizendo que ela nao
                ' recebeu — e o teste do 2.4 mata o processo exatamente aqui.
                CrashInjection.Talvez(CrashInjection.DepoisDeReceberAntesDeMarcarDrenada)

                _writer.MarcarDrenada(g)
                n += 1
            Next
            Return n
        End Function

        ''' <summary>
        ''' A geração em que a fila travou, se travou — a que já falhou
        ''' <paramref name="limite"/> vezes ou mais e continua na cabeça.
        '''
        ''' <b>O bloqueio é deliberado e continua.</b> Marcar como drenada uma
        ''' geração que a UI não recebeu seria perder em silêncio o que este
        ''' desenho existe para não perder, e pular a cabeça entregaria as
        ''' seguintes fora de ordem. O que este método muda não é o bloqueio: é
        ''' ele deixar de ser <b>invisível</b>.
        '''
        ''' Quem chama decide o que fazer — avisar o usuário, registrar, parar
        ''' de tentar. O que não dá é ninguém saber que a fila parou.
        ''' </summary>
        Public Function TravadoEm(Optional limite As Integer = 3) As Long?
            Dim pendentes = _writer.PublicacoesPendentes()
            If pendentes.Count = 0 Then Return Nothing
            Dim cabeca = pendentes(0)
            If _writer.TentativasDeEntrega(cabeca) >= limite Then Return cabeca
            Return Nothing
        End Function

        Public Function UltimoErro(geracao As Long) As String
            Return _writer.UltimoErroDeEntrega(geracao)
        End Function

    End Class

End Namespace

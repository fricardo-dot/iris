Imports System.Collections.Generic
Imports Iris.Cache

Namespace Global.Iris.Integration

    ''' <summary>
    ''' Quem consome uma geração publicada. A UI implementa isto.
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
                consumidor.Receber(g)
                _writer.MarcarDrenada(g)
                n += 1
            Next
            Return n
        End Function

    End Class

End Namespace

Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports Iris.Core
Imports Iris.Model
Imports Iris.Sync

Namespace Global.Iris.Integration.Outlook

    ''' <summary>
    ''' O <see cref="ISweepSource"/> real, sobre o broker do Outlook.
    '''
    ''' É a única peça do caminho de varredura que toca COM, e ela é fina de
    ''' propósito: paginar já é problema resolvido pelo
    ''' <c>GetMessagePageAsync</c>, que a Q1 mediu e a Fase 1 endureceu. Aqui
    ''' só há tradução e uma decisão.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A DECISÃO: de onde vem a contagem</b>
    '''
    ''' O S6 compara o que foi lido com o que a pasta declarava ter, antes e
    ''' depois. Mas <c>MessagePage.TotalAtStart</c> <b>só vem na primeira
    ''' página</b> — o caminho por <c>Table</c> não tem <c>Count</c> barato, e
    ''' pedir a contagem a cada página gastaria uma chamada COM na fila única
    ''' da STA para reafirmar um número já desatualizado quando chegasse.
    '''
    ''' Então <see cref="Contar"/> abre uma travessia de uma página só, lê o
    ''' total e descarta o resto. Custa uma ida ao provider por contagem, duas
    ''' por varredura. É o preço do S6, e ele é o que separa "li tudo" de
    ''' "achei que li tudo".
    ''' </summary>
    Public NotInheritable Class OutlookSweepSource
        Implements ISweepSource

        Private ReadOnly _broker As IOutlookBroker
        Private ReadOnly _pasta As FolderKey
        Private ReadOnly _universo As SweepUniverse
        Private ReadOnly _geracaoDaConsulta As Long

        Public Sub New(broker As IOutlookBroker, pasta As FolderKey,
                       universo As SweepUniverse, geracaoDaConsulta As Long)
            If broker Is Nothing Then Throw New ArgumentNullException(NameOf(broker))
            _broker = broker
            _pasta = pasta
            _universo = universo
            _geracaoDaConsulta = geracaoDaConsulta
        End Sub

        ''' <summary>
        ''' Quantas idas ao provider a varredura custou. Medido para a D5, e
        ''' não estimado.
        ''' </summary>
        Public Property IdasAoProvider As Integer = 0

        ''' <summary>
        ''' Linhas que a paginacao DESCARTOU por nao serem mensagem ou por
        ''' erro de leitura.
        '''
        ''' Isto importa para o S6 e nao e detalhe: a contagem da pasta inclui
        ''' o que a paginacao pula, entao "lidas < contadas" pode ser corrida
        ''' (chegou correio) ou pode ser descarte SISTEMATICO. Se for
        ''' sistematico, o S6 rejeita a pasta SEMPRE e ela nunca publica - e
        ''' isso pareceria "caixa viva" sem ser.
        ''' </summary>
        Public Property Descartadas As Integer = 0

        Public Function Contar(ct As CancellationToken) As SourceCount _
                              Implements ISweepSource.Contar
            Dim p = Pedir(Nothing, 1, ct)
            If Not p.TotalAtStart.HasValue Then
                ' Sem total nao ha S6, e sem S6 nao ha como distinguir "li
                ' tudo" de "a fonte parou de devolver". Falhar aqui e melhor
                ' que publicar sem a guarda.
                Throw New InvalidOperationException(
                    "a pagina nao trouxe TotalAtStart: sem contagem nao ha S6")
            End If
            Return New SourceCount(p.TotalAtStart.Value, _universo)
        End Function

        Public Function LerPagina(cursor As String, tamanho As Integer,
                                  ct As CancellationToken) As SourcePage _
                                  Implements ISweepSource.LerPagina
            Dim p = Pedir(cursor, tamanho, ct)
            Descartadas += p.SkippedCount
            Return New SourcePage(
                p.Items.Select(AddressOf Traduzir).ToList(),
                p.NextCursor,
                fim:=(p.NextCursor Is Nothing),
                universo:=_universo,
                drenadoAlem:=p.DrainedExtra,
                descartadas:=p.SkippedCount)
        End Function

        ' ==============================================================

        Private Function Pedir(cursor As String, alvo As Integer,
                               ct As CancellationToken) As MessagePage
            IdasAoProvider += 1
            Dim q As New MessageQuery(_pasta, MessageSort.ReceivedDesc, _geracaoDaConsulta)
            Dim r = _broker.GetMessagePageAsync(q, cursor, alvo, ct).GetAwaiter().GetResult()
            If Not r.Succeeded Then
                Throw New InvalidOperationException(
                    $"GetMessagePageAsync falhou: {r.Kind} {r.Detail}")
            End If
            Return r.Value
        End Function

        ''' <summary>
        ''' <c>MailSummary</c> para <c>SourceRow</c>.
        '''
        ''' <c>SearchKey</c> e <c>InternetMessageId</c> ficam vazios, e a
        ''' ausência é deliberada: a Q1 mediu que nenhum dos dois vem por
        ''' coluna de <c>Table</c>, e o caminho de listagem lê por Table.
        ''' Preenchê-los exigiria abrir o <c>MailItem</c> de cada mensagem —
        ''' uma ida ao COM por linha, que é exatamente o custo que a paginação
        ''' por Table existe para evitar.
        '''
        ''' <c>ReceivedAt</c> vai em ISO 8601 com offset. "DateTime sem Kind" é
        ''' a origem clássica de mensagem aparecendo com hora errada depois de
        ''' ler do cache, e persistir é justamente o que vai acontecer com
        ''' isto.
        ''' </summary>
        Private Shared Function Traduzir(m As MailSummary) As SourceRow
            Return New SourceRow With {
                .Key = m.Key.EntryId,
                .Subject = m.Subject,
                .SenderName = m.SenderName,
                .ReceivedAt = If(m.ReceivedTime.HasValue, m.ReceivedTime.Value.ToString("o"), Nothing),
                .SizeBytes = CType(m.SizeBytes, Long?),
                .HasAttachments = m.HasAttachments,
                .IsUnread = m.IsUnread,
                .MessageClass = "IPM.Note"}
        End Function

    End Class

End Namespace

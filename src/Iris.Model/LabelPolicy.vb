Namespace Global.Iris.Model

    ''' <summary>
    ''' Sobre quais desfechos de leitura uma política <b>pode</b> decidir.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE LISTA EXPLÍCITA NÃO BASTA</b>
    '''
    ''' O portão da Fase 3 exige que cada <see cref="LabelReadingKind"/> aceito
    ''' esteja listado <b>pelo nome</b> na autorização. Isso resolve o
    ''' esquecimento — estado novo nega até alguém decidir —, e não resolve o
    ''' outro caso: alguém listar um estado que <b>não é prova de nada</b>.
    '''
    ''' <c>Denied</c> quer dizer que a leitura foi negada. <c>Unreadable</c>,
    ''' que não deu para ler. <c>Unknown</c> é o zero do enum, que aparece em
    ''' campo esquecido e desserialização incompleta. Nenhum deles descreve o
    ''' item: descrevem o <b>fracasso de descrevê-lo</b>.
    '''
    ''' Uma cerimônia de ativação que os listasse estaria autorizando com base
    ''' na ausência de informação. Nenhuma assinatura torna isso uma prova
    ''' positiva, então a proibição é <b>estrutural</b>, e não configurável.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E OS QUATRO QUE SOBRAM</b>
    '''
    ''' <c>Present</c>, <c>Absent</c>, <c>Blank</c> e <c>HistoricalOnly</c>
    ''' descrevem o item. Que <b>signifiquem</b> permissão continua sendo
    ''' decisão da política corporativa — a P14 e a P16 seguem abertas, e nada
    ''' aqui autoriza. O que este módulo diz é só que sobre esses quatro dá
    ''' para <b>decidir</b>.
    ''' </summary>
    Public Module LabelPolicy

        ''' <summary>
        ''' Uma política pode decidir sobre este desfecho?
        '''
        ''' <c>False</c> significa "isto não é informação sobre o item", e
        ''' nenhuma autorização pode mudar isso.
        ''' </summary>
        Public Function Elegivel(kind As LabelReadingKind) As Boolean
            Select Case kind
                Case LabelReadingKind.Present,
                     LabelReadingKind.Absent,
                     LabelReadingKind.Blank,
                     LabelReadingKind.HistoricalOnly
                    Return True
                Case Else
                    ' Unknown, Malformed, Conflicting, Unsupported, Denied,
                    ' Restricted, Transient, NotDownloaded, Changed, Unreadable
                    ' — e qualquer membro que apareca depois, ate alguem vir
                    ' aqui e decidir.
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' A forma da leitura bate com o desfecho que ela declara?
        '''
        ''' ------------------------------------------------------------------
        ''' <b>POR QUE O PORTÃO NÃO PODE CONFIAR NO CAMPO</b>
        '''
        ''' <see cref="LabelReading"/> é um DTO público, e nada impede alguém
        ''' de montar um <c>Present</c> <b>sem registro ativo nenhum</b> ou um
        ''' <c>Absent</c> <b>com</b> registro ativo. As duas combinações são
        ''' impossíveis de verdade — o parser nunca as produz —, e as duas
        ''' passariam por um portão que olhasse só o <c>Kind</c>.
        '''
        ''' A primeira é a mais perigosa: um <c>Absent</c> forjado carregando
        ''' um rótulo restritivo entraria pela porta de "sem rótulo".
        ''' </summary>
        Public Function Coerente(leitura As LabelReading) As Boolean
            If leitura Is Nothing Then Return False

            Dim ativos = 0
            Dim total = 0
            Dim declarados = 0
            For Each r In leitura.Registros
                total += 1
                If r.Enabled.HasValue Then declarados += 1
                If r.Ativo Then ativos += 1
            Next

            Select Case leitura.Kind
                Case LabelReadingKind.Present
                    ' Exatamente um ativo. Zero seria "sem rotulo" disfarcado;
                    ' mais de um seria Conflicting.
                    Return ativos = 1
                Case LabelReadingKind.Absent, LabelReadingKind.Blank
                    ' A propriedade nao existe, ou existe vazia. Nos dois casos
                    ' nao ha registro para existir.
                    Return total = 0
                Case LabelReadingKind.HistoricalOnly
                    ' Ha registro, e todos se declaram desligados.
                    Return total > 0 AndAlso ativos = 0 AndAlso declarados > 0
                Case Else
                    ' Desfecho nao elegivel nao chega aqui, e se chegar nao
                    ' passa.
                    Return False
            End Select
        End Function

    End Module

End Namespace

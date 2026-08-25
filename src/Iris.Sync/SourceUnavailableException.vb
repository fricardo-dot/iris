Imports Iris.Model

Namespace Global.Iris.Sync

    ''' <summary>
    ''' A <b>fonte</b> recusou a chamada, e o <see cref="ErrorKind"/> vem junto.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE O TIPO EXISTE</b>
    '''
    ''' O <see cref="SweepRunner"/> descarta a tentativa em qualquer exceção, e
    ''' isso está certo: nunca publica metade. Mas quem revisa depois precisa
    ''' saber se a varredura morreu porque o <b>Outlook</b> recusou ou porque o
    ''' <b>Iris</b> tem defeito — e essas duas coisas não podem ser separadas
    ''' pelo texto da mensagem.
    '''
    ''' A primeira tentativa fez exatamente isso: um teste de integração aceitava
    ''' <c>Falhou</c> quando o motivo continha certas substrings. O Codex
    ''' derrubou mostrando que <c>"GetMessagePageAsync falhou: {Kind}"</c> cobre
    ''' <see cref="ErrorKind.Unexpected"/>, <see cref="ErrorKind.Stale"/>,
    ''' <see cref="ErrorKind.Denied"/> e <see cref="ErrorKind.NotImplemented"/>
    ''' junto — ou seja, uma regressão minha passaria de soluço do ambiente.
    '''
    ''' Aqui a causa é <b>estruturada</b>. Quem decide o que tolerar decide
    ''' sobre o enum, não sobre uma frase.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE NÃO PASSA POR AQUI</b>
    '''
    ''' Falha de contrato da fonte — página sem <c>TotalAtStart</c>, cursor que
    ''' não anda, linha a mais — <b>não é</b> indisponibilidade. É defeito, e
    ''' continua saindo como exceção comum, sem classificação que a desculpe.
    ''' </summary>
    Public NotInheritable Class SourceUnavailableException
        Inherits Exception

        Public ReadOnly Property Kind As ErrorKind

        Public Sub New(kind As ErrorKind, detalhe As String)
            MyBase.New($"a fonte recusou ({kind}): {detalhe}")
            Me.Kind = kind
        End Sub

        ''' <summary>
        ''' O ambiente pode ter produzido esta recusa sozinho.
        '''
        ''' Delega ao <see cref="ErrorPolicy.Transitorio"/>, que é o único lugar
        ''' onde a lista existe. Já foi uma cópia escrita à mão aqui, e havia
        ''' outras duas — a divergência entre cópias de uma regra de tolerância
        ''' só aparece como intermitência, meses depois.
        ''' </summary>
        Public ReadOnly Property DoAmbiente As Boolean
            Get
                Return ErrorPolicy.Transitorio(Kind)
            End Get
        End Property

    End Class

End Namespace

Imports System.Collections.Generic
Imports Iris.Assist

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' O diário quando <b>não há banco</b> — e ele recusa tudo.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO EXISTE EM VEZ DE <c>Nothing</c></b>
    '''
    ''' O cache pode não abrir, e o <c>AssistTransmitter</c> precisa de um
    ''' diário. Passar <c>Nothing</c> viraria <c>NullReferenceException</c> no
    ''' primeiro passo — e "explodiu" e "recusou por decisão" não são a mesma
    ''' coisa para quem lê depois.
    '''
    ''' Recusando, o transmissor para no passo 3 com <c>SemDiario</c>, que é
    ''' exatamente o que aconteceu: não havia onde registrar, e por isso nada
    ''' foi tentado. Transmitir sem poder registrar seria pior que não
    ''' transmitir.
    ''' </summary>
    Friend NotInheritable Class DiarioAusente
        Implements IDisclosureJournal

        Public Function Intencao(c As DisclosureCapability, quando As DateTimeOffset) _
                                 As Boolean Implements IDisclosureJournal.Intencao
            Return False
        End Function

        Public Function Iniciando(requestId As Guid, quando As DateTimeOffset) As Boolean _
                                  Implements IDisclosureJournal.Iniciando
            Return False
        End Function

        Public Function Concluir(requestId As Guid, quando As DateTimeOffset) As Boolean _
                                 Implements IDisclosureJournal.Concluir
            Return False
        End Function

        Public Function Falhar(requestId As Guid, quando As DateTimeOffset,
                               nota As DisclosureNote, podeTerChegado As Boolean,
                               Optional codigoHttp As Integer? = Nothing) As Boolean _
                               Implements IDisclosureJournal.Falhar
            Return False
        End Function

        Public Function NaoEnviou(requestId As Guid, quando As DateTimeOffset,
                                  nota As DisclosureNote,
                                  Optional motivo As DisclosureReason =
                                      DisclosureReason.NaoDecidido) As Boolean _
                                  Implements IDisclosureJournal.NaoEnviou
            Return False
        End Function

        Public Function Reconciliar(quando As DateTimeOffset) As Integer _
                                    Implements IDisclosureJournal.Reconciliar
            Return 0
        End Function

        Public Function Ler(quantas As Integer) As IReadOnlyList(Of DisclosureEntry) _
                            Implements IDisclosureJournal.Ler
            Return Array.Empty(Of DisclosureEntry)()
        End Function

    End Class

End Namespace

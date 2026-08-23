Imports CommunityToolkit.Mvvm.ComponentModel
Imports Iris.Model

Namespace Global.Iris.App.ViewModels

    ''' <summary>
    ''' Uma linha da lista.
    '''
    ''' Existe por um motivo concreto: <see cref="MailSummary"/> é um DTO
    ''' puro, sem notificação de mudança — e tem que continuar assim, porque
    ''' ele mora em Iris.Model, que não conhece WPF nem MVVM.
    '''
    ''' Sem este envelope, marcar uma mensagem como lida alteraria o DTO e a
    ''' linha continuaria em negrito na tela. O usuário abriria a mensagem e
    ''' veria o Iris insistindo que ela é nova.
    '''
    ''' Mantido raso de propósito: cada propriedade aqui é multiplicada por
    ''' cada linha visível.
    ''' </summary>
    Public NotInheritable Class MessageRowViewModel
        Inherits ObservableObject

        Private _isUnread As Boolean

        Public Sub New(summary As MailSummary)
            Me.Summary = summary
            _isUnread = summary.IsUnread
        End Sub

        ''' <summary>
        ''' O DTO original. Quem precisa de dado bruto — chave, tamanho,
        ''' classe — lê daqui, sem duplicar tudo neste envelope.
        ''' </summary>
        Public ReadOnly Property Summary As MailSummary

        Public ReadOnly Property Key As ItemKey
            Get
                Return Summary.Key
            End Get
        End Property

        Public ReadOnly Property Subject As String
            Get
                Return Summary.Subject
            End Get
        End Property

        Public ReadOnly Property SenderName As String
            Get
                Return Summary.SenderName
            End Get
        End Property

        Public ReadOnly Property ReceivedTime As DateTimeOffset?
            Get
                Return Summary.ReceivedTime
            End Get
        End Property

        Public ReadOnly Property HasAttachments As Boolean
            Get
                Return Summary.HasAttachments
            End Get
        End Property

        ''' <summary>
        ''' A única propriedade que muda depois de criada — e a razão de esta
        ''' classe existir.
        ''' </summary>
        Public Property IsUnread As Boolean
            Get
                Return _isUnread
            End Get
            Set(value As Boolean)
                SetProperty(_isUnread, value)
            End Set
        End Property

    End Class

End Namespace

Imports System.Collections.Generic

Namespace Global.Iris.Model

    ''' <summary>
    ''' Identidade de uma tarefa. Tipo próprio pelo mesmo motivo do
    ''' <see cref="DraftKey"/> e do <see cref="AppointmentKey"/>: o compilador
    ''' impede passar uma mensagem para uma operação de tarefa.
    ''' </summary>
    Public NotInheritable Class TaskKey
        Implements IEquatable(Of TaskKey)

        Public ReadOnly Property Item As ItemKey

        Public Sub New(item As ItemKey)
            If item Is Nothing Then Throw New ArgumentNullException(NameOf(item))
            Me.Item = item
        End Sub

        Public Overrides Function ToString() As String
            Return "tarefa " & Item.ToString()
        End Function

        Public Overloads Function Equals(other As TaskKey) As Boolean _
            Implements IEquatable(Of TaskKey).Equals
            If other Is Nothing Then Return False
            Return Equals(Item, other.Item)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, TaskKey))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Item.GetHashCode()
        End Function
    End Class

    ''' <summary>
    ''' <b>O que se quer gravar numa tarefa — e o que ela NÃO tem.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>NÃO HÁ RESPONSÁVEL AQUI, E A AUSÊNCIA É A FUNCIONALIDADE</b>
    '''
    ''' É a mesma armadilha do calendário, num lugar diferente. Uma tarefa com
    ''' responsável é uma <i>atribuição</i>: <c>TaskItem.Assign()</c> seguido de
    ''' <c>Send()</c> manda um pedido de tarefa <b>por e-mail</b>, e as
    ''' atualizações de status passam a ir e voltar por e-mail também.
    '''
    ''' O Iris não envia. Sem campo para preencher, não existe caminho para ele
    ''' atribuir uma tarefa por engano — a invariante fica sustentada pelo
    ''' <b>tipo</b>, e não por alguém lembrar.
    '''
    ''' Quem quiser delegar faz no Outlook, onde vê para quem o pedido vai.
    ''' </summary>
    Public NotInheritable Class TaskDraft
        Public Property Subject As String = ""
        Public Property Body As String = ""

        ''' <summary>
        ''' Vencimento, ou <c>Nothing</c> para "sem prazo".
        '''
        ''' Anulável de propósito: o Outlook representa "sem data" com um
        ''' sentinela (<c>4501-01-01</c>), e traduzir isso para uma data
        ''' qualquer transformaria "não tem prazo" em "vence num dia
        ''' esquisito" — a mesma família de ausência-virando-fato que esta base
        ''' já corrigiu em cinco lugares.
        ''' </summary>
        Public Property Vence As DateTimeOffset?
    End Class

    ''' <summary>O que uma tarefa lida tem, para a tela mostrar.</summary>
    Public NotInheritable Class TaskInfo
        Public Property Key As ItemKey
        Public Property Subject As String = ""
        Public Property Vence As DateTimeOffset?
        Public Property Concluida As Boolean

        ''' <summary>
        ''' Esta tarefa está <b>atribuída</b> — ou seja, mexer nela conversa por
        ''' e-mail com outra pessoa.
        '''
        ''' A tela mostra isso e o Iris recusa mexer. Não é um detalhe de
        ''' apresentação: é o aviso de que aquela linha pertence a uma
        ''' negociação que o Iris não participa.
        ''' </summary>
        Public Property Atribuida As Boolean
    End Class

    ''' <summary>
    ''' O que uma leitura de tarefas conseguiu — e o que ela não conseguiu.
    '''
    ''' Mesma disciplina do <see cref="AppointmentWindow"/>: contagem de
    ''' recusados anulável, porque <c>Nothing</c> é "não contei" e zero é
    ''' "contei e não houve".
    ''' </summary>
    Public NotInheritable Class TaskList
        Public Property Items As New List(Of TaskInfo)()
        Public Property Skipped As Integer?
        Public Property Truncada As Boolean
        Public Property MotivoDoCorte As String = ""
    End Class

End Namespace

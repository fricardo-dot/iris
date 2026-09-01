Imports System.Collections.Generic

Namespace Global.Iris.Model

    ''' <summary>
    ''' <b>Em que pasta esta mensagem está — observado, não declarado.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO PRECISA EXISTIR</b>
    '''
    ''' O portão da divulgação autoriza <b>por pasta</b>: a ativação lista quais
    ''' pastas podem sair, e <c>DisclosurePolicy</c> nega mensagem que não seja da
    ''' pasta do pedido. Só que a pasta de cada mensagem vinha do <i>chamador</i>,
    ''' junto com o pedido — as duas pontas da comparação eram a mesma afirmação,
    ''' e nenhuma delas tinha sido verificada contra o Outlook.
    '''
    ''' No caminho por mensagem isso quase não doía: a seleção <i>era</i> a pasta
    ''' aberta, e a lista viera dela. Com a classificação em lote passou a doer: as
    ''' chaves vêm do <b>cache</b>, que é um retrato de quando a varredura rodou.
    ''' Uma mensagem movida depois disso para uma pasta confidencial continua no
    ''' retrato da pasta antiga — e o corpo dela sairia sob a autorização de uma
    ''' pasta onde ela já não está.
    '''
    ''' Achado por revisão externa em 01/09/2026.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>LIDO DUAS VEZES, DE PROPÓSITO</b>
    '''
    ''' Uma vez aqui, antes de qualquer leitura de corpo — é o que o portão usa
    ''' para decidir. E outra vez presa ao corpo, em
    ''' <see cref="MessageSnapshot.Pasta"/>. As duas visitas ao COM acontecem em
    ''' instantes diferentes, e a mensagem pode se mover entre elas; conferir só na
    ''' primeira deixaria essa janela aberta. É o mesmo desenho que o anexo já
    ''' usa, e pelo mesmo motivo.
    ''' </summary>
    Public NotInheritable Class PastaDoItem

        Public ReadOnly Property Item As ItemKey

        ''' <summary>
        ''' A pasta em que o Outlook diz que a mensagem está.
        '''
        ''' <b>Vazia quer dizer "não deu para saber"</b> — e não "está na pasta
        ''' certa". Uma <see cref="FolderKey"/> vazia nunca é igual à pasta de um
        ''' pedido de verdade, então o portão nega. É o lado seguro, e é o mesmo
        ''' tratamento que "não consegui contar anexo" recebe.
        ''' </summary>
        Public ReadOnly Property Pasta As FolderKey

        Public Sub New(item As ItemKey, pasta As FolderKey)
            Me.Item = item
            Me.Pasta = If(pasta, New FolderKey("", ""))
        End Sub

        ''' <summary>A pasta foi lida?</summary>
        Public ReadOnly Property Sabida As Boolean
            Get
                Return Pasta IsNot Nothing AndAlso Pasta.EntryId.Length > 0
            End Get
        End Property

    End Class

End Namespace

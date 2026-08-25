Imports System.Collections.Generic
Imports System.Linq

Namespace Global.Iris.Model

    ''' <summary>
    ''' <b>Uma mensagem como o provider a entregou — numa leitura só.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE O TIPO EXISTE</b>
    '''
    ''' O pipeline de conteúdo recebia <c>ItemKey</c>, <c>ChangeKey</c>, assunto,
    ''' remetente e corpo como <b>parâmetros separados</b>. Isso preservava o
    ''' par (item, versão) e não provava nada sobre ele: qualquer chamador podia
    ''' passar o item aprovado, a versão aprovada, e um corpo qualquer.
    '''
    ''' Aqui os campos vêm juntos, de uma leitura só, e o construtor é
    ''' <c>Friend</c> — visível só para a borda que fala com o Outlook e para os
    ''' testes.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ISTO PROVA, E O QUE NÃO PROVA</b>
    '''
    ''' Prova <b>qual camada</b> montou o objeto: fora da borda do provider,
    ''' ninguém monta. Não prova que a leitura do COM foi atômica — o Outlook
    ''' não oferece isso, e a §29.2 é justamente a resposta a essa falta: a
    ''' autorização se prende aos bytes, não à confiança na leitura.
    ''' </summary>
    Public NotInheritable Class MessageSnapshot

        Public ReadOnly Property Item As ItemKey
        ''' <summary>A <c>PR_CHANGE_KEY</c> lida <b>na mesma operação</b>.</summary>
        Public ReadOnly Property ChangeKey As String
        Public ReadOnly Property Assunto As String
        Public ReadOnly Property Remetente As String
        Public ReadOnly Property Destinatarios As IReadOnlyList(Of String)
        Public ReadOnly Property Corpo As String
        ''' <summary>O corpo veio como HTML.</summary>
        Public ReadOnly Property EhHtml As Boolean
        ''' <summary>O provider entregou o corpo inteiro.</summary>
        Public ReadOnly Property CorpoCompleto As Boolean

        ''' <summary>
        ''' <b>Tem anexo</b> — lido na <b>mesma visita</b> que o corpo.
        '''
        ''' <c>Nothing</c> quer dizer <b>não deu para saber</b>, e não "não
        ''' tem": a contagem pode falhar por guarda do Object Model, por item de
        ''' classe inesperada, ou por erro de COM. Quem lê isto trata os dois
        ''' casos igual — anexo está fora desta fase por inteiro, e "não sei"
        ''' nunca vira prova de ausência.
        '''
        ''' Ler aqui, e não só na classificação, é o que <b>fecha a corrida</b>:
        ''' o portão classifica numa visita e o corpo é lido em outra, então um
        ''' anexo acrescentado no meio passaria pelo portão. Esta leitura vem
        ''' presa ao corpo que vira bytes.
        ''' </summary>
        Public ReadOnly Property TemAnexo As Boolean?

        Friend Sub New(item As ItemKey, changeKey As String, assunto As String,
                       remetente As String, destinatarios As IEnumerable(Of String),
                       corpo As String, ehHtml As Boolean, corpoCompleto As Boolean,
                       temAnexo As Boolean?)
            Me.Item = item
            Me.ChangeKey = If(changeKey, "")
            Me.Assunto = If(assunto, "")
            Me.Remetente = If(remetente, "")
            Me.Destinatarios = Array.AsReadOnly(
                If(destinatarios, Enumerable.Empty(Of String)()).ToArray())
            Me.Corpo = If(corpo, "")
            Me.EhHtml = ehHtml
            Me.CorpoCompleto = corpoCompleto
            Me.TemAnexo = temAnexo
        End Sub

    End Class

End Namespace

Imports System.Collections.Generic

Namespace Global.Iris.Model

    ''' <summary>
    ''' Como a leitura do rótulo de sensibilidade se saiu.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO NÃO É UM BOOLEANO, NEM UM PAR "TEM/NÃO TEM"</b>
    '''
    ''' O portão da Fase 3 decide se conteúdo pode sair desta máquina. Se
    ''' <i>"não tem rótulo"</i> e <i>"não consegui ler o rótulo"</i> chegarem
    ''' a ele como o mesmo valor, ele libera mensagem classificada toda vez
    ''' que a leitura falhar — e falha de leitura é justamente o que acontece
    ''' com mensagem protegida.
    '''
    ''' Só "ausente versus exceção" também não basta. Uma <c>COMException</c>
    ''' pode significar propriedade inexistente, acesso negado, item não
    ''' baixado, Outlook ocupado ou defeito nosso, e essas cinco coisas
    ''' merecem decisões diferentes. Por isso o tipo enumera os estados que
    ''' o portão precisa distinguir, e nenhum deles é "provavelmente ok".
    '''
    ''' <b>Nenhum valor aqui autoriza nada.</b> Autorização vem da política,
    ''' e a política da Fase 3 não autoriza — ver a §28.2 do FASE3.md.
    ''' </summary>
    Public Enum LabelReadingKind
        ''' <summary>
        ''' Não classificado. É o valor <b>zero</b> de propósito: campo
        ''' esquecido, struct default e desserialização incompleta caem aqui,
        ''' e daqui não se autoriza nada.
        ''' </summary>
        Unknown = 0

        ''' <summary>Lido, com valor presente e bem formado.</summary>
        Present

        ''' <summary>
        ''' Propriedade <b>comprovadamente</b> ausente — o provider disse
        ''' `MAPI_E_NOT_FOUND`, não "deu erro".
        '''
        ''' Ausente NÃO quer dizer permitido. A semântica corporativa de
        ''' ausência é decisão de política, não de leitura.
        ''' </summary>
        Absent

        ''' <summary>Existe, e o valor é string vazia. Distinto de ausente.</summary>
        Blank

        ''' <summary>
        ''' Existe e não dá para interpretar: GUID inválido, campo
        ''' obrigatório faltando, formato desconhecido.
        '''
        ''' Nunca colapsar em <see cref="Absent"/>. Valor que eu não entendo
        ''' é exatamente o caso em que eu não posso decidir.
        ''' </summary>
        Malformed

        ''' <summary>Mais de um rótulo ativo, ou rótulos em conflito.</summary>
        Conflicting

        ''' <summary>Leitura negada — guarda do Object Model, política, IRM.</summary>
        Denied

        ''' <summary>Item protegido. Reservado por ser palavra da linguagem.</summary>
        Restricted

        ''' <summary>Ocupado, RPC, desconectado — pode voltar sozinho.</summary>
        Transient

        ''' <summary>Item não baixado, ou baixado só em parte.</summary>
        NotDownloaded

        ''' <summary>O item mudou entre o começo e o fim da leitura.</summary>
        Changed

        ''' <summary>
        ''' Falhou por outro motivo, com o HRESULT preservado. Caçamba, e por
        ''' isso mesmo <b>nunca</b> permissiva.
        ''' </summary>
        Unreadable
    End Enum

    ''' <summary>
    ''' Em que etapa a leitura parou. Obter o <c>PropertyAccessor</c>,
    ''' resolver a propriedade e ler o valor são falhas diferentes, e
    ''' colapsá-las apaga a informação que diz se o problema é do item, do
    ''' store ou da propriedade.
    ''' </summary>
    Public Enum LabelReadStage
        None = 0
        Item
        Accessor
        [Property]
        Value
        Parse
    End Enum

    ''' <summary>
    ''' O que permite perceber que o item mudou desde que foi classificado.
    '''
    ''' Nenhum destes cobre <b>atomicamente</b> rótulo e corpo — e é por isso
    ''' que a autorização da §29.2 se prende ao hash dos bytes que vão sair,
    ''' e não a esta evidência. Aqui é só o que dá para observar.
    ''' </summary>
    Public NotInheritable Class LabelVersionEvidence
        Public ReadOnly Property EntryId As String
        Public ReadOnly Property LastModified As DateTimeOffset?
        ''' <summary>`PR_CHANGE_KEY` em hex, quando o provider entrega.</summary>
        Public ReadOnly Property ChangeKey As String

        Public Sub New(entryId As String, lastModified As DateTimeOffset?, changeKey As String)
            Me.EntryId = If(entryId, "")
            Me.LastModified = lastModified
            Me.ChangeKey = changeKey
        End Sub

        ''' <summary>Duas evidências que descrevem a MESMA versão do item.</summary>
        Public Function Mesma(outra As LabelVersionEvidence) As Boolean
            If outra Is Nothing Then Return False
            If Not String.Equals(EntryId, outra.EntryId, StringComparison.Ordinal) Then Return False
            If Not String.Equals(If(ChangeKey, ""), If(outra.ChangeKey, ""),
                                 StringComparison.Ordinal) Then Return False
            Return Nullable.Equals(LastModified, outra.LastModified)
        End Function
    End Class

    ''' <summary>
    ''' O resultado da leitura do rótulo de UM item.
    '''
    ''' <b>O conteúdo do rótulo não vem inteiro.</b> Vêm os GUIDs, que são a
    ''' identidade estável, e o comprimento do valor cru. O nome do rótulo é
    ''' texto escolhido pela empresa e pode ele próprio ser sensível — e a
    ''' medição da §32 registra contagens, não amostras.
    ''' </summary>
    Public NotInheritable Class LabelReading
        Public ReadOnly Property Item As ItemKey
        Public ReadOnly Property Kind As LabelReadingKind
        Public ReadOnly Property Stage As LabelReadStage
        Public ReadOnly Property HResult As Integer?
        ''' <summary>GUIDs dos rótulos ATIVOS, em minúsculas, ordenados.</summary>
        Public ReadOnly Property LabelIds As IReadOnlyList(Of String)
        ''' <summary>Comprimento do valor cru, em caracteres. Não o valor.</summary>
        Public ReadOnly Property RawLength As Integer
        Public ReadOnly Property Version As LabelVersionEvidence

        Public Sub New(item As ItemKey, kind As LabelReadingKind,
                       Optional stage As LabelReadStage = LabelReadStage.None,
                       Optional hresult As Integer? = Nothing,
                       Optional labelIds As IReadOnlyList(Of String) = Nothing,
                       Optional rawLength As Integer = 0,
                       Optional version As LabelVersionEvidence = Nothing)
            Me.Item = item
            Me.Kind = kind
            Me.Stage = stage
            Me.HResult = hresult
            Me.LabelIds = If(labelIds, CType(Array.Empty(Of String)(), IReadOnlyList(Of String)))
            Me.RawLength = rawLength
            Me.Version = version
        End Sub

        ''' <summary>
        ''' A leitura produziu um estado sobre o qual dá para <b>raciocinar</b>
        ''' — o que é bem menos do que "pode transmitir".
        '''
        ''' Existe para o portão poder dizer "nem sei o que isto é" com uma
        ''' pergunta só, e não para autorizar coisa nenhuma.
        ''' </summary>
        Public ReadOnly Property Conclusiva As Boolean
            Get
                Return Kind = LabelReadingKind.Present OrElse
                       Kind = LabelReadingKind.Absent OrElse
                       Kind = LabelReadingKind.Blank
            End Get
        End Property
    End Class

End Namespace

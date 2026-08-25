Imports System.Collections.Generic
Imports System.Linq

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

        ''' <summary>
        ''' Mais de um rótulo ativo, <b>ou</b> o mesmo rótulo dizendo duas
        ''' coisas — <c>Enabled=True</c> e <c>Enabled=False</c> para o mesmo
        ''' GUID. A segunda forma é a que a primeira versão do parser deixava
        ''' passar como <see cref="Present"/>, porque o conjunto ordenado
        ''' terminava com um ativo só.
        ''' </summary>
        Conflicting

        ''' <summary>
        ''' Todo registro reconhecido está <c>Enabled=False</c>.
        '''
        ''' Isto descreve o <b>formato</b>, e só ele. Normalmente representa
        ''' rótulo removido — mas atualidade e autoridade não estão provadas, e
        ''' chamar de "foi removido" seria promover leitura plausível a fato.
        '''
        ''' Não é <see cref="Malformed"/>, porque o valor está bem formado; nem
        ''' <see cref="Absent"/>, que é a propriedade não existir. Colapsar em
        ''' qualquer um dos dois apagaria uma distinção que a política
        ''' corporativa pode querer usar.
        ''' </summary>
        HistoricalOnly

        ''' <summary>
        ''' Este caminho não suporta a propriedade — o provider recusou a
        ''' operação, não disse que o item não a tem.
        '''
        ''' Existe porque eu tratava <c>E_INVALIDARG</c> como
        ''' <see cref="Absent"/>, e a própria medição do 3.0 mostrou que esse
        ''' HRESULT é o que a <c>Table</c> devolve para "não aceito este DASL".
        ''' Ausência comprovada nesta conta é <c>MAPI_E_NOT_FOUND</c>, e só ela.
        ''' </summary>
        Unsupported

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

        ''' <summary>
        ''' Os registros do valor, <b>um por GUID</b>, com os campos preservados
        ''' e não interpretados.
        '''
        ''' Por GUID e não achatado: um item pode carregar mais de um registro,
        ''' e o <c>ContentBits</c> pertence ao registro do seu GUID. Um campo
        ''' único por leitura misturaria registros diferentes, e o portão
        ''' decidiria sobre a mistura.
        ''' </summary>
        Public ReadOnly Property Registros As IReadOnlyList(Of LabelRecord)

        ''' <summary>
        ''' Os <b>nomes de campo</b> vistos, de todos os registros juntos.
        ''' Nomes, não valores — o valor de <c>Name</c> é texto escolhido pela
        ''' empresa e pode ele próprio ser sensível.
        ''' </summary>
        Public ReadOnly Property Campos As IReadOnlyList(Of String)
            Get
                Return Registros.SelectMany(Function(r) r.Campos).
                                 Distinct(StringComparer.OrdinalIgnoreCase).
                                 OrderBy(Function(c) c, StringComparer.OrdinalIgnoreCase).
                                 ToList()
            End Get
        End Property
        Public ReadOnly Property Version As LabelVersionEvidence

        Public Sub New(item As ItemKey, kind As LabelReadingKind,
                       Optional stage As LabelReadStage = LabelReadStage.None,
                       Optional hresult As Integer? = Nothing,
                       Optional labelIds As IReadOnlyList(Of String) = Nothing,
                       Optional rawLength As Integer = 0,
                       Optional version As LabelVersionEvidence = Nothing,
                       Optional registros As IReadOnlyList(Of LabelRecord) = Nothing)
            Me.Item = item
            Me.Kind = kind
            Me.Stage = stage
            Me.HResult = hresult
            Me.LabelIds = If(labelIds, CType(Array.Empty(Of String)(), IReadOnlyList(Of String)))
            Me.RawLength = rawLength
            Me.Version = version
            Me.Registros = If(registros,
                              CType(Array.Empty(Of LabelRecord)(), IReadOnlyList(Of LabelRecord)))
        End Sub

        ''' <remarks>
        ''' <b>Aqui havia uma propriedade <c>Conclusiva</c>, e ela foi removida.</b>
        '''
        ''' Ela devolvia <c>True</c> para <see cref="LabelReadingKind.Present"/>,
        ''' <see cref="LabelReadingKind.Absent"/> e
        ''' <see cref="LabelReadingKind.Blank"/>, com um comentário explicando
        ''' que "conclusiva" queria dizer "dá para raciocinar" e não "pode
        ''' transmitir".
        '''
        ''' O comentário não sobreviveria ao primeiro
        ''' <c>If leitura.Conclusiva Then</c>. E os três estados que ela reunia
        ''' são exatamente aqueles cuja política <b>difere</b>: <c>Present</c>
        ''' não tem autoridade positiva confirmada (P16), <c>Absent</c> tem
        ''' semântica corporativa desconhecida (P14), e <c>Blank</c> não diz por
        ''' que a propriedade está vazia.
        '''
        ''' Um booleano que junta justamente os três é um convite a decidir sem
        ''' olhar. Quem decide trata cada membro do enum <b>explicitamente</b>.
        ''' </remarks>
    End Class

End Namespace

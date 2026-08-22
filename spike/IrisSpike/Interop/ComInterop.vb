Imports System.Runtime.InteropServices

Namespace Interop

    ''' <summary>
    ''' IOleMessageFilter — exigido pelo R13 do ESCOPO.md.
    '''
    ''' Sem um message filter registrado, uma chamada COM feita enquanto o
    ''' Outlook está ocupado (diálogo modal aberto, reparando store, em
    ''' sincronização) falha imediatamente com RPC_E_CALL_REJECTED em vez de
    ''' esperar. O filtro decide se a chamada espera e tenta de novo.
    ''' </summary>
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("00000016-0000-0000-C000-000000000046")>
    Friend Interface IOleMessageFilter

        <PreserveSig>
        Function HandleInComingCall(dwCallType As Integer,
                                    hTaskCaller As IntPtr,
                                    dwTickCount As Integer,
                                    lpInterfaceInfo As IntPtr) As Integer

        <PreserveSig>
        Function RetryRejectedCall(hTaskCallee As IntPtr,
                                   dwTickCount As Integer,
                                   dwRejectType As Integer) As Integer

        <PreserveSig>
        Function MessagePending(hTaskCallee As IntPtr,
                                dwTickCount As Integer,
                                dwPendingType As Integer) As Integer
    End Interface

    ''' <summary>
    ''' Implementação do filtro, com orçamento de retry limitado.
    '''
    ''' ATENÇÃO (R13): o filtro é registrado por thread e não sabe qual
    ''' operação está em curso. Por isso <see cref="AllowRetry"/> existe — o
    ''' broker desliga o retry ao redor de operações NÃO idempotentes, com
    ''' destaque para MailItem.Send(). Repetir um envio manda o e-mail duas
    ''' vezes.
    ''' </summary>
    Public NotInheritable Class OutlookMessageFilter
        Implements IOleMessageFilter

        ' Retornos de HandleInComingCall / dwRejectType
        Private Const SERVERCALL_ISHANDLED As Integer = 0
        Private Const SERVERCALL_RETRYLATER As Integer = 2

        ' Retornos de MessagePending
        Private Const PENDINGMSG_WAITDEFPROCESS As Integer = 2

        ' Retornos de RetryRejectedCall
        Private Const CANCEL_CALL As Integer = -1
        Private Const RETRY_AFTER_MS As Integer = 100

        ''' <summary>Teto total de espera antes de desistir da chamada.</summary>
        Public Property RetryBudgetMs As Integer = 5000

        ''' <summary>
        ''' Desligado pelo broker em operações com efeito colateral.
        ''' </summary>
        Public Property AllowRetry As Boolean = True

        Public ReadOnly Property RejectionsSeen As Integer
            Get
                Return _rejections
            End Get
        End Property

        Public ReadOnly Property RetriesIssued As Integer
            Get
                Return _retries
            End Get
        End Property

        Public ReadOnly Property CallsCancelled As Integer
            Get
                Return _cancelled
            End Get
        End Property

        Private _rejections As Integer
        Private _retries As Integer
        Private _cancelled As Integer

        Private Shared _registered As OutlookMessageFilter

        <DllImport("ole32.dll")>
        Private Shared Function CoRegisterMessageFilter(
            newFilter As IOleMessageFilter,
            <Out> ByRef oldFilter As IOleMessageFilter) As Integer
        End Function

        ''' <summary>
        ''' Registra o filtro. DEVE ser chamado de dentro da thread STA do
        ''' broker — o registro é por thread.
        ''' </summary>
        Public Shared Function Register() As OutlookMessageFilter
            Dim filter As New OutlookMessageFilter()
            Dim previous As IOleMessageFilter = Nothing
            Dim hr = CoRegisterMessageFilter(filter, previous)
            If hr <> 0 Then
                Throw New InvalidOperationException(
                    $"CoRegisterMessageFilter falhou (HRESULT 0x{hr:X8}).")
            End If
            _registered = filter
            Return filter
        End Function

        ''' <summary>Remove o filtro. Também por thread.</summary>
        Public Shared Sub Revoke()
            Dim previous As IOleMessageFilter = Nothing
            CoRegisterMessageFilter(Nothing, previous)
            _registered = Nothing
        End Sub

        Public Shared ReadOnly Property Current As OutlookMessageFilter
            Get
                Return _registered
            End Get
        End Property

        ''' <summary>
        ''' Chamada de entrada vinda do Outlook (por exemplo, um evento).
        ''' Sempre aceitar: recusar aqui é como os eventos somem.
        ''' </summary>
        Private Function HandleInComingCall(dwCallType As Integer,
                                            hTaskCaller As IntPtr,
                                            dwTickCount As Integer,
                                            lpInterfaceInfo As IntPtr) As Integer _
            Implements IOleMessageFilter.HandleInComingCall
            Return SERVERCALL_ISHANDLED
        End Function

        ''' <summary>
        ''' O Outlook recusou nossa chamada. dwTickCount é o tempo já
        ''' decorrido desde o início dela.
        ''' </summary>
        Private Function RetryRejectedCall(hTaskCallee As IntPtr,
                                           dwTickCount As Integer,
                                           dwRejectType As Integer) As Integer _
            Implements IOleMessageFilter.RetryRejectedCall

            If dwRejectType <> SERVERCALL_RETRYLATER Then
                _cancelled += 1
                Return CANCEL_CALL
            End If

            _rejections += 1

            If Not AllowRetry Then
                ' Operação não idempotente em curso. Nunca repetir.
                _cancelled += 1
                Return CANCEL_CALL
            End If

            If dwTickCount >= RetryBudgetMs Then
                _cancelled += 1
                Return CANCEL_CALL
            End If

            _retries += 1
            Return RETRY_AFTER_MS
        End Function

        ''' <summary>
        ''' Chegou mensagem do Windows enquanto esperávamos o Outlook.
        ''' WAITDEFPROCESS mantém a thread responsiva sem reentrar na chamada.
        ''' </summary>
        Private Function MessagePending(hTaskCallee As IntPtr,
                                        dwTickCount As Integer,
                                        dwPendingType As Integer) As Integer _
            Implements IOleMessageFilter.MessagePending
            Return PENDINGMSG_WAITDEFPROCESS
        End Function
    End Class

    ''' <summary>
    ''' Utilitários COM que o .NET moderno não fornece mais prontos.
    ''' </summary>
    Friend Module ComHelpers

        ' Marshal.GetActiveObject NÃO existe no .NET Core/5+. Anexar a uma
        ' instância em execução exige P/Invoke direto. Esta é uma pegadinha
        ' concreta da stack escolhida e vale registrar no relatório.
        <DllImport("oleaut32.dll", PreserveSig:=False)>
        Private Sub GetActiveObject(ByRef rclsid As Guid,
                                    pvReserved As IntPtr,
                                    <MarshalAs(UnmanagedType.IUnknown)> ByRef ppunk As Object)
        End Sub

        <DllImport("ole32.dll", PreserveSig:=False)>
        Private Sub CLSIDFromProgID(<MarshalAs(UnmanagedType.LPWStr)> progId As String,
                                    ByRef clsid As Guid)
        End Sub

        ''' <summary>
        ''' Anexa a uma instância JÁ EM EXECUÇÃO. Retorna Nothing se não há
        ''' nenhuma. Nunca inicia o aplicativo — o ESCOPO.md exige que o
        ''' Outlook já esteja aberto, e iniciar por CreateObject produz uma
        ''' instância sem perfil interativo.
        ''' </summary>
        Public Function GetRunningInstance(progId As String) As Object
            Dim clsid As Guid
            Try
                CLSIDFromProgID(progId, clsid)
            Catch ex As COMException
                ' ProgID não registrado: o Outlook clássico não está instalado.
                Return Nothing
            End Try

            Dim instance As Object = Nothing
            Try
                GetActiveObject(clsid, IntPtr.Zero, instance)
            Catch ex As COMException
                ' MK_E_UNAVAILABLE: registrado, mas nenhuma instância rodando.
                Return Nothing
            End Try

            Return instance
        End Function

        ''' <summary>
        ''' Liberação determinística de um RCW. Ver R7: encadear expressões
        ''' cria wrappers intermediários que ninguém libera, e o sintoma é
        ''' OUTLOOK.EXE órfão.
        ''' </summary>
        Public Sub Release(ByRef comObject As Object)
            If comObject Is Nothing Then Return
            Try
                If Marshal.IsComObject(comObject) Then
                    Marshal.ReleaseComObject(comObject)
                End If
            Catch
                ' Liberar nunca deve derrubar o encerramento.
            Finally
                comObject = Nothing
            End Try
        End Sub

        ''' <summary>
        ''' Verifica se um grafo de objetos carrega alguma referência COM
        ''' escondida. É a asserção que sustenta a fronteira da seção 4 do
        ''' ESCOPO.md: só DTOs atravessam, nunca RCW.
        ''' </summary>
        Public Function ContainsComReference(graph As Object,
                                             Optional maxDepth As Integer = 4) As Boolean
            Return ContainsComReferenceCore(graph, maxDepth, New HashSet(Of Object)(ReferenceEqualityComparer.Instance))
        End Function

        Private Function ContainsComReferenceCore(node As Object,
                                                  depth As Integer,
                                                  seen As HashSet(Of Object)) As Boolean
            If node Is Nothing OrElse depth < 0 Then Return False
            If Marshal.IsComObject(node) Then Return True

            Dim t = node.GetType()
            If t.IsPrimitive OrElse TypeOf node Is String OrElse TypeOf node Is DateTime OrElse
               TypeOf node Is Decimal OrElse t.IsEnum Then
                Return False
            End If

            If Not seen.Add(node) Then Return False

            Dim items = TryCast(node, Collections.IEnumerable)
            If items IsNot Nothing Then
                For Each element In items
                    If ContainsComReferenceCore(element, depth - 1, seen) Then Return True
                Next
                Return False
            End If

            For Each prop In t.GetProperties()
                If prop.GetIndexParameters().Length > 0 OrElse Not prop.CanRead Then Continue For
                Dim value As Object
                Try
                    value = prop.GetValue(node)
                Catch
                    Continue For
                End Try
                If ContainsComReferenceCore(value, depth - 1, seen) Then Return True
            Next

            Return False
        End Function
    End Module

End Namespace

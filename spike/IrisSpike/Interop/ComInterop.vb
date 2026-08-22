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

        ''' <summary>
        ''' ThreadStatic porque o registro do message filter é POR THREAD.
        ''' Como Shared comum, dois brokers (ou outro componente que use
        ''' filtro) enxergariam o estado um do outro.
        ''' </summary>
        <ThreadStatic>
        Private Shared _registered As OutlookMessageFilter

        ''' <summary>
        ''' Filtro que estava registrado antes de nós. CoRegisterMessageFilter
        ''' devolve o anterior, e descartá-lo significa não conseguir
        ''' restaurar o estado da thread no encerramento.
        ''' </summary>
        Private _previous As IOleMessageFilter

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
            filter._previous = previous
            _registered = filter
            Return filter
        End Function

        ''' <summary>
        ''' Restaura o filtro anterior. Também por thread. Registrar Nothing
        ''' deixaria a thread sem o filtro que ela tinha antes de nós.
        ''' </summary>
        Public Shared Sub Revoke()
            Dim current = _registered
            Dim restore As IOleMessageFilter = If(current Is Nothing, Nothing, current._previous)
            Dim discarded As IOleMessageFilter = Nothing
            CoRegisterMessageFilter(restore, discarded)
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
    Public Module ComHelpers

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
        ''' Por que não basta "deu certo ou não": um Outlook que está
        ''' ABRINDO, com diálogo modal ou reparando um store recusa a
        ''' chamada, e isso é completamente diferente de não estar aberto.
        ''' Confundir os dois manda o usuário abrir um Outlook que já está
        ''' na tela — foi exatamente o que este spike fez na primeira
        ''' execução contra o Outlook real (R13).
        ''' </summary>
        Public Enum AttachOutcome
            Ok
            NotRegistered
            NotRunning
            Busy
            Failed
        End Enum

        ' HRESULTs relevantes
        Private Const MK_E_UNAVAILABLE As Integer = &H800401E3
        Private Const RPC_E_CALL_REJECTED As Integer = &H80010001
        Private Const RPC_E_SERVERCALL_RETRYLATER As Integer = &H8001010A
        Private Const RPC_E_DISCONNECTED As Integer = &H80010108
        Private Const CO_E_SERVER_EXEC_FAILURE As Integer = &H80080005

        ''' <summary>
        ''' Anexa a uma instância JÁ EM EXECUÇÃO. Nunca inicia o aplicativo —
        ''' o ESCOPO.md exige que o Outlook já esteja aberto, e iniciar por
        ''' CreateObject produz uma instância sem perfil interativo.
        ''' </summary>
        Public Function GetRunningInstance(progId As String) _
            As (Instance As Object, Outcome As AttachOutcome, Hresult As Integer)

            Dim clsid As Guid
            Try
                CLSIDFromProgID(progId, clsid)
            Catch ex As COMException
                Return (Nothing, AttachOutcome.NotRegistered, ex.HResult)
            End Try

            Dim instance As Object = Nothing
            Try
                GetActiveObject(clsid, IntPtr.Zero, instance)
            Catch ex As COMException
                Select Case ex.HResult
                    Case MK_E_UNAVAILABLE
                        Return (Nothing, AttachOutcome.NotRunning, ex.HResult)
                    Case RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER,
                         RPC_E_DISCONNECTED, CO_E_SERVER_EXEC_FAILURE
                        Return (Nothing, AttachOutcome.Busy, ex.HResult)
                    Case Else
                        Return (Nothing, AttachOutcome.Failed, ex.HResult)
                End Select
            End Try

            Return (instance, AttachOutcome.Ok, 0)
        End Function

        ''' <summary>
        ''' Liberação determinística de um RCW. Ver R7: encadear expressões
        ''' cria wrappers intermediários que ninguém libera, e o sintoma é
        ''' OUTLOOK.EXE órfão.
        ''' </summary>
        ''' <remarks>
        ''' ByVal de propósito: com Option Strict On, um parâmetro ByRef As
        ''' Object obrigaria conversão estreitante na volta para qualquer
        ''' campo tipado. O chamador atribui Nothing ao próprio campo.
        ''' </remarks>
        Public Sub Release(comObject As Object)
            If comObject Is Nothing Then Return
            Try
                If Marshal.IsComObject(comObject) Then
                    Marshal.ReleaseComObject(comObject)
                End If
            Catch
                ' Liberar nunca deve derrubar o encerramento.
            End Try
        End Sub

        ''' <summary>
        ''' Verifica se um grafo de objetos carrega alguma referência COM
        ''' escondida. É a asserção que sustenta a fronteira da seção 4 do
        ''' ESCOPO.md: só DTOs atravessam, nunca RCW.
        ''' </summary>
        ''' <remarks>
        ''' LIMITES conhecidos, para ninguém tratar isto como prova formal:
        ''' a profundidade é finita; getters são executados (podem ter efeito
        ''' colateral) e os que lançam são ignorados; um wrapper que esconda
        ''' um RCW atrás de lógica própria pode escapar. A garantia de
        ''' verdade vem do desenho — DTOs de tipos primitivos e uma API que
        ''' não devolve COM — e isto aqui é rede de segurança, não o piso.
        ''' </remarks>
        Public Function ContainsComReference(graph As Object,
                                             Optional maxDepth As Integer = 6) As Boolean
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

            ' Campos também, não só propriedades: um RCW guardado em campo
            ' privado passaria batido numa varredura só de propriedades.
            Const AllInstance As Reflection.BindingFlags =
                Reflection.BindingFlags.Public Or Reflection.BindingFlags.NonPublic Or
                Reflection.BindingFlags.Instance

            For Each field In t.GetFields(AllInstance)
                Dim value As Object
                Try
                    value = field.GetValue(node)
                Catch
                    Continue For
                End Try
                If ContainsComReferenceCore(value, depth - 1, seen) Then Return True
            Next

            For Each prop In t.GetProperties(AllInstance)
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

Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Threading
Imports IrisSpike.Broker
Imports IrisSpike.Interop
Imports IrisSpike.Model

Namespace Checks

    ''' <summary>
    ''' Grupo A da Fase 0 — o broker.
    '''
    ''' Estes critérios NÃO dependem do Outlook e rodam em qualquer máquina.
    ''' Isso é de propósito: o risco que eles cobrem é de arquitetura, não de
    ''' integração. Um spike que só provasse "MailItem funciona" validaria o
    ''' OOM e deixaria o desenho sem prova nenhuma.
    ''' </summary>
    Public NotInheritable Class BrokerChecks

        Private Const Group As String = "A — Broker"

        Private ReadOnly _runner As CheckRunner
        Private ReadOnly _broker As OutlookBroker

        Public Sub New(runner As CheckRunner, broker As OutlookBroker)
            _runner = runner
            _broker = broker
        End Sub

        Public Async Function RunAsync() As Task

            ' ---------------------------------------------------------------
            Await _runner.RunAsync(
                "A1", Group, "Thread do broker é STA dedicada",
                Async Function()
                    Await Task.CompletedTask

                    Dim apartment = _broker.Apartment
                    Dim isolated = _broker.ThreadId <> Environment.CurrentManagedThreadId

                    If apartment <> ApartmentState.STA Then
                        Return (CheckStatus.Fail, $"Apartment = {apartment}, esperado STA (R6).")
                    End If
                    If Not isolated Then
                        Return (CheckStatus.Fail, "Broker está na thread do chamador; não é dedicada.")
                    End If

                    Return (CheckStatus.Pass,
                            $"STA, thread {_broker.ThreadId}, separada da chamadora " &
                            $"({Environment.CurrentManagedThreadId}).")
                End Function)

            ' ---------------------------------------------------------------
            ' O critério que separa "fila" de "message pump". Um
            ' While/Dequeue processaria comandos e nunca faria um
            ' DispatcherTimer disparar, porque timer depende de mensagem de
            ' janela sendo bombeada. É assim que os eventos do Outlook
            ' morrem silenciosamente.
            Await _runner.RunAsync(
                "A2", Group, "Dispatcher despachando na STA (pump estrutural)",
                Async Function()
                    Dim ticked As New TaskCompletionSource(Of Boolean)(
                        TaskCreationOptions.RunContinuationsAsynchronously)
                    Dim timerRef As DispatcherTimer = Nothing

                    Await _broker.InvokeAsync(
                        Sub()
                            timerRef = New DispatcherTimer(DispatcherPriority.Normal,
                                                           Dispatcher.CurrentDispatcher) With {
                                .Interval = TimeSpan.FromMilliseconds(50)
                            }
                            AddHandler timerRef.Tick,
                                Sub()
                                    timerRef.Stop()
                                    ticked.TrySetResult(True)
                                End Sub
                            timerRef.Start()
                        End Sub)

                    Dim completed = Await Task.WhenAny(ticked.Task, Task.Delay(3000))
                    GC.KeepAlive(timerRef)

                    If completed IsNot ticked.Task Then
                        Return (CheckStatus.Fail,
                                "DispatcherTimer não disparou em 3s: a thread não bombeia " &
                                "mensagens. Eventos do Outlook não chegariam.")
                    End If

                    ' Limite honesto: isto prova que o dispatcher despacha e
                    ' que a thread não está presa num Take() bloqueante. NÃO
                    ' prova que um callback COM do Outlook chega, nem em qual
                    ' thread ele chega, nem o comportamento durante uma
                    ' chamada COM longa. Essa prova é do grupo D.
                    Return (CheckStatus.Pass,
                            "DispatcherTimer disparou: dispatcher vivo, thread não bloqueada. " &
                            "Entrega de evento COM é o grupo D.")
                End Function)

            ' ---------------------------------------------------------------
            Await _runner.RunAsync(
                "A3", Group, "InvokeAsync marshala para a thread do broker",
                Async Function()
                    Dim observed = Await _broker.InvokeAsync(
                        Function() Environment.CurrentManagedThreadId)

                    If observed <> _broker.ThreadId Then
                        Return (CheckStatus.Fail,
                                $"Trabalho rodou na thread {observed}, esperado {_broker.ThreadId}.")
                    End If

                    Return (CheckStatus.Pass, $"Trabalho executou na thread {observed}.")
                End Function)

            ' ---------------------------------------------------------------
            Await _runner.RunAsync(
                "A4", Group, "IOleMessageFilter registrado na thread do broker",
                Async Function()
                    Await Task.CompletedTask

                    If _broker.MessageFilter Is Nothing Then
                        Return (CheckStatus.Fail,
                                "Sem message filter: RPC_E_CALL_REJECTED viraria falha imediata (R13).")
                    End If

                    Dim onBrokerThread = Await _broker.InvokeAsync(
                        Function() OutlookMessageFilter.Current IsNot Nothing)

                    If Not onBrokerThread Then
                        Return (CheckStatus.Fail, "Filtro não está registrado NA thread do broker.")
                    End If

                    Return (CheckStatus.Pass,
                            $"Registrado; orçamento de retry {_broker.MessageFilter.RetryBudgetMs} ms.")
                End Function)

            ' ---------------------------------------------------------------
            ' Controle positivo: antes de afirmar que o DTO está limpo, é
            ' preciso provar que o detector consegue detectar. Um detector
            ' que sempre retorna False passaria no A6 sem valer nada.
            Await _runner.RunAsync(
                "A5", Group, "Detector de RCW funciona (controle positivo)",
                Async Function()
                    Await Task.CompletedTask

                    Dim comType = Type.GetTypeFromProgID("Shell.Application")
                    If comType Is Nothing Then
                        Return (CheckStatus.Skipped,
                                "Shell.Application indisponível; sem objeto COM para o controle.")
                    End If

                    Dim comObject As Object = Nothing
                    Try
                        comObject = Activator.CreateInstance(comType)
                        If Not ComHelpers.ContainsComReference(comObject) Then
                            Return (CheckStatus.Fail,
                                    "Detector não reconheceu um RCW real; o A6 não teria valor.")
                        End If
                        Return (CheckStatus.Pass, "Detector reconhece um RCW real.")
                    Finally
                        ComHelpers.Release(comObject)
                    End Try
                End Function)

            ' ---------------------------------------------------------------
            Await _runner.RunAsync(
                "A6", Group, "DTO sintético cruza a fronteira sem RCW (regressão)",
                Async Function()
                    Dim dto = Await _broker.InvokeAsync(
                        Function()
                            Return New MailSummary With {
                                .Key = New ItemKey With {.EntryId = "0000", .StoreId = "0000"},
                                .Subject = "sonda",
                                .SenderName = "spike",
                                .ReceivedTime = DateTime.Now,
                                .SizeBytes = 0,
                                .Content = ContentState.MetadataOnly,
                                .MessageClass = "IPM.Note"
                            }
                        End Function)

                    If ComHelpers.ContainsComReference(dto) Then
                        Return (CheckStatus.Fail,
                                "O DTO carrega RCW escondido: a fronteira da seção 4 está furada.")
                    End If

                    ' Teste fraco por natureza: o DTO é sintético e já se sabe
                    ' por inspeção que não tem COM. Vale como regressão. A
                    ' prova com dado real é o B4, sobre DTOs construídos a
                    ' partir de MailItem.
                    Return (CheckStatus.Pass, "Nenhum RCW no grafo do DTO sintético.")
                End Function)

            ' ---------------------------------------------------------------
            ' A fronteira agora é IMPOSTA, não só convencionada: ReadAsync e
            ' MutateAsync passam o resultado por GuardResult. Antes desta
            ' verificação, devolver um MAPIFolder da lambda compilava e
            ' funcionava.
            Await _runner.RunAsync(
                "A7", Group, "Devolver um RCW pela lambda é bloqueado",
                Async Function()
                    Try
                        Dim leaked = Await _broker.InvokeAsync(
                            Function() CObj(Activator.CreateInstance(
                                Type.GetTypeFromProgID("Shell.Application"))))

                        ComHelpers.Release(leaked)
                        Return (CheckStatus.Fail,
                                "Um RCW atravessou a fronteira do broker sem ser barrado.")
                    Catch ex As InvalidOperationException
                        Return (CheckStatus.Pass, "Tentativa de vazar RCW foi rejeitada.")
                    End Try
                End Function)

            ' ---------------------------------------------------------------
            Await _runner.RunAsync(
                "A7b", Group, "Guarda de thread rejeita chamada externa",
                Async Function()
                    Await Task.CompletedTask
                    Try
                        _broker.AssertOnBrokerThread()
                        Return (CheckStatus.Fail,
                                "A guarda não disparou fora da thread do broker (R6).")
                    Catch ex As InvalidOperationException
                        ' Limite: isto só protege métodos que CHAMAM a guarda.
                        ' Não intercepta uso de um RCW já vazado — quem faz
                        ' isso é o A7.
                        Return (CheckStatus.Pass,
                                "Guarda rejeitou o acesso externo (protege só quem a chama).")
                    End Try
                End Function)

        End Function

        ''' <summary>
        ''' A8 roda por último: derruba o broker e confirma que a thread
        ''' encerra. Fica separado porque destrói o objeto sob teste.
        ''' </summary>
        Public Async Function RunShutdownCheckAsync() As Task
            Await _runner.RunAsync(
                "A8", Group, "Encerramento ordenado da thread do broker",
                Async Function()
                    Await Task.CompletedTask
                    Try
                        _broker.Shutdown(TimeSpan.FromSeconds(10))
                    Catch ex As TimeoutException
                        Return (CheckStatus.Fail,
                                "Thread não encerrou: é o caminho para OUTLOOK.EXE órfão (R7).")
                    End Try

                    If _broker.IsRunning Then
                        Return (CheckStatus.Fail, "Dispatcher ainda ativo após Shutdown().")
                    End If

                    Return (CheckStatus.Pass, "Pump parado e thread encerrada.")
                End Function)
        End Function

    End Class

End Namespace

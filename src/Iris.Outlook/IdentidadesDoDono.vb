Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' <b>Os endereços pelos quais o dono desta caixa envia.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>PARA QUE SERVE, E POR QUE PRECISA DE MAIS DE UMA FORMA</b>
    '''
    ''' A fila de respostas pendentes separa "estou esperando alguém" de
    ''' "alguém está me esperando", e a separação é uma pergunta só: <i>quem
    ''' escreveu a última mensagem desta conversa?</i> A resposta compara o
    ''' remetente lido com este conjunto.
    '''
    ''' O remetente lido não vem numa forma só. Numa organização Exchange,
    ''' <c>SenderEmailAddress</c> de uma mensagem <b>interna</b> é um endereço
    ''' X.500 — e as internas são justamente as que enchem a fila. Então o
    ''' conjunto precisa das duas: o SMTP das contas e o X.500 do usuário da
    ''' sessão. Semear só o SMTP faria as mensagens do próprio dono aparecerem
    ''' como sendo de terceiros, que é o defeito exato que a fila não pode ter.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ISTO NÃO SABE</b>
    '''
    ''' <b>Alias</b> — o Outlook não expõe uma lista deles de um jeito em que se
    ''' possa confiar. <b>Caixa compartilhada</b> em que o dono responde como
    ''' outra pessoa. <b>Delegação.</b> Nada disso sai daqui, e é por isso que a
    ''' semeadura escreve um arquivo que o dono pode corrigir em vez de decidir
    ''' sozinha: o que esta função devolve é um <i>começo</i>, não a verdade.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>R7</b>
    '''
    ''' <c>Accounts</c>, <c>Account</c>, <c>CurrentUser</c> e <c>AddressEntry</c>
    ''' são objetos COM, cada um em sua variável, liberados em ordem inversa. A
    ''' regra já foi violada quatro vezes neste projeto, sempre em código que
    ''' "só lia uma contagem".
    ''' </summary>
    Friend Module IdentidadesDoDono

        ''' <summary>
        ''' Tudo o que dá para saber, sem lançar. Falha parcial devolve o que
        ''' deu — meia lista é melhor que nenhuma, e o que falta o dono
        ''' acrescenta no arquivo.
        ''' </summary>
        Friend Function Ler(ns As OL.NameSpace) As IReadOnlyList(Of String)
            Dim achados As New List(Of String)()
            If ns Is Nothing Then Return achados

            LerAsContas(ns, achados)
            LerOUsuarioDaSessao(ns, achados)
            Return achados
        End Function

        ''' <summary>O SMTP de cada conta configurada.</summary>
        Private Sub LerAsContas(ns As OL.NameSpace, achados As List(Of String))
            Dim contas As OL.Accounts = Nothing
            Try
                contas = ns.Accounts
                For i = 1 To contas.Count
                    Dim conta As OL.Account = Nothing
                    Try
                        conta = contas.Item(i)
                        Acrescentar(achados, Texto(Function() conta.SmtpAddress))
                    Catch
                        ' Uma conta ilegível não impede as outras.
                    Finally
                        ComHelpers.Release(conta)
                    End Try
                Next
            Catch
            Finally
                ComHelpers.Release(contas)
            End Try
        End Sub

        ''' <summary>
        ''' O usuário da sessão, nas duas formas que ele tem.
        '''
        ''' <c>AddressEntry.Address</c> devolve o X.500 numa caixa Exchange, e
        ''' <c>GetExchangeUser().PrimarySmtpAddress</c> devolve o SMTP
        ''' correspondente. As duas entram: qual delas o
        ''' <c>SenderEmailAddress</c> vai trazer depende de a mensagem ser
        ''' interna ou externa, e não é escolha nossa.
        ''' </summary>
        Private Sub LerOUsuarioDaSessao(ns As OL.NameSpace, achados As List(Of String))
            Dim usuario As OL.Recipient = Nothing
            Try
                usuario = ns.CurrentUser
                If usuario Is Nothing Then Return

                Acrescentar(achados, Texto(Function() usuario.Address))

                ' NUNCA ENCADEAR: AddressEntry devolve outro objeto COM, e
                ' GetExchangeUser mais um. Cada um na sua variável (R7).
                Dim entrada As OL.AddressEntry = Nothing
                Try
                    entrada = usuario.AddressEntry
                    If entrada Is Nothing Then Return

                    Acrescentar(achados, Texto(Function() entrada.Address))

                    Dim exchange As OL.ExchangeUser = Nothing
                    Try
                        exchange = entrada.GetExchangeUser()
                        If exchange IsNot Nothing Then
                            Acrescentar(achados, Texto(Function() exchange.PrimarySmtpAddress))
                        End If
                    Catch
                        ' Caixa que não é Exchange: não há usuário a resolver,
                        ' e o SMTP já veio das contas.
                    Finally
                        ComHelpers.Release(exchange)
                    End Try
                Catch
                Finally
                    ComHelpers.Release(entrada)
                End Try
            Catch
            Finally
                ComHelpers.Release(usuario)
            End Try
        End Sub

        ''' <summary>Sem repetir, sem vazio — a normalização é do modelo.</summary>
        Private Sub Acrescentar(achados As List(Of String), valor As String)
            If String.IsNullOrWhiteSpace(valor) Then Return
            If achados.Contains(valor, StringComparer.OrdinalIgnoreCase) Then Return
            achados.Add(valor.Trim())
        End Sub

        ''' <summary>Propriedade que não lê vira vazio, e não exceção.</summary>
        Private Function Texto(ler As Func(Of String)) As String
            Try
                Return If(ler(), "")
            Catch ex As COMException
                Return ""
            Catch ex As InvalidCastException
                Return ""
            End Try
        End Function

    End Module

End Namespace

Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports Iris.Model
Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' O rótulo vem por coluna de <c>Table</c>? (P6)
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ESTE PROBE É O PRIMEIRO CONTATO COM A CAIXA</b>
    '''
    ''' O protocolo da §32 lê por <c>Table</c> <b>antes</b> de abrir item, e
    ''' isso não é otimização: uma <c>Table</c> projeta colunas sem
    ''' materializar 20 RCWs de <c>MailItem</c>, sem tocar em <c>Body</c> e
    ''' sem chance de disparar diálogo de direitos. Com o usuário longe da
    ''' máquina, a ordem "menos invasivo primeiro" é a única defensável.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ACEITAR A COLUNA NÃO É ENTREGAR O VALOR</b>
    '''
    ''' <c>Columns.Add</c> pode aceitar um DASL e as linhas virem vazias ou
    ''' lançando. Por isso este probe conta as três coisas separadamente —
    ''' aceitou, veio valor, deu erro — e o <c>Serve</c> do
    ''' <see cref="LabelColumnProbe"/> exige as três alinhadas. Um probe que
    ''' só olhasse o <c>Add</c> concluiria que o caminho barato existe quando
    ''' ele não existe.
    '''
    ''' <b>Nada aqui lê corpo, anexo ou conversa.</b>
    ''' </summary>
    Friend Module SensitivityLabelColumn

        Public Function Probe(ns As OL.NameSpace, folder As FolderKey,
                              quantas As Integer) As OperationResult(Of LabelColumnProbe)

            Dim pasta As OL.MAPIFolder = Nothing
            Try
                Try
                    pasta = TryCast(ns.GetFolderFromID(folder.EntryId, folder.StoreId), OL.MAPIFolder)
                Catch ex As COMException
                    Return OperationResult(Of LabelColumnProbe).Fail(ErrorKind.NotFound, "pasta")
                End Try
                If pasta Is Nothing Then
                    Return OperationResult(Of LabelColumnProbe).Fail(ErrorKind.NotFound, "pasta")
                End If

                Return OperationResult(Of LabelColumnProbe).Ok(Medir(pasta, quantas))
            Finally
                ComHelpers.Release(pasta)
            End Try
        End Function

        Private Function Medir(pasta As OL.MAPIFolder, quantas As Integer) As LabelColumnProbe
            Dim r As New LabelColumnProbe()
            Dim relogio = Stopwatch.StartNew()

            Dim tabela As OL.Table = Nothing
            Try
                tabela = pasta.GetTable()

                Dim colunas As OL.Columns = Nothing
                Try
                    colunas = tabela.Columns
                    colunas.RemoveAll()
                    r.ColunaAceita = Acrescentar(colunas, SensitivityLabels.DaslMsipLabels, r)
                Finally
                    ComHelpers.Release(colunas)
                End Try

                If Not r.ColunaAceita Then Return Fechar(r, relogio)

                While r.LinhasLidas < quantas AndAlso Not tabela.EndOfTable
                    r.LinhasLidas += 1
                    Try
                        Dim linha As OL.Row = Nothing
                        Try
                            linha = tabela.GetNextRow()
                            If linha Is Nothing Then
                                r.LinhasLidas -= 1
                                Exit While
                            End If
                            Dim valor = TryCast(linha.Item(1), String)
                            If String.IsNullOrWhiteSpace(valor) Then
                                r.LinhasSemValor += 1
                            Else
                                r.LinhasComValor += 1
                            End If
                        Finally
                            ComHelpers.Release(linha)
                        End Try
                    Catch ex As COMException
                        ' A armadilha em pessoa: coluna aceita, linha
                        ' lancando. Conta e SEGUE — parar aqui esconderia
                        ' quantas linhas dao erro, que e justamente o numero
                        ' que decide se o caminho barato existe.
                        r.LinhasComErro += 1
                    End Try
                End While

                Return Fechar(r, relogio)
            Finally
                ComHelpers.Release(tabela)
            End Try
        End Function

        ''' <summary>
        ''' <c>Columns.Add</c> devolve um objeto COM, e ignorar o retorno
        ''' deixa RCW sem dono — a R7, já violada quatro vezes neste projeto,
        ''' sempre em código que "só lia".
        ''' </summary>
        Private Function Acrescentar(colunas As OL.Columns, dasl As String,
                                     r As LabelColumnProbe) As Boolean
            Dim coluna As OL.Column = Nothing
            Try
                coluna = colunas.Add(dasl)
                Return True
            Catch ex As Exception
                ' Excecao COMUM, e nao so COMException: medido nesta conta, o
                ' Outlook recusa um DASL de named property com
                ' ArgumentException, nao com COMException. Capturar so a
                ' segunda derrubava a operacao inteira e transformava a
                ' resposta da P6 - "a coluna nao serve" - num erro do broker.
                r.HResultDoAdd = ex.HResult
                Return False
            Finally
                ComHelpers.Release(coluna)
            End Try
        End Function

        Private Function Fechar(r As LabelColumnProbe, relogio As Stopwatch) As LabelColumnProbe
            relogio.Stop()
            r.MilissegundosTotais = relogio.Elapsed.TotalMilliseconds
            Return r
        End Function

    End Module

End Namespace

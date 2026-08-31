Imports Iris.Outlook.Interop
Imports OL = Microsoft.Office.Interop.Outlook

Namespace Global.Iris.Outlook

    ''' <summary>
    ''' <b>Imagem embutida não é anexo — e a distinção mora aqui, uma vez só.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE FOI MEDIDO, E POR QUE ISTO EXISTE</b>
    '''
    ''' Em 30/08/2026, na pasta "0. E-mails Lidos" de uma caixa corporativa
    ''' real, com 13 mensagens: <b>zero</b> sem anexo nenhum, <b>dez</b> só com
    ''' imagem embutida, três com anexo de verdade.
    '''
    ''' O portão negava qualquer <c>Attachments.Count &gt; 0</c>, então a IA
    ''' recusava 13 de 13. Não era guarda rigorosa: era guarda olhando a coisa
    ''' errada — toda assinatura corporativa tem logo, e logo é anexo com
    ''' <c>Content-ID</c>.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE DISTINGUE</b>
    '''
    ''' <c>PR_ATTACH_CONTENT_ID</c> (<c>0x3712001F</c>). Quem tem, é
    ''' referenciado pelo corpo por <c>cid:</c> — logo de assinatura, imagem no
    ''' meio do texto. Quem não tem, é arquivo que alguém mandou.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E O QUE ELA NÃO DISTINGUE</b>
    '''
    ''' Uma <b>captura de tela colada no corpo</b> é embutida também, e pode
    ''' carregar o teor inteiro da mensagem. Por isso a contagem viaja junto até
    ''' a tela e é declarada no resumo: trocar uma recusa honesta por um resumo
    ''' silenciosamente parcial seria a mesma família de defeito que esta base
    ''' passou a série inteira corrigindo.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE NUM MÓDULO, E NÃO REPETIDO NOS DOIS LUGARES</b>
    '''
    ''' Dois leitores precisam desta regra — o do portão
    ''' (<see cref="AnexosPresentes"/>) e o do corpo
    ''' (<see cref="MessageSnapshots"/>). Duas cópias divergem, e quando
    ''' divergirem o portão vai autorizar por um critério e a captura vai usar
    ''' outro. Este projeto já pagou por duas implementações da mesma regra na
    ''' busca, e lá a divergência ficou escondida até alguém procurar por ela.
    ''' </summary>
    Friend Module ClassificacaoDeAnexo

        ''' <summary>
        ''' Percorre os anexos e separa.
        '''
        ''' <c>Real</c> é <c>Nothing</c> quando não deu para olhar a coleção — e
        ''' <c>Nothing</c> nunca vale como "não tem": quem chama fecha.
        '''
        ''' <b>Anexo que não se deixa inspecionar conta como REAL.</b> Não saber
        ''' o que é vale como o caso que bloqueia.
        ''' </summary>
        Friend Function Contar(anexos As OL.Attachments) _
                               As (Real As Boolean?, Embutidas As Integer?)
            If anexos Is Nothing Then Return (Nothing, Nothing)

            Dim total As Integer
            Try
                total = anexos.Count
            Catch
                Return (Nothing, Nothing)
            End Try

            If total = 0 Then Return (False, 0)

            Dim embutidas = 0
            Dim reais = 0
            For i = 1 To total
                Dim a As OL.Attachment = Nothing
                Try
                    a = anexos.Item(i)
                    If EhEmbutida(a) Then embutidas += 1 Else reais += 1
                Catch
                    reais += 1
                Finally
                    ComHelpers.Release(a)
                End Try
            Next

            Return (reais > 0, embutidas)
        End Function

        ''' <summary>
        ''' O anexo é referenciado pelo corpo?
        '''
        ''' <b>Falha ao ler vale FALSE</b> — ou seja, "é anexo de verdade". A
        ''' propriedade pode simplesmente não existir, e ausência de
        ''' <c>Content-ID</c> é o que caracteriza anexo comum; exceção de leitura
        ''' também cai aqui, e o resultado conservador é o mesmo.
        ''' </summary>
        Private Function EhEmbutida(a As OL.Attachment) As Boolean
            Try
                Dim cid = TryCast(a.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F"), String)
                Return Not String.IsNullOrWhiteSpace(cid)
            Catch
                Return False
            End Try
        End Function

    End Module

End Namespace

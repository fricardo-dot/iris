Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports Iris.Model

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>O que o dono já disse que não precisa de resposta.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO EXISTE ANTES DA IA</b>
    '''
    ''' A fila mostra quem falou por último, e não sabe se aquilo pede resposta.
    ''' Newsletter, notificação automática e "obrigado, recebido" entram como
    ''' pendências de vinte dias — e uma fila com lixo é uma fila que se aprende
    ''' a ignorar, inclusive nas linhas em que ela acertou.
    '''
    ''' Duas ações resolvem a maior parte disso sem chamada nenhuma a modelo
    ''' externo: <b>não exige resposta</b> tira aquela conversa, e <b>ignorar
    ''' remetente</b> vira regra permanente. As duas são do dono, ficam em
    ''' arquivo que ele pode abrir, e não dependem de autorização de divulgação.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>DOIS ARQUIVOS, E NÃO UM</b>
    '''
    ''' São coisas de vidas diferentes. A conversa dispensada é um fato pontual
    ''' que envelhece — a conversa pode nem existir mais. O remetente ignorado é
    ''' uma regra que vale para o que ainda vai chegar. Misturá-los faria a
    ''' limpeza de um apagar o outro.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>FALHA VALE COMO NADA DISPENSADO</b>
    '''
    ''' Não conseguir ler devolve conjunto vazio, e a fila mostra <i>mais</i> do
    ''' que deveria. É o lado certo de errar: linha a mais o dono descarta de
    ''' novo; linha a menos ele nunca vê, e não sabe que não viu.
    ''' </summary>
    Public NotInheritable Class DispensasDaFila

        Private ReadOnly _conversas As String
        Private ReadOnly _remetentes As String

        Public Sub New(Optional pasta As String = Nothing)
            Dim raiz = If(pasta, Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris"))
            _conversas = Path.Combine(raiz, "fila-dispensadas.txt")
            _remetentes = Path.Combine(raiz, "fila-remetentes-ignorados.txt")
        End Sub

        Public ReadOnly Property CaminhoDasConversas As String
            Get
                Return _conversas
            End Get
        End Property

        Public ReadOnly Property CaminhoDosRemetentes As String
            Get
                Return _remetentes
            End Get
        End Property

        ''' <summary>As conversas dispensadas, para a fila filtrar.</summary>
        Public Function Conversas() As IReadOnlyList(Of String)
            Return Ler(_conversas)
        End Function

        ''' <summary>
        ''' Os remetentes ignorados, no mesmo tipo das identidades do dono.
        '''
        ''' <b>Reaproveitar <see cref="MinhasIdentidades"/> não é economia:</b> é
        ''' a mesma pergunta — "este endereço está neste conjunto?" — com o mesmo
        ''' tratamento de caixa, de nome de exibição e de X.500. Um casamento
        ''' escrito de novo aqui divergiria do outro em algum caso, e o dono
        ''' veria a regra funcionar para uns remetentes e não para outros.
        ''' </summary>
        Public Function Remetentes() As MinhasIdentidades
            Return New MinhasIdentidades(Ler(_remetentes))
        End Function

        ''' <summary>
        ''' Dispensa uma conversa. Repetir é inofensivo — o dono pode clicar duas
        ''' vezes, e a fila não pode piorar por isso.
        ''' </summary>
        Public Function DispensarConversa(conversa As String) As Boolean
            Return Acrescentar(_conversas, conversa, CabecalhoDasConversas())
        End Function

        Public Function IgnorarRemetente(endereco As String) As Boolean
            Return Acrescentar(_remetentes, endereco, CabecalhoDosRemetentes())
        End Function

        ' ==============================================================

        Private Shared Function Ler(caminho As String) As IReadOnlyList(Of String)
            Try
                If Not File.Exists(caminho) Then Return Array.Empty(Of String)()

                Return File.ReadAllLines(caminho, Encoding.UTF8).
                       Select(Function(l) l.Trim()).
                       Where(Function(l) l.Length > 0 AndAlso Not l.StartsWith("#")).
                       ToList()
            Catch
                ' Ver o cabecalho: falha vale como nada dispensado, e a fila
                ' mostra mais do que deveria -- o lado certo de errar.
                Return Array.Empty(Of String)()
            End Try
        End Function

        ''' <summary>
        ''' Acrescenta uma linha, sem repetir. Devolve <c>False</c> quando não
        ''' deu para gravar — e quem chama <b>não</b> pode tratar isso como
        ''' sucesso: a linha continuaria na fila e o dono acharia que resolveu.
        ''' </summary>
        Private Shared Function Acrescentar(caminho As String, valor As String,
                                            cabecalho As IEnumerable(Of String)) As Boolean
            If String.IsNullOrWhiteSpace(valor) Then Return False

            Try
                Dim ja = Ler(caminho)
                If ja.Contains(valor.Trim(), StringComparer.OrdinalIgnoreCase) Then Return True

                Dim novo = Not File.Exists(caminho)
                Directory.CreateDirectory(Path.GetDirectoryName(caminho))
                Using escritor = New StreamWriter(caminho, append:=True, encoding:=Encoding.UTF8)
                    If novo Then
                        For Each linha In cabecalho
                            escritor.WriteLine(linha)
                        Next
                    End If
                    escritor.WriteLine(valor.Trim())
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function CabecalhoDasConversas() As String()
            Return {
                "# Conversas que voce marcou como 'nao exige resposta'.",
                "#",
                "# Elas somem da fila de respostas pendentes. Apague uma linha para",
                "# a conversa voltar a aparecer.",
                "#",
                "# Cada linha e um ConversationID do Outlook -- ilegivel de proposito,",
                "# porque e identificador e nao assunto.",
                "#"}
        End Function

        Private Shared Function CabecalhoDosRemetentes() As String()
            Return {
                "# Remetentes cujas mensagens normalmente nao exigem resposta.",
                "#",
                "# Conversa cuja ULTIMA mensagem e de um deles nao aparece na fila.",
                "# Serve para newsletter, notificacao automatica e afins.",
                "#",
                "# Um endereco por linha. Apague uma linha para voltar a ver.",
                "#"}
        End Function

    End Class

End Namespace

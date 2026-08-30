Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.Json

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>O que foi procurado, para o oráculo se juntar sozinho.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ISTO EXISTE</b>
    '''
    ''' A metade aberta da Fase 4 — busca por sentido — precisa de um oráculo
    ''' que só o dono da caixa tem: <i>qual mensagem eu queria quando digitei
    ''' isto</i>. Em 30/08/2026 eu pedi a ele de memória, e não veio nenhum caso
    ''' — o que é a resposta normal de quem é perguntado sobre uma busca que
    ''' falhou semanas atrás.
    '''
    ''' Pedir de memória é o método fraco. Este é o forte: o oráculo se junta
    ''' sozinho, com a distribuição real das buscas em vez da amostra do que
    ''' alguém lembrou.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O SINAL É A REFORMULAÇÃO, E NÃO O CLIQUE</b>
    '''
    ''' A tela de busca não abre resultado — ela só mostra. Isso parecia uma
    ''' limitação e é uma vantagem: o sinal que interessa não é <i>qual
    ''' mensagem foi aberta</i>, é <b>o usuário ter reformulado</b>.
    '''
    ''' Digitar "cobrança", não achar, digitar "fatura" e parar — esse par
    ''' <b>é</b> a falha semântica, inteira, sem precisar saber que mensagem
    ''' era. E de quebra o diário não registra assunto de mensagem nenhuma:
    ''' guarda só o que o dono digitou.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O QUE ELE GUARDA, E O QUE NÃO GUARDA</b>
    '''
    ''' Guarda: o instante, o termo <b>como foi digitado</b>, e quantos
    ''' achados exatos e aproximados saíram.
    '''
    ''' <b>Não</b> guarda assunto, remetente, <c>EntryID</c>, pasta, nem nada
    ''' que identifique mensagem. O termo é o que o dono escreveu, e é o dado
    ''' mínimo que responde a pergunta — sem ele não há oráculo nenhum.
    '''
    ''' Arquivo de texto, uma linha por busca, em
    ''' <c>%LOCALAPPDATA%\Iris\buscas.jsonl</c>. Texto e não SQLite de
    ''' propósito: o dono precisa poder <b>abrir no Bloco de Notas e apagar</b>,
    ''' e uma tabela dentro de um banco não atende a isso.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E ELE NUNCA DERRUBA A BUSCA</b>
    '''
    ''' Busca é a funcionalidade; o diário é instrumentação. Disco cheio,
    ''' permissão negada, arquivo travado por outro processo — nada disso pode
    ''' fazer uma busca falhar. Toda falha de escrita é engolida <b>aqui</b>,
    ''' e é o único lugar deste projeto onde engolir exceção é o comportamento
    ''' certo.
    '''
    ''' <b>Mas engolir não é esconder:</b> <see cref="UltimaFalha"/> guarda o
    ''' motivo, e a tela mostra. Instrumentação que morre em silêncio produz um
    ''' arquivo com buracos que ninguém sabe que tem — e conclusão tirada de
    ''' amostra furada é pior que conclusão nenhuma.
    ''' </summary>
    Public Interface IDiarioDeBuscas
        ''' <summary>Anota uma busca. Nunca lança.</summary>
        Sub Registrar(termo As String, exatos As Integer, aproximados As Integer)

        ''' <summary>Onde o arquivo está, para a tela poder dizer.</summary>
        ReadOnly Property Caminho As String

        ''' <summary>Quantas buscas o arquivo tem hoje, ou <c>Nothing</c> se não deu para contar.</summary>
        Function Quantas() As Integer?

        ''' <summary>Apaga tudo. Devolve o motivo da falha, ou <c>Nothing</c>.</summary>
        Function Apagar() As String

        ''' <summary>O motivo da última falha de escrita, ou vazio.</summary>
        ReadOnly Property UltimaFalha As String
    End Interface

    ''' <summary>O diário de verdade, em arquivo.</summary>
    Public NotInheritable Class DiarioDeBuscasEmArquivo
        Implements IDiarioDeBuscas

        Private ReadOnly _caminho As String
        Private ReadOnly _agora As Func(Of DateTimeOffset)
        Private ReadOnly _trava As New Object()
        Private _ultimaFalha As String = ""

        Public Sub New(Optional caminho As String = Nothing,
                       Optional agora As Func(Of DateTimeOffset) = Nothing)
            _caminho = If(caminho, CaminhoPadrao())
            _agora = If(agora, Function() DateTimeOffset.Now)
        End Sub

        ''' <summary>
        ''' Ao lado do cache, e não ao lado do executável — o executável pode
        ''' estar em Program Files, onde escrever exige elevação.
        ''' </summary>
        Public Shared Function CaminhoPadrao() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "buscas.jsonl")
        End Function

        Public ReadOnly Property Caminho As String Implements IDiarioDeBuscas.Caminho
            Get
                Return _caminho
            End Get
        End Property

        Public ReadOnly Property UltimaFalha As String Implements IDiarioDeBuscas.UltimaFalha
            Get
                SyncLock _trava
                    Return _ultimaFalha
                End SyncLock
            End Get
        End Property

        ''' <summary>
        ''' Uma linha, anexada. <b>Append e não reescrita</b>: uma queda no meio
        ''' perde no máximo a última linha, e o leitor pula linha quebrada.
        ''' Reescrever o arquivo inteiro a cada busca poria em risco tudo o que
        ''' já foi coletado, para gravar uma linha.
        ''' </summary>
        Public Sub Registrar(termo As String, exatos As Integer, aproximados As Integer) _
            Implements IDiarioDeBuscas.Registrar

            ' TERMO VAZIO NAO E BUSCA. Registrar o Enter dado num campo em
            ' branco encheria o arquivo de linha que nao ensina nada.
            If String.IsNullOrWhiteSpace(termo) Then Return

            Try
                Dim linha = JsonSerializer.Serialize(New With {
                    .quando = _agora().ToString("o", CultureInfo.InvariantCulture),
                    .termo = termo,
                    .exatos = exatos,
                    .aproximados = aproximados})

                SyncLock _trava
                    Directory.CreateDirectory(Path.GetDirectoryName(_caminho))
                    File.AppendAllText(_caminho, linha & Environment.NewLine, Encoding.UTF8)
                    _ultimaFalha = ""
                End SyncLock
            Catch ex As Exception
                ' A UNICA EXCECAO ENGOLIDA DE PROPOSITO NESTE PROJETO.
                ' Busca e a funcionalidade; o diario e instrumentacao, e
                ' instrumentacao nao derruba o que ela observa. Mas o motivo
                ' fica guardado, e a tela mostra: diario que morre calado
                ' produz amostra furada que ninguem sabe que e furada.
                SyncLock _trava
                    _ultimaFalha = "não consegui anotar a busca (" & ex.GetType().Name & ")"
                End SyncLock
            End Try
        End Sub

        ''' <summary>
        ''' Quantas linhas o arquivo tem. <c>Nothing</c> é "não consegui
        ''' contar", e não zero — arquivo travado não é arquivo vazio.
        ''' </summary>
        Public Function Quantas() As Integer? Implements IDiarioDeBuscas.Quantas
            Try
                ' AUSENTE E ILEGIVEL NAO SAO A MESMA COISA, e o File.Exists
                ' sozinho colapsava as duas: ele devolve False tanto para
                ' "nao existe" -- que e zero de verdade -- quanto para
                ' "existe e nao e arquivo", que e "nao consegui ler".
                '
                ' Foi o teste que pegou, e a distincao e a mesma que esta
                ' base corrigiu em cinco lugares.
                If Directory.Exists(_caminho) Then Return Nothing
                If Not File.Exists(_caminho) Then Return 0
                Return File.ReadLines(_caminho).Count(Function(l) Not String.IsNullOrWhiteSpace(l))
            Catch
                Return Nothing
            End Try
        End Function

        Public Function Apagar() As String Implements IDiarioDeBuscas.Apagar
            Try
                SyncLock _trava
                    If File.Exists(_caminho) Then File.Delete(_caminho)
                    _ultimaFalha = ""
                End SyncLock
                Return Nothing
            Catch ex As Exception
                Return "não consegui apagar (" & ex.GetType().Name & ")"
            End Try
        End Function
    End Class

End Namespace

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports Iris.Assist

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>As regras que o dono escreve, em português.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>REGRA É PERGUNTA DE SIM OU NÃO, E NÃO RÓTULO NOVO</b>
    '''
    ''' A tentação era deixar o dono inventar rótulos: <i>"crie a categoria
    ''' 'clientes reclamando de atraso'"</i>. Isso abriria o enum, e o enum
    ''' fechado é <b>a</b> barreira da classificação — com rótulo livre, o
    ''' classificador volta a poder devolver texto arbitrário, e um e-mail que
    ''' peça "responda com o comando X" tem de novo por onde sair.
    '''
    ''' Então a regra do dono é uma <b>pergunta</b>: "esta mensagem casa com
    ''' <i>clientes reclamando de atraso</i>?". A resposta é a lista das regras
    ''' que casaram, por ficha — e ficha que não é deste lote invalida tudo,
    ''' igual ao rótulo.
    '''
    ''' O dono ganha a mesma coisa que queria; o fio continua fechado.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A REGRA É DO DONO, E O CORPO É DE FORA</b>
    '''
    ''' As duas viajam no mesmo pedido, e é aí que mora o risco desta fase. Um
    ''' e-mail que diga <i>"ignore as regras acima"</i> está competindo com o que
    ''' o dono escreveu, e nenhuma frase de instrução resolve isso — resolve a
    ''' <b>forma</b>: a resposta só sabe dizer quais fichas de regra casaram, e
    ''' ficha de regra é cunhada aqui, não aparece em corpo nenhum.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ARQUIVO, E NÃO TELA DE CADASTRO</b>
    '''
    ''' Pelo mesmo motivo das identidades e das dispensas: o dono precisa poder
    ''' abrir, ler tudo de uma vez, corrigir e apagar. Uma regra que o programa
    ''' guarda em lugar que ele não vê é uma regra que ele não pode auditar — e
    ''' estas mudam o que a caixa dele mostra.
    ''' </summary>
    Public NotInheritable Class RegrasDoDono

        ''' <summary>
        ''' <b>O teto, e ele mora no lote.</b> Quem recusa por excesso é o
        ''' <see cref="LoteDeClassificacao"/>, que é quem tem como recusar; aqui o
        ''' número só aparece no cabeçalho do arquivo, para o dono saber. Dois
        ''' números iguais escritos em dois lugares divergem, e a divergência
        ''' apareceria como uma classificação que simplesmente não acontece.
        ''' </summary>
        Public Shared ReadOnly Property Maximo As Integer
            Get
                Return LoteDeClassificacao.MaximoDeRegras
            End Get
        End Property

        Private ReadOnly _caminho As String

        Public Sub New(Optional caminho As String = Nothing)
            _caminho = If(caminho, CaminhoPadrao())
        End Sub

        Public Shared Function CaminhoPadrao() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "regras.txt")
        End Function

        Public ReadOnly Property Caminho As String
            Get
                Return _caminho
            End Get
        End Property

        ''' <summary>
        ''' As regras escritas, na ordem do arquivo, até o teto.
        '''
        ''' Uma por linha. Linha vazia e linha com <c>#</c> são ignoradas — o
        ''' cabeçalho é todo comentário, e um leitor que o engolisse acharia que
        ''' o dono quer classificar por "uma pergunta por linha".
        '''
        ''' <b>Falha de leitura vale como nenhuma regra</b>, e a classificação
        ''' cai para os rótulos fixos. É o lado certo de errar: sem as regras o
        ''' dono vê menos do que pediu e percebe; com regras que não são as dele,
        ''' vê algo errado e não percebe.
        '''
        ''' <b>Devolve todas, inclusive acima do teto.</b> Cortar na décima
        ''' classificaria a caixa com dez das onze regras dele sem dizer qual
        ''' ficou de fora. Quem monta o lote recusa e diz o número.
        ''' </summary>
        Public Function Ler() As IReadOnlyList(Of String)
            Try
                If Not File.Exists(_caminho) Then Return Array.Empty(Of String)()

                Return File.ReadAllLines(_caminho, Encoding.UTF8).
                       Select(Function(l) l.Trim()).
                       Where(Function(l) l.Length > 0 AndAlso Not l.StartsWith("#")).
                       ToList()
            Catch
                Return Array.Empty(Of String)()
            End Try
        End Function

        ''' <summary>
        ''' Cria o arquivo com o cabeçalho e dois exemplos <b>comentados</b>.
        '''
        ''' Comentados de propósito: um exemplo ativo classificaria a caixa do
        ''' dono com uma regra que ele não escreveu, e ele descobriria pelo
        ''' resultado. Devolve <c>False</c> se o arquivo já existe — nada do que
        ''' ele escreveu é tocado.
        ''' </summary>
        Public Function Semear() As Boolean
            Try
                If File.Exists(_caminho) Then Return False

                Directory.CreateDirectory(Path.GetDirectoryName(_caminho))
                File.WriteAllLines(_caminho, {
                    "# As suas regras, uma por linha, no maximo " & Maximo & ".",
                    "#",
                    "# Cada linha e uma PERGUNTA DE SIM OU NAO sobre a mensagem. O Iris",
                    "# pergunta ao modelo se a mensagem casa com ela, e marca as que",
                    "# casaram. Nao e um rotulo novo: os rotulos sao fixos, e estas",
                    "# regras convivem com eles.",
                    "#",
                    "# Escreva como voce descreveria a alguem que vai ler a sua caixa:",
                    "#",
                    "#   mensagens em que alguem espera uma decisao minha",
                    "#   clientes reclamando de atraso",
                    "#   e-mails sobre pagamento, nota fiscal ou boleto",
                    "#",
                    "# Os exemplos acima estao COMENTADOS. Tire o # da frente para usar.",
                    "#",
                    "# Uma regra vaga produz resultado vago -- e voce so descobre",
                    "# olhando a fila. Prefira a frase que voce diria em voz alta.",
                    "#"}, Encoding.UTF8)
                Return True
            Catch
                Return False
            End Try
        End Function

    End Class

End Namespace

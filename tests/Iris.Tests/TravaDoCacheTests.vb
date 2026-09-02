Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>Quem usa a conexão compartilhada toma a trava.</b>
'''
''' ------------------------------------------------------------------
''' <b>ESTE DEFEITO JÁ VOLTOU TRÊS VEZES COM NOMES DIFERENTES</b>
'''
''' Uma <c>SqliteConnection</c> não tem contrato de uso simultâneo. O WAL coordena
''' <i>conexões</i>, e não torna uma conexão reentrante — foi essa confusão que
''' produziu a falha rara de 25/08/2026.
'''
''' <list type="number">
''' <item>31/08: "a trava que eu tinha posto não travava nada" — ela protegia a
''' recarga do acervo contra ela mesma, e não contra os outros dois caminhos.</item>
''' <item>01/09: a trava foi para dentro do <c>CacheDatabase</c>, e só o
''' <c>RotulosNoCache</c> passou a tomá-la.</item>
''' <item>02/09: uma revisão externa achou o escritor, o dreno, o sink da
''' varredura, o serviço do acervo e o <b>diário do egresso</b> usando a mesma
''' conexão sem ela — e um grep achou mais quatro que a revisão não citou.</item>
''' </list>
'''
''' Três consertos pontuais e três voltas. O que faltava era a regra ser
''' <b>verificável</b>, e não lembrada.
'''
''' ------------------------------------------------------------------
''' <b>A REGRA, E POR QUE ELA É TEXTUAL</b>
'''
''' Todo método que toca <c>.Connection</c> de um <c>CacheDatabase</c>, ou um campo
''' que veio dela, tem de estar sob <c>SyncLock</c>. Isso é uma propriedade de
''' <i>fluxo</i>, e verificá-la de verdade exigiria análise semântica; o que se lê
''' aqui é a forma. Ela deixa passar sutileza — uma trava que não é a do banco, um
''' <c>SyncLock</c> num ramo morto — e pega o caso que de fato acontece: alguém
''' escreve um método novo e esquece.
'''
''' <b>Um alarme, e não uma prova.</b> É o mesmo desenho do meta-teste do
''' paralelismo, e pelo mesmo motivo.
''' </summary>
<TestClass>
Public Class TravaDoCacheTests

    ''' <summary>
    ''' Arquivos que tocam a conexão e <b>não</b> a compartilham, com a razão.
    ''' Lista curta, e é para ficar curta: cada entrada aqui é um método que este
    ''' teste deixa de conferir.
    ''' </summary>
    Private Shared ReadOnly Isentos As New Dictionary(Of String, String)(
        StringComparer.OrdinalIgnoreCase) From {
        {"CacheDatabase.vb", "é quem PUBLICA a trava; travar-se a si mesmo seria circular"},
        {"Application.xaml.vb", "só passa a instância adiante, na abertura, antes de haver concorrência"},
        {"VarreduraDaPasta.vb", "abre conexão PRÓPRIA para a varredura — o motivo está no XML-doc dela"},
        {"AcervoViewModel.vb", "aponta o serviço e resolve a pasta; as duas rotas travam por dentro"}
    }

    ''' <summary>
    ''' <b>Todo uso de <c>.Connection</c> está sob <c>SyncLock</c>.</b>
    '''
    ''' A busca é por método: acha a linha do uso, sobe até a assinatura, e cobra um
    ''' <c>SyncLock</c> entre as duas.
    ''' </summary>
    <TestMethod>
    Public Sub Quem_usa_a_conexao_COMPARTILHADA_toma_a_trava()
        Dim soltos As New List(Of String)()
        Dim conferidos = 0

        For Each caminho In FontesDoProduto()
            Dim nome = Path.GetFileName(caminho)
            If Isentos.ContainsKey(nome) Then Continue For

            Dim linhas = File.ReadAllLines(caminho)
            For i = 0 To linhas.Length - 1
                If Not UsaAConexao(linhas(i)) Then Continue For
                conferidos += 1

                Dim onde = Envolvente(linhas, i)
                If onde.Travado Then Continue For

                ' AUXILIAR PRIVADO: a trava e do CHAMADOR.
                '
                ' Um Private so e alcancavel pela superficie da propria classe, e
                ' essa superficie ja foi conferida. Mas isso nao basta sozinho --
                ' um metodo publico novo poderia chamar o auxiliar sem travar e
                ' nunca tocar a conexao diretamente, ficando invisivel para a
                ' regra principal. Entao o auxiliar so passa se TODAS as chamadas
                ' a ele, no arquivo, estiverem sob trava.
                If onde.Privado AndAlso TodasAsChamadasSobTrava(linhas, onde.Nome) Then
                    Continue For
                End If

                soltos.Add($"{nome}:{i + 1} — «{linhas(i).Trim()}»")
            Next
        Next

        ' CONTROLE POSITIVO. Sem ele, um erro no caminho ou no padrão faria zero
        ' usos serem conferidos e o teste passaria dizendo nada.
        Assert.IsTrue(conferidos >= 8,
            $"controle: esperava ao menos 8 usos da conexão, conferi {conferidos}")

        Assert.AreEqual(0, soltos.Count,
            "uso da conexão compartilhada fora da trava — duas threads na mesma " &
            "SqliteConnection dão «transaction already active», leitor invalidado, " &
            "ou o registro do egresso falhando:" & Environment.NewLine &
            String.Join(Environment.NewLine, soltos))
    End Sub

    ''' <summary>
    ''' <b>Toda isenção tem razão escrita, e o arquivo existe.</b>
    '''
    ''' O par do teste acima. Sem ele, a lista de isentos vira o lugar onde se
    ''' esconde o que não se quis consertar — e cresceria calada.
    ''' </summary>
    <TestMethod>
    Public Sub Toda_isencao_tem_RAZAO_e_arquivo()
        Dim fontes = FontesDoProduto().Select(AddressOf Path.GetFileName).ToList()

        For Each par In Isentos
            Assert.IsTrue(par.Value.Length > 20,
                $"a isenção de {par.Key} não explica por quê")
            CollectionAssert.Contains(fontes, par.Key,
                $"{par.Key} está isento e não existe mais — isenção órfã " &
                "esconde o dia em que o arquivo voltar com outro conteúdo")
        Next

        Assert.IsTrue(Isentos.Count <= 6,
            "a lista de isentos cresceu: ela é o lugar onde a regra deixa de valer")
    End Sub

    ' ==================================================================
    ' O ANDAIME

    ''' <summary>
    ''' Só <c>.Connection</c> de banco. <c>HttpClient</c>, <c>Connection</c> de
    ''' outra coisa e menção em comentário ficam de fora.
    ''' </summary>
    Private Shared Function UsaAConexao(linha As String) As Boolean
        Dim limpa = SemComentario(linha)
        If Not limpa.Contains(".Connection") Then Return False
        Return Regex.IsMatch(limpa, "\b(_db|db|banco|cache)\.Connection\b",
                             RegexOptions.IgnoreCase)
    End Function

    ''' <summary>
    ''' Há <c>SyncLock</c> entre o começo do método e esta linha?
    '''
    ''' Sobe até a assinatura — <c>Function</c>, <c>Sub</c> ou <c>New</c> — e procura
    ''' no caminho. Construtor conta como isento: ele roda antes de a instância ser
    ''' compartilhada com thread nenhuma.
    ''' </summary>
    ''' <summary>
    ''' Sobe da linha do uso até a assinatura que a contém, dizendo se havia
    ''' <c>SyncLock</c> no caminho, se o método é privado, e como ele se chama.
    '''
    ''' Construtor conta como travado: ele roda antes de a instância ser
    ''' compartilhada com thread nenhuma.
    ''' </summary>
    Private Shared Function Envolvente(linhas As String(), uso As Integer) _
            As (Travado As Boolean, Privado As Boolean, Nome As String)
        For i = uso To 0 Step -1
            Dim t = linhas(i).Trim()
            If t.StartsWith("SyncLock ", StringComparison.OrdinalIgnoreCase) Then
                Return (True, False, "")
            End If
            If Regex.IsMatch(t, "^(Public|Private|Friend|Protected).*\bSub New\b") Then
                Return (True, False, "")
            End If

            Dim m = Regex.Match(t,
                "^(Public|Private|Friend|Protected|Shared)[\w ]*\b(Function|Sub)\s+(\w+)")
            If m.Success Then
                Return (False, m.Groups(1).Value = "Private", m.Groups(3).Value)
            End If
        Next
        Return (False, False, "")
    End Function

    ''' <summary>
    ''' <b>Toda chamada a este auxiliar chega, por alguma cadeia, a uma trava?</b>
    '''
    ''' É o que torna a isenção do <c>Private</c> <i>transitiva</i> em vez de
    ''' confiada. E a busca é <b>recursiva</b>: a cadeia real tem dois níveis —
    ''' método público travado → auxiliar → auxiliar. Uma versão de um nível só
    ''' acusaria o segundo auxiliar como solto, e a resposta seria pôr uma trava
    ''' redundante para calar o teste. Um teste que se cala assim ensina a
    ''' contorná-lo.
    '''
    ''' Auxiliar sem chamador nenhum não passa: código morto não é isenção.
    '''
    ''' <b>O ciclo para a recursão</b> — dois auxiliares privados que se chamem
    ''' não provam nada, e sem o conjunto de visitados isto giraria para sempre.
    ''' </summary>
    Private Shared Function TodasAsChamadasSobTrava(
            linhas As String(), nome As String,
            Optional vistos As HashSet(Of String) = Nothing) As Boolean

        If nome.Length = 0 Then Return False
        If vistos Is Nothing Then vistos = New HashSet(Of String)(StringComparer.Ordinal)
        If Not vistos.Add(nome) Then Return False

        Dim chamadas = 0
        For i = 0 To linhas.Length - 1
            Dim t = SemComentario(linhas(i))
            If Not Regex.IsMatch(t, "\b" & Regex.Escape(nome) & "\s*\(") Then Continue For
            ' A propria assinatura nao e chamada.
            If Regex.IsMatch(t.Trim(), "^(Public|Private|Friend|Protected|Shared)") Then
                Continue For
            End If

            chamadas += 1
            Dim onde = Envolvente(linhas, i)
            If onde.Travado Then Continue For
            If onde.Privado AndAlso
               TodasAsChamadasSobTrava(linhas, onde.Nome, vistos) Then Continue For
            Return False
        Next

        Return chamadas > 0
    End Function

    Private Shared Function SemComentario(linha As String) As String
        Dim corte = linha.IndexOf("'"c)
        Return If(corte >= 0, linha.Substring(0, corte), linha)
    End Function

    Private Shared Function FontesDoProduto() As IEnumerable(Of String)
        Dim raiz = Path.GetFullPath(Path.Combine(PastaDaSuite(), "..", "..", "src"))
        Assert.IsTrue(Directory.Exists(raiz), $"nao achei {raiz}")
        Return Directory.GetFiles(raiz, "*.vb", SearchOption.AllDirectories).
               Where(Function(f) Not f.Contains("\obj\") AndAlso Not f.Contains("\bin\"))
    End Function

    Private Shared Function PastaDaSuite(<CallerFilePath> Optional aqui As String = Nothing) As String
        Return Path.GetDirectoryName(aqui)
    End Function

End Class

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A marca da mutação fica colada no efeito irreversível.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ISTO É UM META-TESTE, E NÃO UM TESTE NORMAL</b>
'''
''' O que se quer garantir é <i>posicional</i>: a marca tem de estar imediatamente
''' antes da primeira coisa que fica no mundo — <c>Save</c>, <c>Send</c>,
''' <c>Delete</c>, <c>Add</c>, <c>SaveAsFile</c> — e não antes.
'''
''' Exercitar isso por chamada exigiria fazer o Outlook falhar num ponto
''' escolhido, que é justamente o que não se controla. A posição, essa se lê.
'''
''' ------------------------------------------------------------------
''' <b>O DEFEITO QUE ISTO IMPEDE DE VOLTAR</b>
'''
''' A marca era posta pelo <b>despacho</b>, antes de o trabalho começar. Uma
''' queda de conexão ao abrir o item — antes de existir qualquer <c>Send</c> —
''' virava <c>Ambiguous</c>, e o compositor dizia <i>"a mensagem pode ter sido
''' enviada, confira Itens Enviados"</i> sobre um envio que não começou.
'''
''' É o erro <b>simétrico</b> do pecado central deste projeto. Custa diferente e
''' custa: o dono procura uma mensagem que nunca existiu e aprende a não
''' acreditar no aviso — o que o torna inútil no dia em que ele for verdadeiro.
''' Achado por revisão externa em 01/09/2026.
''' </summary>
<TestClass>
Public Class MarcaDaMutacaoTests

    ''' <summary>
    ''' As chamadas que deixam efeito no mundo. Lista explícita: uma regra que
    ''' tentasse adivinhar "o que é irreversível" acusaria leitura como escrita.
    ''' </summary>
    Private Shared ReadOnly Efeitos As String() = {
        ".save()", ".send()", ".delete()", ".saveasfile(", ".add("
    }

    ''' <summary>Os módulos que escrevem no Outlook.</summary>
    Private Shared ReadOnly Escritores As String() = {
        "CalendarWriting.vb", "ContactWriting.vb", "TaskWriting.vb",
        "DraftWriting.vb", "MessageReading.vb"
    }

    ''' <summary>
    ''' <b>Toda marca é seguida de um efeito, na linha seguinte.</b>
    '''
    ''' Não "em algum lugar depois": na linha seguinte. Uma marca solta a três
    ''' linhas do efeito volta a abrir a janela que este conserto fechou — e
    ''' seria invisível numa leitura rápida.
    ''' </summary>
    <TestMethod>
    Public Sub Toda_marca_esta_COLADA_no_efeito()
        Dim soltas As New List(Of String)()
        Dim conferidas = 0

        For Each nome In Escritores
            Dim linhas = File.ReadAllLines(Caminho(nome))
            Dim dentroDoEnvolucro = False

            For i = 0 To linhas.Length - 1
                ' DENTRO DO ENVOLUCRO NAO SE COBRA POSICAO.
                '
                ' A chamada ao marcador que mora dentro do "Dim aoComecar As
                ' Action" e a DEFINICAO dele, e nao um acionamento: a linha
                ' seguinte e o End Sub, e cobra-la ali seria cobrar o efeito
                ' dentro da definicao. Quem tem posicao e o "aoComecar()", que e
                ' onde o efeito de fato acontece.
                If linhas(i).Trim().StartsWith("Dim aoComecar As Action") Then
                    dentroDoEnvolucro = True
                    Continue For
                End If
                If dentroDoEnvolucro Then
                    If linhas(i).Trim() = "End Sub" Then dentroDoEnvolucro = False
                    Continue For
                End If

                If Not EhAMarca(linhas(i)) Then Continue For
                conferidas += 1

                Dim seguinte = If(i + 1 < linhas.Length, linhas(i + 1).ToLowerInvariant(), "")
                If Not Efeitos.Any(Function(e) seguinte.Contains(e)) Then
                    soltas.Add($"{nome}:{i + 1} — a linha seguinte é «{seguinte.Trim()}»")
                End If
            Next
        Next

        ' CONTROLE POSITIVO. Sem ele, um erro no caminho ou no padrão faria zero
        ' marcas serem conferidas e o teste passaria dizendo nada.
        Assert.IsTrue(conferidas >= 10,
            $"controle: esperava ao menos 10 marcas, conferi {conferidas}")

        Assert.AreEqual(0, soltas.Count,
            "marca de mutação longe do efeito — a falha entre as duas viraria " &
            "ambiguidade inventada: " & Environment.NewLine &
            String.Join(Environment.NewLine, soltas))
    End Sub

    ''' <summary>
    ''' <b>Todo escritor que recebe o marcador o aciona.</b>
    '''
    ''' O par do teste acima: sem ele, apagar a chamada de dentro de uma função
    ''' deixaria a mutação nunca ser marcada — e aí uma falha <i>depois</i> do
    ''' <c>Send</c> viraria "nada saiu", que é o erro oposto e o pior dos dois.
    ''' </summary>
    <TestMethod>
    Public Sub Todo_escritor_que_RECEBE_o_marcador_o_aciona()
        Dim mudas As New List(Of String)()
        Dim conferidas = 0

        For Each nome In Escritores
            Dim texto = File.ReadAllText(Caminho(nome))
            For Each f In Funcoes(texto)
                If Not f.Cabecalho.Contains("marcar As Action") Then Continue For
                conferidas += 1
                ' ACIONAR OU REPASSAR. SaveAttachment nao aciona: ela entrega o
                ' marcador a GravarComTemporario, que e quem sabe onde o arquivo
                ' comeca a existir. Exigir o acionamento aqui obrigaria a marcar
                ' cedo demais -- o oposto do que este arquivo cobra.
                Dim usa = f.Corpo.Contains("Then marcar()") OrElse
                          f.Corpo.Contains("aoComecar()") OrElse
                          f.Corpo.Contains("(marcar,")
                If Not usa Then mudas.Add($"{nome}:{f.Nome}")
            Next
        Next

        Assert.IsTrue(conferidas >= 10,
            $"controle: esperava ao menos 10 escritores com marcador, achei {conferidas}")
        Assert.AreEqual(0, mudas.Count,
            "escritor que recebe o marcador e nunca o aciona: uma falha depois do " &
            "efeito viraria «nada aconteceu» — " & String.Join(", ", mudas))
    End Sub

    ''' <summary>
    ''' <b>O despacho não marca mais.</b>
    '''
    ''' É a metade que dava o defeito, e ela é conferida por nome: se voltar um
    ''' <c>fase.Marcar()</c> ao <c>RunAsync</c>, tudo o que está acima deixa de
    ''' valer sem que nenhum outro teste perceba.
    ''' </summary>
    <TestMethod>
    Public Sub O_DESPACHO_nao_marca()
        Dim texto = File.ReadAllText(Caminho("OutlookBroker.vb"))
        Dim dentro = Regex.Match(texto,
            "Private Async Function RunAsync.*?\r?\n        End Function",
            RegexOptions.Singleline)

        Assert.IsTrue(dentro.Success, "controle: nao achei o RunAsync")
        Assert.IsFalse(dentro.Value.Contains("fase.Marcar()"),
            "o DESPACHO voltou a marcar a mutação como iniciada antes de ela " &
            "começar — falha ao abrir o item volta a virar «pode ter sido enviada»")
    End Sub

    ' ==================================================================
    ' O ANDAIME

    ''' <summary>
    ''' <b>Esta linha aciona a marca?</b> Há duas formas, e a segunda apareceu
    ''' quando os escritores passaram a precisar saber <i>se</i> o efeito começou.
    '''
    ''' <c>aoComecar()</c> é um envolucro local que liga um sinalizador e chama o
    ''' marcador. Aceitá-lo aqui só vale porque o teste abaixo cobra que o
    ''' envolucro chame mesmo o marcador — senão bastaria batizar qualquer coisa
    ''' de <c>aoComecar</c> para calar este arquivo.
    ''' </summary>
    Private Shared Function EhAMarca(linha As String) As Boolean
        Dim t = linha.Trim()
        Return t = "aoComecar()" OrElse linha.Contains("Then marcar()")
    End Function

    ''' <summary>
    ''' <b>Todo envolucro <c>aoComecar</c> chama o marcador de verdade.</b>
    '''
    ''' Sem isto, a forma nova seria uma porta: um <c>Sub aoComecar</c> que só
    ''' ligasse o sinalizador local satisfaria os dois testes acima e a mutação
    ''' nunca seria marcada — e uma falha depois do efeito viraria "nada
    ''' aconteceu", que é o pior dos dois erros.
    ''' </summary>
    <TestMethod>
    Public Sub Todo_envolucro_chama_o_MARCADOR()
        Dim mudos As New List(Of String)()
        Dim conferidos = 0

        For Each nome In Escritores
            Dim linhas = File.ReadAllLines(Caminho(nome))
            For i = 0 To linhas.Length - 1
                If Not linhas(i).Trim().StartsWith("Dim aoComecar As Action") Then Continue For
                conferidos += 1

                ' O corpo do envolucro vai ate o End Sub que o fecha.
                Dim fim = i
                While fim < linhas.Length AndAlso linhas(fim).Trim() <> "End Sub"
                    fim += 1
                End While

                Dim corpo = String.Join(" ", linhas.Skip(i).Take(fim - i + 1))
                If Not corpo.Contains("marcar()") Then mudos.Add($"{nome}:{i + 1}")
            Next
        Next

        Assert.IsTrue(conferidos >= 4,
            $"controle: esperava ao menos 4 envolucros, achei {conferidos}")
        Assert.AreEqual(0, mudos.Count,
            "envolucro que liga o sinalizador e NAO marca a mutação: uma falha " &
            "depois do efeito viraria «nada aconteceu» — " & String.Join(", ", mudos))
    End Sub

    Private Shared Function Caminho(nome As String) As String
        Dim raiz = Path.GetFullPath(Path.Combine(PastaDaSuite(), "..", ".."))
        Dim achado = Path.Combine(raiz, "src", "Iris.Outlook", nome)
        Assert.IsTrue(File.Exists(achado), $"nao achei {achado}")
        Return achado
    End Function

    Private Shared Function PastaDaSuite(<CallerFilePath> Optional aqui As String = Nothing) As String
        Return Path.GetDirectoryName(aqui)
    End Function

    ''' <summary>
    ''' Parte o arquivo em funções, para o cabeçalho de uma não ser creditado a
    ''' outra. Cabeçalho é da linha da assinatura até o <c>) As</c> que a fecha.
    ''' </summary>
    Private Shared Iterator Function Funcoes(texto As String) _
            As IEnumerable(Of (Nome As String, Cabecalho As String, Corpo As String))
        Dim marcas = Regex.Matches(texto, "^\s*(?:Public|Private|Friend) Function (\w+)\(",
                                   RegexOptions.Multiline)
        For i = 0 To marcas.Count - 1
            Dim ini = marcas(i).Index
            Dim fim = If(i + 1 < marcas.Count, marcas(i + 1).Index, texto.Length)
            Dim corpo = texto.Substring(ini, fim - ini)

            ' O cabecalho vai ate a primeira linha que fecha a assinatura. Uma
            ' assinatura de tres linhas so declara "marcar" na segunda.
            Dim cabecalho = corpo
            Dim fechou = corpo.IndexOf(") _", StringComparison.Ordinal)
            Dim direto = corpo.IndexOf(") As ", StringComparison.Ordinal)
            If fechou >= 0 AndAlso (direto < 0 OrElse fechou < direto) Then
                cabecalho = corpo.Substring(0, fechou)
            ElseIf direto >= 0 Then
                cabecalho = corpo.Substring(0, direto)
            End If

            Yield (marcas(i).Groups(1).Value, cabecalho, corpo)
        Next
    End Function

End Class

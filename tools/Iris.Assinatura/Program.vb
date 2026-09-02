Imports System.Collections.Generic
Imports System.IO

Namespace Global.Iris.Assinatura

    ''' <summary>
    ''' <b>A casca de linha de comando.</b> Chamada pelos dois scripts de
    ''' <c>tools/</c>; não faz nada além de ler argumentos, chamar o
    ''' <see cref="Assinador"/> e gravar arquivo.
    '''
    ''' <b>A chave privada nunca é impressa.</b> Ela vai direto do gerador para o
    ''' arquivo de destino. Passá-la pela saída padrão a deixaria no histórico do
    ''' console, no buffer de rolagem e em qualquer transcrição de sessão.
    '''
    ''' <b>E nunca é aceita por argumento</b>, só por caminho de arquivo: a linha
    ''' de comando de um processo é legível por outros processos do mesmo
    ''' usuário e costuma ir para o log de auditoria do Windows.
    ''' </summary>
    Public Module Program

        Private Const Uso As String =
            "uso:" & vbLf &
            "  Iris.Assinatura gerar   --destino <chave.pem>" & vbLf &
            "  Iris.Assinatura assinar --chave <chave.pem> --arquivo <iris.json>" & vbLf &
            "  Iris.Assinatura publica --chave <chave.pem>"

        Public Function Main(argumentos As String()) As Integer
            Try
                Dim opcoes = Ler(argumentos)
                Select Case If(argumentos.FirstOrDefault(), "")
                    Case "gerar" : Return Gerar(Exigir(opcoes, "destino"))
                    Case "assinar" : Return AssinarArquivo(Exigir(opcoes, "chave"),
                                                           Exigir(opcoes, "arquivo"))
                    Case "publica" : Return Publica(Exigir(opcoes, "chave"))
                    Case Else
                        Console.Error.WriteLine(Uso)
                        Return 2
                End Select
            Catch ex As Exception
                ' SO A MENSAGEM, sem pilha: quem le isto e o dono publicando uma
                ' versao, e a pilha nao lhe diz nada que a mensagem nao diga.
                Console.Error.WriteLine("erro: " & ex.Message)
                Return 1
            End Try
        End Function

        Private Function Gerar(destino As String) As Integer
            ' RECUSA SOBRESCREVER. Gerar por cima invalida todas as versoes ja
            ' publicadas -- as copias do Iris que estao por ai tem a chave
            ' publica antiga embutida e recusam tudo o que for assinado com a
            ' nova. Isso nao pode acontecer por um comando repetido.
            If File.Exists(destino) AndAlso New FileInfo(destino).Length > 0 Then
                Throw New InvalidOperationException(
                    $"já existe uma chave em {destino}; apague-a antes se é isso mesmo que quer")
            End If

            Dim publicaBase64 As String = Nothing
            Dim privada = Assinador.GerarPar(publicaBase64)

            ' O arquivo pode ja existir VAZIO: o script o cria antes, para
            ' aplicar a ACL restritiva enquanto ele ainda nao tem conteudo.
            ' Gravar por cima preserva a ACL de um arquivo existente.
            File.WriteAllText(destino, privada)

            Console.WriteLine(publicaBase64)
            Return 0
        End Function

        Private Function AssinarArquivo(chave As String, arquivo As String) As Integer
            Dim assinatura = Assinador.Assinar(File.ReadAllText(chave),
                                               File.ReadAllBytes(arquivo))
            File.WriteAllBytes(arquivo & ".sig", assinatura)
            Console.WriteLine(assinatura.Length)
            Return 0
        End Function

        Private Function Publica(chave As String) As Integer
            Console.WriteLine(Assinador.PublicaDe(File.ReadAllText(chave)))
            Return 0
        End Function

        ''' <summary>
        ''' Lê pares <c>--nome valor</c>. O verbo é o primeiro argumento e não
        ''' entra aqui.
        ''' </summary>
        Private Function Ler(argumentos As String()) As Dictionary(Of String, String)
            Dim lidas As New Dictionary(Of String, String)(StringComparer.Ordinal)
            Dim i = 1
            While i < argumentos.Length - 1
                If argumentos(i).StartsWith("--", StringComparison.Ordinal) Then
                    lidas(argumentos(i).Substring(2)) = argumentos(i + 1)
                    i += 2
                Else
                    i += 1
                End If
            End While
            Return lidas
        End Function

        Private Function Exigir(opcoes As Dictionary(Of String, String),
                                nome As String) As String
            Dim valor As String = Nothing
            If Not opcoes.TryGetValue(nome, valor) OrElse valor.Length = 0 Then
                Throw New ArgumentException($"falta --{nome}" & vbLf & Uso)
            End If
            Return valor
        End Function

    End Module

End Namespace

Imports System.Security.Cryptography

Namespace Global.Iris.App

    ''' <summary>
    ''' <b>A chave pública que decide de quem o Iris aceita atualização — e o
    ''' endereço onde ele pergunta.</b>
    '''
    ''' As duas coisas moram juntas porque só valem juntas: um endereço sem chave
    ''' é um lugar de onde baixar qualquer coisa, e uma chave sem endereço não
    ''' confere nada.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ENQUANTO ISTO ESTIVER VAZIO, NÃO HÁ VERIFICAÇÃO DE VERSÃO</b>
    '''
    ''' E é assim de propósito. A alternativa — sair perguntando com a chave
    ''' vazia — faria toda resposta virar "a assinatura não confere", que é a
    ''' frase de <i>alguém trocou o arquivo</i> dita quando o que houve foi
    ''' <i>ninguém configurou ainda</i>. São coisas diferentes e a tela tem de
    ''' dizer coisas diferentes.
    '''
    ''' Para preencher: rode <c>tools/gerar-chave-de-assinatura.ps1</c>, uma vez,
    ''' e cole aqui a chave pública que ele imprime. A privada fica com você e
    ''' não entra neste repositório.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE EMBUTIDA, E NÃO NUM ARQUIVO AO LADO</b>
    '''
    ''' Uma chave num arquivo de configuração é uma chave que quem trocar o
    ''' arquivo passa a controlar — e aí a assinatura confere, e não prova nada.
    ''' Ela tem de vir de dentro do executável, junto com o código que a usa.
    ''' </summary>
    Friend NotInheritable Class ChaveDeAtualizacao

        ''' <summary>
        ''' A chave pública em Base64, no formato SubjectPublicKeyInfo — que é o
        ''' que <c>ExportSubjectPublicKeyInfo</c> escreve e
        ''' <c>ImportSubjectPublicKeyInfo</c> lê.
        ''' </summary>
        Public Const PublicaBase64 As String = ""

        ''' <summary>
        ''' Onde o manifesto assinado é publicado. Nas releases do GitHub,
        ''' <c>/releases/latest/download/</c> é um endereço estável que sempre
        ''' aponta para a última — não é preciso saber o número da versão para
        ''' perguntar qual é a última versão.
        ''' </summary>
        Public Const Endereco As String = ""

        ''' <summary>
        ''' <c>False</c> antes de a chave e o endereço serem preenchidos. A tela
        ''' usa isto para dizer que atualizações ainda não foram configuradas, em
        ''' vez de acusar uma assinatura que ninguém tentou fazer.
        ''' </summary>
        Public Shared ReadOnly Property Configurada As Boolean
            Get
                Return PublicaBase64.Length > 0 AndAlso
                       Endereco.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        ''' <summary>
        ''' Os bytes da chave, ou um vetor vazio se ela não estiver preenchida ou
        ''' não for Base64 — e um vetor vazio faz toda conferência recusar, que é
        ''' o desfecho certo para uma chave que não dá para ler.
        ''' </summary>
        Public Shared Function Bytes() As Byte()
            If PublicaBase64.Length = 0 Then Return Array.Empty(Of Byte)()
            Try
                Dim lidos = Convert.FromBase64String(PublicaBase64)

                ' E ELA TEM DE SER UMA CHAVE. Base64 valido nao e chave valida, e
                ' descobrir isso aqui da uma mensagem util; descobrir la dentro
                ' da "a assinatura nao confere", que manda procurar no lugar
                ' errado.
                Using verificador = ECDsa.Create()
                    Dim consumidos = 0
                    verificador.ImportSubjectPublicKeyInfo(lidos, consumidos)

                    ' A SPKI INTEIRA, E P-256. ImportSubjectPublicKeyInfo ignora o
                    ' que sobrar depois da chave, e aceita qualquer curva. As duas
                    ' coisas passariam calado aqui e reapareceriam como "a assinatura
                    ' nao confere" na tela -- mandando procurar no lugar errado.
                    If consumidos <> lidos.Length Then Return Array.Empty(Of Byte)()
                    If verificador.KeySize <> 256 Then Return Array.Empty(Of Byte)()
                End Using

                Return lidos
            Catch
                Return Array.Empty(Of Byte)()
            End Try
        End Function

    End Class

End Namespace

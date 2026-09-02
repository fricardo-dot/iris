Imports System.Security.Cryptography

Namespace Global.Iris.Assinatura

    ''' <summary>
    ''' <b>O par de chaves e a assinatura de um manifesto.</b>
    '''
    ''' Separado do <c>Program</c> para poder ser chamado pela suíte: o teste que
    ''' importa é o de ida e volta — assinar aqui e ler com
    ''' <c>ManifestoDeVersao.Ler</c>. Sem ele, "as duas pontas usam o mesmo
    ''' formato" é uma afirmação sobre a documentação da plataforma, e não sobre
    ''' este programa.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>P-256, E CONFERIDO NOS DOIS LADOS</b>
    '''
    ''' A curva é fixada na geração e <b>exigida</b> na leitura. Sem exigir,
    ''' uma chave de outra curva funcionaria — e o desenho escrito diz P-256.
    ''' Um desenho que o código não faz cumprir não é desenho, é intenção.
    ''' </summary>
    Public NotInheritable Class Assinador

        ''' <summary>Quantos bytes tem uma chave pública P-256 em SPKI.</summary>
        Private Const TamanhoDaPublicaP256 As Integer = 91

        ''' <summary>
        ''' Gera o par. Devolve a chave privada em PEM (PKCS#8) e a pública em
        ''' Base64 (SubjectPublicKeyInfo) em <paramref name="publicaBase64"/>.
        ''' </summary>
        Public Shared Function GerarPar(ByRef publicaBase64 As String) As String
            ' "novo", e nao "ecdsa": um local chamado ecdsa eclipsaria o TIPO
            ' ECDsa, porque VB e insensivel a maiusculas. CLAUDE.md, secao 1.
            Using novo = ECDsa.Create(ECCurve.NamedCurves.nistP256)
                publicaBase64 = Convert.ToBase64String(novo.ExportSubjectPublicKeyInfo())
                Return novo.ExportPkcs8PrivateKeyPem()
            End Using
        End Function

        ''' <summary>
        ''' Assina <paramref name="dados"/> com a chave privada em PEM, e
        ''' <b>confere a própria assinatura antes de devolvê-la</b>.
        '''
        ''' Assinar e não verificar deixa um par trocado ser descoberto lá na
        ''' frente, na máquina de destino, com a mensagem "este arquivo não foi
        ''' publicado por você" — dita sobre um arquivo que foi.
        ''' </summary>
        Public Shared Function Assinar(privadaEmPem As String, dados As Byte()) As Byte()
            Using dono = ECDsa.Create()
                dono.ImportFromPem(privadaEmPem)
                ExigirP256(dono)

                Dim feita = dono.SignData(dados, HashAlgorithmName.SHA256)
                If Not dono.VerifyData(dados, feita, HashAlgorithmName.SHA256) Then
                    Throw New CryptographicException(
                        "a assinatura não confere contra a própria chave que a fez")
                End If
                Return feita
            End Using
        End Function

        ''' <summary>
        ''' A chave pública correspondente a uma privada, em Base64. Serve para o
        ''' dono reimprimir a pública sem gerar par novo — gerar outro invalidaria
        ''' todas as versões já publicadas.
        ''' </summary>
        Public Shared Function PublicaDe(privadaEmPem As String) As String
            Using dono = ECDsa.Create()
                dono.ImportFromPem(privadaEmPem)
                ExigirP256(dono)
                Return Convert.ToBase64String(dono.ExportSubjectPublicKeyInfo())
            End Using
        End Function

        Private Shared Sub ExigirP256(quem As ECDsa)
            If quem.KeySize <> 256 Then
                Throw New CryptographicException(
                    $"a chave tem {quem.KeySize} bits; o Iris só confere P-256")
            End If

            ' O TAMANHO DA SPKI, e nao so o KeySize: ha outras curvas de 256
            ' bits, e o cliente traz uma chave P-256 embutida. Uma curva
            ' diferente produziria assinatura que nenhuma copia do Iris confere,
            ' e o erro apareceria como "nao foi publicado por voce".
            '
            ' ESTA LINHA NAO TEM CONTROLE NEGATIVO, e vale dize-lo. Apaga-la
            ' deixa a suite verde: o unico teste de curva errada usa P-384, que
            ' o KeySize acima ja recusa. Provar esta aqui exigiria uma curva de
            ' 256 bits que nao a P-256 -- brainpool, por exemplo -- e a
            ' disponibilidade dela depende do CNG da maquina, o que faria o
            ' teste passar aqui e falhar noutro lugar por motivo nenhum.
            '
            ' Fica como defesa em profundidade DECLARADA. A alternativa era
            ' apaga-la por nao ser testavel, e ela custa uma comparacao no
            ' caminho que decide de quem o Iris aceita codigo.
            If quem.ExportSubjectPublicKeyInfo().Length <> TamanhoDaPublicaP256 Then
                Throw New CryptographicException("a chave não é da curva P-256")
            End If
        End Sub

    End Class

End Namespace

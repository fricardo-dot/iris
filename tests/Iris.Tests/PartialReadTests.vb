Imports Iris.Core
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' O que uma leitura incompleta impede.
'''
''' O caso perigoso não é a leitura que FALHA — essa aparece. É a que vem
''' pela metade: três destinatários certos na tela, aprovados por quem
''' conferiu, e a mensagem indo para menos gente do que deveria. O que falta
''' é invisível por definição.
'''
''' Por isso a regra não é visual. Sinalizar com um ícone e deixar tudo
''' funcionando seria o pior dos mundos: um aviso que ninguém sabe
''' interpretar, e a ação irreversível liberada do mesmo jeito.
''' </summary>
<TestClass>
Public Class PartialReadTests

    ' ================================================================
    ' O estado de cada parte
    ' ================================================================

    ''' <summary>
    ''' A distinção que motivou tudo: lista vazia porque não tem, contra
    ''' lista vazia porque não deu para ler. Antes as duas eram a mesma
    ''' coisa na tela.
    ''' </summary>
    <TestMethod>
    Public Sub Vazio_completo_e_diferente_de_vazio_por_falha()
        Dim naoTem = PartStatus.CompleteWith(0)
        Dim naoLeu = PartStatus.Missing(ErrorKind.Denied)

        Assert.IsTrue(naoTem.IsTrustworthy, "Zero destinatários lidos de zero é informação completa.")
        Assert.IsFalse(naoLeu.IsTrustworthy)
        Assert.AreEqual(ErrorKind.Denied, naoLeu.Reason)
    End Sub

    ''' <summary>
    ''' Incompleto NÃO é confiável, e este é o ponto todo. Há conteúdo na
    ''' tela, ele está certo, e mesmo assim agir sobre ele é agir sobre
    ''' menos do que existe.
    ''' </summary>
    <TestMethod>
    Public Sub Incompleto_nao_e_confiavel_mesmo_com_conteudo()
        Dim tresDeCinco = PartStatus.IncompleteWith(5, 3, ErrorKind.Denied)

        Assert.AreEqual(PartState.Incomplete, tresDeCinco.State)
        Assert.IsFalse(tresDeCinco.IsTrustworthy,
            "Três endereços certos continuam sendo menos que cinco.")
        Assert.AreEqual(5, tresDeCinco.Expected)
        Assert.AreEqual(3, tresDeCinco.Obtained)
    End Sub

    <TestMethod>
    Public Sub Completo_e_confiavel()
        Assert.IsTrue(PartStatus.Full.IsTrustworthy)
        Assert.IsTrue(PartStatus.CompleteWith(7).IsTrustworthy)
    End Sub

    ' ================================================================
    ' O que fica bloqueado
    ' ================================================================

    <TestMethod>
    Public Sub Destinatarios_incompletos_bloqueiam_responder()
        Assert.IsFalse(ReplyReadiness.CanReply(PartStatus.IncompleteWith(5, 3, ErrorKind.Denied)),
            "Responder a todos responderia a menos gente do que a mensagem tem.")
        Assert.IsFalse(ReplyReadiness.CanReply(PartStatus.Missing(ErrorKind.Denied)))
        Assert.IsFalse(ReplyReadiness.CanReply(Nothing))
    End Sub

    ''' <summary>
    ''' Controle negativo: leitura completa LIBERA. Sem isto, uma regra que
    ''' bloqueasse tudo passaria — e o Iris não responderia mensagem nenhuma.
    ''' </summary>
    <TestMethod>
    Public Sub Destinatarios_completos_liberam_responder()
        Assert.IsTrue(ReplyReadiness.CanReply(PartStatus.CompleteWith(5)))
        Assert.IsTrue(ReplyReadiness.CanReply(PartStatus.CompleteWith(0)),
            "Mensagem sem destinatário lido por inteiro ainda é leitura confiável.")
    End Sub

    <TestMethod>
    Public Sub Anexos_incompletos_bloqueiam_encaminhar()
        Assert.IsFalse(ReplyReadiness.CanForward(PartStatus.IncompleteWith(3, 1, ErrorKind.Busy)))
        Assert.IsFalse(ReplyReadiness.CanForward(PartStatus.Missing(ErrorKind.Denied)))
        Assert.IsTrue(ReplyReadiness.CanForward(PartStatus.CompleteWith(3)))
    End Sub

    ''' <summary>
    ''' Corpo incompleto NÃO bloqueia responder, e isso é uma escolha.
    '''
    ''' O corpo aparece na tela: um texto truncado é perceptível, e
    ''' responder a uma mensagem que se leu pela metade é decisão de quem
    ''' responde. A lista de destinatários é o oposto — é insumo que
    ''' ninguém confere linha a linha antes de clicar.
    ''' </summary>
    <TestMethod>
    Public Sub Corpo_incompleto_nao_bloqueia_responder()
        Assert.IsFalse(ReplyReadiness.BodyBlocksReply(
            PartStatus.IncompleteWith(1, 0, ErrorKind.NotDownloaded)))
        Assert.IsFalse(ReplyReadiness.BodyBlocksReply(PartStatus.Missing(ErrorKind.Denied)))
    End Sub

    ' ================================================================
    ' O que o usuário lê
    ' ================================================================

    ''' <summary>
    ''' Aviso sem consequência clara é ruído; consequência sem explicação é
    ''' frustração. A mensagem tem de dizer QUANTO faltou.
    ''' </summary>
    <TestMethod>
    Public Sub A_explicacao_diz_quanto_faltou()
        Dim texto = ReplyReadiness.DescribeBlock("os destinatários",
                                                 PartStatus.IncompleteWith(5, 3, ErrorKind.Denied))

        StringAssert.Contains(texto, "3")
        StringAssert.Contains(texto, "5")
        StringAssert.Contains(texto, "destinatários")
    End Sub

    <TestMethod>
    Public Sub Leitura_completa_nao_gera_aviso()
        Assert.AreEqual("", ReplyReadiness.DescribeBlock("os destinatários", PartStatus.Full))
        Assert.AreEqual("", ReplyReadiness.DescribeBlock("os anexos", PartStatus.CompleteWith(2)))
    End Sub

    <TestMethod>
    Public Sub Parte_indisponivel_diz_que_nao_deu_para_ler()
        Dim texto = ReplyReadiness.DescribeBlock("os anexos", PartStatus.Missing(ErrorKind.Denied))

        Assert.AreNotEqual("", texto)
        StringAssert.Contains(texto, "anexos")
    End Sub


    ' ================================================================
    ' Contagem dupla
    ' ================================================================

    ''' <summary>
    ''' Contagem que mudou no meio do percurso invalida o snapshot.
    '''
    ''' Nao existe prova de completude em cima de uma colecao COM que muda
    ''' sozinha. O que da para fazer e fechar para o lado seguro: se a
    ''' colecao tinha 5 e passou a ter 6, os 5 lidos nao sao "todos".
    ''' </summary>
    <TestMethod>
    Public Sub Contagem_que_muda_no_meio_nao_e_completa()
        Dim mudou = PartStatus.FromCounts(esperadoAntes:=5, esperadoDepois:=6,
                                          obtidos:=5, ultimaFalha:=ErrorKind.None)

        Assert.IsFalse(mudou.IsTrustworthy,
            "Cinco de cinco, mas a colecao virou seis: o snapshot nao vale.")
        Assert.AreEqual(PartState.Unavailable, mudou.State)
    End Sub

    ''' <summary>
    ''' Controle negativo: contagem estavel e tudo lido E completo. Sem
    ''' isto, uma regra que nunca aprovasse passaria.
    ''' </summary>
    <TestMethod>
    Public Sub Contagem_estavel_com_tudo_lido_e_completa()
        Dim ok = PartStatus.FromCounts(5, 5, 5, ErrorKind.None)
        Assert.IsTrue(ok.IsTrustworthy)
        Assert.AreEqual(5, ok.Obtained)
    End Sub

    <TestMethod>
    Public Sub Contagem_estavel_com_item_faltando_e_incompleta()
        Dim faltou = PartStatus.FromCounts(5, 5, 3, ErrorKind.Denied)
        Assert.AreEqual(PartState.Incomplete, faltou.State)
        Assert.AreEqual(3, faltou.Obtained)
        Assert.AreEqual(5, faltou.Expected)
    End Sub

    ''' <summary>
    ''' As fabricas recusam estados que nao fazem sentido. Sem estas
    ''' guardas dava para fabricar um "incompleto" que na verdade e
    ''' completo, ou um completo sem prova nenhuma.
    ''' </summary>
    <TestMethod>
    Public Sub As_fabricas_recusam_o_que_nao_faz_sentido()
        Assert.ThrowsException(Of ArgumentOutOfRangeException)(
            Function() PartStatus.IncompleteWith(expected:=3, obtained:=5, reason:=ErrorKind.Denied))

        Assert.ThrowsException(Of ArgumentException)(
            Function() PartStatus.IncompleteWith(5, 3, ErrorKind.None))

        Assert.ThrowsException(Of ArgumentException)(
            Function() PartStatus.Missing(ErrorKind.None))
    End Sub

    ''' <summary>
    ''' Nothing e tratado como NAO confiavel. E o que torna o default
    ''' ausente seguro: quem esquecer de preencher bloqueia, em vez de
    ''' declarar completude que nao provou.
    ''' </summary>
    <TestMethod>
    Public Sub Status_ausente_nao_e_confiavel()
        Assert.IsFalse(ReplyReadiness.CanReply(Nothing))
        Assert.IsFalse(ReplyReadiness.CanForward(Nothing))
        Assert.AreEqual("", ReplyReadiness.DescribeBlock("os anexos", Nothing))
    End Sub

End Class

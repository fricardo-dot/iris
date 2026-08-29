Imports System.Collections.Generic
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports Iris.App.ViewModels
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>GUARDAS DO LEITOR DE MENSAGEM — a primeira suite deste ViewModel.</b>
'''
''' ------------------------------------------------------------------
''' O relatorio da Fase 2 chamou esta de "a unica linha que precisa de arquivo
''' de teste novo: nao existe suite para o leitor". Era literal — o
''' <c>MessageDetailViewModel</c> nunca tinha sido instanciado por um teste.
'''
''' Sao tres caminhos, e todos tem a mesma forma: <b>uma operacao lenta volta
''' quando o mundo ja mudou</b>. O que muda e o que mudou — a mensagem, ou a
''' janela.
'''
''' ------------------------------------------------------------------
''' <b>O EFEITO NAO SE DESFAZ; A TELA E QUE PARA DE SER ESCRITA</b>
'''
''' Vale para os tres. O anexo pode ter sido gravado, a marcacao pode ter
''' valido — cancelar nao desfaz escrita em disco nem chamada ao Outlook que
''' ja saiu. O que as guardas impedem e <b>anunciar o desfecho no lugar
''' errado</b>: no leitor de outra mensagem, ou num leitor que ja saiu da
''' tela.
''' </summary>
<TestClass>
<DoNotParallelize>
Public Class LeitorDeMensagemTests

    Private Shared Function Chave(n As Integer) As ItemKey
        Return New ItemKey($"E-{n}", "store-1")
    End Function

    Private Shared Function LinhaDe(n As Integer, Optional naoLida As Boolean = False) As MessageRowViewModel
        Return New MessageRowViewModel(New MailSummary With {
            .Key = Chave(n), .Subject = $"assunto {n}",
            .SenderName = "quem", .IsUnread = naoLida})
    End Function

    Private Shared Function Anexo(n As Integer) As AttachmentInfo
        Return New AttachmentInfo With {
            .Key = New AttachmentKey(Chave(n), 0, $"arquivo-{n}.txt", 10),
            .FileName = $"arquivo-{n}.txt", .SizeBytes = 10}
    End Function

    Private Shared Function Broker() As FakeBroker
        Dim b As New FakeBroker()
        b.LeitorLigado = True
        For n = 1 To 2
            b.ComDetalhe(New MessageDetail With {
                .Key = Chave(n), .Subject = $"assunto {n}",
                .SenderName = "quem", .SenderAddress = "quem@x.invalido",
                .Content = ContentState.AttachmentsAvailable, .Format = BodyFormat.PlainText,
                .TextBody = $"corpo {n}",
                .RecipientsStatus = PartStatus.Full,
                .AttachmentsStatus = PartStatus.Full,
                .Attachments = New List(Of AttachmentInfo)() From {Anexo(n)}})
        Next
        Return b
    End Function

    Private Shared Sub NoDispatcherAsync(corpo As Func(Of Dispatcher, Task))
        Dim erro As Exception = Nothing
        Dim t As New Thread(
            Sub()
                Dim d = Dispatcher.CurrentDispatcher
                d.BeginInvoke(
                    Async Sub()
                        Try
                            Await corpo(d)
                        Catch ex As Exception
                            erro = ex
                        Finally
                            d.InvokeShutdown()
                        End Try
                    End Sub)
                Dispatcher.Run()
            End Sub)
        t.SetApartmentState(ApartmentState.STA)
        t.Start()
        Assert.IsTrue(t.Join(TimeSpan.FromSeconds(30)), "a thread STA nao terminou")
        If erro IsNot Nothing Then Throw erro
    End Sub

    Private Shared Function AbrirLeitor(b As FakeBroker, d As Dispatcher,
                                   destino As String) As MessageDetailViewModel
        Return New MessageDetailViewModel(b, d, Sub(t As Task, nome As String)
                                                End Sub,
                                          New FakeSaveFile() With {.Escolha = destino})
    End Function

    Private Shared Async Function Assentar(condicao As Func(Of Boolean),
                                           Optional voltas As Integer = 300) As Task
        For i = 1 To voltas
            If condicao() Then Return
            Await Task.Delay(5)
        Next
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>Controle: salvar anexo sem interferencia anuncia o sucesso.</b>
    '''
    ''' Sem ele, um leitor que nunca escreve <c>AttachmentStatus</c> passaria
    ''' nos tres testes seguintes.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_salvar_anexo_anuncia_onde_salvou()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                Using leitor = AbrirLeitor(b, d, "C:\destino\arquivo-1.txt")
                    leitor.Show(LinhaDe(1))
                    Await Assentar(Function() leitor.HasMessage)

                    Await leitor.SaveAttachmentCommand.ExecuteAsync(Anexo(1))

                    StringAssert.Contains(leitor.AttachmentStatus, "Salvo em",
                        "controle: a gravacao normal tinha de anunciar o destino")
                End Using
            End Function)
    End Sub

    ''' <summary>
    ''' <b>Trocar de mensagem durante a gravacao NAO anuncia no leitor novo.</b>
    '''
    ''' O comentario da producao conta o defeito que originou a guarda:
    ''' <c>_disposed</c> sozinho fecha so o fechamento da janela, e sem geracao
    ''' trocar de mensagem durante a gravacao publicava "Salvo em ..." — ou a
    ''' falha do anexo da mensagem A — no leitor da B, com o ViewModel vivo e
    ''' ninguem para desconfiar.
    '''
    ''' <b>Controle negativo confirmado:</b> desligando as duas conferencias de
    ''' <c>_disposed OrElse geracao</c> do <c>SalvarAnexoAsync</c>, este teste e
    ''' o irmao abaixo caem — 2 falhas.
    ''' </summary>
    <TestMethod>
    Public Sub Trocar_de_mensagem_durante_a_gravacao_NAO_anuncia()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                b.TravaDoAnexo = New TaskCompletionSource(Of Boolean)()
                Using leitor = AbrirLeitor(b, d, "C:\destino\arquivo-1.txt")
                    leitor.Show(LinhaDe(1))
                    Await Assentar(Function() leitor.HasMessage)

                    Dim voo = leitor.SaveAttachmentCommand.ExecuteAsync(Anexo(1))
                    Await Assentar(Function() b.Chamadas.Contains("SaveAttachment"))
                    Assert.IsTrue(b.Chamadas.Contains("SaveAttachment"),
                        "controle: a gravacao tinha de estar parada no broker")

                    ' O usuario clicou em outra mensagem.
                    leitor.Show(LinhaDe(2))

                    b.TravaDoAnexo.SetResult(True)
                    Await voo

                    Assert.IsFalse(leitor.AttachmentStatus.Contains("Salvo em"),
                        $"anunciou no leitor da outra mensagem: {leitor.AttachmentStatus}")
                End Using
            End Function)
    End Sub

    ''' <summary>
    ''' <b>Fechar o leitor durante a gravacao NAO escreve na tela que saiu.</b>
    '''
    ''' Irmao do anterior, e a outra metade da mesma condicao composta. Aqui
    ''' quem muda nao e a mensagem: e a existencia da janela.
    ''' </summary>
    <TestMethod>
    Public Sub Fechar_o_leitor_durante_a_gravacao_NAO_anuncia()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                b.TravaDoAnexo = New TaskCompletionSource(Of Boolean)()
                Dim leitor = AbrirLeitor(b, d, "C:\destino\arquivo-1.txt")
                leitor.Show(LinhaDe(1))
                Await Assentar(Function() leitor.HasMessage)

                Dim voo = leitor.SaveAttachmentCommand.ExecuteAsync(Anexo(1))
                Await Assentar(Function() b.Chamadas.Contains("SaveAttachment"))
                Assert.IsTrue(b.Chamadas.Contains("SaveAttachment"),
                    "controle: a gravacao tinha de estar parada no broker")

                leitor.Dispose()

                b.TravaDoAnexo.SetResult(True)
                Await voo

                Assert.IsFalse(leitor.AttachmentStatus.Contains("Salvo em"),
                    $"anunciou num leitor descartado: {leitor.AttachmentStatus}")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>Marcar como lida falhando DEPOIS do fechamento nao reverte a linha.</b>
    '''
    ''' A marcacao e otimista: a linha perde o negrito na hora, e so volta a
    ''' ficar negrito se o Outlook recusar. Falha AMBIGUA tambem nao reverte,
    ''' porque a marcacao pode ter sido aplicada — e desfazer na tela mentiria
    ''' tanto quanto manter.
    '''
    ''' Aqui a falha e dura (<c>Denied</c>), que reverteria; o que impede e o
    ''' fechamento. E a guarda tem <b>duas</b> conferencias: antes de agendar
    ''' no dispatcher e DENTRO do delegate, porque entre agendar e executar
    ''' cabe um <c>Dispose</c>.
    '''
    ''' <b>Controle negativo, medido nos tres estados:</b>
    '''
    '''   • so a conferencia de antes removida .... <b>passa</b>
    '''   • so a de dentro do delegate removida ... <b>passa</b>
    '''   • as duas removidas .................... <b>falha</b>
    '''
    ''' Mesmo padrao da arvore: sao independentemente suficientes, e o que este
    ''' teste prova e a propriedade do par. A de dentro do delegate so seria
    ''' alcancavel sozinha se o <c>Dispose</c> acontecesse entre o agendamento
    ''' e a execucao no dispatcher, e um teste que roda NO dispatcher nao
    ''' consegue se colocar nesse intervalo.
    ''' </summary>
    <TestMethod>
    Public Sub Marcar_como_lida_falhando_depois_do_fechamento_NAO_reverte()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                b.FalhaAoMarcarLida = ErrorKind.Denied
                b.TravaDaLeitura = New TaskCompletionSource(Of Boolean)()
                Dim leitor = AbrirLeitor(b, d, Nothing)

                Dim linha = LinhaDe(1, naoLida:=True)
                leitor.Show(linha)
                Await Assentar(Function() leitor.HasMessage)

                ' O temporizador de "ficou exibida tempo suficiente" e de 1 s.
                Await Assentar(Function() b.Chamadas.Contains("MarkRead"), voltas:=600)
                Assert.IsTrue(b.Chamadas.Contains("MarkRead"),
                    "controle: a marcacao tinha de ter comecado")
                Assert.IsFalse(linha.IsUnread,
                    "controle: a marcacao e otimista, a linha perde o negrito na hora")

                leitor.Dispose()

                b.TravaDaLeitura.SetResult(True)

                ' MARCO, E NAO RELOGIO. Sem o "MarkRead-fim" so daria para
                ' esperar um tempo e torcer: uma continuacao atrasada faria o
                ' teste concluir "nao reverteu" antes de a reversao ter tido
                ' chance de acontecer, e ele passaria ate com a guarda fora.
                Await Assentar(Function() b.Chamadas.Contains("MarkRead-fim"))
                Assert.IsTrue(b.Chamadas.Contains("MarkRead-fim"),
                    "controle: a marcacao tinha de ter terminado")
                Await Assentar(Function() linha.IsUnread, voltas:=40)

                Assert.IsFalse(linha.IsUnread,
                    "reverteu a linha numa lista que ja nao esta em lugar nenhum")
            End Function)
    End Sub

    ''' <summary>
    ''' <b>Controle do irmao acima: sem fechar, a falha dura REVERTE mesmo.</b>
    '''
    ''' Este e o controle negativo embutido, e ele importa mais que o normal:
    ''' sem ele, um leitor que simplesmente nunca reverte passaria no teste
    ''' anterior, e o teste anterior nao provaria nada.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_falha_dura_sem_fechamento_REVERTE_a_linha()
        NoDispatcherAsync(
            Async Function(d)
                Dim b = Broker()
                b.FalhaAoMarcarLida = ErrorKind.Denied
                Using leitor = AbrirLeitor(b, d, Nothing)
                    Dim linha = LinhaDe(1, naoLida:=True)
                    leitor.Show(linha)
                    Await Assentar(Function() leitor.HasMessage)

                    ' NAO se mede o estado otimista aqui. Sem trava, a recusa
                    ' volta na mesma tacada: a linha perde e recupera o negrito
                    ' entre duas voltas do laco, e cobrar o intervalo seria
                    ' cobrar velocidade de agendamento. O irmao com trava e o
                    ' lugar certo para medir o estado otimista.
                    Await Assentar(Function() b.Chamadas.Contains("MarkRead"), voltas:=600)
                    Assert.IsTrue(b.Chamadas.Contains("MarkRead"),
                        "controle: a marcacao tinha de ter comecado")

                    Await Assentar(Function() linha.IsUnread, voltas:=200)
                    Assert.IsTrue(linha.IsUnread,
                        "controle: falha dura sem fechamento tinha de reverter o negrito")
                End Using
            End Function)
    End Sub

    ''' <summary>
    ''' <b>FALHA AO CONFERIR A IDENTIDADE DO ANEXO NÃO PODE VIRAR "É O MESMO".</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>OS AUXILIARES TOLERANTES ANULAVAM A GUARDA</b>
    '''
    ''' O índice de um anexo é instável, então antes de gravar o
    ''' <c>SaveAttachment</c> confere nome <b>e</b> tamanho. Só que os dois
    ''' lados da comparação eram lidos com os auxiliares tolerantes: exceção
    ''' vira <c>""</c> e <c>0</c>. Se a leitura falhasse nos dois momentos —
    ''' ao indexar e ao gravar — <c>""/0</c> casava com <c>""/0</c>, a guarda
    ''' passava e <b>o anexo errado era gravado com o nome certo</b>.
    '''
    ''' É o dano exato que essa guarda existe para impedir, chegando por dentro
    ''' dela. Agora a leitura precisa ser conclusiva: <c>Nothing</c> em qualquer
    ''' dos dois fecha.
    '''
    ''' <b>Controle negativo:</b> fazendo a função aceitar <c>Nothing</c> como
    ''' vazio/zero, as duas primeiras linhas caem.
    ''' </summary>
    <TestMethod>
    Public Sub Identidade_de_anexo_ILEGIVEL_nao_casa()
        Dim dono As New ItemKey("E-1", "store-1")

        ' A CHAVE FABRICADA: nome vazio e tamanho zero, que e o que os
        ' auxiliares tolerantes produzem quando a leitura falha.
        Dim fabricada As New AttachmentKey(dono, 1, "", 0, identidadeConhecida:=False)

        ' A LEITURA DE AGORA, contra uma chave CONCLUSIVA.
        '
        ' Estas duas usavam a chave fabricada, e por isso passavam pelo motivo
        ' errado: a funcao retornava no IdentidadeConhecida e nem chegava a
        ' olhar os argumentos. Apagar as duas guardas da leitura atual deixaria
        ' as assercoes verdes. A revisao externa pegou.
        Dim conclusiva As New AttachmentKey(dono, 1, "", 0)
        Assert.IsFalse(MessageReading.MesmaIdentidade(Nothing, 0, conclusiva),
            "nome ilegivel casou -- o anexo errado seria gravado")
        Assert.IsFalse(MessageReading.MesmaIdentidade("", Nothing, conclusiva),
            "tamanho ilegivel casou")

        ' O OUTRO LADO, e ele escapou da primeira versao desta guarda: a
        ' leitura de AGORA conclui, e a chave e que foi fabricada. Sem o
        ' IdentidadeConhecida, ""/0 de hoje casava com ""/0 de ontem -- e
        ' "x.dat"/0 casava com uma chave em que so o nome tinha sido lido.
        Assert.IsFalse(MessageReading.MesmaIdentidade("", 0, fabricada),
            "leitura conclusiva casou com uma chave que nunca foi conclusiva")

        Dim soONome As New AttachmentKey(dono, 1, "x.dat", 0, identidadeConhecida:=False)
        Assert.IsFalse(MessageReading.MesmaIdentidade("x.dat", 0, soONome),
            "o tamanho da chave era o zero de quem nao leu, e passou")

        ' CONTROLE POSITIVO: leitura conclusiva e igual continua casando,
        ' senao nenhum anexo seria salvo nunca.
        Dim boa As New AttachmentKey(dono, 1, "contrato.pdf", 4096)
        Assert.IsTrue(MessageReading.MesmaIdentidade("contrato.pdf", 4096, boa))

        ' E diferente continua nao casando, nos dois campos.
        Assert.IsFalse(MessageReading.MesmaIdentidade("outro.pdf", 4096, boa))
        Assert.IsFalse(MessageReading.MesmaIdentidade("contrato.pdf", 4095, boa))
    End Sub

    ''' <summary>
    ''' <b>O AVISO DIZ O QUE DE FATO BLOQUEOU.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>AVISO QUE EXAGERA ENSINA A IGNORAR AVISO</b>
    '''
    ''' A frase era fixa — <i>"Responder e encaminhar ficam bloqueados"</i> —
    ''' e as duas ações olham coisas diferentes: responder olha os
    ''' <b>destinatários</b>, encaminhar olha os <b>anexos</b>. Com só os
    ''' anexos incompletos, a tela afirmava que responder estava bloqueado, e
    ''' não estava.
    '''
    ''' Ficou muito mais comum em 29/08, quando identidade de anexo não
    ''' conferida passou a derrubar a completude da lista — ou seja, o meu
    ''' conserto anterior aumentou a frequência de uma frase falsa.
    '''
    ''' <b>Controle negativo:</b> devolvendo a frase fixa, a primeira asserção
    ''' cai.
    ''' </summary>
    <TestMethod>
    Public Sub O_aviso_de_leitura_parcial_nomeia_so_o_que_bloqueou()
        NoDispatcherAsync(
            Async Function(d)
                Dim b As New FakeBroker() With {.LeitorLigado = True}
                b.ComDetalhe(New MessageDetail With {
                    .Key = Chave(1), .Subject = "assunto",
                    .SenderName = "quem", .SenderAddress = "quem@x.invalido",
                    .Content = ContentState.AttachmentsAvailable,
                    .Format = BodyFormat.PlainText, .TextBody = "corpo",
                    .RecipientsStatus = PartStatus.Full,
                    .AttachmentsStatus = PartStatus.IncompleteWith(3, 1, ErrorKind.Denied),
                    .Attachments = New List(Of AttachmentInfo)()})

                Using leitor = AbrirLeitor(b, d, "C:\destino\x.txt")
                    leitor.Show(LinhaDe(1))
                    Await Assentar(Function() leitor.HasPartialRead)

                    ' SO OS ANEXOS ESTAO INCOMPLETOS.
                    Assert.IsTrue(leitor.CanReply, "controle: responder continua liberado")
                    Assert.IsFalse(leitor.CanForward, "controle: encaminhar tinha de bloquear")

                    Assert.IsFalse(leitor.PartialReadNotice.Contains("Responder"),
                        "o aviso diz que responder esta bloqueado, e ele nao esta: " &
                        leitor.PartialReadNotice)
                    StringAssert.Contains(leitor.PartialReadNotice, "encaminhar fica bloqueado")
                End Using
            End Function)
    End Sub

    ''' <summary>
    ''' <b>A REGRA "A IDENTIDADE FOI LIDA" MORA NUM LUGAR SÓ.</b>
    '''
    ''' Ela decide duas coisas em dois arquivos: se a chave do anexo nasce
    ''' confiável, e se o anexo conta como <i>obtido</i> na completude da
    ''' lista. Estava repetida em quatro pontos — e repetição é como um deles
    ''' fica para trás, que foi exatamente o que aconteceu com o caminho dos
    ''' rascunhos.
    '''
    ''' <b>Controle negativo:</b> fazendo a função devolver sempre
    ''' <c>True</c>, este teste cai — e com ele a completude da lista volta a
    ''' fechar sobre anexo que ninguém conferiu.
    ''' </summary>
    <TestMethod>
    Public Sub Identidade_lida_exige_os_DOIS_valores()
        Assert.IsTrue(MessageReading.IdentidadeLida("contrato.pdf", 4096))

        ' Vazio e zero LIDOS continuam sendo leitura conclusiva: um anexo
        ' pode legitimamente ter nome vazio ou tamanho zero.
        Assert.IsTrue(MessageReading.IdentidadeLida("", 0))

        Assert.IsFalse(MessageReading.IdentidadeLida(Nothing, 4096),
            "sem o nome, a identidade nao foi conferida")
        Assert.IsFalse(MessageReading.IdentidadeLida("contrato.pdf", Nothing),
            "sem o tamanho, a identidade nao foi conferida")
        Assert.IsFalse(MessageReading.IdentidadeLida(Nothing, Nothing))
    End Sub

End Class

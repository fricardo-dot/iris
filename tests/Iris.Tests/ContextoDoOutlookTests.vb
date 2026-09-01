Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports Iris.App.ViewModels
Imports Iris.Assist
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Integration.Assist.Http
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace Global.Iris.Tests

    ''' <summary>
    ''' <b>O contexto de produção — o que liga a tela ao Outlook.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ESTE ARQUIVO EXISTE</b>
    '''
    ''' A segunda passada do Codex encontrou o caminho de produção usando
    ''' <c>ContextoIndisponivel</c>: tudo o que a suíte provava sobre "o fluxo"
    ''' era provado sobre a imitação, e trocar a produção de volta para a
    ''' imitação deixaria a suíte verde. A terceira passada encontrou, dentro do
    ''' contexto já ligado, um <c>temAnexo:=False</c> <b>fixo</b> — o portão nega
    ''' mensagem com anexo, e o caminho de produção lhe afirmava que não havia.
    '''
    ''' Aqui o <see cref="ContextoDoOutlook"/> é montado de verdade, contra um
    ''' broker que responde. É o único lugar onde essa afirmação é conferida.
    ''' </summary>
    <TestClass>
    Public Class ContextoDoOutlookTests

        Private Shared ReadOnly Pasta As New FolderKey("store-1", "pasta-1")
        Private Shared ReadOnly Item As New ItemKey("E-1", "store-1")

        Private Shared Function Destino() As AssistDestination
            Return New AssistDestination("provedor-de-teste",
                                         "https://exemplo.invalido/v1", "modelo-de-teste")
        End Function

        Private Shared Function Rotulo() As LabelReading
            Return New LabelReading(Item, LabelReadingKind.Absent, LabelReadStage.Parse,
                                    version:=New LabelVersionEvidence(
                                        "E-1", DateTimeOffset.UnixEpoch, "CK-1"))
        End Function

        ''' <summary>Um contexto de produção ligado a este broker.</summary>
        Private Shared Function Montar(b As FakeBroker) As ContextoDoOutlook
            b.Rotulos = Function(itens) OperationResult(Of IReadOnlyList(Of LabelReading)).
                                        Ok({Rotulo()})
            Return New ContextoDoOutlook(
                b, Destino(),
                Function() (Pasta, CType({Item}, IReadOnlyList(Of ItemKey))))
        End Function

        Private Shared Function Anexo(tem As Boolean?) _
                                      As OperationResult(Of IReadOnlyList(Of AttachmentPresence))
            Return OperationResult(Of IReadOnlyList(Of AttachmentPresence)).
                   Ok({New AttachmentPresence(Item, tem)})
        End Function

        ' ==================================================================
        ' O anexo é lido, e não suposto

        ''' <summary>
        ''' <b>Anexo lido como "tem" chega ao portão como "tem".</b>
        '''
        ''' O caso direto, e o que estava quebrado: a classificação dizia
        ''' <c>False</c> sem consultar nada.
        ''' </summary>
        <TestMethod>
        Public Sub Anexo_lido_como_TEM_chega_ao_portao_como_tem()
            Dim b As New FakeBroker() With {.PastaDeTodos = Pasta}
            b.Anexos = Function(itens) Anexo(True)

            Dim c = Montar(b).Classificar()

            Assert.AreEqual(1, c.Count)
            Assert.IsTrue(c(0).TemAnexo, "o portao precisa saber que tem anexo para negar")
        End Sub

        ''' <summary>
        ''' <b>"Não deu para contar" conta como "tem".</b>
        '''
        ''' Guarda do Object Model, item de classe inesperada, erro de COM: a
        ''' contagem falha e o resultado é <c>Nothing</c>. Ler isso como
        ''' ausência seria a mesma falha aberta que o 3.0 já custou uma vez, com
        ''' <c>E_INVALIDARG</c> virando "não tem rótulo".
        ''' </summary>
        <TestMethod>
        Public Sub Anexo_que_nao_deu_para_contar_conta_como_TEM()
            Dim b As New FakeBroker() With {.PastaDeTodos = Pasta}
            b.Anexos = Function(itens) Anexo(Nothing)

            Assert.IsTrue(Montar(b).Classificar()(0).TemAnexo,
                          "nao consegui contar nunca vira prova de ausencia")
        End Sub

        ''' <summary>
        ''' <b>A chamada inteira falhando também conta como "tem".</b>
        '''
        ''' O item nem aparece na resposta. Ele não pode desaparecer da
        ''' classificação — item a menos faria a thread parecer menor do que é —
        ''' nem entrar como limpo.
        ''' </summary>
        <TestMethod>
        Public Sub Falha_na_leitura_de_anexo_conta_como_TEM()
            Dim b As New FakeBroker() With {.PastaDeTodos = Pasta}
            b.Anexos = Function(itens) OperationResult(Of IReadOnlyList(Of AttachmentPresence)).
                                       Fail(ErrorKind.Busy, "sem resposta")

            Dim c = Montar(b).Classificar()

            Assert.AreEqual(1, c.Count, "o item nao pode sumir da classificacao")
            Assert.IsTrue(c(0).TemAnexo)
        End Sub

        ''' <summary>
        ''' Controle negativo: sem anexo, a classificação diz que não tem.
        '''
        ''' Sem ele, um contexto que respondesse <c>True</c> sempre passaria nos
        ''' três testes de cima — e a IA ficaria inutilizável pelo motivo errado,
        ''' que é a outra forma de estar quebrado.
        ''' </summary>
        <TestMethod>
        Public Sub Controle_sem_anexo_a_classificacao_diz_que_NAO_tem()
            Dim b As New FakeBroker() With {.PastaDeTodos = Pasta}
            b.Anexos = Function(itens) Anexo(False)

            Assert.IsFalse(Montar(b).Classificar()(0).TemAnexo)
        End Sub

        ''' <summary>
        ''' <b>Sem seleção, o contexto não vai ao COM.</b>
        '''
        ''' O preflight existe para não gastar leitura sem autorização; ir ao
        ''' Outlook para classificar lista vazia seria gastar do mesmo jeito.
        ''' </summary>
        <TestMethod>
        Public Sub Sem_selecao_nao_ha_ida_ao_COM()
            Dim b As New FakeBroker() With {.PastaDeTodos = Pasta}
            Dim c As New ContextoDoOutlook(
                b, Destino(),
                Function() (Pasta, CType(Array.Empty(Of ItemKey)(), IReadOnlyList(Of ItemKey))))

            Assert.AreEqual(0, c.Classificar().Count)
            Assert.AreEqual(0, b.Chamadas.Count, "nao podia ter ido ao Outlook")
        End Sub

        ''' <summary>
        ''' <b>O pedido declara a operação e a pasta aberta.</b>
        '''
        ''' É por ele que o portão decide, e uma pasta errada aqui seria uma
        ''' autorização conferida contra o lugar errado.
        ''' </summary>
        <TestMethod>
        Public Sub O_pedido_carrega_a_operacao_e_a_pasta()
            Dim p = Montar(New FakeBroker()).Pedido(AssistOperation.Redigir)

            Assert.AreEqual(AssistOperation.Redigir, p.Operacao)
            Assert.AreEqual(Pasta.EntryId, p.Pasta.EntryId)
            Assert.AreEqual("provedor-de-teste", p.Destino.Provedor)
        End Sub

        ' ==================================================================
        ' E a produção usa ESTE contexto

        ''' <summary>
        ''' <b>A produção monta o contexto do Outlook, e não a imitação.</b>
        '''
        ''' Este teste lê o <b>código-fonte</b> do <c>MainViewModel</c>, como os
        ''' de binding leem o XAML, e pela mesma razão: o que se quer provar não
        ''' é um comportamento do objeto, e sim <b>qual objeto foi escolhido</b>
        ''' na montagem. Construir um <c>MainViewModel</c> de verdade exigiria
        ''' Outlook, cache e Dispatcher — e o teste passaria a provar outra
        ''' coisa.
        '''
        ''' Sem ele, voltar a produção para <c>ContextoIndisponivel</c> deixaria
        ''' a suíte inteira verde, que foi exatamente o que aconteceu.
        ''' </summary>
        <TestMethod>
        Public Sub A_producao_monta_o_contexto_do_Outlook()
            Dim fonte = LerFonte("MainViewModel.vb")

            StringAssert.Contains(fonte, "New ContextoDoOutlook(",
                "a producao tem de montar o contexto de verdade")
            Assert.IsFalse(fonte.Contains("New ContextoIndisponivel()"),
                "ContextoIndisponivel na producao faz o caminho central ate o " &
                "Outlook faltar mesmo depois da cerimonia de ativacao")
        End Sub

        ''' <summary>
        ''' <b>A fábrica de provedor é fechada: o que não reconhece, recusa.</b>
        '''
        ''' Era um <c>If</c> que instanciava o OpenRouter para <b>qualquer</b>
        ''' ativação carregada. Uma que declarasse outro provedor, com outro
        ''' endereço, receberia o protocolo do OpenRouter e a credencial guardada
        ''' sob o nome dele.
        '''
        ''' O que se cobra: que provedor desconhecido vire
        ''' <c>AssistenteIndisponivel</c> — e não um adaptador genérico que
        ''' "tenta assim mesmo".
        ''' </summary>
        <TestMethod>
        Public Sub A_fabrica_de_provedor_e_FECHADA()
            Dim desconhecido = Ativacao("provedor-que-ninguem-implementou")
            Dim p = MainViewModel.ProvedorPara(ActivationLoadResult.Ok(desconhecido))

            Assert.IsInstanceOfType(p, GetType(AssistenteIndisponivel),
                                    "provedor desconhecido nao pode virar adaptador")
            ' O destino existe e e VAZIO -- nao Nothing. E o vazio que faz o
            ' portao recusar: endpoint em branco nao e HTTPS e nao casa com
            ' autorizacao nenhuma.
            Assert.AreEqual("", p.Destino.Endpoint,
                            "endpoint em branco e o que faz o portao recusar antes de ler")
            Assert.AreEqual("", p.Destino.Provedor)
        End Sub

        ''' <summary>
        ''' Controle: o provedor <b>reconhecido</b> vira o adaptador dele.
        '''
        ''' Sem ele, uma fábrica que devolvesse <c>AssistenteIndisponivel</c>
        ''' sempre passaria no teste de cima — e a IA nunca funcionaria, pelo
        ''' motivo errado.
        ''' </summary>
        <TestMethod>
        Public Sub Controle_o_provedor_RECONHECIDO_vira_adaptador()
            Dim p = MainViewModel.ProvedorPara(ActivationLoadResult.Ok(Ativacao("openrouter")))

            Assert.IsInstanceOfType(p, GetType(OpenRouterAssistantProvider))
        End Sub

        ''' <summary>
        ''' <b>Sem ativação carregada, não há provedor.</b> O caso de produção
        ''' enquanto ninguém escreveu a cerimônia.
        ''' </summary>
        <TestMethod>
        Public Sub Sem_ativacao_carregada_nao_ha_provedor()
            Dim p = MainViewModel.ProvedorPara(
                ActivationLoadResult.Nao(ActivationLoadFailure.Ausente))

            Assert.IsInstanceOfType(p, GetType(AssistenteIndisponivel))
        End Sub

        ''' <summary>Uma ativação válida que declara o provedor pedido.</summary>
        Private Shared Function Ativacao(provedor As String) As ActivationRecord
            Return New ActivationRecord(
                "ativacao-1", 1, "quem", DateTimeOffset.UnixEpoch,
                provedor, "https://exemplo.invalido/v1", "modelo-x",
                "local", "sem retenção",
                {AssistOperation.Resumir}, {Pasta}, Array.Empty(Of String)(),
                {LabelReadingKind.Absent}, {0},
                ate:=DateTimeOffset.UnixEpoch.AddDays(30),
                provedoresPermitidos:={"algum"})
        End Function

        ''' <summary>
        ''' Controle: a busca no fonte REALMENTE acusa.
        '''
        ''' Sem isto, um caminho errado faria os dois asserts de cima passarem
        ''' para sempre — o primeiro por não achar nada e o segundo por não achar
        ''' nada também.
        ''' </summary>
        <TestMethod>
        Public Sub Controle_a_leitura_do_fonte_encontra_o_que_esta_la()
            Dim fonte = LerFonte("MainViewModel.vb")

            Assert.IsTrue(fonte.Length > 1000, "o arquivo lido nao pode estar vazio")
            StringAssert.Contains(fonte, "Class MainViewModel",
                "se nem isto e encontrado, a busca nao esta lendo o arquivo certo")
        End Sub

        Private Shared Function LerFonte(arquivo As String) As String
            Dim d = New DirectoryInfo(AppContext.BaseDirectory)
            While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
                d = d.Parent
            End While
            Assert.IsNotNull(d, "nao achei a raiz do repositorio")

            Dim caminho = Path.Combine(d.FullName, "src", "Iris.App", "ViewModels", arquivo)
            Assert.IsTrue(File.Exists(caminho), arquivo & " nao encontrado em " & caminho)
            Return File.ReadAllText(caminho)
        End Function

    End Class

End Namespace

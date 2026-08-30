Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports Iris.App.ViewModels
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Os caminhos de <c>Binding</c> do XAML resolvem contra os ViewModels.
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ISTO EXISTE</b>
'''
''' Binding com caminho errado no WPF <b>falha em silêncio</b>. A propriedade
''' não existe, o controle fica vazio, nada é lançado, nada aparece no log — e
''' a suíte continua verde porque nenhum teste toca em XAML.
'''
''' É "verde mas quebrado" na forma mais pura, e é o formato de erro que esta
''' fase inteira persegue. A faixa do acervo é o caso concreto: se
''' <c>Acervo.Ressalva</c> virasse <c>Acervo.Ressalvas</c> num refactor, a
''' ressalva que a §23 obriga a mostrar simplesmente <b>sumiria da tela</b>,
''' e o produto voltaria a exibir o acervo como se fosse o estado corrente da
''' caixa — sem nenhum sinal de que algo quebrou.
''' </summary>
<TestClass>
Public Class BindingsDaJanelaTests

    ''' <summary>
    ''' Raízes conhecidas: o prefixo do caminho e o tipo em que ele começa.
    '''
    ''' Só as raízes que este teste sabe resolver. Um <c>Binding</c> dentro de
    ''' <c>DataTemplate</c> tem outro DataContext e não é verificável assim —
    ''' por isso a lista é explícita em vez de "tudo o que aparecer".
    ''' </summary>
    Private Shared Function Raizes() As Dictionary(Of String, Type)
        Return New Dictionary(Of String, Type) From {
            {"Acervo.", GetType(AcervoViewModel)},
            {"Busca.", GetType(BuscaViewModel)},
            {"Agenda.", GetType(AgendaViewModel)},
            {"Connection.", GetType(ConnectionViewModel)},
            {"Composer.", GetType(ComposerViewModel)},
            {"Detail.", GetType(MessageDetailViewModel)},
            {"Messages.", GetType(MessageListViewModel)},
            {"Folders.", GetType(FolderTreeViewModel)}}
    End Function

    <TestMethod>
    Public Sub Todo_binding_conhecido_resolve_no_ViewModel()
        Dim xaml = LerXaml()
        Dim conhecidas = Raizes()
        Dim quebrados As New List(Of String)()
        Dim conferidos = 0

        For Each caminho In CaminhosDeBinding(xaml)
            Dim raiz = conhecidas.Keys.FirstOrDefault(Function(k) caminho.StartsWith(k, StringComparison.Ordinal))
            If raiz Is Nothing Then Continue For

            conferidos += 1
            Dim membro = caminho.Substring(raiz.Length)
            ' So o primeiro segmento: "A.B.C" resolve A, e B/C dependem do tipo
            ' de A, que este teste nao persegue.
            Dim primeiro = membro.Split("."c)(0)
            If primeiro.Length = 0 Then Continue For

            If conhecidas(raiz).GetProperty(primeiro,
                    BindingFlags.Public Or BindingFlags.Instance) Is Nothing Then
                quebrados.Add($"{caminho}  (nao existe em {conhecidas(raiz).Name})")
            End If
        Next

        Assert.IsTrue(conferidos > 5,
            $"so {conferidos} bindings conferidos — o teste nao esta encontrando o XAML")
        Assert.AreEqual(0, quebrados.Count,
            "binding com caminho errado falha em SILENCIO no WPF: " &
            Environment.NewLine & String.Join(Environment.NewLine, quebrados))
    End Sub

    ''' <summary>
    ''' A faixa do acervo existe, e mostra a RESSALVA.
    '''
    ''' Não basta o binding resolver: a §23 obriga a ressalva a aparecer junto
    ''' do acervo, e um refactor que removesse o <c>TextBlock</c> passaria no
    ''' teste acima sem problema nenhum — não haveria binding quebrado, haveria
    ''' binding ausente.
    ''' </summary>
    <TestMethod>
    Public Sub A_janela_mostra_a_ressalva_do_acervo()
        Dim xaml = LerXaml()
        StringAssert.Contains(xaml, "Acervo.Ressalva",
            "a ressalva da §23 tem de estar na janela, nao so no ViewModel")
        StringAssert.Contains(xaml, "AcervoIndisponivel",
            "cache que nao abre tem de aparecer — vazio silencioso e " &
            "indistinguivel de 'nao ha nada guardado'")
    End Sub

    ''' <summary>
    ''' <b>A varredura tem botao na janela, e ele nao some.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' A faixa do acervo era visivel so quando havia ressalva
    ''' (<c>TemAlgoADizer</c>). Com o botao dentro dela, ele sumiria justamente
    ''' quando nao ha nada a ressalvar — que e quando alguem quer varrer.
    '''
    ''' E o mesmo defeito que a faixa da IA ja teve, com o mesmo custo: botao
    ''' que some esconde a funcionalidade E o motivo de ela estar
    ''' indisponivel. A visibilidade passou a ser "existe acervo".
    ''' </summary>
    <TestMethod>
    Public Sub A_varredura_tem_BOTAO_na_janela()
        Dim xaml = LerXaml()

        StringAssert.Contains(xaml, "Acervo.VarrerCommand",
            "sem botao, o SweepRunner continua sendo codigo que ninguem chama")
        StringAssert.Contains(xaml, "Acervo.Varrendo",
            "varrer bloqueia e demora: sem sinal, 'lendo' e 'travou' sao iguais")

        Assert.IsFalse(xaml.Contains("Acervo.TemAlgoADizer"),
            "a faixa do acervo voltou a sumir quando nao ha ressalva, e leva o " &
            "botao de varrer junto")
    End Sub

    ''' <summary>
    ''' <b>A faixa do acervo e a da IA não moram na mesma linha da grade.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O DEFEITO QUE ESTE TESTE FECHA</b>
    '''
    ''' As duas ficavam em <c>Grid.Row="2"</c>. Num <c>Grid</c> isso significa
    ''' <b>empilhadas</b>, e a da IA — declarada depois, e com fundo próprio —
    ''' pintava por cima. <b>A faixa do acervo nunca foi vista na tela</b>, e a
    ''' pendência da Fase 2 que dizia exatamente isso não era falta de dado: era
    ''' uma linha de grade faltando.
    '''
    ''' Ninguém notou porque as duas tinham visibilidade condicional: a do
    ''' acervo só aparecia havendo ressalva, e a da IA cobria justamente quando
    ''' aparecia. Duas condições escondendo uma sobreposição.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE ESTE TESTE LÊ TEXTO, E NÃO PIXEL</b>
    '''
    ''' O jeito honesto seria instanciar a <c>MainWindow</c> e medir os
    ''' retângulos, como o <c>Aviso_e_resultado_NAO_se_cobrem</c> faz dentro da
    ''' faixa da IA. A janela exige broker, cache e sessão, e montar tudo isso
    ''' para conferir um número de linha custaria mais do que vale.
    '''
    ''' Então ele lê o XAML e cobra a propriedade que causou o defeito: as duas
    ''' faixas, que aparecem <b>ao mesmo tempo</b>, declaram linhas diferentes.
    ''' Não prova ausência de sobreposição em geral — prova que esta não voltou.
    ''' </summary>
    <TestMethod>
    Public Sub O_acervo_e_a_IA_nao_dividem_a_LINHA_da_grade()
        Dim xaml = LerXaml()

        Dim daIa = Text.RegularExpressions.Regex.Match(
            xaml, "<local:FaixaDaIa\s+Grid\.Row=""(\d+)""")
        Assert.IsTrue(daIa.Success, "nao achei a faixa da IA na janela")

        Dim doAcervo = Text.RegularExpressions.Regex.Match(
            xaml, "<Border Grid\.Row=""(\d+)""\s*?
\s*Visibility=""\{Binding MostrarAcervo,")
        Assert.IsTrue(doAcervo.Success, "nao achei a faixa do acervo na janela")

        Assert.AreNotEqual(doAcervo.Groups(1).Value, daIa.Groups(1).Value,
            "as duas faixas voltaram para a mesma linha da grade, e a de baixo " &
            "cobre a de cima -- foi assim que a faixa do acervo passou a fase " &
            "inteira sem nunca ter sido vista")
    End Sub


    ''' <summary>
    ''' <b>A agenda e o acervo dividem a linha 2, e por isso a exclusao tem
    ''' de ser DECLARADA.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' Quando a Fase 6 entrou, a agenda passou a ocupar a mesma linha do
    ''' acervo. E deliberado: sao o mesmo lugar da janela mudando de assunto
    ''' conforme a pasta selecionada. E e exatamente a situacao que produziu
    ''' o defeito de 2026 -- duas bordas na mesma linha, cada uma com sua
    ''' condicao, esperando nao coincidir.
    '''
    ''' A diferenca e que agora a exclusao nao e esperanca: o acervo depende
    ''' de <c>MostrarAcervo</c>, que e <c>Acervo IsNot Nothing AndAlso Not
    ''' Agenda.TemPasta</c>. Uma condicao, num lugar, que da para ler.
    ''' </summary>
    <TestMethod>
    Public Sub A_agenda_e_o_acervo_se_EXCLUEM_na_mesma_linha()
        Dim xaml = LerXaml()

        Dim daAgenda = Text.RegularExpressions.Regex.Match(
            xaml, "<Border Grid\.Row=""(\d+)""\s*
\s*Visibility=""\{Binding Agenda\.TemPasta,")
        Assert.IsTrue(daAgenda.Success,
            "nao achei a faixa da agenda, ou ela deixou de depender de Agenda.TemPasta")

        Dim doAcervo = Text.RegularExpressions.Regex.Match(
            xaml, "<Border Grid\.Row=""(\d+)""\s*
\s*Visibility=""\{Binding MostrarAcervo,")
        Assert.IsTrue(doAcervo.Success,
            "a faixa do acervo voltou a depender de algo que nao exclui a agenda")

        Assert.AreEqual(daAgenda.Groups(1).Value, doAcervo.Groups(1).Value,
            "controle: as duas TEM de estar na mesma linha. Em linhas diferentes, " &
            "este teste deixa de proteger o que ele diz proteger")

        ' E a exclusao existe do lado do ViewModel, e nao so no XAML.
        Assert.IsNotNull(GetType(MainViewModel).GetProperty("MostrarAcervo"),
            "MostrarAcervo sumiu do MainViewModel, e a exclusao virou coincidencia")
    End Sub

    ''' <summary>
    ''' <b>A janela hospeda a faixa da IA, com o contexto certo.</b>
    '''
    ''' A faixa é um <c>UserControl</c> próprio — foi extraída para que o teste
    ''' de renderização pudesse instanciar a <b>faixa de verdade</b> em vez de
    ''' uma imitação. O que a janela precisa fazer é hospedá-la e dar a ela o
    ''' <c>DataContext</c> certo; sem isso, todos os bindings de dentro
    ''' resolveriam contra o <c>MainViewModel</c> e ficariam vazios em silêncio.
    ''' </summary>
    <TestMethod>
    Public Sub A_janela_hospeda_a_faixa_da_IA()
        Dim xaml = LerXaml()
        StringAssert.Contains(xaml, "<local:FaixaDaIa",
            "a faixa da IA tem de estar na janela")
        StringAssert.Contains(xaml, "DataContext=" & Q & "{Binding Assistente}" & Q,
            "sem o contexto certo, os bindings de dentro resolvem contra o " &
            "MainViewModel e ficam vazios em silencio")
    End Sub

    ''' <summary>
    ''' <b>Todo binding da faixa resolve no <c>AssistenteViewModel</c>.</b>
    '''
    ''' O <c>DataContext</c> da faixa é o assistente, então os caminhos são
    ''' diretos — <c>Aviso</c>, e não <c>Assistente.Aviso</c>. Um refactor que
    ''' renomeasse uma propriedade deixaria o binding perfeito e o texto
    ''' invisível: binding com caminho errado no WPF <b>falha em silêncio</b>.
    ''' </summary>
    <TestMethod>
    Public Sub Todo_binding_da_faixa_resolve_no_AssistenteViewModel()
        Dim quebrados As New List(Of String)()

        For Each caminho In CaminhosDeBinding(LerFaixa())
            Dim partes = caminho.Split("."c)
            Dim alvo As Type = GetType(AssistenteViewModel)

            For Each membro In partes
                If alvo Is Nothing Then Exit For
                Dim p = alvo.GetProperty(membro)
                If p Is Nothing Then
                    quebrados.Add(caminho)
                    Exit For
                End If
                alvo = p.PropertyType
            Next
        Next

        Assert.AreEqual(0, quebrados.Count,
            "caminho que nao resolve fica vazio em silencio: " &
            String.Join(", ", quebrados))
    End Sub

    ''' <summary>
    ''' <b>A faixa mostra a situação da IA — e o que ficou sem desfecho.</b>
    '''
    ''' Binding ausente não é binding quebrado, e passaria pelo teste de cima sem
    ''' reclamar. O que a §28.2 obriga a mostrar é o motivo de a IA não estar
    ''' habilitada; o que a §29.6 obriga é o número de envios que ficaram
    ''' ambíguos numa execução anterior.
    '''
    ''' "Pode ter saído conteúdo desta caixa e ninguém sabe" não pode viver só no
    ''' banco.
    ''' </summary>
    <TestMethod>
    Public Sub A_faixa_mostra_a_situacao_da_IA()
        Dim xaml = LerFaixa()
        StringAssert.Contains(xaml, "{Binding Aviso}",
            "o motivo de a IA nao estar habilitada tem de aparecer")
        StringAssert.Contains(xaml, "{Binding Reconciliacao.Aviso}",
            "envios sem desfecho conhecido nao podem viver so no banco")
        StringAssert.Contains(xaml, "{Binding Resultado}",
            "e a resposta do modelo tem de ter onde aparecer")
    End Sub

    ''' <summary>
    ''' <b>A ação existe na faixa.</b>
    '''
    ''' Sem os botões, o 3.5 seria uma tela de status: os comandos existiriam no
    ''' ViewModel e ninguém os alcançaria, e nem uma ativação futura tornaria a
    ''' funcionalidade utilizável.
    ''' </summary>
    <TestMethod>
    Public Sub A_acao_da_IA_existe_na_faixa()
        Dim xaml = LerFaixa()
        For Each comando In {"ResumirCommand", "RedigirCommand",
                             "DesfazerCommand", "CancelarCommand"}
            StringAssert.Contains(xaml, "{Binding " & comando & "}",
                comando & " nao esta na faixa — o comando existiria sem ninguem alcancar")
        Next
    End Sub

    ''' <summary>
    ''' <b>A resposta do modelo aparece num <c>TextBlock</c>.</b>
    '''
    ''' Não num controle que interprete Markdown, HTML ou link: ela vem de um
    ''' lugar que leu o e-mail, que por sua vez veio de fora. A barreira da §29.5
    ''' é estrutural, e este teste é onde ela fica presa ao XAML.
    ''' </summary>
    <TestMethod>
    Public Sub A_resposta_do_modelo_aparece_em_TEXTBLOCK()
        Dim xaml = LerFaixa()
        Dim i = xaml.IndexOf("{Binding Resultado}", StringComparison.Ordinal)
        Assert.IsTrue(i > 0, "o binding tem de existir")

        Dim antes = xaml.Substring(0, i)
        Dim elemento = antes.Substring(antes.LastIndexOf("<"c))

        StringAssert.StartsWith(elemento, "<TextBlock",
            "a resposta do modelo nao pode ir para um controle que INTERPRETE: " &
            "ela e dado, e dado que veio de fora")
    End Sub

    ''' <summary>
    ''' Controle: a busca por caminho quebrado REALMENTE acusa.
    '''
    ''' Sem isto, um extrator que não achasse binding nenhum faria o teste
    ''' principal passar para sempre.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_um_caminho_inventado_seria_acusado()
        Assert.IsNull(GetType(AcervoViewModel).GetProperty("RessalvaQueNaoExiste",
                          BindingFlags.Public Or BindingFlags.Instance),
                      "a propriedade inventada nao pode existir de verdade")

        Dim achados = CaminhosDeBinding(
            "<TextBlock Text=""{Binding Acervo.RessalvaQueNaoExiste}"" />").ToList()
        CollectionAssert.Contains(achados, "Acervo.RessalvaQueNaoExiste",
            "o extrator nao encontra nem um caminho plantado — ele nao extrai nada")
    End Sub

    ' ==================================================================

    Private Shared Iterator Function CaminhosDeBinding(xaml As String) As IEnumerable(Of String)
        ' {Binding Caminho} e {Binding Path=Caminho} e {Binding Caminho, ...}
        For Each m As Match In Regex.Matches(xaml, "\{Binding\s+(?:Path=)?([A-Za-z_][\w.]*)")
            Yield m.Groups(1).Value
        Next
    End Function

    Private Shared _xaml As String

    ''' <summary>Aspas duplas, sem duplicar aspas dentro do literal.</summary>
    Private Const Q As String = """"

    ''' <summary>O XAML da faixa da IA, que é um <c>UserControl</c> próprio.</summary>
    Private Shared Function LerFaixa() As String
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
            d = d.Parent
        End While
        Assert.IsNotNull(d, "nao achei a raiz do repositorio")
        Dim caminho = Path.Combine(d.FullName, "src", "Iris.App", "Views", "FaixaDaIa.xaml")
        Assert.IsTrue(File.Exists(caminho), "FaixaDaIa.xaml nao encontrado em " & caminho)
        Return File.ReadAllText(caminho)
    End Function

    ''' <summary>
    ''' <b>O TEXTO DA PASTA VAZIA É CALCULADO, e não literal no XAML.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>PROVA DE ALCANCE, e não de leitura</b>
    '''
    ''' Os dois testes do <c>MessageListViewModel</c> provam que o
    ''' <c>EmptyMessage</c> diz a coisa certa. Eles <b>não</b> provam que a tela
    ''' mostra o <c>EmptyMessage</c>: se o XAML voltasse para
    ''' <c>Text="Esta pasta está vazia"</c>, os dois continuariam verdes — e a
    ''' revisão externa apontou que o comentário deles anunciava justamente esse
    ''' controle negativo, que não existia.
    '''
    ''' É a mesma lição do calendário escondido pela política de visibilidade:
    ''' <i>prova de leitura não é prova de alcance</i>. Este teste fecha o
    ''' último metro, lendo o XAML de verdade.
    ''' </summary>
    <TestMethod>
    Public Sub O_texto_da_pasta_vazia_vem_do_ViewModel()
        Dim xaml = LerXaml()

        StringAssert.Contains(xaml, "Text=""{Binding Messages.EmptyMessage}""",
            "o TextBlock da pasta vazia parou de ler o EmptyMessage -- e o " &
            "texto literal afirma que a pasta esta vazia mesmo quando a " &
            "leitura perdeu itens")

        ' E o literal NAO pode voltar por outro caminho.
        Assert.IsFalse(xaml.Contains("Text=""Esta pasta está vazia"""),
            "o texto literal voltou ao XAML")
    End Sub

    ''' <summary>
    ''' <b>A PERGUNTA ANTES DE APAGAR ESTÁ NA TELA.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>PROVA DE ALCANCE, e não de leitura</b>
    '''
    ''' Os testes do <c>AgendaViewModel</c> provam que a confirmação existe e
    ''' que o comando de apagar não executa antes dela. Eles <b>não</b> provam
    ''' que a tela mostra a pergunta: se o XAML não tivesse o painel, o botão
    ''' "Apagar" pediria a confirmação e nada apareceria — e o compromisso
    ''' ficaria sem apagar, sem explicação.
    '''
    ''' É a mesma lição do texto da pasta vazia, e a mesma do calendário que a
    ''' política de visibilidade escondia: <i>prova de leitura não é prova de
    ''' alcance</i>. Este teste fecha o último metro, lendo o XAML.
    ''' </summary>
    <TestMethod>
    Public Sub A_pergunta_antes_de_apagar_esta_na_tela()
        Dim xaml = LerXaml()

        StringAssert.Contains(xaml, "{Binding Agenda.PerguntaDaExclusao}",
            "a pergunta nao aparece na tela: o usuario clicaria em Apagar e " &
            "nada aconteceria")
        StringAssert.Contains(xaml, "{Binding Agenda.ApagarCommand}",
            "nao ha como confirmar")
        StringAssert.Contains(xaml, "{Binding Agenda.CancelarExclusaoCommand}",
            "nao ha como desistir depois de perguntar")

        ' E o formulario de criar tambem tem de estar ligado.
        StringAssert.Contains(xaml, "{Binding Agenda.CriarCommand}")
        StringAssert.Contains(xaml, "Agenda.NovoAssunto")
    End Sub

    ''' <summary>
    ''' <b>A FAIXA DE ESCRITA SÓ APARECE SE A AGENDA SOUBER ESCREVER.</b>
    '''
    ''' Botão que não funciona é pior que botão ausente: ele promete uma coisa
    ''' que o objeto não sabe fazer. A visibilidade está presa ao
    ''' <c>PodeEscrever</c>, que é falso quando não há escritor.
    ''' </summary>
    <TestMethod>
    Public Sub A_faixa_de_escrita_depende_de_PodeEscrever()
        Dim xaml = LerXaml()
        StringAssert.Contains(xaml, "{Binding Agenda.PodeEscrever,",
            "a faixa de escrita aparece mesmo quando a agenda nao escreve")
    End Sub

    ''' <summary>
    ''' <b>A FAIXA DE TAREFAS ESTÁ NA TELA — e os dois botões são dois.</b>
    '''
    ''' A regra da Fase 5 é "a sugestão preenche, você confirma, o Iris cria".
    ''' Um <c>TarefasViewModel</c> perfeito num XAML que não o alcança é
    ''' funcionalidade que não existe — foi o que aconteceu com a faixa do
    ''' acervo, que ficou dias escondida atrás de outra borda.
    '''
    ''' A asserção que importa é a do meio: <c>ProporTarefaCommand</c> e
    ''' <c>Tarefas.CriarCommand</c> são <b>dois</b> comandos ligados a dois
    ''' botões. Um XAML que ligasse o botão "Desta mensagem" direto ao criar
    ''' teria a tela criando tarefa sem confirmação, com o ViewModel inteiro
    ''' correto e todos os outros testes verdes.
    ''' </summary>
    <TestMethod>
    Public Sub A_faixa_de_tarefas_tem_propor_E_criar_separados()
        Dim xaml = LerXaml()

        StringAssert.Contains(xaml, "{Binding Tarefas.AbrirCommand}",
            "nao ha como abrir as tarefas: a pasta delas nao aparece na arvore")
        ' DOIS ELEMENTOS <Button>, e nao duas strings no arquivo.
        '
        ' A revisao externa pegou a diferenca: procurar os dois bindings no
        ' texto passaria com os dois no MESMO controle, ou com o botao "Desta
        ' mensagem" ligado ao criar e o ProporTarefaCommand sobrevivendo num
        ' elemento morto em outro canto. O comentario prometia "dois comandos
        ' ligados a dois botoes" e o teste provava bem menos que isso.
        '
        ' O recorte por "<Button" e um proxy -- e por isso a assercao de
        ' controle logo abaixo confere que ele achou botoes de verdade.
        ' Split simples, e nao Regex: a primeira versao usava um padrao com
        ' borda de palavra, e o escape virou um BACKSPACE de verdade no
        ' fonte -- o padrao nunca casava e o recorte achava zero botoes.
        ' So nao passou em silencio porque a assercao de controle existe.
        Dim botoes = xaml.Split({"<Button"}, StringSplitOptions.None).Skip(1).
                     Select(Function(t) t.Substring(0, Math.Min(t.Length, 800))).ToList()

        Assert.IsTrue(botoes.Count > 4,
            "o recorte por <Button achou " & botoes.Count & " botao(oes): ele " &
            "nao esta olhando a janela, e as assercoes abaixo viraram fumaca")

        Dim propoe = botoes.Where(Function(b) b.Contains("{Binding ProporTarefaCommand}")).Count()
        Dim cria = botoes.Where(Function(b) b.Contains("{Binding Tarefas.CriarCommand}")).Count()

        Assert.AreEqual(1, propoe,
            "achei " & propoe & " botao(oes) de propor; a proposta a partir da " &
            "mensagem tem de estar na tela, e uma vez so")
        Assert.AreEqual(1, cria,
            "achei " & cria & " botao(oes) de criar; propor sem criar deixaria " &
            "a tarefa presa no formulario")

        ' E O QUE IMPORTA: nao sao o mesmo botao. Um controle que fizesse as
        ' duas coisas juntaria as duas etapas que a Fase 5 separa de proposito.
        Assert.IsFalse(
            botoes.Any(Function(b) b.Contains("{Binding ProporTarefaCommand}") AndAlso
                                   b.Contains("{Binding Tarefas.CriarCommand}")),
            "propor e criar no MESMO botao: a sugestao viraria tarefa sem " &
            "confirmacao, que e a unica coisa que esta fase existe para impedir")
        StringAssert.Contains(xaml, "{Binding Tarefas.AvisoDaSelecionada}",
            "o botao de concluir pode ficar desabilitado sem a tela dizer por que")
    End Sub

    ''' <summary>
    ''' <b>A faixa de tarefas tem linha própria.</b>
    '''
    ''' O acervo e a agenda dividem a linha 2 porque são mutuamente exclusivos
    ''' — a pasta selecionada é de calendário ou não é. As tarefas não têm essa
    ''' exclusão: a pasta delas nem aparece na árvore. Numa linha compartilhada
    ''' elas empilhariam sobre o vizinho, que é <i>exatamente</i> o defeito que
    ''' escondeu a faixa do acervo.
    ''' </summary>
    <TestMethod>
    Public Sub A_faixa_de_tarefas_NAO_divide_linha_com_ninguem()
        ' A GRADE RAIZ, E NAO O ARQUIVO INTEIRO. O primeiro corte deste teste
        ' contava Grid.Row="4" no XAML todo e falhou por causa de uma grade
        ' ANINHADA -- a do compositor, que tem linhas proprias e nada a ver com
        ' esta. Um teste que acusa colisao onde nao ha e um teste que sera
        ' desligado na primeira vez que atrapalhar.
        '
        ' O recorte usa a INDENTACAO: filho direto da grade raiz abre com doze
        ' espacos. E um proxy, e nao um parser -- e por isso a mensagem de
        ' falha manda conferir, em vez de afirmar.
        Dim raiz = LerXaml().Split(CChar(vbLf)).
                   Where(Function(l) l.StartsWith("            <") AndAlso
                                     Not l.StartsWith("             ")).ToList()

        Dim ocupantes = raiz.Where(Function(l) l.Contains("Grid.Row=""4""")).Count()

        Assert.AreEqual(1, ocupantes,
            "a linha 4 da grade raiz tem " & ocupantes & " ocupante(s) diretos. " &
            "Duas bordas na mesma linha se empilham e a de cima cobre a de " &
            "baixo -- foi assim que a faixa do acervo ficou dias invisivel. " &
            "Confira se o recorte por indentacao ainda vale antes de mexer no numero.")

        ' CONTROLE: o recorte acha os vizinhos conhecidos. Sem isto, uma
        ' mudanca de indentacao esvaziaria a lista e o teste passaria por
        ' nao olhar nada -- o bloqueio que nunca bloqueia.
        Assert.AreEqual(1, raiz.Where(Function(l) l.Contains("Grid.Row=""3""")).Count(),
            "o recorte por indentacao parou de achar a faixa da IA: ele nao " &
            "esta mais olhando a grade raiz, e a assercao acima virou fumaca")
    End Sub

    ''' <summary>
    ''' <b>A RESSALVA DOS CONTATOS TEM LUGAR FIXO NA TELA.</b>
    '''
    ''' Nesta caixa a pasta pessoal de Contatos tem zero itens, e a
    ''' organizacao inteira e enderecavel pelo GAL, que esta fora de escopo.
    ''' Entao a lista vazia e o caso NORMAL desta faixa -- e uma faixa que
    ''' mostrasse o vazio e calasse afirmaria ausencia a partir de nao ter
    ''' olhado, sobre pessoas.
    '''
    ''' O ViewModel podia estar perfeito: se o XAML nao alcancasse a
    ''' ressalva, ela nao existiria para quem usa.
    ''' </summary>
    <TestMethod>
    Public Sub A_ressalva_dos_contatos_esta_na_tela()
        Dim xaml = LerXaml()

        StringAssert.Contains(xaml, "{Binding Contatos.Ressalva}",
            "a ressalva do GAL nao chegou a tela: a lista vazia passa a " &
            "significar ausencia de contatos, que e falso nesta caixa")
        StringAssert.Contains(xaml, "{Binding Contatos.TemRessalva,",
            "sem a visibilidade a ressalva fica sempre visivel ou nunca; " &
            "as duas coisas sao piores que a condicional")
        StringAssert.Contains(xaml, "{Binding Contatos.Aviso}",
            "o aviso de ficha repetida nao chegou a tela")
    End Sub

    ''' <summary>
    ''' <b>Propor e criar contato tambem sao dois botoes.</b>
    '''
    ''' Mesma prova da Fase 5, e pelo mesmo motivo -- so que aqui o estrago
    ''' dura mais: compromisso repetido alguem apaga, ficha repetida fica no
    ''' catalogo com dados divergentes.
    ''' </summary>
    <TestMethod>
    Public Sub A_faixa_de_contatos_tem_propor_E_criar_separados()
        Dim xaml = LerXaml()
        Dim botoes = xaml.Split({"<Button"}, StringSplitOptions.None).Skip(1).
                     Select(Function(t) t.Substring(0, Math.Min(t.Length, 800))).ToList()

        Assert.IsTrue(botoes.Count > 4,
            "o recorte por <Button achou " & botoes.Count & " botao(oes): ele " &
            "nao esta olhando a janela, e as assercoes abaixo viraram fumaca")

        Assert.AreEqual(1, botoes.Where(Function(b) b.Contains("{Binding ProporContatoCommand}")).Count(),
            "a proposta a partir do remetente tem de estar na tela, e uma vez so")
        Assert.AreEqual(1, botoes.Where(Function(b) b.Contains("{Binding Contatos.CriarCommand}")).Count(),
            "propor sem criar deixaria o contato preso no formulario")

        Assert.IsFalse(
            botoes.Any(Function(b) b.Contains("{Binding ProporContatoCommand}") AndAlso
                                   b.Contains("{Binding Contatos.CriarCommand}")),
            "propor e criar no MESMO botao: a sugestao viraria ficha sem " &
            "confirmacao, e ficha repetida nao some sozinha")
    End Sub

    ''' <summary>
    ''' <b>A faixa de contatos tem linha propria.</b>
    '''
    ''' Mesmo recorte por indentacao do teste das tarefas, e mesmo controle:
    ''' sem ele, uma mudanca de formatacao esvaziaria a lista e o teste
    ''' passaria por nao olhar nada.
    ''' </summary>
    <TestMethod>
    Public Sub A_faixa_de_contatos_NAO_divide_linha_com_ninguem()
        Dim raiz = LerXaml().Split(CChar(vbLf)).
                   Where(Function(l) l.StartsWith("            <") AndAlso
                                     Not l.StartsWith("             ")).ToList()

        Dim ocupantes = raiz.Where(Function(l) l.Contains("Grid.Row=""5""")).Count()

        Assert.AreEqual(1, ocupantes,
            "a linha 5 da grade raiz tem " & ocupantes & " ocupante(s) diretos. " &
            "Duas bordas na mesma linha se empilham e a de cima cobre a de baixo.")

        Assert.AreEqual(1, raiz.Where(Function(l) l.Contains("Grid.Row=""4""")).Count(),
            "o recorte por indentacao parou de achar a faixa das tarefas: ele " &
            "nao esta mais olhando a grade raiz, e a assercao acima virou fumaca")
    End Sub

    ''' <summary>
    ''' <b>O GRAU DO ACHADO CHEGA NA TELA — nos dois lugares.</b>
    '''
    ''' A busca ganhou um segundo passe tolerante a erro de digitação e a
    ''' flexão. Medido em 29/08 sobre o acervo real: erro de digitação foi de
    ''' 0,4% para 93,8%, e flexão de 0% para 100%.
    '''
    ''' O ganho tem preço, e o preço é ruído. Um achado aproximado é um palpite
    ''' bom, e palpite misturado com certeza é a família de defeito que este
    ''' projeto passou uma série de revisões corrigindo — aqui ela apareceria
    ''' como "a busca achou", quando o certo é "a busca achou algo parecido".
    '''
    ''' <b>Dois lugares, e os dois são necessários.</b> O rodapé diz quantos
    ''' são palpite; a marca na linha diz <i>quais</i>. Só o rodapé deixaria o
    ''' usuário sabendo que há três palpites entre quinze linhas, sem saber
    ''' onde eles estão.
    ''' </summary>
    <TestMethod>
    Public Sub O_grau_do_achado_chega_na_tela()
        Dim xaml = LerXaml()

        StringAssert.Contains(xaml, "{Binding Busca.FraseDosAproximados}",
            "a tela nao diz quantos achados sao aproximados: palpite e achado " &
            "ficam indistinguiveis na mesma lista")
        StringAssert.Contains(xaml, "{Binding Busca.TemAproximados,",
            "sem a visibilidade a frase fica sempre visivel ou nunca")
        StringAssert.Contains(xaml, "{Binding Aproximado,",
            "a marca nao chegou na LINHA: o rodape diz quantos sao palpite e " &
            "nao diz quais")
    End Sub

    ''' <summary>
    ''' <b>A COLETA SE DECLARA NA TELA.</b>
    '''
    ''' O dono autorizou o diário de buscas em 30/08. Autorizar não é o mesmo
    ''' que querer esquecer que ele existe — e um registro de comportamento que
    ''' some da vista é exatamente o que ninguém deveria aceitar de software
    ''' nenhum, inclusive deste.
    '''
    ''' Três coisas, e as três são necessárias: que está anotando, <b>onde</b> o
    ''' arquivo está, e um botão de apagar. Sem o caminho, o dono não consegue
    ''' localizar o que autorizou; sem o botão, não consegue desfazer. Poder
    ''' apagar era metade do acordo.
    ''' </summary>
    <TestMethod>
    Public Sub O_diario_de_buscas_se_declara_na_tela()
        Dim xaml = LerXaml()

        StringAssert.Contains(xaml, "{Binding Busca.FraseDoDiario}",
            "a tela nao diz que esta anotando as buscas: coleta autorizada " &
            "virou coleta silenciosa")
        StringAssert.Contains(xaml, "{Binding Busca.ApagarDiarioCommand}",
            "nao ha como apagar o registro pela tela, e poder apagar era " &
            "metade do acordo com o dono")
        StringAssert.Contains(xaml, "{Binding Busca.RegistrandoBuscas,",
            "sem a visibilidade a faixa aparece mesmo quando nao ha diario, " &
            "e passa a anunciar uma coleta que nao acontece")
        StringAssert.Contains(xaml, "{Binding Busca.AlternarDiarioCommand}",
            "nao ha como DESLIGAR a coleta pela tela. Apagar nao serve: a " &
            "busca seguinte recria o arquivo, entao apagar nunca chega a ser " &
            "'pare de coletar'.")
        StringAssert.Contains(xaml, "{Binding Busca.RotuloDoInterruptor}",
            "o botao nao diz em que estado a coleta esta")
        StringAssert.Contains(xaml, "{Binding Busca.TemFalhaDoDiario,",
            "a falha do diario nao chega a tela: amostra furada que ninguem " &
            "sabe que e furada")
    End Sub

    Private Shared Function LerXaml() As String
        If _xaml IsNot Nothing Then Return _xaml
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
            d = d.Parent
        End While
        Assert.IsNotNull(d, "nao achei a raiz do repositorio")
        Dim caminho = Path.Combine(d.FullName, "src", "Iris.App", "MainWindow.xaml")
        Assert.IsTrue(File.Exists(caminho), "MainWindow.xaml nao encontrado em " & caminho)
        _xaml = File.ReadAllText(caminho)
        Return _xaml
    End Function

End Class

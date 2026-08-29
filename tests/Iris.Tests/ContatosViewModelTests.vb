Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.App.ViewModels
Imports Iris.Core
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A TELA DE CONTATOS — e a frase que ela não pode dizer.</b>
'''
''' O primeiro teste daqui é o mais importante da fase, e é sobre uma lista
''' <b>vazia</b>: nesta caixa a pasta pessoal de Contatos tem 0 itens, e a
''' organização inteira é endereçável pelo GAL, que está fora de escopo.
'''
''' Uma tela que mostrasse a lista vazia e calasse afirmaria ausência a partir
''' de não ter olhado — sobre pessoas. Tudo o mais nesta fase é consequência
''' disso.
''' </summary>
<TestClass>
Public Class ContatosViewModelTests

    Private Shared ReadOnly Pasta As New FolderKey("contatos-1", "store-1")

    Private NotInheritable Class BrokerDeContatos
        Implements IContatosBroker

        Friend ReadOnly Chamadas As New List(Of String)()
        Friend Itens As New List(Of ContactInfo)()
        Friend Recusados As Integer? = 0
        Friend Truncada As Boolean
        Friend Recusa As String
        Friend FalhaAoLer As Boolean
        Friend UltimoRascunho As ContactDraft
        Friend UltimaPasta As FolderKey
        Friend Trava As TaskCompletionSource(Of Boolean)
        Friend TravaDaPasta As TaskCompletionSource(Of Boolean)
        Friend TravaDaLeitura As TaskCompletionSource(Of Boolean)
        Friend LancaAoLer As Boolean
        Friend SemPasta As Boolean

        ''' <summary>
        ''' O leitor devolve SUCESSO com a ressalva VAZIA.
        '''
        ''' Existe porque a revisao externa mostrou o buraco: o duplo
        ''' sempre injetava a ressalva certa, entao o teste central da
        ''' fase nao conseguia enxergar a tela apagando-a no caminho de
        ''' sucesso. Um duplo que so sabe fazer a coisa certa nao prova
        ''' que a tela faz a coisa certa.
        ''' </summary>
        Friend SemRessalva As Boolean

        Public Async Function GetDefaultContactsFolderAsync(cancel As CancellationToken) _
            As Task(Of OperationResult(Of FolderKey)) _
            Implements IContatosBroker.GetDefaultContactsFolderAsync

            Chamadas.Add("pasta")
            If TravaDaPasta IsNot Nothing Then Await TravaDaPasta.Task
            If SemPasta Then
                Return OperationResult(Of FolderKey).Fail(ErrorKind.NotFound, "sem pasta")
            End If
            Return OperationResult(Of FolderKey).Ok(Pasta)
        End Function

        Public Async Function GetContactsAsync(folder As FolderKey, teto As Integer,
                                               cancel As CancellationToken) _
            As Task(Of OperationResult(Of ContactList)) _
            Implements IContatosBroker.GetContactsAsync

            Chamadas.Add("ler")
            If TravaDaLeitura IsNot Nothing Then Await TravaDaLeitura.Task
            If LancaAoLer Then Throw New InvalidOperationException("o COM caiu")
            If FalhaAoLer Then
                Return OperationResult(Of ContactList).Fail(ErrorKind.Busy, "ocupado")
            End If

            ' O DUPLO IMITA O LEITOR DE VERDADE: quem marca a ressalva e o
            ' ContactWriting, e um duplo que a esquecesse faria o teste da
            ' ressalva provar a tela em vez do desenho.
            Dim lista As New ContactList With {
                .Skipped = Recusados,
                .Truncada = Truncada,
                .ForaDoAlcance = If(SemRessalva, "", RegrasDeContato.ForaDoAlcance)}
            lista.Items.AddRange(Itens)
            Return OperationResult(Of ContactList).Ok(lista)
        End Function

        Public Async Function CreateContactAsync(folder As FolderKey, rascunho As ContactDraft,
                                                 cancel As CancellationToken) _
            As Task(Of OperationResult(Of ContactInfo)) _
            Implements IContatosBroker.CreateContactAsync

            Chamadas.Add("criar")
            UltimaPasta = folder
            UltimoRascunho = rascunho
            If Trava IsNot Nothing Then Await Trava.Task
            If Recusa IsNot Nothing Then
                Return OperationResult(Of ContactInfo).Fail(ErrorKind.Denied, Recusa)
            End If

            Dim novo = New ContactInfo With {
                .Key = New ItemKey("novo", "store-1"),
                .Nome = rascunho.Nome,
                .Email = rascunho.Email}
            Itens.Add(novo)
            Return OperationResult(Of ContactInfo).Ok(novo)
        End Function
    End Class

    Private Shared Function Contato(nome As String, email As String) As ContactInfo
        Return New ContactInfo With {
            .Key = New ItemKey(nome, "store-1"), .Nome = nome, .Email = email}
    End Function

    Private Shared Async Function Aberta(b As BrokerDeContatos) As Task(Of ContatosViewModel)
        Dim vm As New ContatosViewModel(b)
        Await vm.AbrirCommand.ExecuteAsync(Nothing)
        Return vm
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>LISTA VAZIA NÃO DIZ "NENHUM CONTATO" — E CARREGA A RESSALVA.</b>
    '''
    ''' O teste central da Fase 7. Nesta caixa a pasta pessoal tem 0 itens e a
    ''' organização inteira é endereçável: o vazio da pasta não é vazio do
    ''' catálogo, e uma tela que não disser isso mente por silêncio.
    '''
    ''' <b>Controle negativo:</b> fazendo o <c>CarregarAsync</c> deixar de
    ''' copiar <c>ForaDoAlcance</c>, a segunda asserção cai.
    ''' </summary>
    <TestMethod>
    Public Async Function Pasta_vazia_ressalva_o_GAL_e_nao_afirma_ausencia() As Task
        Dim b As New BrokerDeContatos()
        Dim vm = Await Aberta(b)

        Assert.AreEqual(0, vm.Contatos.Count, "o duplo devia vir vazio")

        StringAssert.Contains(vm.Resumo, "nenhum contato LIDO",
            "a tela afirmou ausência a partir de não ter olhado: " & vm.Resumo)
        Assert.IsTrue(vm.TemRessalva, "a lista veio vazia e a tela não ressalvou nada")
        StringAssert.Contains(vm.Ressalva, "GAL",
            "a ressalva não nomeia o catálogo que está fora do alcance")
    End Function

    ''' <summary>
    ''' <b>E NÃO SOME NEM QUANDO O LEITOR ESQUECE DE MANDÁ-LA.</b>
    '''
    ''' Este teste é o que faltava, e a revisão externa o nomeou: o duplo
    ''' sempre injetava a ressalva certa, então o teste da lista vazia não
    ''' conseguia enxergar a tela apagando-a no caminho de sucesso.
    '''
    ''' A invariante é <i>a tela ressalva sempre que houve leitura</i>, e não
    ''' <i>a tela repassa o que o leitor mandou</i>. A diferença aparece
    ''' exatamente aqui — num leitor futuro, num duplo, numa implementação que
    ''' ninguém escreveu ainda.
    '''
    ''' <b>Controle negativo:</b> voltando a atribuição para
    ''' <c>Ressalva = r.Value.ForaDoAlcance</c>, este teste cai.
    ''' </summary>
    <TestMethod>
    Public Async Function A_ressalva_sobrevive_a_um_leitor_que_a_esquece() As Task
        Dim b As New BrokerDeContatos() With {.SemRessalva = True}
        b.Itens.Add(Contato("Ana", "ana@empresa.com"))
        Dim vm = Await Aberta(b)

        Assert.IsTrue(vm.TemRessalva,
            "o leitor mandou sucesso com a ressalva vazia e a tela calou: " &
            "lista na tela sem dizer o que ela não alcança")
        StringAssert.Contains(vm.Ressalva, "GAL")
    End Function

    ''' <summary>
    ''' <b>Leitura que EXPLODE também mantém a ressalva, e limpa o resumo.</b>
    '''
    ''' O caminho do <c>Catch</c> não tinha teste: o duplo só sabia devolver
    ''' <c>Fail</c>. Uma exceção deixava o resumo antigo na tela, afirmando uma
    ''' contagem sobre uma leitura que não aconteceu.
    ''' </summary>
    <TestMethod>
    Public Async Function Leitura_que_explode_mantem_a_ressalva_e_limpa_o_resumo() As Task
        Dim b As New BrokerDeContatos()
        b.Itens.Add(Contato("Ana", "ana@empresa.com"))
        Dim vm = Await Aberta(b)
        Assert.IsTrue(vm.Resumo.Contains("1 contato"), "controle: a primeira leitura contou")

        b.LancaAoLer = True
        Await vm.AtualizarCommand.ExecuteAsync(Nothing)

        Assert.IsTrue(vm.TemErro, "a exceção não apareceu")
        Assert.IsTrue(vm.TemRessalva, "a ressalva sumiu no caminho da exceção")
        Assert.AreEqual("", vm.Resumo,
            "o resumo antigo ficou afirmando uma contagem que esta leitura não fez")
    End Function

    ''' <summary>
    ''' <b>Depois de uma atualização falhar, o aviso diz de quando ele é.</b>
    '''
    ''' A lista fica — apagá-la perderia informação verdadeira. Mas o aviso de
    ''' ficha repetida é uma afirmação sobre o que foi lido, e continuar
    ''' falando no presente sobre uma leitura que já não vale é a mesma família
    ''' de defeito da fase: dizer mais do que se sabe.
    ''' </summary>
    <TestMethod>
    Public Async Function Aviso_apos_falha_diz_que_e_da_leitura_anterior() As Task
        Dim b As New BrokerDeContatos()
        b.Itens.Add(Contato("Ana Lima", "ana@empresa.com"))
        Dim vm = Await Aberta(b)

        vm.ProporDoRemetente("Ana L.", "ana@empresa.com")
        Assert.IsTrue(vm.TemAviso, "controle: o aviso apareceu com leitura boa")
        Assert.IsFalse(vm.Aviso.Contains("pode estar desatualizado"),
                       "controle: com leitura boa o aviso não se desculpa")

        b.FalhaAoLer = True
        Await vm.AtualizarCommand.ExecuteAsync(Nothing)

        Assert.IsTrue(vm.TemAviso, "o aviso sumiu, e a informação era verdadeira")
        StringAssert.Contains(vm.Aviso, "pode estar desatualizado",
            "o aviso continuou falando no presente sobre uma leitura que falhou")
    End Function

    ''' <summary>
    ''' <b>Descartar durante a ABERTURA não notifica a tela descartada.</b>
    '''
    ''' A geração já protegia o corpo do <c>CarregarAsync</c>; o que não estava
    ''' protegido era o <c>Finally</c> do <c>AbrirAsync</c>, que mexia em
    ''' <c>Carregando</c> e nos comandos depois do <c>Dispose</c>.
    ''' </summary>
    <TestMethod>
    Public Async Function Descartar_durante_a_abertura_nao_notifica() As Task
        Dim b As New BrokerDeContatos() With {
            .TravaDaPasta = New TaskCompletionSource(Of Boolean)()}
        Dim vm As New ContatosViewModel(b)

        Dim mudancas = 0
        AddHandler vm.PropertyChanged, Sub(remetente, args) mudancas += 1

        Dim emVoo = vm.AbrirCommand.ExecuteAsync(Nothing)
        vm.Dispose()
        Dim antes = mudancas
        b.TravaDaPasta.SetResult(True)
        Await emVoo

        Assert.AreEqual(antes, mudancas,
            "a tela descartada recebeu notificação: o Finally do AbrirAsync " &
            "não conferia _disposed")
        Assert.IsFalse(vm.TemPasta, "instalou a pasta numa tela já descartada")
    End Function

    ''' <summary>
    ''' <b>E a ressalva NÃO some quando a leitura falha.</b>
    '''
    ''' É o momento em que ela mais importa: sem lista nenhuma na tela, um
    ''' usuário lê "não consegui ler os contatos" e conclui que os contatos
    ''' estão ali, atrás da falha. Metade deles nunca esteve.
    ''' </summary>
    <TestMethod>
    Public Async Function A_ressalva_sobrevive_a_falha_de_leitura() As Task
        Dim b As New BrokerDeContatos() With {.FalhaAoLer = True}
        Dim vm = Await Aberta(b)

        Assert.IsTrue(vm.TemErro, "a falha não apareceu")
        Assert.IsTrue(vm.TemRessalva,
            "a ressalva sumiu justamente quando não há lista para contradizê-la")
    End Function

    ''' <summary>
    ''' <b>CONTROLE POSITIVO: com contatos lidos, a tela conta.</b>
    '''
    ''' Sem ele, um resumo que dissesse "nenhum contato LIDO" sempre passaria no
    ''' teste de cima — o bloqueio que nunca deixa passar nada.
    ''' </summary>
    <TestMethod>
    Public Async Function Controle_com_contatos_a_tela_conta() As Task
        Dim b As New BrokerDeContatos()
        b.Itens.Add(Contato("Ana", "ana@empresa.com"))
        b.Itens.Add(Contato("Bruno", "bruno@empresa.com"))
        Dim vm = Await Aberta(b)

        Assert.AreEqual(2, vm.Contatos.Count)
        StringAssert.Contains(vm.Resumo, "2 contato(s) lido(s)")
        Assert.IsTrue(vm.TemRessalva, "a ressalva vale com lista cheia também")
    End Function

    ''' <summary>
    ''' <b>PROPOR NÃO CRIA.</b>
    '''
    ''' A mesma invariante da Fase 5, num catálogo onde o estrago dura mais:
    ''' compromisso repetido alguém apaga, ficha repetida fica.
    ''' </summary>
    <TestMethod>
    Public Async Function Propor_do_remetente_NAO_cria() As Task
        Dim b As New BrokerDeContatos()
        Dim vm = Await Aberta(b)

        vm.ProporDoRemetente("Ana Lima", "ana@empresa.com")

        Assert.AreEqual("Ana Lima", vm.NovoNome)
        Assert.AreEqual("ana@empresa.com", vm.NovoEmail)
        CollectionAssert.DoesNotContain(b.Chamadas, "criar",
            "propor criou o contato sozinho -- e é a única coisa que esta " &
            "separação existe para impedir")
        Assert.IsTrue(vm.TemProposta)
    End Function

    ''' <summary>
    ''' Remetente sem nome de exibição vira uma ficha com o endereço no nome, e
    ''' não uma ficha em branco que o escritor recusaria depois do clique.
    ''' </summary>
    <TestMethod>
    Public Async Function Remetente_sem_nome_propoe_o_endereco() As Task
        Dim b As New BrokerDeContatos()
        Dim vm = Await Aberta(b)

        vm.ProporDoRemetente("", "anonimo@empresa.com")

        Assert.AreEqual("anonimo@empresa.com", vm.NovoNome)
        Assert.IsTrue(vm.PodeCriar, "a proposta ficaria presa num botão desabilitado")
    End Function

    ''' <summary>
    ''' <b>O AVISO DE REPETIDO APARECE — E NÃO BLOQUEIA.</b>
    '''
    ''' Avisar é honesto; recusar não seria. A busca só enxerga o que esta
    ''' leitura trouxe: nunca o GAL, e nem a pasta inteira quando houve
    ''' truncamento. Recusar com essa base transformaria "não encontrei" em
    ''' "não existe" — o erro da fase, virado do avesso.
    ''' </summary>
    <TestMethod>
    Public Async Function Contato_repetido_AVISA_mas_deixa_criar() As Task
        Dim b As New BrokerDeContatos()
        b.Itens.Add(Contato("Ana Lima", "ana@empresa.com"))
        Dim vm = Await Aberta(b)

        vm.ProporDoRemetente("Ana L.", "ANA@empresa.com")

        Assert.IsTrue(vm.TemAviso, "não avisou que já há ficha com este endereço")
        StringAssert.Contains(vm.Aviso, "Ana Lima", "o aviso não diz de quem é a ficha")
        Assert.IsTrue(vm.PodeCriar,
            "o aviso virou bloqueio: a busca só viu o que esta leitura trouxe, " &
            "e homônimo é comum e legítimo")
    End Function

    ''' <summary>
    ''' <b>CONTROLE: endereço novo não gera aviso.</b>
    '''
    ''' Sem ele, um aviso que aparecesse sempre passaria no teste de cima.
    ''' </summary>
    <TestMethod>
    Public Async Function Controle_endereco_novo_nao_avisa() As Task
        Dim b As New BrokerDeContatos()
        b.Itens.Add(Contato("Ana Lima", "ana@empresa.com"))
        Dim vm = Await Aberta(b)

        vm.ProporDoRemetente("Carlos", "carlos@empresa.com")

        Assert.IsFalse(vm.TemAviso, "avisou repetido sobre um endereço novo: " & vm.Aviso)
    End Function

    ''' <summary>Criar manda para a pasta que a abertura descobriu, e recarrega.</summary>
    <TestMethod>
    Public Async Function Criar_usa_a_pasta_de_contatos_e_recarrega() As Task
        Dim b As New BrokerDeContatos()
        Dim vm = Await Aberta(b)

        vm.ProporDoRemetente("Ana Lima", "ana@empresa.com")
        Await vm.CriarCommand.ExecuteAsync(Nothing)

        Assert.AreEqual(Pasta, b.UltimaPasta, "criou noutra pasta")
        Assert.AreEqual("Ana Lima", b.UltimoRascunho.Nome)
        Assert.AreEqual("ana@empresa.com", b.UltimoRascunho.Email)
        Assert.AreEqual(2, b.Chamadas.Where(Function(c) c = "ler").Count(),
                        "não recarregou depois de criar")
        Assert.AreEqual("", vm.NovoNome, "o formulário não limpou")
    End Function

    ''' <summary>A recusa do escritor aparece com as palavras dele.</summary>
    <TestMethod>
    Public Async Function A_recusa_do_escritor_aparece_na_tela() As Task
        Dim b As New BrokerDeContatos() With {
            .Recusa = "um contato sem nome vira uma ficha anônima no catálogo"}
        Dim vm = Await Aberta(b)

        vm.ProporDoRemetente("Ana", "ana@empresa.com")
        Await vm.CriarCommand.ExecuteAsync(Nothing)

        Assert.IsTrue(vm.TemErro)
        StringAssert.Contains(vm.Erro, "ficha anônima", "a tela engoliu o motivo: " & vm.Erro)
    End Function

    ''' <summary>
    ''' <b>O segundo clique não cria a segunda ficha.</b>
    '''
    ''' A segunda chamada vai <b>direto ao método</b>, e não pelo comando: o
    ''' <c>AsyncRelayCommand</c> serializa sozinho, e um teste que passasse por
    ''' ele provaria o toolkit. Foi a lição da revisão da Fase 5.
    ''' </summary>
    <TestMethod>
    Public Async Function O_segundo_clique_nao_cria_de_novo() As Task
        Dim b As New BrokerDeContatos() With {.Trava = New TaskCompletionSource(Of Boolean)()}
        Dim vm = Await Aberta(b)

        vm.ProporDoRemetente("Ana Lima", "ana@empresa.com")
        Dim emVoo = vm.CriarCommand.ExecuteAsync(Nothing)
        Dim segundo = vm.CriarAsync()

        b.Trava.SetResult(True)
        Await emVoo
        Await segundo

        Assert.AreEqual(1, b.Chamadas.Where(Function(c) c = "criar").Count(),
                        "criou duas fichas da mesma pessoa")
    End Function

    ''' <summary>Dois cliques em "Abrir" não descobrem a pasta duas vezes.</summary>
    <TestMethod>
    Public Async Function Dois_cliques_em_abrir_descobrem_a_pasta_uma_vez() As Task
        Dim b As New BrokerDeContatos() With {
            .TravaDaPasta = New TaskCompletionSource(Of Boolean)()}
        Dim vm As New ContatosViewModel(b)

        Dim um = vm.AbrirCommand.ExecuteAsync(Nothing)
        Dim dois = vm.AbrirCommand.ExecuteAsync(Nothing)

        b.TravaDaPasta.SetResult(True)
        Await um
        Await dois

        Assert.AreEqual(1, b.Chamadas.Where(Function(c) c = "pasta").Count(),
                        "duas descobertas concorrentes atribuindo _pasta em corrida")
    End Function

    ''' <summary>Descartar durante a gravação não ressuscita a tela.</summary>
    <TestMethod>
    Public Async Function Descartar_durante_a_gravacao_nao_recarrega() As Task
        Dim b As New BrokerDeContatos() With {.Trava = New TaskCompletionSource(Of Boolean)()}
        Dim vm = Await Aberta(b)
        Dim lidasAntes = b.Chamadas.Where(Function(c) c = "ler").Count()

        vm.ProporDoRemetente("Ana Lima", "ana@empresa.com")
        Dim emVoo = vm.CriarCommand.ExecuteAsync(Nothing)

        vm.Dispose()
        b.Trava.SetResult(True)
        Await emVoo

        Assert.AreEqual(lidasAntes, b.Chamadas.Where(Function(c) c = "ler").Count(),
                        "recarregou depois do Dispose")
        Assert.AreEqual("Ana Lima", vm.NovoNome,
                        "mexeu no formulário de uma tela já descartada")
    End Function

    ''' <summary>
    ''' <b>Recusados desconhecidos não viram zero.</b>
    '''
    ''' <c>Nothing</c> é "não contei"; zero é "contei e não houve".
    ''' </summary>
    <TestMethod>
    Public Async Function Recusados_desconhecidos_nao_viram_zero() As Task
        Dim b As New BrokerDeContatos() With {.Recusados = Nothing}
        b.Itens.Add(Contato("Ana", "ana@empresa.com"))
        Dim vm = Await Aberta(b)

        StringAssert.Contains(vm.Resumo, "não sei quantos itens foram recusados")

        Dim b2 As New BrokerDeContatos() With {.Recusados = 0}
        b2.Itens.Add(Contato("Ana", "ana@empresa.com"))
        Dim vm2 = Await Aberta(b2)
        Assert.IsFalse(vm2.Resumo.Contains("não sei quantos"))
    End Function

    ''' <summary>
    ''' <b>Campo ilegível não vira campo vazio na tela.</b>
    '''
    ''' "sem e-mail" e "não consegui ler" são frases diferentes porque são
    ''' coisas diferentes — num catálogo de endereços, a primeira é uma
    ''' afirmação sobre o cadastro de uma pessoa.
    ''' </summary>
    <TestMethod>
    Public Async Function Campo_ilegivel_e_campo_vazio_sao_frases_diferentes() As Task
        Dim b As New BrokerDeContatos()
        b.Itens.Add(New ContactInfo With {.Nome = "Ana", .Email = "", .Empresa = Nothing})
        Dim vm = Await Aberta(b)

        Dim linha = vm.Contatos.Single()
        Assert.AreEqual("sem e-mail", linha.Email, "vazio de verdade tem frase própria")
        Assert.AreEqual("não consegui ler", linha.Empresa,
                        "campo ilegível virou afirmação sobre o cadastro")
    End Function

End Class

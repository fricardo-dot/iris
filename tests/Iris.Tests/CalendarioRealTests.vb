Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Threading
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A LEITURA DO CALENDÁRIO, CONTRA O OUTLOOK DE VERDADE.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE NÃO DÁ PARA PROVAR ISTO COM DUPLO</b>
'''
''' O que a leitura do calendário faz de específico é <b>pedir ao Outlook</b>
''' que expanda séries dentro de uma janela. Um duplo que devolvesse uma
''' lista provaria a tradução do DTO e nada mais — e a tradução não é onde
''' mora o risco.
'''
''' O risco mora em três lugares, e nenhum deles é testável sem COM:
'''
'''   • a <b>ordem</b> <c>Sort → IncludeRecurrences → Restrict</c>, que o
'''     Outlook aceita fora de ordem e responde errado, sem erro;
'''   • o <b>formato da data</b> no <c>Restrict</c>, que não é ISO e não pode
'''     depender do idioma do Office;
'''   • o <b>fuso</b>. É a mesma armadilha que a Q1 pegou na paginação de
'''     mensagens, onde o filtro DASL interpretou a data como UTC, a Table
'''     devolveu hora local, e a paginação perdeu 803 de 1.003 mensagens
'''     <i>terminando parecendo completa</i>.
'''
''' ------------------------------------------------------------------
''' <b>SOMENTE LEITURA</b>
'''
''' Nenhum teste aqui cria, move, apaga ou responde convite. A caixa do dono
''' sai destes testes exatamente como entrou.
'''
''' Requer Outlook clássico aberto. Sem ele: <b>inconclusivo</b>, nunca
''' verde — "não deu para verificar" não pode virar "verificado" (§22.12).
''' </summary>
<TestClass>
Public Class CalendarioRealTests

    Private Shared Async Function AbrirAsync() As Task(Of OutlookBroker)
        Dim broker As New OutlookBroker(New NullLog())
        broker.Start()
        Dim estado = Await broker.ConnectAsync(CancellationToken.None)

        Select Case PagingIntegrationTests.Decidir(estado = SessionState.Connected,
                                                   PagingIntegrationTests.OutlookInstalado())
            Case PagingIntegrationTests.SemOutlook.Prosseguir
                Return broker
            Case PagingIntegrationTests.SemOutlook.Falhar
                broker.Dispose()
                Assert.Fail("O Outlook esta INSTALADO e nao respondeu. Abra-o e rode de novo.")
                Return Nothing
            Case Else
                broker.Dispose()
                Assert.Inconclusive("Outlook nao instalado: ambiente sem suporte.")
                Return Nothing
        End Select
    End Function

    ''' <summary>A pasta de calendário padrão, pelo caminho de produção.</summary>
    Private Shared Async Function CalendarioAsync(b As OutlookBroker) As Task(Of FolderKey)
        Dim stores = Await b.GetStoresAsync(CancellationToken.None)
        Assert.IsTrue(stores.Succeeded, $"nao consegui listar stores: {stores.Kind}")

        For Each store In stores.Value
            Dim filhas = Await b.GetFolderChildrenAsync(store.RootFolder, CancellationToken.None)
            If Not filhas.Succeeded Then Continue For
            Dim cal = filhas.Value.FirstOrDefault(
                Function(f) f.ContentKind = FolderContentKind.Calendar)
            If cal IsNot Nothing Then Return cal.Key
        Next

        Assert.Inconclusive("Nenhuma pasta de calendario nesta conta.")
        Return Nothing
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>CONTROLE POSITIVO: uma janela larga devolve compromissos.</b>
    '''
    ''' Vem primeiro, e sem ele os outros não valem nada: uma leitura que
    ''' sempre devolve vazio passa em todo teste de "não devolve demais".
    '''
    ''' A janela é o mês corrente inteiro, porque a medição de 28/08/2026
    ''' contou <b>434 compromissos</b> nesta caixa — um mês qualquer tem o
    ''' que mostrar.
    ''' </summary>
    <TestMethod>
    Public Async Function Controle_a_janela_do_mes_devolve_compromissos() As Task
        Dim b = Await AbrirAsync()
        Using b
            Dim cal = Await CalendarioAsync(b)
            Dim hoje = DateTimeOffset.Now
            Dim de = New DateTimeOffset(hoje.Year, hoje.Month, 1, 0, 0, 0, hoje.Offset)
            Dim ate = de.AddMonths(1)

            Dim r = Await b.GetAppointmentsAsync(cal, de, ate, CancellationToken.None)

            Assert.IsTrue(r.Succeeded, $"a leitura falhou: {r.Kind} / {r.Detail}")
            Assert.IsTrue(r.Value.Items.Count > 0,
                "controle: o mes corrente desta caixa tinha de ter compromisso. " &
                "Se ele estiver mesmo vazio, este teste precisa de outra janela.")
            Assert.AreEqual(de, r.Value.De)
            Assert.AreEqual(ate, r.Value.Ate)
        End Using
    End Function

    ''' <summary>
    ''' <b>Nenhum compromisso escapa da janela pedida.</b>
    '''
    ''' É o teste do FUSO, e ele é o motivo deste arquivo existir. Se o
    ''' <c>Restrict</c> interpretar a data num fuso diferente do que a leitura
    ''' devolve, o resultado vem deslocado — e vem <b>parecendo certo</b>,
    ''' porque continua sendo uma lista de compromissos plausíveis.
    '''
    ''' A comparação é sobre a intersecção, e não sobre <c>Start</c> sozinho:
    ''' um compromisso que começa antes da janela e termina dentro dela
    ''' pertence à janela, e é justamente o caso que um filtro só por
    ''' <c>[Start]</c> perderia.
    '''
    ''' <b>CONTROLE NEGATIVO CONFIRMADO, e o número surpreendeu.</b>
    ''' Invertendo a ordem em <c>CalendarReading</c> — <c>Restrict</c> antes
    ''' de <c>IncludeRecurrences</c> — este teste falha com <b>65</b>
    ''' compromissos fora da janela, ocorrências de janeiro numa janela de
    ''' agosto. Eu esperava que a ordem errada <i>perdesse</i> ocorrências;
    ''' ela faz a expansão ignorar o filtro.
    '''
    ''' <b>Este teste, sozinho, é o que segura a ordem.</b> Todos os outros
    ''' deste arquivo continuam verdes nesse controle — inclusive o da
    ''' expansão, e o comentário dele explica por quê.
    '''
    ''' Ele também pega <c>IncludeRecurrences</c> desligado, com <b>5</b>
    ''' fora da janela: sem expansão os mestres voltam nas datas originais,
    ''' que estão fora. Dois defeitos diferentes, o mesmo sintoma, números
    ''' diferentes.
    ''' </summary>
    <TestMethod>
    Public Async Function Nenhum_compromisso_cai_fora_da_janela() As Task
        Dim b = Await AbrirAsync()
        Using b
            Dim cal = Await CalendarioAsync(b)
            Dim de = DateTimeOffset.Now.Date
            Dim janelaDe = New DateTimeOffset(de, DateTimeOffset.Now.Offset).AddDays(-30)
            Dim janelaAte = janelaDe.AddDays(60)

            Dim r = Await b.GetAppointmentsAsync(cal, janelaDe, janelaAte, CancellationToken.None)
            Assert.IsTrue(r.Succeeded, $"a leitura falhou: {r.Kind} / {r.Detail}")

            Dim fora = r.Value.Items.
                       Where(Function(a) a.End <= janelaDe OrElse a.Start >= janelaAte).
                       ToList()

            Assert.AreEqual(0, fora.Count,
                "compromisso fora da janela pedida — o filtro do Restrict esta " &
                "interpretando a data noutro fuso. Primeiros: " &
                String.Join(" | ", fora.Take(3).Select(
                    Function(a) $"{a.Subject} {a.Start:yyyy-MM-dd HH:mm}")))
        End Using
    End Function

    ''' <summary>
    ''' <b>A expansão de séries acontece — provada pela mesma série aparecendo
    ''' em dias diferentes.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A PRIMEIRA VERSÃO DESTE TESTE NÃO PROVAVA NADA</b>
    '''
    ''' Ela cobrava <c>metade1 + metade2 &gt;= todo</c>, chamando isso de prova
    ''' por cruzamento. A revisão externa mostrou que a desigualdade vale
    ''' <b>com ou sem</b> expansão — e o meu próprio controle negativo já tinha
    ''' dito isso e eu não vi: invertendo a ordem em <c>CalendarReading</c>,
    ''' aquele teste ficou <b>verde</b> enquanto o irmão da janela caía com 65
    ''' compromissos fora do lugar.
    '''
    ''' O que separa expansão ligada de desligada é a <b>mesma série aparecendo
    ''' mais de uma vez</b> na janela. Sem expansão, o mestre aparece uma vez
    ''' só, na data original, e nenhum assunto se repete em dois dias.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>E A SEGUNDA VERSÃO TAMBÉM NÃO GUARDA A ORDEM. MEDIDO.</b>
    '''
    ''' Eu troquei o teste achando que ele passaria a pegar a ordem invertida.
    ''' Não pega — e o motivo é interessante: <b>com a ordem invertida a
    ''' expansão continua acontecendo</b>. O que quebra é o filtro. Foi por
    ''' isso que o controle daquele dia produziu 65 compromissos <i>fora da
    ''' janela</i> em vez de ocorrências faltando.
    '''
    ''' Os três estados, medidos contra a caixa real em 28/08/2026:
    '''
    ''' <code>
    '''                                  este teste   o da janela
    '''   ordem certa .................... passa ....... passa
    '''   Restrict antes do Include ...... passa ....... FALHA (65 fora)
    '''   IncludeRecurrences desligado ... FALHA ....... FALHA (5 fora)
    ''' </code>
    '''
    ''' Então: <b>este</b> teste guarda que a expansão está LIGADA. O
    ''' <c>Nenhum_compromisso_cai_fora_da_janela</c> guarda a ORDEM. Nenhum dos
    ''' dois guarda os dois, e dizer o contrário seria a terceira vez que eu
    ''' afirmo cobertura que não medi neste mesmo arquivo.
    '''
    ''' <b>Sem série na janela desta caixa</b>, o teste fica inconclusivo. Ele
    ''' não vira verde por ausência de matéria.
    ''' </summary>
    <TestMethod>
    Public Async Function A_serie_aparece_MAIS_DE_UMA_VEZ_na_janela() As Task
        Dim b = Await AbrirAsync()
        Using b
            Dim cal = Await CalendarioAsync(b)
            Dim inicio = New DateTimeOffset(DateTimeOffset.Now.Date, DateTimeOffset.Now.Offset).AddDays(-60)
            Dim fim = inicio.AddDays(120)

            Dim r = Await b.GetAppointmentsAsync(cal, inicio, fim, CancellationToken.None)
            Assert.IsTrue(r.Succeeded, $"a leitura falhou: {r.Kind} / {r.Detail}")

            Dim series = r.Value.Items.Where(Function(a) a.IsRecurring).ToList()
            If series.Count = 0 Then
                Assert.Inconclusive(
                    "Nenhum compromisso de serie nesta janela desta caixa. " &
                    "Sem serie, o teste nao distingue expansao ligada de desligada.")
            End If

            ' O SINAL: mesmo assunto, dias diferentes.
            Dim repetidas = series.
                GroupBy(Function(a) a.Subject).
                Where(Function(g) g.Select(Function(a) a.Start.LocalDateTime.Date).
                                    Distinct().Count() > 1).
                ToList()

            Assert.IsTrue(repetidas.Count > 0,
                $"{series.Count} compromisso(s) marcados como serie, e nenhum aparece em " &
                "mais de um dia da janela: a expansao nao esta acontecendo. " &
                "Confira a ordem Sort -> IncludeRecurrences -> Restrict.")
        End Using
    End Function

    ''' <summary>
    ''' <b>Leitura truncada NÃO se apresenta como completa.</b>
    '''
    ''' Até 28/08/2026 dois caminhos devolviam <c>Ok</c> com lista incompleta:
    ''' exceção no <c>GetNext</c> virava fim da coleção, e o teto do laço virava
    ''' sucesso silencioso. Agora existe <c>Truncada</c>, e este teste cobra que
    ''' o caminho normal <b>não</b> a levante — senão o campo nasceria inútil,
    ''' ligado sempre, e a tela avisaria "lista incompleta" o tempo todo até
    ''' alguém aprender a ignorar o aviso.
    ''' </summary>
    <TestMethod>
    Public Async Function Leitura_normal_NAO_vem_truncada() As Task
        Dim b = Await AbrirAsync()
        Using b
            Dim cal = Await CalendarioAsync(b)
            Dim de = DateTimeOffset.Now
            Dim r = Await b.GetAppointmentsAsync(cal, de, de.AddDays(7), CancellationToken.None)

            Assert.IsTrue(r.Succeeded, $"a leitura falhou: {r.Kind}")
            Assert.IsFalse(r.Value.Truncada,
                $"uma janela de sete dias veio truncada: {r.Value.MotivoDoCorte}")
            Assert.AreEqual("", r.Value.MotivoDoCorte)
        End Using
    End Function

    ''' <summary>
    ''' <b>Janela invertida é recusa, e não lista vazia.</b>
    '''
    ''' Devolver vazio para <c>ate &lt;= de</c> faria um erro de chamada
    ''' parecer uma agenda livre. É a mesma distinção que a busca do acervo
    ''' faz entre "não achei" e "não consegui olhar".
    ''' </summary>
    <TestMethod>
    Public Async Function Janela_invertida_RECUSA() As Task
        Dim b = Await AbrirAsync()
        Using b
            Dim cal = Await CalendarioAsync(b)
            Dim agora = DateTimeOffset.Now

            Dim r = Await b.GetAppointmentsAsync(cal, agora, agora.AddDays(-1), CancellationToken.None)

            Assert.IsFalse(r.Succeeded, "janela invertida devolveu sucesso")
        End Using
    End Function

    ''' <summary>
    ''' <b>O filtro é montado em cultura invariante.</b>
    '''
    ''' Este é o único teste do arquivo que não precisa do Outlook, e está
    ''' aqui porque o assunto é o mesmo. O <c>Restrict</c> não entende ISO, e
    ''' usar a cultura da máquina faria o filtro depender do idioma do Office
    ''' — numa máquina em português, <c>28/08/2026</c> em vez de
    ''' <c>08/28/2026</c>.
    '''
    ''' Já é a segunda vez neste projeto que a cultura ambiente entra onde não
    ''' devia: a primeira foi um teste que media a máquina em vez do código.
    ''' </summary>
    <TestMethod>
    Public Sub O_filtro_do_Restrict_nao_depende_da_cultura()
        Dim de = New DateTimeOffset(2026, 8, 28, 9, 30, 0, TimeSpan.Zero)
        Dim ate = New DateTimeOffset(2026, 9, 4, 9, 30, 0, TimeSpan.Zero)

        Dim antes = Threading.Thread.CurrentThread.CurrentCulture
        Try
            Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("pt-BR")
            Dim emPortugues = CalendarFilter.Janela(de, ate)

            Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US")
            Dim emIngles = CalendarFilter.Janela(de, ate)

            Assert.AreEqual(emIngles, emPortugues,
                "o filtro mudou com a cultura da maquina — e o Outlook nao muda junto")
            StringAssert.Contains(emPortugues, "[Start] <")
            StringAssert.Contains(emPortugues, "[End] >")
        Finally
            Threading.Thread.CurrentThread.CurrentCulture = antes
        End Try
    End Sub

End Class

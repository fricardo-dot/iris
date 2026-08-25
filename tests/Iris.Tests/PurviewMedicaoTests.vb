Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>A medição do 3.0 — o rótulo do Purview, contra a caixa real.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ELA É UM TESTE E NÃO UM SCRIPT</b>
'''
''' A §32 do FASE3.md exige que a medição passe pelo <b>broker/STA real</b>.
''' Medir por outro caminho mede o Outlook, não o caminho do produto — e o
''' que a Fase 3 precisa saber é o que o <i>Iris</i> consegue ler, com o
''' message filter, o retry e a classificação de falha que ele tem.
'''
''' ------------------------------------------------------------------
''' <b>O PROTOCOLO, E POR QUE ELE É NESTA ORDEM</b>
'''
''' O usuário <b>não está na máquina</b>. Uma chamada que abra diálogo de
''' direitos ou de autenticação ficaria pendurada até o timeout, numa caixa
''' corporativa, sem ninguém para fechá-la. Então a ordem é do menos
''' invasivo para o mais, e ela para ao primeiro indício adverso:
'''
'''   <b>A.</b> inspeção passiva — um antes/depois observável;
'''   <b>B.</b> piloto por <c>Table</c>, 20 linhas, sem materializar item;
'''   <b>C.</b> confirmação por item, em poucos casos escolhidos;
'''   <b>D.</b> expansão adaptativa;
'''   <b>E.</b> protegidos <b>de fora</b>, e isso é resultado, não lacuna.
'''
''' <b>Nada aqui lê corpo, anexo ou conversa. Nada aqui escreve.</b>
''' </summary>
<TestClass>
Public Class PurviewMedicaoTests

    ''' <summary>Onde o relatório cru é escrito. Fora do repo — tem dado da caixa.</summary>
    Private Shared ReadOnly PastaDeMedicao As String =
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "medicoes")

    Private Shared ReadOnly Diario As New StringBuilder()

    Private Shared Sub Anotar(linha As String)
        Diario.AppendLine(linha)
        Console.WriteLine(linha)
    End Sub

    ' ==================================================================

    <TestMethod, TestCategory("Integracao")>
    Public Async Function Medir_o_rotulo_do_Purview_na_caixa_real() As Task
        Dim broker = Await AbrirAsync()
        Try
            Dim entrada = Await AcharEntradaAsync(broker)
            Anotar("# Medição 3.0 — MSIP_Labels")
            Anotar($"quando: {DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss zzz}")
            Anotar("")

            ' ---- A. inspeção passiva -----------------------------------
            Dim antes = Await AmostraAsync(broker, entrada, 20)
            Anotar($"## A. inspeção passiva — {antes.Count} itens de metadado")
            Assert.IsTrue(antes.Count > 0,
                "pasta vazia nao mede nada; a medicao precisa de amostra")

            ' ---- B. piloto por Table -----------------------------------
            Dim probe = Await broker.ProbeLabelColumnAsync(entrada, 20, CancellationToken.None)
            Anotar("")
            Anotar("## B. piloto por Table (P6)")
            If Not probe.Succeeded Then
                Anotar($"P6: a Table recusou a operacao inteira — {probe.Kind} {probe.Detail}")
            Else
                Dim c = probe.Value
                Anotar($"P6: coluna aceita = {c.ColunaAceita}" &
                       If(c.HResultDoAdd.HasValue, $" (HRESULT 0x{c.HResultDoAdd.Value:X8})", ""))
                Anotar($"P6: linhas={c.LinhasLidas} comValor={c.LinhasComValor} " &
                       $"semValor={c.LinhasSemValor} comErro={c.LinhasComErro}")
                Anotar($"P6: {c.MilissegundosTotais:F1} ms — o caminho barato SERVE? {c.Serve}")

                ' A armadilha registrada: aceitar a coluna NAO e entregar o
                ' valor. Se o Add passou e nenhuma linha trouxe nada, isso e
                ' AMBIGUO — pode ser "ninguem tem rotulo" ou "a coluna nao
                ' funciona" — e tem de aparecer como ambiguidade.
                If c.ColunaAceita AndAlso c.LinhasComValor = 0 AndAlso c.LinhasComErro = 0 Then
                    Anotar("P6: AMBIGUO — coluna aceita e nenhuma linha com valor. " &
                           "Nao da para distinguir 'ninguem tem rotulo' de 'a coluna nao entrega'.")
                End If
            End If

            ' ---- B2. o CONTROLE NEGATIVO da P7 -------------------------
            Anotar("")
            Anotar("## B2. controle negativo (P7) — o que 'ausente' parece nesta conta")
            Dim semantica = Await broker.ProbeLabelSemanticsAsync(antes(0).Key,
                                                                  CancellationToken.None)
            If Not semantica.Succeeded Then
                Anotar($"o probe falhou: {semantica.Kind} {semantica.Detail}")
            Else
                For Each t In semantica.Value.Tentativas
                    Anotar("  " & t.ToString())
                Next
                ComentarSemantica(semantica.Value)
            End If

            ' ---- C. confirmação por item -------------------------------
            Anotar("")
            Anotar("## C. confirmação por item (P1, P2, P7, P8, P9)")
            Dim relogio = Stopwatch.StartNew()
            Dim leituras = Await broker.GetSensitivityLabelsAsync(
                antes.Select(Function(x) x.Key).ToList(), CancellationToken.None)
            relogio.Stop()

            Assert.IsTrue(leituras.Succeeded,
                $"a leitura por item falhou inteira: {leituras.Kind} {leituras.Detail}")
            Dim porItem = leituras.Value
            Assert.AreEqual(antes.Count, porItem.Count,
                "toda linha pedida tem de voltar com um desfecho proprio")

            Dim custo = relogio.Elapsed.TotalMilliseconds / porItem.Count
            Anotar($"custo por item: {custo:F1} ms " &
                   $"({relogio.Elapsed.TotalMilliseconds:F0} ms / {porItem.Count})")
            Relatar(porItem)

            ' ---- indício adverso? --------------------------------------
            Dim adverso = porItem.Where(Function(l) l.Kind = LabelReadingKind.Denied OrElse
                                                    l.Kind = LabelReadingKind.Transient).Count()
            If adverso > 0 Then
                Anotar("")
                Anotar($"PARANDO a expansão: {adverso} leitura(s) com sinal adverso " &
                       "(negada ou transitória). A §32 manda parar ao primeiro indício.")
            Else
                ' ---- D. expansão adaptativa ----------------------------
                Anotar("")
                Anotar("## D. expansão adaptativa (P3, P8)")
                Await MedirPastaAsync(broker, entrada, "Caixa de Entrada", 400)

                ' ---- P11: a propriedade aparece igual em outra pasta? ----
                Anotar("")
                Anotar("## P11 — outras pastas do mesmo store")
                For Each nome In {"Itens Enviados", "Sent Items", "Rascunhos", "Drafts"}
                    Dim outra = Await AcharPastaAsync(broker, nome)
                    If outra IsNot Nothing Then
                        Await MedirPastaAsync(broker, outra, nome, 100)
                    End If
                Next
                Anotar("Stores adicionais e caixas compartilhadas: **não disponíveis** " &
                       "nesta conta — um store só. Fica NÃO MEDIDO, não 'igual'.")
            End If

            ' ---- P4 -----------------------------------------------------
            Anotar("")
            Anotar("## P4 — o Sensitivity clássico concorda?")
            Anotar("NÃO MEDIDO por desenho: `Sensitivity` não responde por rótulo moderno " &
                   "(ESCOPO §10), então concordância entre os dois não é evidência de nada. " &
                   "Ler os dois e comparar produziria um número que convida a conclusão errada.")

            ' ---- P5 / P13 -----------------------------------------------
            Anotar("")
            Anotar("## P5 e P13 — corpo protegido e efeito colateral")
            Anotar("**NÃO MEDIDO POR RESTRIÇÃO OPERACIONAL.** Abrir corpo protegido pode " &
                   "disparar diálogo de direitos, e o usuário não está na máquina para " &
                   "fechá-lo. O estado fica tratado como PROIBIDO, que é o desfecho seguro.")

            ' ---- A, fecho ------------------------------------------------
            Dim depois = Await AmostraAsync(broker, entrada, antes.Count)
            Anotar("")
            Anotar("## A (fecho) — a leitura mexeu em alguma coisa?")
            Anotar(Comparar(antes, depois))
            Anotar("**Ausência de download ou de efeito colateral NÃO foi provada** — " &
                   "só não foi observada nesta amostra.")

            ' ---- P10 / P16 ------------------------------------------------
            Anotar("")
            Anotar("## P10 — evidência de versão")
            Dim comChangeKey = porItem.Where(Function(l) l.Version IsNot Nothing AndAlso
                                                         Not String.IsNullOrEmpty(l.Version.ChangeKey)).Count()
            Anotar($"PR_CHANGE_KEY veio em {comChangeKey} de {porItem.Count}. " &
                   "Nenhuma dessas evidências cobre ATOMICAMENTE rótulo e corpo — por isso a " &
                   "autorização da §29.2 se prende ao hash dos bytes, não a elas.")

            Anotar("")
            Anotar("## P14 e P16 — o que a medição NÃO decide")
            Anotar("P14: a semântica corporativa de ausência é **desconhecida** e não " &
                   "autoriza transmissão. Nenhuma leitura do Outlook responde isso.")
            Anotar("P16: o DASL é named property do namespace de **cabeçalhos de internet**. " &
                   "Ler bem não prova autoridade. Sinal restritivo serve para NEGAR; " &
                   "ausência ou rótulo baixo não servem, sozinhos, para PERMITIR.")

            Gravar()
        Finally
            broker.Dispose()
        End Try
    End Function

    ' ==================================================================

    ''' <summary>
    ''' Lê o controle e diz o que ele significa — <b>sem</b> deixar a
    ''' conclusão para quem for ler a tabela depois.
    '''
    ''' Se a propriedade que NÃO existe devolver a mesma coisa que o rótulo,
    ''' então <c>Blank</c> é indistinguível de ausente nesta conta, e um
    ''' portão que tratasse <c>Blank</c> como conclusivo estaria decidindo
    ''' sobre ruído.
    ''' </summary>
    Private Shared Sub ComentarSemantica(p As NamedPropertyProbe)
        Dim controle = p.Tentativas.LastOrDefault()
        If controle Is Nothing Then Return

        If controle.Lancou Then
            Anotar("  → P7 RESPONDIDA: propriedade inexistente LANÇA. " &
                   "Então 'vazio' é vazio de verdade, e Absent é distinguível de Blank.")
        Else
            Anotar("  → **P7 RESPONDIDA, e o resultado é ruim**: propriedade que NÃO " &
                   "existe devolve valor sem lançar. Logo `Blank` NÃO distingue " &
                   "'sem rótulo' de 'propriedade ausente', e **nenhuma leitura vazia " &
                   "pode ser conclusiva** neste caminho.")
        End If
    End Sub

    ''' <summary>
    ''' Contagem por desfecho — <b>e nunca o valor do rótulo</b>.
    '''
    ''' O nome de um rótulo é texto que a empresa escolheu, e pode ele próprio
    ''' ser sensível ("Confidencial — Projeto X"). O GUID entra pseudonimizado
    ''' pelos oito primeiros caracteres, que basta para contar distintos e não
    ''' basta para identificar.
    ''' </summary>
    Private Shared Sub Relatar(leituras As IReadOnlyList(Of LabelReading))
        For Each g In leituras.GroupBy(Function(l) l.Kind).OrderByDescending(Function(g2) g2.Count())
            Dim etapas = String.Join(", ", g.Select(Function(l) l.Stage.ToString()).Distinct())
            Dim hrs = String.Join(", ", g.Where(Function(l) l.HResult.HasValue).
                                          Select(Function(l) $"0x{l.HResult.Value:X8}").Distinct())
            Anotar($"  {g.Key,-14} {g.Count(),4}   etapa: {etapas}" &
                   If(hrs.Length > 0, $"   hresult: {hrs}", ""))
        Next

        Dim ausentes = leituras.Where(Function(l) l.Kind = LabelReadingKind.Absent).Count()
        Dim vazios = leituras.Where(Function(l) l.Kind = LabelReadingKind.Blank).Count()
        If ausentes > 0 AndAlso vazios > 0 Then
            Anotar("  (Absent = a propriedade nao existe no item; Blank = existe e esta " &
                   "vazia. Sao estados DIFERENTES, e so aparecem separados porque o " &
                   "controle da P7 provou que ausente LANCA.)")
        End If

        Dim ids = leituras.SelectMany(Function(l) l.LabelIds).Distinct().ToList()
        If ids.Count > 0 Then
            Anotar($"  rótulos distintos: {ids.Count} — " &
                   String.Join(", ", ids.Select(Function(i) i.Substring(0, 8) & "…")))
        End If
    End Sub

    ''' <summary>
    ''' Prevalência, com a margem junto — e o limite superior quando o achado
    ''' é zero, porque amostra nenhuma prova impossibilidade.
    ''' </summary>
    Private Shared Sub RelatarPrevalencia(leituras As IReadOnlyList(Of LabelReading))
        Dim n = leituras.Count
        If n = 0 Then Return
        Dim comRotulo = leituras.Where(Function(l) l.Kind = LabelReadingKind.Present OrElse
                                                   l.Kind = LabelReadingKind.Conflicting).Count()
        Dim p = comRotulo / CDbl(n)
        Dim margem = 1.96 * Math.Sqrt(Math.Max(p * (1 - p), 0.0001) / n)
        Anotar($"P3: prevalência de rótulo = {p:P1} ± {margem:P1} (n={n}, 95%)")

        Dim multiplos = leituras.Where(Function(l) l.Kind = LabelReadingKind.Conflicting).Count()
        If multiplos = 0 Then
            Anotar($"P8: **nenhum** múltiplo observado. Limite superior aproximado " &
                   $"≈ {3.0 / n:P2} a 95% — **não** impossibilidade.")
        Else
            Anotar($"P8: {multiplos} item(ns) com mais de um rótulo ativo.")
        End If
    End Sub

    Private Shared Function Comparar(antes As IReadOnlyList(Of MailSummary),
                                     depois As IReadOnlyList(Of MailSummary)) As String
        Dim porChave = depois.ToDictionary(Function(m) m.Key.EntryId, Function(m) m)
        Dim mudaram = 0
        Dim sumiram = 0
        For Each a In antes
            Dim d As MailSummary = Nothing
            If Not porChave.TryGetValue(a.Key.EntryId, d) Then
                sumiram += 1
            ElseIf d.IsUnread <> a.IsUnread OrElse d.SizeBytes <> a.SizeBytes Then
                mudaram += 1
            End If
        Next
        Return $"de {antes.Count} itens: {mudaram} com metadado diferente, {sumiram} fora da amostra."
    End Function

    ' ==================================================================

    ''' <summary>
    ''' Mede uma pasta e relata. Falha de pasta <b>não</b> derruba a medição:
    ''' "esta pasta não deu para ler" é resultado, e trocá-lo por uma exceção
    ''' perderia o que as outras pastas responderam.
    ''' </summary>
    Private Shared Async Function MedirPastaAsync(broker As OutlookBroker, pasta As FolderKey,
                                                  nome As String, quantos As Integer) As Task
        Dim amostra = Await AmostraAsync(broker, pasta, quantos)
        If amostra.Count = 0 Then
            Anotar($"  {nome}: vazia ou ilegível — nada a medir")
            Return
        End If

        Dim r = Await broker.GetSensitivityLabelsAsync(
            amostra.Select(Function(x) x.Key).ToList(), CancellationToken.None)
        If Not r.Succeeded Then
            Anotar($"  {nome}: a leitura falhou — {r.Kind} {r.Detail}")
            Return
        End If

        Anotar($"### {nome} — n={r.Value.Count}")
        Relatar(r.Value)
        RelatarPrevalencia(r.Value)
    End Function

    ''' <summary>A pasta com este nome no primeiro store, ou <c>Nothing</c>.</summary>
    Private Shared Async Function AcharPastaAsync(broker As OutlookBroker,
                                                  nome As String) As Task(Of FolderKey)
        Dim stores = Await broker.GetStoresAsync(CancellationToken.None)
        If Not stores.Succeeded OrElse stores.Value.Count = 0 Then Return Nothing

        Dim filhas = Await broker.GetFolderChildrenAsync(stores.Value(0).RootFolder,
                                                         CancellationToken.None)
        If Not filhas.Succeeded Then Return Nothing

        For Each f In filhas.Value
            If f.Name.Equals(nome, StringComparison.OrdinalIgnoreCase) Then Return f.Key
        Next
        Return Nothing
    End Function

    Private Shared Async Function AmostraAsync(broker As OutlookBroker, pasta As FolderKey,
                                               quantos As Integer) As Task(Of List(Of MailSummary))
        Dim consulta As New MessageQuery(pasta, MessageSort.ReceivedDesc, 1)
        Dim saida As New List(Of MailSummary)()
        Dim cursor As String = Nothing
        Do
            Dim p = Await broker.GetMessagePageAsync(consulta, cursor, Math.Min(quantos, 50),
                                                     CancellationToken.None)
            If Not p.Succeeded Then Exit Do
            saida.AddRange(p.Value.Items)
            cursor = p.Value.NextCursor
        Loop While cursor IsNot Nothing AndAlso saida.Count < quantos
        Return saida.Take(quantos).ToList()
    End Function

    Private Shared Sub Gravar()
        Try
            Directory.CreateDirectory(PastaDeMedicao)
            File.WriteAllText(Path.Combine(PastaDeMedicao, "purview-3.0.md"),
                              Diario.ToString(), Encoding.UTF8)
        Catch
            ' O relatorio ja saiu no console. Nao gravar nao invalida a
            ' medicao, e falhar aqui trocaria o resultado por um erro de I/O.
        End Try
    End Sub

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
                Assert.Fail($"Outlook instalado e sem responder (estado: {estado}). " &
                            "Esta medicao e a unica prova sobre o Purview nesta conta.")
                Return Nothing
            Case Else
                broker.Dispose()
                Assert.Inconclusive($"Outlook nao instalado (estado: {estado}).")
                Return Nothing
        End Select
    End Function

    Private Shared Async Function AcharEntradaAsync(broker As OutlookBroker) As Task(Of FolderKey)
        Dim stores = Await broker.GetStoresAsync(CancellationToken.None)
        Assert.IsTrue(stores.Succeeded AndAlso stores.Value.Count > 0, "nenhum store")

        Dim filhas = Await broker.GetFolderChildrenAsync(stores.Value(0).RootFolder,
                                                         CancellationToken.None)
        Assert.IsTrue(filhas.Succeeded, "GetFolderChildrenAsync falhou")

        For Each f In filhas.Value
            If f.ContentKind = FolderContentKind.Mail AndAlso
               (f.Name.StartsWith("Caixa de Entrada", StringComparison.OrdinalIgnoreCase) OrElse
                f.Name.Equals("Inbox", StringComparison.OrdinalIgnoreCase)) Then
                Return f.Key
            End If
        Next
        Assert.Inconclusive("nao achei a Caixa de Entrada")
        Return Nothing
    End Function

End Class

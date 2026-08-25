Imports System.Linq
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>O parser do <c>MSIP_Labels</c>, nos casos que esta caixa não tem.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ISTO PRECISA EXISTIR SEPARADO DA MEDIÇÃO</b>
'''
''' A medição do 3.0 achou 10 itens rotulados em 554, com três GUIDs, e
''' <b>nenhum</b> caso difícil: nenhum múltiplo, nenhum conflito, nenhum
''' histórico, nenhum valor corrompido. Ou seja, a caixa real exercita
''' exatamente o caminho fácil.
'''
''' Um parser cuja única prova é a caixa real é um parser cujos ramos
''' perigosos <b>nunca rodam</b>. E foram justamente esses os dois que o
''' Codex derrubou na primeira versão:
'''
'''   • valor <b>meio corrompido</b> — um GUID bom e um inválido — saía
'''     <c>Present</c>, porque pares não reconhecidos eram ignorados em
'''     silêncio;
'''   • o <b>mesmo</b> GUID com <c>Enabled=True</c> e <c>Enabled=False</c>
'''     saía <c>Present</c>, porque o conjunto ordenado terminava com um
'''     ativo só. O comentário prometia detectar conflito; o código detectava
'''     apenas <i>mais de um GUID</i>.
'''
''' Os dois liberariam conteúdo de um item cuja classificação eu não sei ler.
''' </summary>
<TestClass>
Public Class RotuloParserTests

    Private Const G1 As String = "baea3331-0000-4000-8000-000000000001"
    Private Const G2 As String = "0d35eb3e-0000-4000-8000-000000000002"

    Private Shared Function Registro(guid As String, campo As String, valor As String) As String
        Return $"MSIP_Label_{guid}_{campo}={valor}"
    End Function

    Private Shared Function Ler(texto As String) _
                                As (Tipo As LabelReadingKind,
                                    Ativos As IReadOnlyList(Of String),
                                    Campos As IReadOnlyList(Of String))
        Return SensitivityLabels.Analisar(texto)
    End Function

    ' ==================================================================
    ' Controle positivo — sem isto, um parser que devolvesse Malformed
    ' sempre passaria em todos os testes de "não conclui indevidamente".

    <TestMethod>
    Public Sub Um_rotulo_ativo_e_Present()
        Dim r = Ler(Registro(G1, "Enabled", "True") & ";" &
                    Registro(G1, "SetDate", "2026-01-01T00:00:00Z"))

        Assert.AreEqual(LabelReadingKind.Present, r.Tipo)
        CollectionAssert.AreEqual({G1}, r.Ativos.ToArray())
    End Sub

    ' ==================================================================

    ''' <summary>Dois rótulos ativos ao mesmo tempo é conflito.</summary>
    <TestMethod>
    Public Sub Dois_ativos_e_Conflicting()
        Dim r = Ler(Registro(G1, "Enabled", "True") & ";" & Registro(G2, "Enabled", "True"))

        Assert.AreEqual(LabelReadingKind.Conflicting, r.Tipo)
        Assert.AreEqual(2, r.Ativos.Count)
    End Sub

    ''' <summary>
    ''' <b>O MESMO GUID dizendo as duas coisas</b> — o caso que passava como
    ''' <c>Present</c>.
    '''
    ''' O conjunto de ativos tem um elemento só, então uma contagem
    ''' <c>&gt; 1</c> não pega. Só pega quem comparar ativos com inativos.
    ''' </summary>
    <TestMethod>
    Public Sub O_MESMO_guid_ativo_e_inativo_e_Conflicting()
        Dim r = Ler(Registro(G1, "Enabled", "True") & ";" & Registro(G1, "Enabled", "False"))

        Assert.AreEqual(LabelReadingKind.Conflicting, r.Tipo,
            "um GUID so nos ativos nao pode virar Present quando ele TAMBEM " &
            "aparece desligado — o valor diz duas coisas")
    End Sub

    ''' <summary>
    ''' <b>Meio corrompido</b> — o outro caso que passava.
    '''
    ''' Um GUID bom e um inválido. A versão anterior ignorava o inválido e
    ''' devolvia <c>Present</c> para o bom, como se o valor estivesse inteiro.
    ''' </summary>
    <TestMethod>
    Public Sub Um_guid_bom_e_um_INVALIDO_contamina_tudo()
        Dim r = Ler(Registro(G1, "Enabled", "True") & ";" &
                    "MSIP_Label_nao-sou-um-guid_Enabled=True")

        Assert.AreEqual(LabelReadingKind.Malformed, r.Tipo,
            "valor que eu nao entendo INTEIRO e o caso em que eu NAO posso decidir")
        Assert.AreEqual(0, r.Ativos.Count, "e nao pode sair rotulo de um valor contaminado")
    End Sub

    ''' <summary><c>Enabled</c> com valor que não é nem True nem False.</summary>
    <TestMethod>
    Public Sub Enabled_com_valor_estranho_contamina()
        Assert.AreEqual(LabelReadingKind.Malformed,
                        Ler(Registro(G1, "Enabled", "Talvez")).Tipo)
    End Sub

    ''' <summary>Chave de rótulo sem o campo depois do GUID.</summary>
    <TestMethod>
    Public Sub Chave_de_rotulo_truncada_contamina()
        Assert.AreEqual(LabelReadingKind.Malformed, Ler("MSIP_Label_=True").Tipo)
    End Sub

    ''' <summary>Registro que se diz de rótulo e não tem nem <c>=</c>.</summary>
    <TestMethod>
    Public Sub Registro_de_rotulo_sem_igual_contamina()
        Assert.AreEqual(LabelReadingKind.Malformed,
                        Ler(Registro(G1, "Enabled", "True") & ";MSIP_Label_lixo").Tipo)
    End Sub

    ''' <summary>
    ''' Rótulo <b>removido</b>: forma boa, histórico presente, nenhum ativo.
    '''
    ''' Não é <c>Malformed</c> nem <c>Absent</c>. É uma informação específica —
    ''' "houve rótulo aqui" — que a política corporativa pode querer usar.
    ''' </summary>
    <TestMethod>
    Public Sub So_historico_e_HistoricalOnly()
        Dim r = Ler(Registro(G1, "Enabled", "False") & ";" &
                    Registro(G1, "SetDate", "2026-01-01T00:00:00Z"))

        Assert.AreEqual(LabelReadingKind.HistoricalOnly, r.Tipo)
        Assert.AreEqual(0, r.Ativos.Count)
    End Sub

    ''' <summary>Registros de rótulo sem nenhum campo <c>Enabled</c>.</summary>
    <TestMethod>
    Public Sub Rotulo_sem_campo_Enabled_e_Malformed()
        Assert.AreEqual(LabelReadingKind.Malformed,
                        Ler(Registro(G1, "SetDate", "2026-01-01T00:00:00Z")).Tipo)
    End Sub

    ''' <summary>Texto que não menciona rótulo nenhum.</summary>
    <TestMethod>
    Public Sub Texto_sem_registro_de_rotulo_e_Malformed()
        Assert.AreEqual(LabelReadingKind.Malformed, Ler("qualquer=coisa;outra=coisa").Tipo)
    End Sub

    ''' <summary>
    ''' Lixo que <b>não menciona</b> <c>MSIP_Label</c> junto de um rótulo bom
    ''' não contamina.
    '''
    ''' O contraponto: se qualquer fragmento estranho contaminasse, um
    ''' cabeçalho colado no fim do valor derrubaria a leitura de todo item
    ''' legitimamente rotulado — e o portão negaria por ruído, não por
    ''' classificação.
    ''' </summary>
    <TestMethod>
    Public Sub Lixo_que_nao_se_diz_rotulo_NAO_contamina()
        Dim r = Ler(Registro(G1, "Enabled", "True") & ";;   ;outra-coisa-qualquer")

        Assert.AreEqual(LabelReadingKind.Present, r.Tipo)
        CollectionAssert.AreEqual({G1}, r.Ativos.ToArray())
    End Sub

    ''' <summary>
    ''' Os <b>nomes de campo</b> saem, e os valores não.
    '''
    ''' É o que responde a P9 — "proteção vem separada da classificação?" —
    ''' estruturalmente, sem que texto de rótulo apareça em relatório nenhum.
    ''' </summary>
    <TestMethod>
    Public Sub Os_nomes_de_campo_saem_e_os_valores_nao()
        Dim r = Ler(Registro(G1, "Enabled", "True") & ";" &
                    Registro(G1, "SetDate", "2026-01-01T00:00:00Z") & ";" &
                    Registro(G1, "ContentBits", "2"))

        CollectionAssert.AreEquivalent({"Enabled", "SetDate", "ContentBits"},
                                       r.Campos.ToArray())
    End Sub

    ''' <summary>Maiúsculas no nome do campo não mudam a leitura.</summary>
    <TestMethod>
    Public Sub A_caixa_do_campo_nao_importa()
        Assert.AreEqual(LabelReadingKind.Present,
                        Ler(Registro(G1, "enabled", "TRUE")).Tipo)
    End Sub

End Class

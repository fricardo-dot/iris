Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions
Imports Iris.Sync
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Q8 — a matriz de providers, e o que fazer fora dela.
'''
''' A matriz está incompleta e continua incompleta depois destes testes: a
''' §19.3 mediu que levantar as outras linhas custa horas e dezenas de GB
''' nesta máquina. O que estes testes cobram é a outra metade da resposta —
''' que ambiente não demonstrado DEGRADE, em vez de ser tratado como
''' "provavelmente igual".
'''
''' O ponto que estes testes aprenderam depois: <b>reconhecer não é
''' autorizar</b>. A primeira versão desta política reconhecia o ambiente do
''' usuário e daí concedia as três inferências, o que transformava "medi
''' quantos itens o OOM alcança" em "medi que ausência, cobertura e
''' incremental funcionam". Nenhuma das três estava medida.
''' </summary>
<TestClass>
Public Class EnvironmentPolicyTests

    ''' <summary>
    ''' O token da janela medido em 2026-08-24. Não é "1 mês" nem número nenhum
    ''' de meses: é o valor CRU do perfil. Ver
    ''' <see cref="EnvironmentFingerprint.WindowToken"/>.
    ''' </summary>
    Private Const TokenMedido As String = "84-09-00-00"

    Private Shared Function Medido() As EnvironmentFingerprint
        Return New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, TokenMedido)
    End Function

    ' ==================================================================
    ' O estado de HOJE: reconhecido, nada autorizado

    <TestMethod>
    Public Sub O_ambiente_do_usuario_esta_reconhecido()
        Dim linha = EnvironmentPolicy.Medido(Medido())
        Assert.IsNotNull(linha, "o ambiente medido sumiu da matriz")
        StringAssert.Contains(linha.Evidence, "§22.3")
    End Sub

    ''' <summary>
    ''' E não autoriza nada — e isso é o estado honesto, não uma pendência
    ''' escondida.
    '''
    ''' Medir que o OOM alcança 1.979 mensagens diz quanto se enxergou. Não diz
    ''' que "não encontrei" significa excluído, nem que uma varredura teve
    ''' cobertura completa, nem que o incremental não perde itens. Três
    ''' inferências, três demonstrações — e nenhuma foi feita.
    ''' </summary>
    <TestMethod>
    Public Sub Reconhecer_nao_e_autorizar()
        Dim c = EnvironmentPolicy.Capacidades(Medido())
        Assert.IsFalse(c.PodeConcluirAusencia,
            "alcance medido nao demonstra que 'nao encontrei' significa excluido")
        Assert.IsFalse(c.PodeAfirmarCoberturaCompleta,
            "a §19.2 mediu pastas cheias reportando zero — cobertura nao vem de contagem")
        Assert.IsFalse(c.PodeUsarIncremental)
        Assert.IsTrue(c.Degradado)
    End Sub

    ''' <summary>
    ''' O motivo da negação nomeia a lacuna, em vez de dizer só "não pode".
    ''' Sem isso, quem for destravar vai destravar no lugar errado.
    ''' </summary>
    <TestMethod>
    Public Sub A_negacao_nomeia_a_lacuna()
        StringAssert.Contains(EnvironmentPolicy.Capacidades(Medido()).Reason, "sensibilidade")
    End Sub

    ' ==================================================================
    ' CONTROLE POSITIVO

    ''' <summary>
    ''' O controle positivo, e sem ele nada aqui prova coisa alguma.
    '''
    ''' Com a matriz de produção autorizando zero, uma política que
    ''' simplesmente negasse tudo — um <c>Return Negar(...)</c> na primeira
    ''' linha — passaria em todos os outros testes desta classe. Este exercita
    ''' o caminho de AUTORIZAÇÃO, com uma matriz injetada.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_positivo_uma_linha_demonstrada_AUTORIZA()
        Dim c = EnvironmentPolicy.Capacidades(Medido(), MatrizSintetica(
            Inference.ConcluirAusencia, Inference.AfirmarCoberturaCompleta,
            Inference.UsarIncremental))

        Assert.IsTrue(c.PodeConcluirAusencia, "a politica nao esta decidindo nada — nega sempre")
        Assert.IsTrue(c.PodeAfirmarCoberturaCompleta)
        Assert.IsTrue(c.PodeUsarIncremental)
        Assert.IsFalse(c.Degradado)
    End Sub

    ''' <summary>
    ''' Autorização é POR INFERÊNCIA: demonstrar uma não libera as outras.
    ''' </summary>
    <TestMethod>
    Public Sub Demonstrar_uma_inferencia_nao_libera_as_outras()
        Dim c = EnvironmentPolicy.Capacidades(Medido(),
            MatrizSintetica(Inference.UsarIncremental))

        Assert.IsTrue(c.PodeUsarIncremental)
        Assert.IsFalse(c.PodeConcluirAusencia, "cada inferencia exige demonstracao propria")
        Assert.IsFalse(c.PodeAfirmarCoberturaCompleta)
        Assert.IsTrue(c.Degradado)
        StringAssert.Contains(c.Reason, "parcial")
    End Sub

    ''' <summary>
    ''' Token não validado bloqueia MESMO com inferências demonstradas.
    '''
    ''' Se o token não muda quando a janela muda, a impressão digital não
    ''' distingue os dois universos — e a autorização concedida ao antigo
    ''' continua valendo no novo, sem ninguém notar.
    ''' </summary>
    <TestMethod>
    Public Sub Token_nao_validado_bloqueia_mesmo_com_evidencia()
        Dim linha = New MeasuredEnvironment(Medido(), "FASE2 §22.3 — teste",
            New Date(2024, 10, 9), tokenValidado:=False,
            grants:={New GrantedInference(Inference.ConcluirAusencia, "FASE2 §22.3 — teste")})

        Dim c = EnvironmentPolicy.Capacidades(Medido(), {linha})
        Assert.IsFalse(c.PodeConcluirAusencia)
        StringAssert.Contains(c.Reason, "sensibilidade")
    End Sub

    ''' <summary>Inferência autorizada sem evidência não se constrói.</summary>
    <TestMethod>
    Public Sub Inferencia_autorizada_exige_evidencia()
        Assert.ThrowsException(Of ArgumentException)(
            Sub()
                Dim x = New GrantedInference(Inference.ConcluirAusencia, "  ")
            End Sub)
    End Sub

    ''' <summary>
    ''' Inferência autorizada DUAS VEZES não se constrói, e o motivo não é
    ''' estética.
    '''
    ''' A autorização é deduplicada por conjunto, mas a razão da degradação era
    ''' calculada da contagem da LISTA. Três grants iguais de
    ''' <c>UsarIncremental</c> davam uma permissão só, contagem três, e daí
    ''' <c>Reason = ""</c> com <c>Degradado = True</c>: degradado sem motivo
    ''' escrito. É o pior estado possível — errado e silencioso, porque quem
    ''' fosse investigar não teria o que ler.
    ''' </summary>
    <TestMethod>
    Public Sub Inferencia_autorizada_duas_vezes_nao_se_constroi()
        Assert.ThrowsException(Of ArgumentException)(
            Sub()
                Dim x = New MeasuredEnvironment(Medido(), "FASE2 §22.3 — teste",
                    Nothing, tokenValidado:=True,
                    grants:={New GrantedInference(Inference.UsarIncremental, "FASE2 §22.3 — a"),
                             New GrantedInference(Inference.UsarIncremental, "FASE2 §22.3 — b")})
            End Sub)
    End Sub

    ''' <summary>
    ''' E valor fora do enum também não. <c>CType(999, Inference)</c> compila
    ''' em VB sem reclamar — o enum não é um conjunto fechado em tempo de
    ''' execução, e um valor inválido passaria direto para o <c>HashSet</c>,
    ''' virando uma permissão que nenhuma checagem consegue negar.
    ''' </summary>
    <TestMethod>
    Public Sub Valor_fora_do_enum_nao_se_constroi()
        Assert.ThrowsException(Of ArgumentException)(
            Sub()
                Dim x = New MeasuredEnvironment(Medido(), "FASE2 §22.3 — teste",
                    Nothing, tokenValidado:=True,
                    grants:={New GrantedInference(CType(999, Inference), "FASE2 §22.3 — teste")})
            End Sub)
    End Sub

    ''' <summary>
    ''' A razão da degradação nunca fica vazia enquanto <c>Degradado</c> é
    ''' verdadeiro. É o invariante que o defeito das duplicatas violava.
    '''
    ''' Cobre <b>todos</b> os retornos de <c>Capacidades</c>, um por um, e não
    ''' uma amostra que eu escolhi. Um invariante transversal testado sobre
    ''' quatro casos escolhidos a dedo continua verdadeiro se eu acrescentar um
    ''' quinto degrau e esquecer o texto — e é exatamente aí que ele quebraria.
    ''' </summary>
    <TestMethod>
    Public Sub Degradado_sempre_vem_com_motivo_escrito()
        Dim casos As New List(Of (Nome As String, Cap As EnvironmentCapabilities))() From {
            ("fp nulo", EnvironmentPolicy.Capacidades(Nothing)),
            ("provider desconhecido", EnvironmentPolicy.Capacidades(
                New EnvironmentFingerprint(ProviderKind.Desconhecido, True, TokenMedido))),
            ("cached sem janela", EnvironmentPolicy.Capacidades(
                New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))),
            ("fora da matriz", EnvironmentPolicy.Capacidades(
                New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "00-00-00-00"))),
            ("token nao validado", EnvironmentPolicy.Capacidades(Medido())),
            ("grants vazios", EnvironmentPolicy.Capacidades(Medido(), {
                New MeasuredEnvironment(Medido(), "FASE2 §22.3 — t", Nothing, tokenValidado:=True)})),
            ("autorizacao parcial", EnvironmentPolicy.Capacidades(Medido(),
                MatrizSintetica(Inference.UsarIncremental))),
            ("autorizacao completa", EnvironmentPolicy.Capacidades(Medido(),
                MatrizSintetica(Inference.ConcluirAusencia,
                                Inference.AfirmarCoberturaCompleta,
                                Inference.UsarIncremental)))}

        ' Um por degrau de Capacidades, mais o caminho que autoriza tudo.
        Assert.AreEqual(8, casos.Count)

        Dim degradados = 0
        For Each c In casos
            If c.Cap.Degradado Then
                degradados += 1
                Assert.IsFalse(String.IsNullOrWhiteSpace(c.Cap.Reason),
                    $"'{c.Nome}' degradou sem motivo escrito: quem investigar nao tem o que ler")
            Else
                Assert.AreEqual("", c.Cap.Reason,
                    $"'{c.Nome}' nao degradou mas trouxe motivo")
            End If
        Next

        ' Controle: nem todos degradam, senao o laco acima nao verificaria nada.
        Assert.AreEqual(7, degradados)
    End Sub

    ' ==================================================================
    ' Degradação fora da matriz

    <TestMethod>
    Public Sub Ambiente_fora_da_matriz_nao_pode_concluir_ausencia()
        Dim fora = New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "00-00-00-00")
        Dim c = EnvironmentPolicy.Capacidades(fora)
        Assert.IsFalse(c.PodeConcluirAusencia,
            "num ambiente nao medido, 'nao encontrei' nao distingue excluido de fora-da-janela")
        StringAssert.Contains(c.Reason, "fora da matriz")
    End Sub

    <TestMethod>
    Public Sub Provider_desconhecido_degrada()
        Assert.IsTrue(EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.Desconhecido, True, TokenMedido)).Degradado)
    End Sub

    <TestMethod>
    Public Sub Ambiente_nulo_degrada()
        Assert.IsTrue(EnvironmentPolicy.Capacidades(Nothing).Degradado)
    End Sub

    ''' <summary>
    ''' Cached com janela não lida é o caso mais traiçoeiro: sabe-se que há
    ''' janela e não se sabe qual. Pior que não saber nada, porque parece
    ''' identificado.
    ''' </summary>
    <TestMethod>
    Public Sub Cached_com_janela_nao_lida_degrada()
        Dim c = EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
        Assert.IsFalse(c.PodeConcluirAusencia)
        StringAssert.Contains(c.Reason, "janela")
    End Sub

    <TestMethod>
    Public Sub PST_nao_herda_do_Exchange()
        Assert.IsTrue(EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.PstLocal, False, "sem-janela")).Degradado,
            "PST nunca foi medido neste projeto")
    End Sub

    ' ==================================================================
    ' A janela faz parte da identidade do ambiente

    ''' <summary>
    ''' Mesma caixa, mesma conta, janela diferente: ambiente DIFERENTE.
    '''
    ''' É a §18.4 virada em código. A janela muda o que EXISTE, não só o que
    ''' custa: em 2026-08-22 o OOM alcançava 1.004 itens numa caixa de 17.668;
    ''' em 2026-08-24, com a janela maior, alcança 1.979 até 2024-10-09. Mesma
    ''' caixa, mesma conta, dois universos.
    ''' </summary>
    <TestMethod>
    Public Sub Trocar_a_janela_muda_o_ambiente()
        Dim um = New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, TokenMedido)
        Dim outro = New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "FF-FF-00-00")
        Assert.AreNotEqual(um.Value(), outro.Value())
        Assert.IsTrue(EnvironmentPolicy.ExigeReconciliacao(um, outro))
    End Sub

    ''' <summary>
    ''' AUMENTAR a janela também invalida, e este é o caso que a intuição erra.
    '''
    ''' "Encolher esconde, aumentar só revela — aumentar é seguro" é falso:
    ''' aumentar revela itens que já haviam sido concluídos AUSENTES, e essa
    ''' conclusão anterior estava errada. Deixá-la de pé por ser "só um
    ''' aumento" preserva justamente o erro.
    ''' </summary>
    <TestMethod>
    Public Sub Aumentar_a_janela_tambem_exige_reconciliacao()
        Assert.IsTrue(EnvironmentPolicy.ExigeReconciliacao(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, TokenMedido),
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "00-00-00-00")),
            "aumentar revela itens ja concluidos ausentes — a conclusao anterior estava errada")
    End Sub

    <TestMethod>
    Public Sub Ambiente_igual_nao_exige_reconciliacao()
        Assert.IsFalse(EnvironmentPolicy.ExigeReconciliacao(Medido(), Medido()),
            "senao toda abertura reconciliaria a caixa inteira")
    End Sub

    <TestMethod>
    Public Sub Ambiente_desconhecido_de_um_lado_exige_reconciliacao()
        Assert.IsTrue(EnvironmentPolicy.ExigeReconciliacao(Nothing, Medido()))
        Assert.IsTrue(EnvironmentPolicy.ExigeReconciliacao(Medido(), Nothing))
    End Sub

    ' ==================================================================
    ' A evidência aponta para algo que EXISTE

    ''' <summary>
    ''' Toda seção citada pela matriz existe de fato no FASE2.md.
    '''
    ''' O teste anterior exigia só que a evidência contivesse "§" — e passou
    ''' feliz enquanto a linha citava a §22.3, que naquele momento não existia:
    ''' o documento parava na §21. Uma evidência que aponta para o nada é pior
    ''' que evidência nenhuma, porque parece verificável.
    ''' </summary>
    <TestMethod>
    Public Sub Toda_secao_citada_pela_matriz_existe_no_FASE2()
        Dim fase2 = LerFase2()
        Dim citadas As New List(Of String)()

        For Each linha In EnvironmentPolicy.Matriz
            Assert.IsFalse(String.IsNullOrWhiteSpace(linha.Evidence),
                $"linha sem evidencia: {linha.Fingerprint.Value()}")
            citadas.AddRange(Secoes(linha.Evidence))
            For Each g In linha.Grants
                citadas.AddRange(Secoes(g.Evidence))
            Next
        Next

        Assert.IsTrue(citadas.Count > 0, "nenhuma secao citada — o teste nao verificaria nada")

        For Each s In citadas.Distinct()
            Dim numero = s.Substring(1)
            Assert.IsTrue(Regex.IsMatch(fase2, "^#+\s*" & Regex.Escape(numero) & "(\D|$)",
                                        RegexOptions.Multiline),
                $"a matriz cita a {s}, que nao existe no FASE2.md")
        Next
    End Sub

    ' ==================================================================

    Private Shared Function MatrizSintetica(ParamArray quais As Inference()) _
            As IReadOnlyList(Of MeasuredEnvironment)
        Return {New MeasuredEnvironment(Medido(), "FASE2 §22.3 — sintetico",
                    New Date(2024, 10, 9), tokenValidado:=True,
                    grants:=quais.Select(Function(i) New GrantedInference(i, "FASE2 §22.3 — sintetico")))}
    End Function

    Private Shared Function Secoes(evidencia As String) As IEnumerable(Of String)
        Return Regex.Matches(evidencia, "§\d+(\.\d+)?").Cast(Of Match)().
               Select(Function(m) m.Value).ToList()
    End Function

    Private Shared _fase2 As String

    Private Shared Function LerFase2() As String
        If _fase2 IsNot Nothing Then Return _fase2
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "FASE2.md"))
            d = d.Parent
        End While
        Assert.IsNotNull(d, "nao achei o FASE2.md a partir de " & AppContext.BaseDirectory)
        _fase2 = File.ReadAllText(Path.Combine(d.FullName, "FASE2.md"))
        Return _fase2
    End Function

End Class

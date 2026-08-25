Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports Iris.Sync
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' A orquestração contra uma fonte que MUTA DURANTE A LEITURA.
'''
''' Cada cenário aqui aconteceu, ou pode acontecer, na caixa do usuário: item
''' movido no meio da varredura, pasta cheia reportando zero, contagem que não
''' bate, universo trocando, o Outlook morrendo. O <see cref="SweepModel"/> já
''' decide o que é decidível sem tocar no mundo; o que estes testes cobram é
''' que o <see cref="SweepRunner"/> <b>obedeça</b> — e que nunca publique
''' metade.
''' </summary>
<TestClass>
Public Class SweepRunnerTests

    Private Shared Function Universo(Optional filtro As String = "f") As SweepUniverse
        Return New SweepUniverse("store-1", "pasta-1", filtro, Nothing, 1, "amb-1")
    End Function

    ''' <summary>Ambiente que autoriza tudo — o que a produção hoje não tem.</summary>
    Private Shared Function Autorizado() As EnvironmentCapabilities
        Dim fp = New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "hipotetica")
        Return EnvironmentPolicy.Capacidades(fp, {
            New MeasuredEnvironment(fp, "FASE2 §22.3 — sintetico", Nothing, tokenValidado:=True,
                grants:={New GrantedInference(Inference.ConcluirAusencia, "FASE2 §22.3 — s"),
                         New GrantedInference(Inference.AfirmarCoberturaCompleta, "FASE2 §22.3 — s"),
                         New GrantedInference(Inference.UsarIncremental, "FASE2 §22.3 — s")})})
    End Function

    ''' <summary>O ambiente REAL do usuário: cached, janela não legível (§23).</summary>
    Private Shared Function Real() As EnvironmentCapabilities
        Return EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))
    End Function

    Private Shared Function Rodar(f As FonteFalsaMutavel, d As DestinoFalso,
                                  Optional tamanho As Integer = 2,
                                  Optional ct As CancellationToken = Nothing,
                                  Optional cap As EnvironmentCapabilities = Nothing) As SweepResult
        Dim r As New SweepRunner(f, d, tamanho)
        Return r.Executar(Universo(), 0, 1, If(cap, Autorizado()), ct)
    End Function

    ' ==================================================================
    ' Controle positivo

    <TestMethod>
    Public Sub Controle_positivo_varredura_limpa_PUBLICA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d", "e")
        Dim d As New DestinoFalso()

        Dim r = Rodar(f, d)

        Assert.IsTrue(r.Publicou, $"deveria publicar. motivo: {r.Motivo}")
        Assert.AreEqual(1, d.Publicadas)
        CollectionAssert.AreEquivalent({"a", "b", "c", "d", "e"}, d.ChavesGravadas)
        Assert.AreEqual(5, r.Attempt.RowsRead)
        Assert.AreEqual(FolderCoverage.Completa, d.CoberturaPublicada)
    End Sub

    ' ==================================================================
    ' §23 — o ambiente limitado OPERA, e publica parcial

    ''' <summary>
    ''' <b>O teste que substituiu o errado.</b>
    '''
    ''' A primeira versão afirmava que ambiente não autorizado NEM ABRE, e
    ''' isso contradizia a §23: na caixa do usuário — cached, janela não
    ''' legível — o Iris não varreria nada e o produto não faria nada.
    '''
    ''' A §23 bloqueou três <em>inferências</em>, não a operação. Então varre,
    ''' encena, publica — como <b>parcial</b>.
    ''' </summary>
    <TestMethod>
    Public Sub Cached_sem_janela_VARRE_e_publica_PARCIAL()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()

        Dim r = Rodar(f, d, cap:=Real())

        Assert.IsTrue(r.Publicou, $"o ambiente do usuario TEM de varrer. motivo: {r.Motivo}")
        Assert.IsTrue(f.PaginasLidas > 0, "a fonte tem de ter sido lida")
        Assert.AreEqual(4, d.ChavesGravadas.Count, "as linhas tem de ter sido encenadas")
        Assert.AreEqual(FolderCoverage.Parcial, d.CoberturaPublicada,
            "sem autorizacao para afirmar cobertura completa, a geracao sai PARCIAL")
        Assert.AreEqual(FolderCoverage.Parcial, r.Cobertura)
    End Sub

    ''' <summary>
    ''' E marcar não-vistos como suspeitos sai SEMPRE, inclusive parcial.
    '''
    ''' <c>Suspeito</c> não é conclusão negativa: é exatamente "não vi e não
    ''' posso concluir por quê". Condicioná-lo à cobertura completa bloquearia
    ''' a resposta honesta e deixaria o item como <c>Presente</c> — afirmação
    ''' mais forte do que a evidência sustenta.
    ''' </summary>
    <TestMethod>
    Public Sub Suspeito_e_emitido_mesmo_com_cobertura_parcial()
        Dim a = SweepModel.Abrir(Universo(), 0, 1, True).State
        a = SweepModel.ContagemInicial(a, 1, Universo()).State
        a = SweepModel.Pagina(a, {"x"}, "1", Universo()).State
        a = SweepModel.ContagemFinal(a, 1, Universo()).State

        Dim r = SweepModel.Publicar(a, 0, podeAfirmarCoberturaCompleta:=False)

        Assert.IsFalse(r.Rejected)
        Assert.AreEqual(FolderCoverage.Parcial, r.Cobertura)
        Assert.IsTrue(r.Commands.Contains(SweepCommand.MarcarNaoVistosComoSuspeitos),
            "suspeita e a resposta honesta, nao uma conclusao negativa")
    End Sub

    ''' <summary>Ambiente NÃO IDENTIFICADO continua não abrindo — este é o gate que resta.</summary>
    <TestMethod>
    Public Sub Ambiente_nao_identificado_NEM_ABRE()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b")
        Dim d As New DestinoFalso()

        ' Sem o helper: ele tem default, e default nao distingue "omitido" de
        ' "explicitamente nulo" — foi assim que este teste passou verde
        ' afirmando o contrario do que acontecia.
        Dim runner As New SweepRunner(f, d, 2)
        Dim r = runner.Executar(Universo(), 0, 1, Nothing, CancellationToken.None)

        Assert.AreEqual(SweepConclusion.Rejeitada, r.Conclusion)
        Assert.IsNull(r.Attempt)
        Assert.AreEqual(0, f.PaginasLidas, "nao pode nem ter lido a fonte")
        Assert.AreEqual(0, d.Abertas, "nao pode nem ter aberto tentativa no destino")
    End Sub

    ' ==================================================================
    ' Mutação durante a leitura

    <TestMethod>
    Public Sub Item_removido_no_meio_REJEITA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.Agenda(1) = Sub(x) x.Remover("d")

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou)
        Assert.AreEqual(0, d.Publicadas, "nao pode publicar meia varredura")
        StringAssert.Contains(r.Motivo, "percorridas 3 (lidas 3 + descartadas 0) <> antes 4")
        Assert.AreEqual(1, d.Descartadas.Count, "a tentativa tem de ser descartada no destino")
    End Sub

    <TestMethod>
    Public Sub Item_acrescentado_no_meio_REJEITA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.Agenda(1) = Sub(x) x.Acrescentar("novo")

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou)
        Assert.AreEqual(0, d.Publicadas)
    End Sub

    ''' <summary>
    ''' <b>Descarte DECLARADO fecha a conta do S6.</b>
    '''
    ''' Medido na Caixa de Entrada do usuário: 1.022 itens declarados, 1.013
    ''' lidos, 9 descartados por não serem mensagem — <b>estável nas três
    ''' execuções</b>. Comparando só o que foi guardado, o S6 rejeitava a caixa
    ''' principal TODAS as vezes, e o sintoma — "lidas 1013 &lt;&gt; antes
    ''' 1022" — era indistinguível de uma mensagem chegando no meio. Defeito
    ''' permanente disfarçado do comportamento correto.
    ''' </summary>
    <TestMethod>
    Public Sub Descarte_declarado_fecha_a_conta_do_S6()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d") With {.DescartarPorPagina = 1}
        Dim d As New DestinoFalso()
        ' A fonte diz que a pasta tem 6: 4 que ela devolve + 2 que ela descarta
        ' (1 por pagina, em 2 paginas).
        f.ContagemDeclarada = 6

        Dim r = Rodar(f, d)

        Assert.IsTrue(r.Publicou, $"descarte declarado tem de fechar a conta. motivo: {r.Motivo}")
        Assert.AreEqual(4, r.Attempt.RowsRead)
        Assert.AreEqual(2, r.Attempt.Discarded)
        Assert.AreEqual(6, r.Attempt.RowsTraversed)
    End Sub

    ''' <summary>
    ''' E descarte NÃO declarado continua rejeitando — senão a correção teria
    ''' virado um buraco: bastaria a fonte omitir linhas em silêncio.
    ''' </summary>
    <TestMethod>
    Public Sub Descarte_NAO_declarado_continua_REJEITANDO()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.ContagemDeclarada = 6   ' diz 6, entrega 4, nao declara descarte

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou,
            "sem declarar o descarte, a diferenca continua sendo linha que sumiu")
        StringAssert.Contains(r.Motivo, "descartadas 0")
    End Sub

    ''' <summary>
    ''' MUTAÇÃO BALANCEADA — e este teste documenta um BURACO, não uma defesa.
    '''
    ''' A primeira versão deste teste <b>não provava o que afirmava</b>: ela
    ''' trocava um item ainda não lido, e o manifesto resultante coincidia com
    ''' o estado da fonte depois da mutação. Ou seja, poderia ter existido.
    '''
    ''' Para produzir um manifesto <b>impossível</b> é preciso remover algo
    ''' <b>já lido</b>:
    '''
    ''' <code>
    '''   inicial: a,b,c,d    pagina 1: a,b    remove "a", acrescenta "z"
    '''   resto por offset: d,z            manifesto: a,b,d,z
    ''' </code>
    '''
    ''' Esse conjunto não existiu antes nem depois — tem o "a" que foi
    ''' removido, tem o "z" que só chegou depois, e perdeu o "c" pelo
    ''' deslocamento do offset. As contagens fecham, o S6 aprova, e a
    ''' varredura publica um retrato de um mundo que nunca houve.
    '''
    ''' O S6 garante que as <b>contagens</b> fecham, que é bem menos que
    ''' integridade do conjunto. Está como teste executável, e não como
    ''' comentário, para que ninguém olhe o S6 e ache que ele garante mais.
    ''' </summary>
    <TestMethod>
    Public Sub Mutacao_balanceada_produz_manifesto_IMPOSSIVEL_e_o_S6_aprova()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.Agenda(1) = Sub(x)
                          x.Remover("a")      ' ja foi lido
                          x.Acrescentar("z")  ' nao existia
                      End Sub

        Dim r = Rodar(f, d)

        Assert.IsTrue(r.Publicou, "o S6 compara CONTAGENS, e elas fecham")

        Dim manifesto = d.ChavesGravadas
        Assert.IsTrue(manifesto.Contains("a"),
            "'a' esta no manifesto e ja nao existia quando a varredura terminou")
        Assert.IsTrue(manifesto.Contains("z"),
            "'z' esta no manifesto e nao existia quando ela comecou")
        Assert.IsFalse(manifesto.Contains("c"),
            "'c' existia o tempo todo e ficou de fora, pelo deslocamento do offset")

        Dim agora = f.Estado
        Assert.IsFalse(manifesto.OrderBy(Function(x) x).SequenceEqual(agora.OrderBy(Function(x) x)),
            "o manifesto nao corresponde ao estado final da fonte")
    End Sub

    <TestMethod>
    Public Sub Universo_trocado_no_meio_DESCARTA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.Agenda(1) = Sub(x) x.TrocarUniverso(
            New SweepUniverse("store-1", "pasta-1", "OUTRO FILTRO", Nothing, 1, "amb-1"))

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou)
        StringAssert.Contains(r.Motivo.ToLowerInvariant(), "universo")
    End Sub

    ''' <summary>
    ''' Universo AUSENTE é rejeição, não dispensa da guarda.
    '''
    ''' Antes a comparação era condicional ao universo não ser nulo, então uma
    ''' fonte que devolvesse <c>Nothing</c> <b>desligava a verificação</b> e
    ''' passava como se nada tivesse mudado. É o formato de defeito que esta
    ''' fase inteira persegue: a proteção some junto com o dado que ela
    ''' protegia, e o resultado parece igual ao de um caso legítimo.
    ''' </summary>
    <TestMethod>
    Public Sub Universo_ausente_REJEITA_em_vez_de_desligar_a_guarda()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b") With {.UniversoNulo = True}
        Dim d As New DestinoFalso()

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou, "universo ausente nao pode passar como 'nao mudou'")
        StringAssert.Contains(r.Motivo, "universo")
    End Sub

    ' ==================================================================
    ' Fonte que mente

    <TestMethod>
    Public Sub Truncamento_REJEITA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d", "e", "f")
        Dim d As New DestinoFalso()
        f.TruncarApos = 1

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou, "truncar e o caso que PARECE sucesso")
        StringAssert.Contains(r.Motivo, "antes")
    End Sub

    <TestMethod>
    Public Sub Zero_enganoso_REJEITA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.ContagemDeclarada = 0

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou,
            "aceitar o zero publicaria uma geracao vazia e apagaria a pasta (§19.2)")
    End Sub

    <TestMethod>
    Public Sub Pasta_genuinamente_vazia_PUBLICA()
        Dim f As New FonteFalsaMutavel(Universo())
        Dim d As New DestinoFalso()

        Dim r = Rodar(f, d)

        Assert.IsTrue(r.Publicou, $"pasta vazia de verdade e valida. motivo: {r.Motivo}")
        Assert.AreEqual(0, r.Attempt.RowsRead)
    End Sub

    <TestMethod>
    Public Sub Cursor_que_nao_avanca_FALHA_em_vez_de_travar()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.CursorTravado = True

        Dim r = Rodar(f, d)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        StringAssert.Contains(r.Motivo, "cursor")
        Assert.IsTrue(f.PaginasLidas < 10, $"leu {f.PaginasLidas} paginas — deveria abortar cedo")
    End Sub

    ''' <summary>
    ''' Página VAZIA sem fim e com cursor parado também aborta cedo.
    '''
    ''' A guarda antiga só olhava páginas não vazias, então esta fonte rodava
    ''' 100.001 vezes antes de o teto pegar. Progresso é obrigatório com ou
    ''' sem linhas.
    ''' </summary>
    <TestMethod>
    Public Sub Pagina_vazia_sem_progresso_ABORTA_cedo()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b") With {.VaziaSemFim = True}
        Dim d As New DestinoFalso()

        Dim r = Rodar(f, d)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.IsTrue(f.PaginasLidas <= 2, $"leu {f.PaginasLidas} paginas vazias antes de abortar")
    End Sub

    ''' <summary>
    ''' Fonte que devolve MAIS linhas do que o pedido é recusada.
    '''
    ''' O tamanho da página é o orçamento de tempo da fila da STA (D5). Uma
    ''' fonte que devolve milhares de linhas de uma vez trava a UI mesmo sem
    ''' laço infinito nenhum.
    ''' </summary>
    <TestMethod>
    Public Sub Pagina_maior_que_o_pedido_FALHA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d", "e", "f")
        Dim d As New DestinoFalso()
        f.EstourarLote = 6

        Dim r = Rodar(f, d, tamanho:=2)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        StringAssert.Contains(r.Motivo, "linhas")
    End Sub

    <TestMethod>
    Public Sub Pagina_nula_FALHA()
        Dim f As New FonteFalsaMutavel(Universo(), "a") With {.PaginaNula = True}
        Dim d As New DestinoFalso()
        Dim r = Rodar(f, d)
        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        StringAssert.Contains(r.Motivo, "nula")
    End Sub

    <TestMethod>
    Public Sub Contagem_nula_FALHA()
        Dim f As New FonteFalsaMutavel(Universo(), "a") With {.ContagemNula = True}
        Dim d As New DestinoFalso()
        Dim r = Rodar(f, d)
        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        StringAssert.Contains(r.Motivo, "nula")
    End Sub

    ' ==================================================================
    ' Falha em cada fronteira

    <TestMethod>
    Public Sub Fonte_que_lanca_na_pagina_FALHA_sem_publicar()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d", "e", "f")
        Dim d As New DestinoFalso()
        f.LancarNaPagina = 2

        Dim r = Rodar(f, d)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.AreEqual(0, d.Publicadas)
        Assert.AreEqual(1, d.Paginas.Count, "a pagina 1 ficou gravada; a 2 nao chegou")
    End Sub

    <TestMethod>
    Public Sub Fonte_que_lanca_na_contagem_inicial_FALHA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b") With {.LancarNaContagem = 1}
        Dim d As New DestinoFalso()
        Dim r = Rodar(f, d)
        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.AreEqual(0, f.PaginasLidas)
    End Sub

    <TestMethod>
    Public Sub Fonte_que_lanca_na_contagem_final_FALHA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b") With {.LancarNaContagem = 2}
        Dim d As New DestinoFalso()
        Dim r = Rodar(f, d)
        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.AreEqual(0, d.Publicadas)
    End Sub

    <TestMethod>
    Public Sub Destino_que_lanca_ao_gravar_FALHA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso() With {.LancarAoGravar = True}
        Dim r = Rodar(f, d)
        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.AreEqual(0, d.Publicadas)
    End Sub

    <TestMethod>
    Public Sub Destino_que_lanca_na_epoca_FALHA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b")
        Dim d As New DestinoFalso() With {.LancarNaEpoca = True}
        Dim r = Rodar(f, d)
        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.AreEqual(0, d.Publicadas)
    End Sub

    ''' <summary>
    ''' O destino falhando ao publicar não pode virar "publicou" — e, sobretudo,
    ''' não pode transformar uma tentativa publicada em descartada.
    '''
    ''' Era esse o bug: o estado <c>Publicada</c> era instalado antes de o
    ''' efeito acontecer, e o <c>Catch</c> depois o convertia em
    ''' <c>Descartada</c>. Publicação é imutável — o 2.1 inteiro construiu
    ''' isso, e a orquestração desfazia em três linhas.
    ''' </summary>
    <TestMethod>
    Public Sub Destino_que_lanca_ao_publicar_nao_transforma_publicada_em_descartada()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b")
        Dim d As New DestinoFalso() With {.LancarAoPublicar = True}

        Dim r = Rodar(f, d)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.AreEqual(0, d.Publicadas)
        Assert.AreNotEqual(AttemptStage.Publicada, r.Attempt.Stage,
            "nunca chegou a publicar, entao nao pode constar como publicada")
    End Sub

    ''' <summary>
    ''' A persistência tem fencing próprio e pode recusar DEPOIS de o modelo
    ''' aprovar. O runner tem de respeitar a recusa dela.
    ''' </summary>
    <TestMethod>
    Public Sub Persistencia_que_recusa_por_ordem_nao_conta_como_publicado()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b")
        Dim d As New DestinoFalso() With {.RespostaAoPublicar = SinkPublishResult.RecusadaPorOrdem}

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou)
        Assert.AreEqual(0, d.Publicadas)
        StringAssert.Contains(r.Motivo, "RecusadaPorOrdem")
    End Sub

    <TestMethod>
    Public Sub Cancelamento_no_meio_nao_publica_nada()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d", "e", "f")
        Dim d As New DestinoFalso()
        Using cts As New CancellationTokenSource()
            f.Agenda(1) = Sub(x) cts.Cancel()
            Dim r = Rodar(f, d, ct:=cts.Token)

            Assert.AreEqual(SweepConclusion.Cancelada, r.Conclusion)
            Assert.AreEqual(0, d.Publicadas, "cancelar nunca publica metade")
            Assert.AreEqual(1, d.Descartadas.Count)
        End Using
    End Sub

    ' ==================================================================
    ' Época

    <TestMethod>
    Public Sub Epoca_mudou_durante_a_varredura_RECUSA_publicar()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.Agenda(1) = Sub(x) d.Epoca = 7

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou)
        Assert.AreEqual(0, d.Publicadas)
        StringAssert.Contains(r.Motivo.ToLowerInvariant(), "epoca")
    End Sub

    ' ==================================================================

    <TestMethod>
    Public Sub Pagina_rejeitada_nao_e_gravada()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.Agenda(1) = Sub(x) x.TrocarUniverso(
            New SweepUniverse("store-1", "pasta-1", "OUTRO", Nothing, 1, "amb-1"))

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou)
        Assert.AreEqual(1, d.Paginas.Count,
            "so a pagina 1, aceita antes da troca; a 2 foi rejeitada e nao gravada")
    End Sub

    ''' <summary>Toda rejeição descarta a tentativa no destino — nada fica aberto.</summary>
    <TestMethod>
    Public Sub Toda_rejeicao_descarta_a_tentativa_no_destino()
        For Each cenario In New Action(Of FonteFalsaMutavel, DestinoFalso)() {
            Sub(f, d) f.ContagemDeclarada = 0,
            Sub(f, d) f.TruncarApos = 1,
            Sub(f, d) f.CursorTravado = True,
            Sub(f, d) f.UniversoNulo = True,
            Sub(f, d) d.Epoca = 99}

            Dim fonte As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
            Dim destino As New DestinoFalso()
            cenario(fonte, destino)

            Dim r = Rodar(fonte, destino)

            Assert.IsFalse(r.Publicou)
            Assert.AreEqual(1, destino.Descartadas.Count,
                "tentativa aberta e nao publicada tem de ser descartada, senao fica orfa no banco")
        Next
    End Sub

End Class

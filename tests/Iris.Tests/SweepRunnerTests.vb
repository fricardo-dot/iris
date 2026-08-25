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

    ''' <summary>Capacidades que AUTORIZAM — o caminho que a produção hoje não tem.</summary>
    Private Shared Function Autorizado() As EnvironmentCapabilities
        Return EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "hipotetica"),
            {New MeasuredEnvironment(
                New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, "hipotetica"),
                "FASE2 §22.3 — sintetico", Nothing, tokenValidado:=True,
                grants:={New GrantedInference(Inference.ConcluirAusencia, "FASE2 §22.3 — s"),
                         New GrantedInference(Inference.AfirmarCoberturaCompleta, "FASE2 §22.3 — s"),
                         New GrantedInference(Inference.UsarIncremental, "FASE2 §22.3 — s")})})
    End Function

    Private Shared Function Rodar(f As FonteFalsaMutavel, d As DestinoFalso,
                                  Optional tamanho As Integer = 2,
                                  Optional ct As CancellationToken = Nothing,
                                  Optional cap As EnvironmentCapabilities = Nothing) As SweepResult
        Dim r As New SweepRunner(f, d, tamanho)
        Return r.Executar(f.UniversoAgora(), 0, 1, If(cap, Autorizado()), ct)
    End Function

    ' ==================================================================
    ' Controle positivo

    ''' <summary>
    ''' Sem defeito nenhum, publica — e o destino recebeu tudo.
    '''
    ''' É o controle positivo, e sem ele todos os testes abaixo passariam num
    ''' runner que simplesmente nunca publica.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_positivo_varredura_limpa_PUBLICA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d", "e")
        Dim d As New DestinoFalso()

        Dim r = Rodar(f, d)

        Assert.IsTrue(r.Publicou, $"deveria publicar. motivo: {r.Motivo}")
        Assert.AreEqual(1, d.Publicadas)
        CollectionAssert.AreEquivalent({"a", "b", "c", "d", "e"}, d.ChavesGravadas,
            "o destino tem de ter recebido todas as chaves")
        Assert.AreEqual(5, r.Attempt.RowsRead)
        Assert.AreEqual(5, r.Attempt.DistinctKeys)
    End Sub

    ' ==================================================================
    ' Mutação durante a leitura

    ''' <summary>
    ''' Item SOME entre a página 1 e o fim: a contagem final não bate e a
    ''' varredura é rejeitada.
    '''
    ''' É o caso do usuário apagando um e-mail enquanto o Iris varre.
    ''' </summary>
    <TestMethod>
    Public Sub Item_removido_no_meio_REJEITA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.Agenda(1) = Sub(x) x.Remover("d")

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou)
        Assert.AreEqual(0, d.Publicadas, "nao pode publicar meia varredura")
        ' O S6 pega na comparacao com a contagem INICIAL: leu 3, a pasta dizia
        ' 4 quando comecou. A contagem final tambem nao bateria, mas a
        ' primeira guarda dispara antes.
        StringAssert.Contains(r.Motivo, "lidas 3 <> antes 4")
    End Sub

    ''' <summary>
    ''' Item CHEGA no meio: idem. Mensagem nova durante a varredura é o caso
    ''' mais comum de todos.
    ''' </summary>
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
    ''' MUTAÇÃO BALANCEADA — e este teste documenta um BURACO, não uma defesa.
    '''
    ''' Um item sai e outro entra entre as duas contagens. Os números batem,
    ''' o S6 passa, e a varredura publica um manifesto que nunca existiu: tem
    ''' o item que saiu e não tem o que entrou, ou vice-versa.
    '''
    ''' O S6 <b>não pega isso</b>, e é importante que esteja escrito como
    ''' teste em vez de como comentário — um dia alguém vai olhar o S6 e achar
    ''' que ele garante integridade do conjunto. Ele garante que as contagens
    ''' fecham, que é bem menos.
    ''' </summary>
    <TestMethod>
    Public Sub Mutacao_balanceada_PASSA_e_isso_e_um_buraco_conhecido()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        ' Depois da pagina 1 (a,b), "d" some e "z" chega. Total continua 4.
        f.Agenda(1) = Sub(x) x.Trocar("d", "z")

        Dim r = Rodar(f, d)

        Assert.IsTrue(r.Publicou,
            "o S6 compara CONTAGENS; troca balanceada mantem as contagens e passa")
        Assert.IsTrue(d.ChavesGravadas.Contains("z"))
        Assert.IsFalse(d.ChavesGravadas.Contains("d"),
            "o manifesto publicado nao contem 'd', que existia quando a varredura comecou")
    End Sub

    ''' <summary>
    ''' O universo troca no meio: descarta, e o motivo diz universo — não
    ''' "estágio errado", que era o diagnóstico mascarado de antes.
    ''' </summary>
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

    ' ==================================================================
    ' Fonte que mente

    ''' <summary>
    ''' TRUNCAMENTO: a fonte para de devolver e diz que acabou.
    '''
    ''' É o mais perigoso de todos porque <b>parece sucesso</b> — a paginação
    ''' termina normalmente, sem erro. Só o S6 pega, comparando o que foi lido
    ''' com o que a pasta dizia ter.
    ''' </summary>
    <TestMethod>
    Public Sub Truncamento_REJEITA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d", "e", "f")
        Dim d As New DestinoFalso()
        f.TruncarApos = 1   ' devolve so (a,b) e declara fim

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou, "truncar e o caso que PARECE sucesso")
        Assert.AreEqual(0, d.Publicadas)
        StringAssert.Contains(r.Motivo, "antes")
    End Sub

    ''' <summary>
    ''' ZERO ENGANOSO: a pasta declara 0 itens e tem 4.
    '''
    ''' A §19.2 mediu isto na caixa do usuário — dezenas de pastas cheias
    ''' reportando <c>Count = 0</c>. Uma varredura que aceita o zero publica
    ''' uma geração vazia e <b>apaga a pasta</b> do ponto de vista do Iris.
    ''' </summary>
    <TestMethod>
    Public Sub Zero_enganoso_REJEITA()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.ContagemDeclarada = 0

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou,
            "aceitar o zero publicaria uma geracao vazia e apagaria a pasta")
        Assert.AreEqual(0, d.Publicadas)
    End Sub

    ''' <summary>
    ''' Pasta genuinamente vazia PUBLICA — o contraponto do teste acima.
    '''
    ''' Sem ele, um runner que rejeitasse toda contagem zero passaria no
    ''' anterior e quebraria a pasta vazia legítima, que é caso comum.
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_genuinamente_vazia_PUBLICA()
        Dim f As New FonteFalsaMutavel(Universo())
        Dim d As New DestinoFalso()

        Dim r = Rodar(f, d)

        Assert.IsTrue(r.Publicou, $"pasta vazia de verdade e valida. motivo: {r.Motivo}")
        Assert.AreEqual(0, r.Attempt.RowsRead)
    End Sub

    ''' <summary>
    ''' Cursor que não avança é laço infinito, e laço infinito na fila da STA
    ''' trava a UI. Não é defesa hipotética: é o modo de falha desta
    ''' arquitetura.
    ''' </summary>
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

    ' ==================================================================
    ' Falha e cancelamento

    <TestMethod>
    Public Sub Fonte_que_lanca_FALHA_sem_publicar()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d", "e", "f")
        Dim d As New DestinoFalso()
        f.LancarNaPagina = 2

        Dim r = Rodar(f, d)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.AreEqual(0, d.Publicadas)
        Assert.AreEqual(1, d.Paginas.Count, "a pagina 1 ficou gravada; a 2 nao chegou")
    End Sub

    ''' <summary>
    ''' O destino falhando ao publicar não pode virar "publicou".
    ''' </summary>
    <TestMethod>
    Public Sub Destino_que_lanca_ao_publicar_NAO_conta_como_publicado()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b")
        Dim d As New DestinoFalso() With {.LancarAoPublicar = True}

        Dim r = Rodar(f, d)

        Assert.AreEqual(SweepConclusion.Falhou, r.Conclusion)
        Assert.AreEqual(0, d.Publicadas)
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
        End Using
    End Sub

    ' ==================================================================
    ' Época e ambiente

    ''' <summary>
    ''' Outra geração publicou enquanto esta corria: a época mudou e esta
    ''' perde. É o critério 10 chegando na orquestração.
    ''' </summary>
    <TestMethod>
    Public Sub Epoca_mudou_durante_a_varredura_RECUSA_publicar()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        f.Agenda(1) = Sub(x) d.Epoca = 7   ' alguem invalidou o universo

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou)
        Assert.AreEqual(0, d.Publicadas)
        StringAssert.Contains(r.Motivo.ToLowerInvariant(), "epoca")
    End Sub

    ''' <summary>
    ''' Ambiente sem autorização NEM ABRE — e é o estado da produção hoje
    ''' (§23).
    '''
    ''' Recusar antes de abrir importa: uma tentativa aberta e abandonada fica
    ''' no banco para alguém retomar depois, e retomar uma varredura que nunca
    ''' deveria ter começado é pior que não ter começado.
    ''' </summary>
    <TestMethod>
    Public Sub Ambiente_nao_autorizado_NEM_ABRE()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b")
        Dim d As New DestinoFalso()

        ' O ambiente REAL do usuario: cached, janela nao legivel.
        Dim real = EnvironmentPolicy.Capacidades(
            New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing))

        Dim r = Rodar(f, d, cap:=real)

        Assert.AreEqual(SweepConclusion.Rejeitada, r.Conclusion)
        Assert.IsNull(r.Attempt, "nao pode nem existir tentativa")
        Assert.AreEqual(0, f.PaginasLidas, "nao pode nem ter lido a fonte")
        Assert.AreEqual(0, d.Paginas.Count)
        StringAssert.Contains(r.Motivo, "janela")
    End Sub

    <TestMethod>
    Public Sub Capacidades_nulas_NEM_ABREM()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b")
        Dim d As New DestinoFalso()
        Dim r = Rodar(f, d, cap:=EnvironmentPolicy.Capacidades(Nothing))
        Assert.AreEqual(SweepConclusion.Rejeitada, r.Conclusion)
        Assert.AreEqual(0, f.PaginasLidas)
    End Sub

    ' ==================================================================

    ''' <summary>
    ''' O efeito só acontece se o modelo mandar: quando a página é rejeitada,
    ''' ela NÃO é gravada.
    ''' </summary>
    <TestMethod>
    Public Sub Pagina_rejeitada_nao_e_gravada()
        Dim f As New FonteFalsaMutavel(Universo(), "a", "b", "c", "d")
        Dim d As New DestinoFalso()
        ' Universo troca antes da pagina 2 ser aceita.
        f.Agenda(1) = Sub(x) x.TrocarUniverso(
            New SweepUniverse("store-1", "pasta-1", "OUTRO", Nothing, 1, "amb-1"))

        Dim r = Rodar(f, d)

        Assert.IsFalse(r.Publicou)
        Assert.AreEqual(1, d.Paginas.Count,
            "so a pagina 1, aceita antes da troca; a 2 foi rejeitada e nao gravada")
    End Sub

End Class

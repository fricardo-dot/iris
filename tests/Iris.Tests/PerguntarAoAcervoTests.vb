Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Threading
Imports Iris.Cache
Imports Iris.Core
Imports Iris.Integration
Imports Iris.Model
Imports Iris.Sync
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>PERGUNTAR AO ACERVO — Fase 10.</b>
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTE ARQUIVO PRENDE</b>
'''
''' Que quem escolhe o que sai da máquina seja a <b>primeira</b> etapa — a que
''' roda aqui dentro, sobre metadado, antes de qualquer byte sair. Nada do que
''' o modelo devolva muda quem foi escolhido, porque quando ele fala a escolha
''' já foi feita.
'''
''' ------------------------------------------------------------------
''' <b>OS CONTROLES NEGATIVOS</b>
'''
''' <see cref="Sem_candidato_no_acervo_NAO_pergunta_nada"/> — sem ele, um modelo
''' sem fonte nenhuma responderia do que sabe do mundo, e o dono leria isso como
''' se fosse a caixa dele.
'''
''' <see cref="Ficha_inventada_na_citacao_zera_a_citacao_e_NAO_a_resposta"/> —
''' sem ele, uma citação errada mandaria o dono conferir a mensagem errada e
''' voltar achando que conferiu.
''' </summary>
' NAO PARALELIZAR: cada teste abre um SQLite proprio.
<TestClass>
<DoNotParallelize>
Public Class PerguntarAoAcervoTests

    Private Shared ReadOnly Quando As New DateTimeOffset(2026, 8, 31, 12, 0, 0, TimeSpan.Zero)

    ' ==================================================================

    ''' <summary>
    ''' <b>A etapa 1 escolhe, e a etapa 2 leva só o que ela escolheu.</b>
    ''' </summary>
    <TestMethod>
    Public Sub A_etapa_2_leva_SO_o_que_a_etapa_1_escolheu()
        Comigo(Sub(db)
                   Varrer(db, "f-1", {"contrato do joao", "almoço de sexta"})
                   Dim levadas As IReadOnlyList(Of PedidoDeParte) = Nothing

                   Dim r = Perguntador(db).Responder(
                       "contrato",
                       Function(pergunta, fontes)
                           levadas = fontes
                           Return "ele disse que assina na terça"
                       End Function)

                   Assert.AreEqual(MotivoDaResposta.Respondeu, r.Motivo)
                   Assert.AreEqual(1, levadas.Count, "levou mensagem que não casava")
                   StringAssert.Contains(levadas.Single().Chave.EntryId, "contrato")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O controle negativo principal.</b> Sem candidato, não se pergunta
    ''' nada: um modelo sem fonte nenhuma responde do que sabe do mundo, e o dono
    ''' leria isso como se fosse a caixa dele.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_candidato_no_acervo_NAO_pergunta_nada()
        Comigo(Sub(db)
                   Varrer(db, "f-1", {"almoço de sexta"})
                   Dim perguntou = False

                   Dim r = Perguntador(db).Responder(
                       "zzzzzz",
                       Function(pergunta, fontes)
                           perguntou = True
                           Return "claro, o contrato foi assinado"
                       End Function)

                   Assert.IsFalse(perguntou, "mandou a pergunta sem fonte nenhuma")
                   Assert.AreEqual(MotivoDaResposta.NadaNoAcervo, r.Motivo)
                   Assert.AreEqual("", r.Texto)
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>A cobertura anda junto, inclusive quando deu certo.</b> Ela é a única
    ''' coisa aqui que diz o que <i>não</i> foi olhado, e a resposta sem ela
    ''' parece completa.
    ''' </summary>
    <TestMethod>
    Public Sub A_cobertura_vem_sempre()
        Comigo(Sub(db)
                   Varrer(db, "f-1", {"contrato do joao"})

                   Dim r = Perguntador(db).Responder(
                       "contrato", Function(pergunta, fontes) "respondi")

                   Assert.IsTrue(r.Cobertura.Length > 0, "resposta sem cobertura")
                   StringAssert.Contains(r.Cobertura, "pasta")
               End Sub)
    End Sub

    ''' <summary>
    ''' Pasta conhecida e nunca varrida não é pasta vazia. "Não achei nada" e
    ''' "ninguém olhou" não podem chegar ao dono com a mesma cara.
    ''' </summary>
    <TestMethod>
    Public Sub Pasta_nunca_varrida_aparece_na_cobertura()
        Comigo(Sub(db)
                   ' Duas pastas conhecidas; so uma varrida.
                   Varrer(db, "f-1", {"contrato do joao"})
                   ' NEW ... .Metodo() NAO E STATEMENT em VB. Armadilha do CLAUDE.md.
                   Dim resolvedor As New ResolvedorDoAcervo(db)
                   resolvedor.Pasta("store-1", "f-2", "Pasta f-2")

                   Dim r = Perguntador(db).Responder(
                       "contrato", Function(pergunta, fontes) "respondi")

                   StringAssert.Contains(r.Cobertura, "1 de 2")
                   StringAssert.Contains(r.Cobertura, "nunca")
               End Sub)
    End Sub

    ''' <summary>
    ''' O teto existe para o dono conseguir conferir o que saiu olhando a lista.
    ''' </summary>
    <TestMethod>
    Public Sub O_teto_de_fontes_e_respeitado()
        Comigo(Sub(db)
                   Varrer(db, "f-1",
                       Enumerable.Range(1, PerguntarAoAcervo.MaximoDeFontes + 5).
                       Select(Function(i) "contrato " & i).ToArray())
                   Dim levadas = 0

                   Dim r = Perguntador(db).Responder(
                       "contrato",
                       Function(pergunta, fontes)
                           levadas = fontes.Count
                           Return "respondi"
                       End Function)

                   Assert.AreEqual(PerguntarAoAcervo.MaximoDeFontes, levadas)
                   Assert.AreEqual(PerguntarAoAcervo.MaximoDeFontes, r.Fontes.Count)
               End Sub)
    End Sub

    ' ==================================================================
    ' A CITAÇÃO

    <TestMethod>
    Public Sub A_citacao_volta_traduzida_e_sai_do_texto()
        Comigo(Sub(db)
                   Varrer(db, "f-1", {"contrato do joao"})

                   Dim r = Perguntador(db).Responder(
                       "contrato",
                       Function(pergunta, fontes)
                           Return "ele assina na terça." & Environment.NewLine &
                                  PerguntarAoAcervo.MarcaDasFontes & " " &
                                  fontes.Single().Ficha
                       End Function)

                   Assert.AreEqual(1, r.Citadas.Count)
                   Assert.AreEqual(r.Fontes.Single(), r.Citadas.Single())
                   Assert.IsFalse(r.Texto.Contains(PerguntarAoAcervo.MarcaDasFontes),
                                  "a linha das fontes ficou no texto")
               End Sub)
    End Sub

    ''' <summary>
    ''' <b>O outro controle negativo.</b> Ficha inventada zera a citação e
    ''' <b>não</b> a resposta: uma citação errada manda o dono conferir a
    ''' mensagem errada e voltar achando que conferiu — e descartar a resposta
    ''' daria a um e-mail o poder de apagar a resposta a uma pergunta que ele
    ''' fez.
    ''' </summary>
    <TestMethod>
    Public Sub Ficha_inventada_na_citacao_zera_a_citacao_e_NAO_a_resposta()
        Comigo(Sub(db)
                   Varrer(db, "f-1", {"contrato do joao"})

                   Dim r = Perguntador(db).Responder(
                       "contrato",
                       Function(pergunta, fontes)
                           Return "ele assina na terça." & Environment.NewLine &
                                  PerguntarAoAcervo.MarcaDasFontes & " iXXXX"
                       End Function)

                   Assert.AreEqual(MotivoDaResposta.Respondeu, r.Motivo)
                   StringAssert.Contains(r.Texto, "assina na terça")
                   Assert.AreEqual(0, r.Citadas.Count)
                   ' E o que SAIU daqui continua sabido.
                   Assert.AreEqual(1, r.Fontes.Count)
               End Sub)
    End Sub

    ''' <summary>
    ''' Sem a linha das fontes não há citação — e isso não é erro. O modelo pode
    ''' não citar; o que ele não pode é citar o que não recebeu.
    ''' </summary>
    <TestMethod>
    Public Sub Resposta_sem_a_linha_das_fontes_continua_valendo()
        Comigo(Sub(db)
                   Varrer(db, "f-1", {"contrato do joao"})

                   Dim r = Perguntador(db).Responder(
                       "contrato", Function(pergunta, fontes) "ele assina na terça")

                   Assert.AreEqual(MotivoDaResposta.Respondeu, r.Motivo)
                   Assert.AreEqual(0, r.Citadas.Count)
                   Assert.AreEqual(1, r.Fontes.Count)
               End Sub)
    End Sub

    ' ==================================================================
    ' QUANDO DÁ ERRADO

    <TestMethod>
    Public Sub Pergunta_vazia_nao_sai_daqui()
        Comigo(Sub(db)
                   Dim perguntou = False
                   Dim r = Perguntador(db).Responder(
                       "   ",
                       Function(pergunta, fontes)
                           perguntou = True
                           Return "algo"
                       End Function)

                   Assert.IsFalse(perguntou)
                   Assert.AreEqual(MotivoDaResposta.PerguntaVazia, r.Motivo)
               End Sub)
    End Sub

    <TestMethod>
    Public Sub Borda_que_estoura_nao_derruba_a_pergunta()
        Comigo(Sub(db)
                   Varrer(db, "f-1", {"contrato do joao"})

                   Dim r = Perguntador(db).Responder(
                       "contrato",
                       Function(pergunta, fontes) As String
                           Throw New InvalidOperationException("a rede caiu")
                       End Function)

                   Assert.AreEqual(MotivoDaResposta.SemResposta, r.Motivo)
                   Assert.IsTrue(r.Cobertura.Length > 0)
               End Sub)
    End Sub

    ''' <summary>
    ''' A pergunta do dono chega inteira à borda. Ela é a instrução dele, e
    ''' recortá-la aqui seria o programa reescrevendo o que ele perguntou.
    ''' </summary>
    <TestMethod>
    Public Sub A_pergunta_chega_inteira_na_borda()
        Comigo(Sub(db)
                   Varrer(db, "f-1", {"contrato do joao"})
                   Dim recebida = ""

                   Perguntador(db).Responder(
                       "contrato: quando o joão disse que assina?",
                       Function(pergunta, fontes)
                           recebida = pergunta
                           Return "na terça"
                       End Function)

                   Assert.AreEqual("contrato: quando o joão disse que assina?", recebida)
               End Sub)
    End Sub

    ' ==================================================================
    ' O ANDAIME

    Private Shared Function Perguntador(db As CacheDatabase) As PerguntarAoAcervo
        ' A MESMA SEQUENCIA DA PRODUCAO: o acervo nasce vazio, o dreno o enche,
        ' e so se nada veio e que se carrega a mao.
        Dim todas As New AcervoDeTodasAsPastas(db)
        Dim dreno As New PublicationDrain(db)
        dreno.Drenar(todas)
        If todas.Recarregado = 0 Then todas.Recarregar()

        Return New PerguntarAoAcervo(New BuscaNoAcervo(todas, dreno), todas)
    End Function

    Private Shared Function Impressao() As EnvironmentFingerprint
        Return New EnvironmentFingerprint(ProviderKind.ExchangeCached, True, Nothing)
    End Function

    Private Shared Function Varrer(db As CacheDatabase, entryId As String,
                                   assuntos As String()) As Long
        Dim resolvedor As New ResolvedorDoAcervo(db)
        Dim chave = resolvedor.Pasta("store-1", entryId, "Pasta " & entryId)
        Dim amb = resolvedor.Ambiente(Impressao())

        Dim universo As New SweepUniverse("store-1", entryId, "f", Nothing, 1, "amb-1")
        Dim fonte As New FonteDeLinhas(universo, assuntos.Select(
            Function(a) New SourceRow With {
                .Key = a,
                .Subject = a,
                .SenderName = "quem",
                .ReceivedAt = Quando.ToString("o"),
                .MessageClass = "IPM.Note"}))

        Dim sink As New SqliteSweepSink(db, chave, amb.Chave)
        Dim r = New SweepRunner(fonte, sink, 50).
                Executar(universo, 0, 1, EnvironmentPolicy.Capacidades(Impressao()),
                         CancellationToken.None)
        Assert.IsTrue(r.Publicou, $"controle: a varredura tinha de publicar. motivo: {r.Motivo}")
        Return chave
    End Function

    Private Shared Sub Comigo(corpo As Action(Of CacheDatabase))
        Dim caminho = Path.Combine(Path.GetTempPath(),
                                   "iris-perguntar-" & Guid.NewGuid().ToString("N") & ".db")
        Try
            Dim falha As OpenFailure = Nothing
            Using db = CacheDatabase.Open(caminho, CacheSchema.Intended(), falha)
                Assert.IsNotNull(db, $"{falha}")
                corpo(db)
            End Using
        Finally
            SqliteConnection.ClearAllPools()
            For Each sufixo In {"", "-wal", "-shm"}
                If File.Exists(caminho & sufixo) Then File.Delete(caminho & sufixo)
            Next
        End Try
    End Sub

    ''' <summary>
    ''' <b>A etapa 1 procura por palavra, e não pela frase inteira.</b>
    '''
    ''' A busca do acervo casa quem tem TODAS as palavras do termo — o certo
    ''' para quem digita uma busca, e o errado para uma pergunta: sete palavras
    ''' exigiriam um assunto com as sete, e não acham nada nunca.
    '''
    ''' Sem este teste, "perguntar ao acervo" respondia <i>nada no acervo</i> a
    ''' toda pergunta escrita como pergunta.
    ''' </summary>
    <TestMethod>
    Public Sub Pergunta_com_muitas_palavras_ainda_acha()
        Comigo(Sub(db)
                   Varrer(db, "f-1", {"contrato do joao"})
                   Dim levadas = 0

                   Perguntador(db).Responder(
                       "quando o joao disse que assina o contrato?",
                       Function(pergunta, fontes)
                           levadas = fontes.Count
                           Return "na terça"
                       End Function)

                   Assert.AreEqual(1, levadas)
               End Sub)
    End Sub

    ''' <summary>
    ''' Quem casa mais palavras da pergunta vem primeiro. As vagas são oito, e
    ''' gastá-las com quem casou uma palavra qualquer deixaria de fora a mensagem
    ''' que era a resposta.
    ''' </summary>
    <TestMethod>
    Public Sub Quem_casa_MAIS_palavras_vem_primeiro()
        Comigo(Sub(db)
                   Varrer(db, "f-1", {"contrato assinado joao", "contrato qualquer"})
                   Dim primeira = ""

                   Perguntador(db).Responder(
                       "contrato joao",
                       Function(pergunta, fontes)
                           primeira = fontes.First().Chave.EntryId
                           Return "respondi"
                       End Function)

                   Assert.AreEqual("contrato assinado joao", primeira)
               End Sub)
    End Sub

End Class

Imports System.Collections.Generic
Imports System.Linq
Imports Iris.Core
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>Qual mutação repete e qual mutação pula — demonstrado, não afirmado.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE ESTE ARQUIVO EXISTE</b>
'''
''' Eu escrevi no código, com confiança, que offset sobre coleção viva repete
''' quando algo é <b>removido</b> e pula quando algo é <b>inserido</b>. É o
''' contrário. O Codex mostrou com quatro linhas de exemplo, e a afirmação
''' errada já tinha virado regra de teste — "repetir com as pontas
''' concordando é defeito, porque deslocar offset exige remoção".
''' Premissa falsa, regra falsa.
'''
''' O que faltava não era revisão: era um lugar onde a direção pudesse ser
''' <b>exercitada</b>. Contra a caixa real ela nunca seria — a mutação
''' acontece quando quer, e o teste ou passa ou falha por motivo que ninguém
''' controla.
'''
''' Aqui a coleção é uma lista, a mutação é agendada, e o resultado é o que
''' é.
''' </summary>
<TestClass>
Public Class OffsetPagingTests

    ''' <summary>
    ''' Percorre <paramref name="mundo"/> em páginas de <paramref name="tamanho"/>,
    ''' aplicando <paramref name="mexer"/> depois da primeira página.
    ''' </summary>
    Private Shared Function Percorrer(mundo As List(Of String), tamanho As Integer,
                                      mexer As Action(Of List(Of String))) As List(Of String)
        Dim visto As New List(Of String)()
        Dim offset = 0
        Dim primeira = True
        Do
            Dim j = OffsetPaging.Janela(offset, tamanho, mundo.Count)
            For i = j.Primeiro To j.Ultimo
                visto.Add(mundo(i - 1))
            Next
            If primeira AndAlso mexer IsNot Nothing Then
                mexer(mundo)
                primeira = False
            End If
            If Not j.Proximo.HasValue Then Exit Do
            offset = j.Proximo.Value
        Loop
        Return visto
    End Function

    ' ==================================================================

    ''' <summary>Controle: coleção parada entrega tudo, uma vez cada.</summary>
    <TestMethod>
    Public Sub Colecao_PARADA_entrega_tudo_uma_vez()
        Dim mundo = New List(Of String)({"A", "B", "C", "D"})

        Dim visto = Percorrer(mundo, 2, Nothing)

        CollectionAssert.AreEqual({"A", "B", "C", "D"}, visto)
    End Sub

    ''' <summary>
    ''' <b>INSERÇÃO antes do offset REPETE.</b>
    '''
    ''' Página de 2 sobre <c>[A B C D]</c>: a primeira dá A e B, o cursor
    ''' guarda 2. Chega X na frente — <c>[X A B C D]</c> — e as posições 3 e 4
    ''' agora são <b>B</b> e C.
    ''' </summary>
    <TestMethod>
    Public Sub INSERCAO_antes_do_offset_REPETE()
        Dim mundo = New List(Of String)({"A", "B", "C", "D"})

        Dim visto = Percorrer(mundo, 2, Sub(m) m.Insert(0, "X"))

        Assert.IsTrue(visto.Count > visto.Distinct().Count(),
                      "inserir antes do offset TEM de repetir: " & String.Join(",", visto))
        CollectionAssert.Contains(visto.Skip(2).ToList(), "B",
                                  "e o repetido e justamente o ultimo da pagina anterior")
    End Sub

    ''' <summary>
    ''' <b>REMOÇÃO antes do offset PULA — em silêncio.</b>
    '''
    ''' Mesma travessia. Some A, e as posições 3 e 4 passam a ser D e nada.
    ''' <b>C nunca é visitado</b>, e a travessia termina parecendo completa.
    '''
    ''' Este é o pior dos dois, e é exatamente o sintoma que a Q1 existe para
    ''' pegar: perder mensagem sem que nada acuse.
    ''' </summary>
    <TestMethod>
    Public Sub REMOCAO_antes_do_offset_PULA_em_silencio()
        Dim mundo = New List(Of String)({"A", "B", "C", "D"})

        Dim visto = Percorrer(mundo, 2, Sub(m) m.Remove("A"))

        CollectionAssert.DoesNotContain(visto, "C",
            "remover antes do offset TEM de pular: " & String.Join(",", visto))
        Assert.AreEqual(visto.Count, visto.Distinct().Count(),
                        "e remover NAO repete — era isto que eu tinha escrito ao contrario")
    End Sub

    ''' <summary>
    ''' <b>Reordenar sem mudar o conjunto também repete.</b>
    '''
    ''' Importa porque derruba a segunda metade da minha afirmação errada:
    ''' "as pontas concordarem prova que nada deslocou". Aqui as pontas
    ''' concordam — o conjunto é idêntico antes e depois — e a travessia do
    ''' meio repete assim mesmo.
    '''
    ''' Numa pasta real isso é um <c>Subject</c> mudando com a ordenação em
    ''' <c>SubjectAsc</c>: o item atravessa a fronteira do offset sem que
    ''' nenhuma chave entre ou saia.
    ''' </summary>
    <TestMethod>
    Public Sub REORDENAR_sem_mudar_o_conjunto_TAMBEM_repete()
        Dim mundo = New List(Of String)({"A", "B", "C", "D"})

        Dim visto = Percorrer(mundo, 2, Sub(m)
                                            m.Remove("D")
                                            m.Insert(0, "D")
                                        End Sub)

        CollectionAssert.AreEquivalent({"A", "B", "C", "D"}, mundo,
            "o conjunto tem de ficar IDENTICO — e o ponto do teste")
        Assert.IsTrue(visto.Count > visto.Distinct().Count(),
            "conjunto identico nas pontas e repeticao no meio: " & String.Join(",", visto))
    End Sub

    ''' <summary>O cursor avança por posições EXAMINADAS, não por devolvidas.</summary>
    <TestMethod>
    Public Sub O_cursor_avanca_por_posicoes_examinadas()
        Dim j = OffsetPaging.Janela(offset:=0, quantas:=3, total:=10)
        Assert.AreEqual(1, j.Primeiro)
        Assert.AreEqual(3, j.Ultimo)
        Assert.AreEqual(3, j.Proximo.Value)
    End Sub

    ''' <summary>A última página não pede continuação.</summary>
    <TestMethod>
    Public Sub A_ultima_pagina_nao_pede_continuacao()
        Dim j = OffsetPaging.Janela(offset:=8, quantas:=5, total:=10)
        Assert.AreEqual(9, j.Primeiro)
        Assert.AreEqual(10, j.Ultimo)
        Assert.IsFalse(j.Proximo.HasValue)
    End Sub

    ''' <summary>Pasta que encolheu abaixo do offset não lê posição nenhuma.</summary>
    <TestMethod>
    Public Sub Pasta_que_encolheu_abaixo_do_offset_nao_le_nada()
        Dim j = OffsetPaging.Janela(offset:=50, quantas:=10, total:=3)
        Assert.IsTrue(j.Ultimo < j.Primeiro, "janela vazia, e o laco nao roda")
        Assert.IsFalse(j.Proximo.HasValue)
    End Sub

End Class

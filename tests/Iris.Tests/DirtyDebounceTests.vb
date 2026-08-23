Imports Iris.Core
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' O acumulador que decide QUANDO recarregar uma pasta suja.
'''
''' Estes testes não usam relógio: o instante entra como parâmetro. Testar
''' isto com <c>DispatcherTimer</c> e 2 s de espera real transformaria a
''' suíte num teste de carga da máquina — e "estável em N execuções" viraria
''' sorte, não prova. Este projeto já teve um teste intermitente e ele custou
''' caro.
'''
''' Cada teste nomeia a MUTAÇÃO que pega. Um teste de "vinte eventos geram
''' uma recarga" não distinguiria a implementação atual da versão errada que
''' este código já teve, porque as duas fariam uma recarga só.
''' </summary>
<TestClass>
Public Class DirtyDebounceTests

    Private Shared ReadOnly T0 As New DateTimeOffset(2026, 8, 23, 12, 0, 0, TimeSpan.Zero)

    Private Shared Function Em(ms As Integer) As DateTimeOffset
        Return T0.AddMilliseconds(ms)
    End Function

    <TestMethod>
    Public Sub Sem_evento_nao_ha_o_que_recarregar()
        Dim d As New DirtyDebounce()
        Assert.IsFalse(d.IsDirty)
        Assert.IsFalse(d.ShouldFlush(Em(999999)))
    End Sub

    ''' <summary>
    ''' PEGA: debounce que conta do PRIMEIRO evento.
    '''
    ''' É a implementação que este projeto já teve. Com ela, um evento em
    ''' t=0 e outro em t=400 recarregariam em t=450 — 50 ms depois do último
    ''' evento, no meio da rajada. O atraso tem de contar do ÚLTIMO.
    ''' </summary>
    <TestMethod>
    Public Sub O_atraso_conta_do_ultimo_evento_e_nao_do_primeiro()
        Dim d As New DirtyDebounce()

        d.Mark(Em(0))
        d.Mark(Em(400))

        Assert.IsFalse(d.ShouldFlush(Em(450)),
            "450 ms depois do PRIMEIRO evento, mas só 50 ms depois do último.")
        Assert.IsFalse(d.ShouldFlush(Em(800)), "ainda 400 ms de silêncio")
        Assert.IsTrue(d.ShouldFlush(Em(850)), "450 ms de silêncio depois do último")
    End Sub

    ''' <summary>
    ''' Controle positivo do teste acima: sem evento novo, o silêncio de
    ''' 450 ms libera. Sem isto, um acumulador que NUNCA liberasse passaria.
    ''' </summary>
    <TestMethod>
    Public Sub Evento_isolado_recarrega_apos_o_silencio()
        Dim d As New DirtyDebounce()
        d.Mark(Em(0))

        Assert.IsFalse(d.ShouldFlush(Em(449)))
        Assert.IsTrue(d.ShouldFlush(Em(450)))
    End Sub

    ''' <summary>
    ''' PEGA: teto que não participa da decisão.
    '''
    ''' O defeito original foi comparar a MESMA variável duas vezes —
    ''' <c>desde &lt; 450 AndAlso desde &lt; 2000</c> — o que se reduz a
    ''' <c>desde &lt; 450</c> e deixa o teto como letra morta. Numa rajada
    ''' que não para — sincronização longa do Exchange — a recarga era
    ''' adiada indefinidamente, e a lista ficava velha justamente quando
    ''' mais muda.
    ''' </summary>
    <TestMethod>
    Public Sub Rajada_continua_recarrega_no_teto()
        Dim d As New DirtyDebounce()

        ' Um evento a cada 100 ms: o silêncio nunca chega a 450.
        Dim t = 0
        While t < 1900
            d.Mark(Em(t))
            Assert.IsFalse(d.ShouldFlush(Em(t)),
                $"em t={t} ainda não deu nem silêncio nem teto")
            t += 100
        End While

        d.Mark(Em(1900))
        Assert.IsFalse(d.ShouldFlush(Em(1950)), "1950 ms desde o primeiro: teto ainda não")
        Assert.IsTrue(d.ShouldFlush(Em(2000)),
            "2 s desde o PRIMEIRO evento libera mesmo sem silêncio.")
    End Sub

    ''' <summary>
    ''' PEGA: acumulador que trava depois da primeira descarga.
    '''
    ''' Uma segunda rajada precisa gerar uma segunda recarga, e a contagem
    ''' recomeça do zero — o teto da rajada nova não pode herdar o instante
    ''' da anterior.
    ''' </summary>
    <TestMethod>
    Public Sub Segunda_rajada_gera_segunda_recarga()
        Dim d As New DirtyDebounce()

        d.Mark(Em(0))
        Assert.IsTrue(d.ShouldFlush(Em(500)))
        d.Clear()
        Assert.IsFalse(d.IsDirty)
        Assert.IsFalse(d.ShouldFlush(Em(5000)), "limpo não recarrega de novo sozinho")

        d.Mark(Em(5000))
        Assert.IsFalse(d.ShouldFlush(Em(5100)),
            "a rajada nova recomeça a contagem; não herda o teto da anterior")
        Assert.IsTrue(d.ShouldFlush(Em(5450)))
    End Sub

    ''' <summary>
    ''' Consultar não pode mudar o que está sendo medido. Se
    ''' <c>ShouldFlush</c> alterasse estado, o timer — que pergunta a cada
    ''' 150 ms — corromperia a própria contagem.
    ''' </summary>
    <TestMethod>
    Public Sub Perguntar_nao_altera_o_estado()
        Dim d As New DirtyDebounce()
        d.Mark(Em(0))

        For i = 1 To 20
            Assert.IsFalse(d.ShouldFlush(Em(100)))
        Next
        Assert.IsTrue(d.IsDirty)

        For i = 1 To 20
            Assert.IsTrue(d.ShouldFlush(Em(500)))
        Next
        Assert.IsTrue(d.IsDirty, "só Clear limpa; perguntar não.")
    End Sub

    ''' <summary>
    ''' Limpar no meio de uma rajada zera de verdade: o evento seguinte
    ''' começa uma rajada nova. É o que acontece quando o usuário troca de
    ''' pasta ou a sessão é substituída.
    ''' </summary>
    <TestMethod>
    Public Sub Limpar_no_meio_da_rajada_recomeca_a_contagem()
        Dim d As New DirtyDebounce()

        d.Mark(Em(0))
        d.Mark(Em(100))
        d.Clear()

        d.Mark(Em(200))
        Assert.IsFalse(d.ShouldFlush(Em(600)),
            "400 ms desde o evento novo — a contagem é dele, não da rajada abandonada")
        Assert.IsTrue(d.ShouldFlush(Em(650)))
    End Sub

    ''' <summary>
    ''' Os limites são configuráveis, e o teste confere que eles são
    ''' realmente usados — não que existem como constante decorativa.
    ''' </summary>
    <TestMethod>
    Public Sub Os_limites_configurados_valem()
        Dim d As New DirtyDebounce(debounceMs:=50, tetoMs:=200)

        d.Mark(Em(0))
        Assert.IsFalse(d.ShouldFlush(Em(49)))
        Assert.IsTrue(d.ShouldFlush(Em(50)), "debounce curto libera antes")

        Dim r As New DirtyDebounce(debounceMs:=50, tetoMs:=200)
        Dim t = 0
        While t <= 190
            r.Mark(Em(t))
            t += 25
        End While
        Assert.IsTrue(r.ShouldFlush(Em(200)), "teto curto libera a rajada contínua antes")
    End Sub

End Class

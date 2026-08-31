Imports System.IO
Imports Iris.Integration
Imports Iris.Model
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>QUEM É "EU" NUMA MENSAGEM — Fase 2, primeira fatia.</b>
'''
''' ------------------------------------------------------------------
''' <b>A REGRA QUE ESTE ARQUIVO IMPEDE DE VOLTAR</b>
'''
''' <i>"Está em Itens Enviados, logo fui eu."</i> Ela funciona na caixa de quem
''' tem uma conta só e nunca respondeu por outra pessoa, e quebra em alias,
''' conta adicional, caixa compartilhada e regra de movimentação. O sintoma não
''' é erro: é a fila de respostas pendentes cobrando do dono mensagens que ele
''' mesmo escreveu.
'''
''' ------------------------------------------------------------------
''' <b>O CONTROLE NEGATIVO</b>
'''
''' <see cref="Sem_identidade_nenhuma_a_resposta_e_NAO_SEI"/>. Sem ele, uma
''' implementação que respondesse <c>DoOutro</c> para tudo passaria em quase
''' todos os outros testes daqui — e seria exatamente a que enche a fila de
''' dívida falsa.
''' </summary>
<TestClass>
Public Class MinhasIdentidadesTests

    Private Shared Function Minhas(ParamArray enderecos As String()) As MinhasIdentidades
        Return New MinhasIdentidades(enderecos)
    End Function

    <TestMethod>
    Public Sub O_remetente_que_sou_eu_e_MINHA()
        Assert.AreEqual(Direcao.Minha,
                        Minhas("ricardo@empresa.com").DirecaoDe("ricardo@empresa.com"))
    End Sub

    <TestMethod>
    Public Sub O_remetente_que_nao_sou_eu_e_DO_OUTRO()
        Assert.AreEqual(Direcao.DoOutro,
                        Minhas("ricardo@empresa.com").DirecaoDe("caroline@outra.com"))
    End Sub

    ''' <summary>
    ''' <b>O controle negativo.</b> Vazio responde "não sei" para tudo — nunca
    ''' "do outro". Palpite aqui vira cobrança contra o dono por mensagem que
    ''' ele escreveu.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_identidade_nenhuma_a_resposta_e_NAO_SEI()
        Dim vazio = Minhas()

        Assert.IsTrue(vazio.Vazio)
        Assert.AreEqual(Direcao.Desconhecida, vazio.DirecaoDe("ricardo@empresa.com"),
            "sem identidade declarada, ate a minha propria mensagem e 'nao sei'")
        Assert.AreEqual(Direcao.Desconhecida, vazio.DirecaoDe("caroline@outra.com"),
            "e a dos outros tambem: o vazio nao pode virar 'do outro por " &
            "eliminacao', senao a fila cobra o dono por mensagem dele mesmo")
    End Sub

    ''' <summary>
    ''' Remetente ilegível é leitura que falhou, e leitura que falhou não é
    ''' evidência de nada.
    ''' </summary>
    <TestMethod>
    Public Sub Remetente_ausente_e_NAO_SEI()
        Dim eu = Minhas("ricardo@empresa.com")

        Assert.AreEqual(Direcao.Desconhecida, eu.DirecaoDe(Nothing))
        Assert.AreEqual(Direcao.Desconhecida, eu.DirecaoDe(""))
        Assert.AreEqual(Direcao.Desconhecida, eu.DirecaoDe("   "))
    End Sub

    ''' <summary>
    ''' <b>O alias é o caso que motivou tudo isto.</b> Uma mensagem enviada por
    ''' um alias é minha, e o alias não aparece em conta nenhuma do Outlook —
    ''' é por isso que o conjunto é explícito e editável.
    ''' </summary>
    <TestMethod>
    Public Sub O_ALIAS_tambem_sou_eu()
        Dim eu = Minhas("ricardo@empresa.com", "regulatorio@empresa.com")

        Assert.AreEqual(Direcao.Minha, eu.DirecaoDe("regulatorio@empresa.com"))
    End Sub

    ''' <summary>
    ''' Numa organização Exchange o remetente interno chega como X.500, e não
    ''' como SMTP. Se essa forma não casar, as mensagens internas do próprio
    ''' dono — as que mais enchem a fila — apareceriam como sendo de terceiros.
    ''' </summary>
    <TestMethod>
    Public Sub O_endereco_X500_tambem_casa()
        Const x500 = "/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=abc123"
        Dim eu = Minhas("ricardo@empresa.com", x500)

        Assert.AreEqual(Direcao.Minha, eu.DirecaoDe(x500))
        Assert.AreEqual(Direcao.Minha, eu.DirecaoDe(x500.ToLowerInvariant()),
            "o X.500 chega com caixa variada, e a comparacao ignora caixa")
    End Sub

    <TestMethod>
    Public Sub A_caixa_das_letras_nao_muda_quem_sou()
        Dim eu = Minhas("Ricardo@Empresa.COM")

        Assert.AreEqual(Direcao.Minha, eu.DirecaoDe("ricardo@empresa.com"))
        Assert.AreEqual(Direcao.Minha, eu.DirecaoDe("RICARDO@EMPRESA.COM"))
    End Sub

    ''' <summary>
    ''' <c>Fulano &lt;f@x&gt;</c> é a forma que o Outlook devolve em vários
    ''' campos, e comparar a cadeia inteira faria o dono virar terceiro.
    ''' </summary>
    <TestMethod>
    Public Sub O_nome_de_exibicao_e_descartado()
        Dim eu = Minhas("ricardo@empresa.com")

        Assert.AreEqual(Direcao.Minha, eu.DirecaoDe("Ricardo Fernandes <ricardo@empresa.com>"))
        Assert.AreEqual(Direcao.Minha, eu.DirecaoDe("  <ricardo@empresa.com>  "))
    End Sub

    ''' <summary>
    ''' Um <c>&lt;</c> sem par é lixo, e cortar por ele <b>inventaria</b> um
    ''' endereço. Melhor não casar do que casar com o que não está escrito.
    ''' </summary>
    <TestMethod>
    Public Sub Sinal_solto_nao_inventa_endereco()
        Dim eu = Minhas("ricardo@empresa.com")

        Assert.AreEqual(Direcao.DoOutro, eu.DirecaoDe("Ricardo <ricardo@empresa.com"),
            "sem o par completo, a cadeia fica como veio")
    End Sub

    ' ==================================================================
    ' O ARQUIVO

    Private Shared Function Temporario() As String
        Return Path.Combine(Path.GetTempPath(),
                            "iris-identidades-" & Guid.NewGuid().ToString("N") & ".txt")
    End Function

    <TestMethod>
    Public Sub Semear_escreve_e_o_que_foi_escrito_volta()
        Dim caminho = Temporario()
        Try
            Dim arquivo As New IdentidadesEmArquivo(caminho)

            Assert.IsTrue(arquivo.Semear({"ricardo@empresa.com", "regulatorio@empresa.com"}))

            Dim eu = arquivo.Ler()
            Assert.AreEqual(2, eu.Quantas, "o cabecalho de comentarios nao pode virar endereco")
            Assert.AreEqual(Direcao.Minha, eu.DirecaoDe("regulatorio@empresa.com"))
        Finally
            If File.Exists(caminho) Then File.Delete(caminho)
        End Try
    End Sub

    ''' <summary>
    ''' <b>Semeadura não desfaz correção.</b> O dono apaga uma linha porque
    ''' aquele endereço não é dele; uma semeadura que insistisse a poria de
    ''' volta na abertura seguinte, e a correção nunca pegaria.
    ''' </summary>
    <TestMethod>
    Public Sub Semear_de_novo_NAO_toca_no_arquivo()
        Dim caminho = Temporario()
        Try
            Dim arquivo As New IdentidadesEmArquivo(caminho)
            arquivo.Semear({"ricardo@empresa.com", "enganado@empresa.com"})

            ' O dono corrige a mao: sobra uma.
            File.WriteAllLines(caminho, {"ricardo@empresa.com"})

            Assert.IsFalse(arquivo.Semear({"ricardo@empresa.com", "enganado@empresa.com"}),
                "semear de novo tem de recusar")
            Assert.AreEqual(1, arquivo.Ler().Quantas, "a correcao do dono foi desfeita")
        Finally
            If File.Exists(caminho) Then File.Delete(caminho)
        End Try
    End Sub

    ''' <summary>
    ''' Sem endereço nenhum para semear, não escreve. Um arquivo só de
    ''' comentários pareceria semeado, e a semeadura de verdade — quando o
    ''' Outlook enfim responder — nunca mais aconteceria.
    ''' </summary>
    <TestMethod>
    Public Sub Semear_o_vazio_nao_cria_arquivo()
        Dim caminho = Temporario()
        Try
            Dim arquivo As New IdentidadesEmArquivo(caminho)

            Assert.IsFalse(arquivo.Semear({}))
            Assert.IsFalse(File.Exists(caminho))
            Assert.IsFalse(arquivo.Existe)
        Finally
            If File.Exists(caminho) Then File.Delete(caminho)
        End Try
    End Sub

    ''' <summary>
    ''' Arquivo que não dá para ler vale como conjunto vazio, e conjunto vazio
    ''' responde "não sei". A fila mostra linha incerta em vez de linha errada.
    ''' </summary>
    <TestMethod>
    Public Sub Arquivo_ausente_vale_como_VAZIO()
        Dim eu = New IdentidadesEmArquivo(Temporario()).Ler()

        Assert.IsTrue(eu.Vazio)
        Assert.AreEqual(Direcao.Desconhecida, eu.DirecaoDe("ricardo@empresa.com"))
    End Sub

    ''' <summary>
    ''' <b>Cadeia sem forma de endereço é "não sei", e não "do outro".</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>O PISO DO VAZIO ENTRANDO PELA PORTA DOS FUNDOS</b>
    '''
    ''' <c>Normalizar</c> tira o nome de exibição de <c>Fulano &lt;f@x&gt;</c>.
    ''' Numa cadeia como <c>Diretoria &lt;Regulatorio&gt;</c> ele produz
    ''' <c>regulatorio</c> — que não está no conjunto e virava <c>DoOutro</c>.
    '''
    ''' Ou seja: uma leitura que não deu certo produzia uma <b>afirmação</b>,
    ''' e a fila cobrava do dono uma resposta por causa dela. É exatamente o
    ''' engano que <see cref="Sem_identidade_nenhuma_a_resposta_e_NAO_SEI"/>
    ''' impede pela porta da frente.
    '''
    ''' Achado por revisão externa em 31/08/2026.
    ''' </summary>
    <TestMethod>
    Public Sub Cadeia_sem_forma_de_endereco_e_NAO_SEI()
        Dim eu = Minhas("ricardo@empresa.com")

        Assert.AreEqual(Direcao.Desconhecida, eu.DirecaoDe("Diretoria <Regulatorio>"),
            "extraiu o miolo de um par de sinais e afirmou que e de outra " &
            "pessoa -- uma leitura que nao deu certo virando fato")
        Assert.AreEqual(Direcao.Desconhecida, eu.DirecaoDe("Caroline Abreu"),
            "nome sem endereco nenhum nao identifica ninguem")
        Assert.AreEqual(Direcao.Desconhecida, eu.DirecaoDe("(remetente ilegivel)"))
    End Sub

    ''' <summary>
    ''' <b>O controle da porta:</b> a exigência de forma não pode fechar as
    ''' duas formas que valem. SMTP tem <c>@</c>; X.500 começa com <c>/</c>.
    ''' Sem isto, um endurecimento a mais transformaria a fila inteira em
    ''' "não sei" e o arquivo continuaria verde.
    ''' </summary>
    <TestMethod>
    Public Sub A_exigencia_de_forma_NAO_fecha_as_duas_que_valem()
        Const x500 = "/O=EXCHANGELABS/OU=EAG/CN=RECIPIENTS/CN=abc123"
        Dim eu = Minhas("ricardo@empresa.com", x500)

        Assert.AreEqual(Direcao.Minha, eu.DirecaoDe("ricardo@empresa.com"))
        Assert.AreEqual(Direcao.Minha, eu.DirecaoDe(x500))
        Assert.AreEqual(Direcao.DoOutro, eu.DirecaoDe("caroline@outra.com"),
            "endereco valido de outra pessoa continua sendo dela, e nao um " &
            "nao-sei por excesso de cautela")
    End Sub

    ''' <summary>
    ''' Linha solta no arquivo não vira identidade. Ele é editado à mão, e
    ''' uma anotação virando "identidade" faria o dono casar com qualquer
    ''' remetente que se lesse igual.
    ''' </summary>
    <TestMethod>
    Public Sub Anotacao_solta_no_arquivo_NAO_vira_identidade()
        Dim eu = Minhas("ricardo@empresa.com", "meu email do trabalho")

        Assert.AreEqual(1, eu.Quantas, "a anotacao entrou no conjunto")
    End Sub

End Class

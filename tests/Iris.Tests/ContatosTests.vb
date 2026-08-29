Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>CONTATOS — e o que uma pasta vazia não prova.</b>
'''
''' ------------------------------------------------------------------
''' <b>O ASSUNTO DA FASE 7</b>
'''
''' Medido em 28/08/2026: a pasta padrão de Contatos desta caixa tem
''' <b>0 itens</b>. Não porque não haja contatos — porque numa conta
''' corporativa eles vivem no <b>GAL</b>, que a §8 põe fora de escopo.
'''
''' Então esta é a fase em que a leitura mais correta do mundo produz o
''' resultado mais enganoso do projeto, se a tela não disser o que não olhou.
''' É a família ausência-virando-fato que esta base corrigiu em cinco lugares,
''' desta vez sobre <b>pessoas</b>.
'''
''' ------------------------------------------------------------------
''' <b>E A DIFERENÇA ENTRE VAZIO E ILEGÍVEL</b>
'''
''' Um <c>ContactItem</c> tem mais de cem propriedades, e qualquer uma pode
''' falhar na leitura. Devolver cadeia vazia nesse caso diria "este contato não
''' tem empresa" sobre um campo que ninguém conseguiu ler — e num catálogo de
''' endereços isso não é um detalhe de apresentação, é cadastro errado.
''' </summary>
<TestClass>
Public Class ContatosTests

    ' ==================================================================
    ' A ressalva

    ''' <summary>
    ''' <b>A ressalva fala do GAL, e diz que vazio não é ausência.</b>
    '''
    ''' Este teste prende as <i>palavras</i>, e não só a existência do campo.
    ''' Uma ressalva que dissesse "leitura parcial" seria verdadeira e inútil:
    ''' quem lê precisa saber que existe um catálogo inteiro fora do alcance.
    ''' </summary>
    <TestMethod>
    Public Sub A_ressalva_nomeia_o_GAL_e_nega_a_ausencia()
        Dim r = RegrasDeContato.ForaDoAlcance

        StringAssert.Contains(r, "GAL", "a ressalva não nomeia o catálogo que falta")
        StringAssert.Contains(r, "não quer dizer que não haja contatos",
            "a ressalva não desfaz a leitura de ausência, que é a única coisa " &
            "que ela existe para desfazer")
    End Sub

    ' ==================================================================
    ' A busca por repetido

    <TestMethod>
    Public Sub Controle_acha_o_contato_com_o_mesmo_endereco()
        Dim lidos = {Contato("Ana", "ana@empresa.com"),
                     Contato("Bruno", "bruno@empresa.com")}

        Dim achado = RegrasDeContato.Procurar(lidos, "BRUNO@empresa.com")

        Assert.IsNotNull(achado, "não achou um endereço que está na lista")
        Assert.AreEqual("Bruno", achado.Nome)
    End Sub

    <TestMethod>
    Public Sub Endereco_que_nao_esta_na_lista_nao_e_achado()
        Dim lidos = {Contato("Ana", "ana@empresa.com")}

        Assert.IsNull(RegrasDeContato.Procurar(lidos, "carlos@empresa.com"))
    End Sub

    ''' <summary>
    ''' <b>DOIS CONTATOS ILEGÍVEIS NÃO SÃO O MESMO CONTATO.</b>
    '''
    ''' O caso que quase passou despercebido. Se <c>Email</c> ilegível virasse
    ''' cadeia vazia, dois contatos que ninguém conseguiu ler casariam entre si
    ''' — e a tela avisaria "já existe um contato com este endereço" sobre um
    ''' endereço que ela não leu.
    '''
    ''' <b>Controle negativo:</b> trocando o <c>Nothing</c> do <c>Texto</c> por
    ''' <c>""</c>, este teste cai.
    ''' </summary>
    <TestMethod>
    Public Sub Contato_com_email_ILEGIVEL_nao_casa_com_nada()
        Dim ilegivel = New ContactInfo With {.Nome = "quem sabe", .Email = Nothing}
        Dim lidos = {ilegivel}

        Assert.IsNull(RegrasDeContato.Procurar(lidos, ""),
                      "casou uma busca vazia com um contato ilegível")
        Assert.IsNull(RegrasDeContato.Procurar(lidos, "   "),
                      "casou espaço em branco com um contato ilegível")
        Assert.IsNull(RegrasDeContato.Procurar(lidos, "alguem@empresa.com"))
    End Sub

    ''' <summary>
    ''' E contato com e-mail <b>vazio de verdade</b> também não casa com busca
    ''' vazia — porque busca vazia não é busca.
    ''' </summary>
    <TestMethod>
    Public Sub Busca_sem_endereco_nao_acha_ninguem()
        Dim lidos = {Contato("Sem endereço", "")}

        Assert.IsNull(RegrasDeContato.Procurar(lidos, ""))
        Assert.IsNull(RegrasDeContato.Procurar(Nothing, "ana@empresa.com"))
    End Sub

    ' ==================================================================
    ' A guarda de criação

    <TestMethod>
    Public Sub Controle_um_rascunho_comum_e_aceito()
        Assert.IsNull(ContactWriting.RecusarRascunho(
            New ContactDraft With {.Nome = "Ana Lima", .Email = "ana@empresa.com"}))
    End Sub

    ''' <summary>
    ''' <b>Contato sem nome não entra.</b>
    '''
    ''' Uma ficha em branco no catálogo não diz de quem é, e quem a encontrar
    ''' depois não tem como saber — o Iris não estava lá quando ela nasceu.
    ''' </summary>
    <DataTestMethod>
    <DataRow("")>
    <DataRow("   ")>
    Public Sub Sem_nome_RECUSA(nome As String)
        Dim motivo = ContactWriting.RecusarRascunho(New ContactDraft With {.Nome = nome})

        Assert.IsNotNull(motivo, "aceitou contato sem nome")
        StringAssert.Contains(motivo, "nome")
    End Sub

    ''' <summary>
    ''' <b>Mas contato SEM E-MAIL entra.</b>
    '''
    ''' Contato de telefone é contato. Exigir endereço faria a guarda recusar
    ''' cadastro legítimo — e uma guarda que recusa demais é tão defeituosa
    ''' quanto uma que recusa de menos, só que ninguém reclama dela.
    ''' </summary>
    <TestMethod>
    Public Sub Sem_email_e_aceito()
        Assert.IsNull(ContactWriting.RecusarRascunho(
            New ContactDraft With {.Nome = "Ana Lima", .Email = ""}))
    End Sub

    <TestMethod>
    Public Sub Rascunho_nulo_RECUSA()
        Assert.IsNotNull(ContactWriting.RecusarRascunho(Nothing))
    End Sub

    ' ==================================================================
    ' A invariante, presa pelo tipo e pelo fonte

    ''' <summary>
    ''' <b>O RASCUNHO NÃO TEM NOTA, NEM CORPO.</b>
    '''
    ''' O que o Iris cria a partir de uma mensagem é só o que a mensagem já
    ''' dizia. Um campo de nota convidaria a pôr no catálogo de endereços um
    ''' texto que o assistente escreveu — e aí um dado gerado viraria dado de
    ''' cadastro, num lugar que outras ferramentas leem como verdade conferida.
    ''' </summary>
    <TestMethod>
    Public Sub O_rascunho_de_contato_NAO_tem_nota_nem_corpo()
        Dim campos = GetType(ContactDraft).GetProperties().
                     Select(Function(p) p.Name.ToLowerInvariant()).ToList()

        For Each proibido In {"body", "nota", "notes", "observacao", "resumo",
                              "personalhomepage", "comentario"}
            Assert.IsFalse(campos.Contains(proibido),
                $"ContactDraft ganhou '{proibido}': o que o Iris grava no " &
                "catálogo é o que a mensagem já dizia, e não texto gerado. " &
                "Leia o comentário do ContactDraft antes de mexer nisto.")
        Next
    End Sub

    ''' <summary>
    ''' <b>O ESCRITOR NÃO ENCAMINHA CARTÃO.</b>
    '''
    ''' Salvar contato é escrita local — este objeto não tem a armadilha da
    ''' reunião nem a da tarefa atribuída. Tem uma só:
    ''' <c>ForwardAsVcard()</c> devolve um <c>MailItem</c> pronto para sair.
    '''
    ''' Como na Fase 5, a varredura lê o <b>fonte</b>, e é um proxy — a
    ''' alternativa seria um duplo de <c>ContactItem</c>, e <c>OL.ContactItem</c>
    ''' é interface COM que o teste não instancia. Entre um proxy e nenhuma
    ''' prova, o proxy.
    ''' </summary>
    <TestMethod>
    Public Sub O_escritor_de_contatos_NAO_encaminha_nem_envia()
        Dim fonte = FonteSemComentarios()

        Assert.IsFalse(fonte.Contains("Forward"),
            "ContactWriting encaminha: ForwardAsVcard devolve um MailItem, e é " &
            "o único caminho de envio que um contato tem. O Iris não envia.")
        Assert.IsFalse(fonte.Contains(".Send("),
            "ContactWriting chama Send: nada sai por e-mail sem o usuário mandar.")

        ' CONTROLE: a varredura acha o que existe. Sem ele, um caminho errado
        ' deixaria as duas asserções passarem por lerem vazio.
        StringAssert.Contains(fonte, ".Save()",
            "a varredura não está lendo o ContactWriting: as asserções acima " &
            "estariam passando por não olhar nada")
    End Sub

    ''' <summary>
    ''' <b>A ressalva é preenchida PELO LEITOR, e não pela tela.</b>
    '''
    ''' Este teste prende a decisão de desenho. Se o <c>Ler</c> deixasse
    ''' <c>ForaDoAlcance</c> vazio e a tela montasse o texto, a ressalva sumiria
    ''' na primeira tela nova que alguém escrevesse — e sumiria em silêncio,
    ''' porque nada quebra quando uma ressalva deixa de aparecer.
    ''' </summary>
    <TestMethod>
    Public Sub O_leitor_marca_a_ressalva_no_resultado()
        Dim fonte = FonteSemComentarios()

        StringAssert.Contains(fonte, ".ForaDoAlcance = RegrasDeContato.ForaDoAlcance",
            "o leitor parou de marcar a ressalva no resultado: ela passou a " &
            "depender de a tela lembrar, e ressalva que depende de lembrar some")
    End Sub

    ' ==================================================================

    ''' <summary>
    ''' <b>LEITURA QUE FALHA DEVOLVE Nothing, E NAO CADEIA VAZIA.</b>
    '''
    ''' Este teste nasceu de um controle negativo que passou quando devia
    ''' falhar: trocar o <c>Nothing</c> por <c>""</c> aqui nao derrubava
    ''' nada, porque todo teste de leitura monta o <c>ContactInfo</c> a mao
    ''' e nunca passa por esta funcao. A guarda existia sem prova.
    '''
    ''' A diferenca importa duas vezes: a tela mostra frases diferentes, e
    ''' a busca por repetido nao casa dois contatos ilegiveis entre si.
    ''' </summary>
    <TestMethod>
    Public Sub Campo_que_nao_se_deixa_ler_devolve_Nothing()
        Dim quebrado = ContactWriting.Texto(
            Function() As String
                Throw New InvalidOperationException("o COM recusou")
            End Function)

        Assert.IsNull(quebrado,
            "campo ilegivel virou cadeia vazia -- e cadeia vazia diz " &
            "''este contato nao tem empresa'' sobre um campo que ninguem leu")
    End Sub

    <TestMethod>
    Public Sub Controle_campo_legivel_atravessa()
        Assert.AreEqual("Ana Lima", ContactWriting.Texto(Function() "Ana Lima"))
        Assert.AreEqual("", ContactWriting.Texto(Function() ""),
                        "vazio de verdade tem de continuar vazio, e nao virar Nothing")
    End Sub

    Private Shared Function Contato(nome As String, email As String) As ContactInfo
        Return New ContactInfo With {.Nome = nome, .Email = email}
    End Function

    ''' <summary>
    ''' O fonte do <c>ContactWriting</c> <b>sem os comentários</b>.
    '''
    ''' Pela mesma razão da Fase 5: o comentário do módulo menciona
    ''' <c>ForwardAsVcard</c> justamente para explicar por que ele não pode
    ''' aparecer. Uma varredura que lê o comentário acusa o texto que protege.
    ''' </summary>
    Private Shared Function FonteSemComentarios() As String
        Dim d = New DirectoryInfo(AppContext.BaseDirectory)
        While d IsNot Nothing AndAlso Not File.Exists(Path.Combine(d.FullName, "Iris.slnx"))
            d = d.Parent
        End While
        Assert.IsNotNull(d, "não achei a raiz do repositório")

        Dim caminho = Path.Combine(d.FullName, "src", "Iris.Outlook", "ContactWriting.vb")
        Assert.IsTrue(File.Exists(caminho), "ContactWriting.vb não encontrado em " & caminho)

        Dim linhas = File.ReadAllLines(caminho).
                     Where(Function(l) Not l.TrimStart().StartsWith("'"))
        Return String.Join(vbLf, linhas)
    End Function

End Class

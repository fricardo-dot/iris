Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports Iris.Assist
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>Todo egress de IA mora num assembly só.</b>
'''
''' ------------------------------------------------------------------
''' <b>POR QUE A REGRA MUDOU</b>
'''
''' O plano dizia, na primeira versão, que haveria um teste sobre o IL do
''' <c>Iris.App</c> provando que ele não instancia provedor de rede. O Codex
''' derrubou: isso prova que <b>uma</b> camada não chama, e não prova que
''' nenhuma outra abre rede. Um adaptador esquecido no <c>Iris.Integration</c>
''' passaria.
'''
''' A regra que vale é sobre a <b>capacidade</b>, não sobre o uso: só um
''' assembly pode falar HTTP, e é o que existe para isso.
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTE TESTE PROVA, E O QUE NÃO PROVA</b>
'''
''' Prova que os assemblies de domínio <b>não referenciam</b> a biblioteca de
''' HTTP — então não há como um deles abrir conexão sem que a referência
''' apareça, e ela apareceria aqui.
'''
''' Não prova ausência de qualquer egress concebível: sockets crus, um
''' processo externo, um COM que busque URL. Essas seriam outras portas, e
''' fechá-las exigiria outras provas. O que está fechado é a porta que a Fase 3
''' abre.
''' </summary>
<TestClass>
Public Class EgressArquiteturaTests

    ''' <summary>As bibliotecas que dão acesso à rede.</summary>
    Private Shared ReadOnly DeRede As String() = {"System.Net.Http", "System.Net.Sockets",
                                                  "System.Net.Requests", "System.Net.WebClient"}

    ''' <summary>
    ''' Todos os assemblies de <b>produção</b> — <b>descobertos</b>, não listados.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE NÃO É UMA LISTA</b>
    '''
    ''' Era. E uma lista escrita à mão prova o que está nela: um projeto novo, ou
    ''' um que alguém esqueceu de acrescentar, passaria calado — que é exatamente
    ''' o caso em que a regra importa.
    '''
    ''' Aqui a varredura é sobre os arquivos <c>Iris.*.dll</c> ao lado do teste.
    ''' Ficam de fora só os que são <b>aparato</b>: este assembly e o harness de
    ''' crash.
    ''' </summary>
    Private Shared Function Producao() As IEnumerable(Of Assembly)
        Dim aparato = {"Iris.Tests", "Iris.CrashHarness"}

        ' "Iris*.dll", e nao "Iris.*.dll": o Iris.App produz Iris.dll, e o
        ' padrao com ponto o deixava de fora — justamente a camada que compoe
        ' tudo e que o teste anterior AFIRMAVA cobrir.
        Return IO.Directory.GetFiles(AppContext.BaseDirectory, "Iris*.dll").
               Select(Function(f) IO.Path.GetFileNameWithoutExtension(f)).
               Where(Function(n) Not aparato.Contains(n, StringComparer.OrdinalIgnoreCase)).
               Select(Function(n) Assembly.Load(n)).
               ToList()
    End Function

    Private Shared Function ReferenciasDeRede(a As Assembly) As IEnumerable(Of String)
        Return a.GetReferencedAssemblies().
                 Select(Function(r) r.Name).
                 Where(Function(n) DeRede.Contains(n, StringComparer.OrdinalIgnoreCase))
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>Exatamente DOIS assemblies de produção falam HTTP, e são estes.</b>
    '''
    ''' Não "nenhum além do esperado": a lista, pelo nome. Zero significaria que
    ''' a busca está procurando a coisa errada — nome trocado, biblioteca que o
    ''' .NET moderno não usa mais — e passaria em qualquer base.
    '''
    ''' Inclui o <c>Iris.Assist</c> entre os que não podem: definir o contrato de
    ''' "chamar o modelo" não pode dar a ele a capacidade de chamar.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>ERA UM. POR QUE PASSOU A SER DOIS</b>
    '''
    ''' Verificar atualizações é rede. Havia três saídas, e duas eram piores:
    '''
    ''' <list type="number">
    ''' <item>pôr o <c>HttpClient</c> da atualização dentro do assembly de egress
    '''       de IA — que juntaria numa caixa só o canal que <i>manda</i>
    '''       conteúdo do dono e o que só <i>busca</i> arquivo público, e
    '''       apagaria a distinção que faz o portão de divulgação existir;</item>
    ''' <item>escondê-lo em qualquer outro projeto — o que faria este teste
    '''       reprovar, e a tentação seguinte seria afrouxar o teste;</item>
    ''' <item>declarar o segundo, aqui, com o motivo. É o que está feito.</item>
    ''' </list>
    '''
    ''' A diferença entre os dois, que é o que autoriza a exceção: o de IA
    ''' <b>manda</b> conteúdo da caixa e por isso tem portão, cofre e diário; o
    ''' de atualização não carrega nada de dentro — ele só busca, e o que ele
    ''' tem no lugar é <b>assinatura</b>, porque o risco dele é o que chega.
    '''
    ''' Se um terceiro aparecer, é aqui que se decide se ele entra.
    ''' </summary>
    <TestMethod>
    Public Sub EXATAMENTE_DOIS_assemblies_de_producao_falam_HTTP()
        Dim comRede = Producao().
            Where(Function(a) ReferenciasDeRede(a).Any()).
            Select(Function(a) a.GetName().Name).
            OrderBy(Function(n) n).
            ToList()

        CollectionAssert.AreEqual({"Iris.Integration.Assist.Http", "Iris.Update"}, comRede,
            "quem abre socket na producao e lista fechada — achei: " &
            String.Join(", ", comRede))
    End Sub

    ''' <summary>
    ''' E a varredura <b>acha</b> os assemblies: uma busca que voltasse vazia
    ''' passaria no teste de cima por não olhar nada.
    ''' </summary>
    <TestMethod>
    Public Sub A_varredura_acha_os_assemblies_de_producao()
        Dim nomes = Producao().Select(Function(a) a.GetName().Name).ToList()

        Assert.IsTrue(nomes.Count >= 8, $"achei so {nomes.Count}: " & String.Join(", ", nomes))
        ' 'Iris' e o assembly do Iris.App — o nome do assembly nao e o nome do
        ' projeto, e foi por isso que ele ficou de fora da primeira varredura.
        For Each esperado In {"Iris.Model", "Iris.Core", "Iris.Assist",
                              "Iris.Integration.Assist.Http", "Iris.Update", "Iris"}
            CollectionAssert.Contains(nomes, esperado)
        Next
    End Sub

    ''' <summary>
    ''' <b>E o construtor que aceita <c>http://</c> não é público.</b>
    '''
    ''' Ele era, com padrão <c>False</c>, e havia teste provando o padrão — o que
    ''' prova o padrão e não impede a produção de passar <c>True</c>. A superfície
    ''' pública tem de ser <b>incapaz</b> de aceitar HTTP.
    ''' </summary>
    <TestMethod>
    Public Sub Nenhum_construtor_PUBLICO_aceita_http()
        Dim publicos = GetType(Integration.Assist.Http.HttpAssistantProvider).
            GetConstructors(BindingFlags.Public Or BindingFlags.Instance)

        For Each c In publicos
            Assert.IsFalse(c.GetParameters().Any(Function(p) p.ParameterType Is GetType(Boolean)),
                "construtor publico com booleano e por onde a excecao de loopback volta")
        Next

        ' E o Friend existe — senao os testes de transporte nao rodariam, e este
        ' teste estaria protegendo uma porta que ninguem usa.
        Assert.IsTrue(GetType(Integration.Assist.Http.HttpAssistantProvider).
                      GetConstructors(BindingFlags.NonPublic Or BindingFlags.Instance).Any(),
                      "o caminho de teste tem de existir, so nao pode ser publico")
    End Sub

    ''' <summary>
    ''' <b>Só o composition root depende do assembly de egress.</b>
    '''
    ''' Depender dele é ganhar a capacidade de egress de segunda mão, e a regra
    ''' é sobre <b>capacidade</b>, não sobre uso.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>A EXCEÇÃO, E POR QUE ELA É UMA SÓ</b>
    '''
    ''' A regra era "ninguém". Ela mudou quando a IA foi ativada, porque alguém
    ''' tem de escolher o provedor e construí-lo — e esse alguém é o
    ''' <c>Iris.App</c>, onde a composição acontece e onde a decisão do usuário
    ''' vira objeto.
    '''
    ''' A exceção está escrita <b>aqui</b> e não em lugar nenhum, de propósito:
    ''' uma regra que passa a admitir exceções sem que a lista delas fique num
    ''' lugar visível é uma regra que perde a primeira e depois todas. Se um
    ''' segundo projeto aparecer nesta lista, este teste é que decide se ele
    ''' entra.
    ''' </summary>
    <TestMethod>
    Public Sub SO_o_composition_root_depende_do_assembly_de_egress()
        SoOCompositionRootDepende("Iris.Integration.Assist.Http",
                                  "nao ha provedor de IA a montar")
    End Sub

    ''' <summary>
    ''' <b>E a mesma regra para o assembly de atualização.</b>
    '''
    ''' Não é a mesma capacidade — ele busca em vez de mandar — mas é a mesma
    ''' razão para prendê-lo no composition root: quem o referencia ganha um
    ''' <c>HttpClient</c> de segunda mão, e a lista de quem pode abrir socket
    ''' deixa de ser lista.
    ''' </summary>
    <TestMethod>
    Public Sub SO_o_composition_root_depende_do_assembly_de_atualizacao()
        SoOCompositionRootDepende("Iris.Update",
                                  "nao ha verificacao de versao a montar")
    End Sub

    ''' <summary>
    ''' As duas metades da regra: <b>ninguém além do root depende</b>, e <b>o
    ''' root depende mesmo</b>.
    '''
    ''' A segunda metade não é zelo: sem ela, apagar a referência do
    ''' <c>Iris.App</c> deixaria o teste verde e o recurso desmontado — a
    ''' permissão viraria uma lista de nomes que não descreve nada.
    ''' </summary>
    Private Shared Sub SoOCompositionRootDepende(assembly As String,
                                                 seFaltar As String)
        ' "Iris", e nao "Iris.App": o projeto Iris.App produz um assembly
        ' chamado Iris. Essa diferenca ja quebrou um teste desta suite antes,
        ' que dizia cobrir o Iris.App e nao cobria nada.
        Dim permitido = New HashSet(Of String)(StringComparer.Ordinal) From {"Iris"}

        Dim dependentes = Producao().
            Where(Function(a) a.GetName().Name <> assembly).
            Where(Function(a) a.GetReferencedAssemblies().
                                Any(Function(r) r.Name = assembly)).
            Select(Function(a) a.GetName().Name).ToList()

        Dim culpados = dependentes.Where(Function(n) Not permitido.Contains(n)).ToList()
        Assert.AreEqual(0, culpados.Count,
            $"so o composition root pode adquirir {assembly} de segunda mao: " &
            String.Join(", ", culpados))

        CollectionAssert.Contains(dependentes, "Iris",
            $"se o composition root nao referencia {assembly}, {seFaltar}")
    End Sub

    ''' <summary>
    ''' <b>A porta não conhece o transporte.</b>
    '''
    ''' <c>IAssistantProvider</c> recebe <c>Byte()</c>, e não um DTO que alguém
    ''' serialize do outro lado — se recebesse, haveria uma segunda
    ''' serialização, e o que foi autorizado deixaria de ser o que sai.
    ''' </summary>
    <TestMethod>
    Public Sub A_porta_recebe_BYTES()
        Dim enviar = GetType(IAssistantProvider).GetMethod("Enviar")

        Assert.AreEqual(GetType(Byte()), enviar.GetParameters()(0).ParameterType,
            "receber outra coisa abriria uma segunda serializacao")
    End Sub

End Class

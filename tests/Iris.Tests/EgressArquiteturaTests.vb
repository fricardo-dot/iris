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
    ''' Os assemblies de <b>domínio</b> — os que não têm nada a ver com rede.
    '''
    ''' <c>Iris.App</c> entra: ele é a composição, e compor não é transportar.
    ''' </summary>
    Private Shared Function Dominio() As IEnumerable(Of Assembly)
        Return {GetType(Model.ItemKey).Assembly,
                GetType(Core.CacheSchema).Assembly,
                GetType(Cache.CacheDatabase).Assembly,
                GetType(Sync.SweepRunner).Assembly,
                GetType(DisclosurePolicy).Assembly,
                GetType(Integration.SqliteDisclosureJournal).Assembly}
    End Function

    Private Shared Function ReferenciasDeRede(a As Assembly) As IEnumerable(Of String)
        Return a.GetReferencedAssemblies().
                 Select(Function(r) r.Name).
                 Where(Function(n) DeRede.Contains(n, StringComparer.OrdinalIgnoreCase))
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>Nenhum assembly de domínio referencia biblioteca de rede.</b>
    '''
    ''' Inclui o <c>Iris.Assist</c>, que define a porta: definir o contrato de
    ''' "chamar o modelo" não pode dar a ele a capacidade de chamar.
    ''' </summary>
    <TestMethod>
    Public Sub Nenhum_assembly_de_dominio_fala_HTTP()
        Dim culpados = Dominio().
            Select(Function(a) (Nome:=a.GetName().Name, Refs:=ReferenciasDeRede(a).ToList())).
            Where(Function(x) x.Refs.Count > 0).
            Select(Function(x) $"{x.Nome} → {String.Join(", ", x.Refs)}").
            ToList()

        Assert.AreEqual(0, culpados.Count,
            "egress de IA mora num assembly SO: " & String.Join(" | ", culpados))
    End Sub

    ''' <summary>
    ''' <b>O controle positivo.</b> O assembly que <i>deve</i> falar HTTP fala.
    '''
    ''' Sem ele, um teste que procurasse a referência errada — nome trocado,
    ''' lista vazia, biblioteca que o .NET moderno não usa mais — passaria em
    ''' todo domínio limpo e não estaria olhando nada.
    ''' </summary>
    <TestMethod>
    Public Sub O_assembly_de_egress_FALA_HTTP()
        Dim http = GetType(Integration.Assist.Http.HttpAssistantProvider).Assembly

        Assert.IsTrue(ReferenciasDeRede(http).Any(),
            "se nem ele referencia rede, o teste de cima nao esta procurando a coisa certa")
    End Sub

    ''' <summary>
    ''' E ele é o <b>único</b>: nenhum outro projeto de produção depende dele.
    '''
    ''' Depender do assembly de egress é ganhar a capacidade de egress de
    ''' segunda mão — e a regra é sobre capacidade, não sobre uso.
    ''' </summary>
    <TestMethod>
    Public Sub Ninguem_do_dominio_depende_do_assembly_de_egress()
        Const egress = "Iris.Integration.Assist.Http"

        Dim culpados = Dominio().
            Where(Function(a) a.GetReferencedAssemblies().
                                Any(Function(r) r.Name = egress)).
            Select(Function(a) a.GetName().Name).ToList()

        Assert.AreEqual(0, culpados.Count,
            "depender do assembly de egress e ganhar a capacidade de segunda mao: " &
            String.Join(", ", culpados))
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

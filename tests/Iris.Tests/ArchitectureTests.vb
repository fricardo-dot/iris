Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Testes que IMPÕEM a fronteira do COM.
'''
''' O plano da Fase 1 afirmava, na v1, que o grafo de projetos "impõe" a
''' fronteira. Não impõe: referências de projeto fluem transitivamente, e um
''' DTO com uma propriedade <c>Object</c> carrega um RCW sem que o assembly
''' sequer conheça os tipos do Interop.
'''
''' Estes testes são o que transforma a intenção em garantia. Se algum deles
''' falhar, alguém abriu um caminho para o COM chegar onde não deveria.
''' </summary>
<TestClass>
Public Class ArchitectureTests

    Private Shared ReadOnly InteropMarkers As String() = {
        "Microsoft.Office.Interop", "Office", "Outlook"
    }

    Private Shared Function ModelAssembly() As Assembly
        Return GetType(Iris.Model.ItemKey).Assembly
    End Function

    Private Shared Function CoreAssembly() As Assembly
        Return GetType(Iris.Core.IOutlookBroker).Assembly
    End Function

    ' ===================================================================
    ' Referências
    ' ===================================================================

    <TestMethod>
    Public Sub Model_nao_referencia_interop()
        AssertSemInterop(ModelAssembly())
    End Sub

    <TestMethod>
    Public Sub Core_nao_referencia_interop()
        AssertSemInterop(CoreAssembly())
    End Sub

    ''' <summary>
    ''' Model e Core miram net10.0 SEM -windows. O interop do Outlook é
    ''' Windows-only, então este TFM torna a fronteira uma impossibilidade
    ''' de compilação, não uma convenção. Se alguém trocar o TFM, este teste
    ''' avisa antes de o problema aparecer.
    ''' </summary>
    <TestMethod>
    Public Sub Model_e_Core_nao_miram_windows()
        For Each asm In {ModelAssembly(), CoreAssembly()}
            Dim tfm = asm.GetCustomAttribute(Of Runtime.Versioning.TargetFrameworkAttribute)()
            Assert.IsNotNull(tfm, $"{asm.GetName().Name} sem TargetFramework.")
            StringAssert.Contains(tfm.FrameworkName, "v10.0",
                                  $"{asm.GetName().Name}: TFM inesperado.")
            Assert.IsFalse(tfm.FrameworkName.Contains("Windows", StringComparison.OrdinalIgnoreCase),
                           $"{asm.GetName().Name} passou a mirar Windows. A fronteira do COM " &
                           "dependia deste TFM; ela virou apenas convenção.")
        Next
    End Sub

    Private Shared Sub AssertSemInterop(assembly As Assembly)
        Dim ofensores = assembly.GetReferencedAssemblies().
            Where(Function(r) r.Name IsNot Nothing AndAlso
                              r.Name.StartsWith("Microsoft.Office", StringComparison.OrdinalIgnoreCase)).
            Select(Function(r) r.Name).ToList()

        Assert.AreEqual(0, ofensores.Count,
            $"{assembly.GetName().Name} referencia interop do Office: " &
            String.Join(", ", ofensores))
    End Sub

    ' ===================================================================
    ' Superfície pública
    ' ===================================================================

    ''' <summary>
    ''' Nenhum membro público de Model ou Core pode usar tipo do Interop.
    '''
    ''' Com EmbedInteropTypes, os tipos são EMBUTIDOS no assembly e não
    ''' aparecem na lista de referências — por isso a checagem olha os tipos
    ''' usados, não só o que o assembly referencia.
    ''' </summary>
    <TestMethod>
    Public Sub Membros_publicos_nao_usam_tipos_de_interop()
        For Each asm In {ModelAssembly(), CoreAssembly()}
            For Each t In asm.GetExportedTypes()
                For Each tipo In TiposUsadosNaSuperficie(t)
                    Dim ns = If(tipo.Namespace, "")
                    Dim ehInterop = ns.StartsWith("Microsoft.Office", StringComparison.OrdinalIgnoreCase) OrElse
                                    tipo.GetCustomAttributes().Any(
                                        Function(a) a.GetType().Name = "TypeIdentifierAttribute")
                    Assert.IsFalse(ehInterop,
                        $"{asm.GetName().Name}.{t.Name} expõe o tipo de interop {tipo.FullName}.")
                Next
            Next
        Next
    End Sub

    ''' <summary>
    ''' Nenhum DTO pode expor Object, delegate ou tipo genérico demais.
    '''
    ''' Este é o furo que a revisão apontou: um RCW é atribuível a Object.
    ''' Uma propriedade <c>As Object</c> deixaria o COM atravessar a
    ''' fronteira sem o Model conhecer um único tipo do Interop.
    ''' </summary>
    <TestMethod>
    Public Sub Model_nao_expoe_Object_nem_delegate()
        Dim problemas As New List(Of String)()

        For Each t In ModelAssembly().GetExportedTypes()
            If t.IsEnum Then Continue For

            For Each p In t.GetProperties(BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.Static)
                VerificarTipoProibido(p.PropertyType, $"{t.Name}.{p.Name}", problemas)
            Next
            For Each f In t.GetFields(BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.Static)
                VerificarTipoProibido(f.FieldType, $"{t.Name}.{f.Name}", problemas)
            Next
        Next

        Assert.AreEqual(0, problemas.Count,
            "DTOs com tipo proibido (Object aceita um RCW):" & Environment.NewLine &
            String.Join(Environment.NewLine, problemas))
    End Sub

    ''' <summary>
    ''' CONTROLE NEGATIVO. Um teste que nunca falhou não prova nada: se
    ''' VerificarTipoProibido tivesse um bug e nunca acusasse, o teste acima
    ''' passaria para sempre sem valer nada.
    '''
    ''' Aqui um tipo deliberadamente ruim é submetido ao mesmo verificador,
    ''' e ele PRECISA ser reprovado.
    ''' </summary>
    <TestMethod>
    Public Sub Verificador_reprova_tipo_deliberadamente_ruim()
        Dim problemas As New List(Of String)()

        For Each p In GetType(TipoRuimDePropósito).GetProperties()
            VerificarTipoProibido(p.PropertyType, $"controle.{p.Name}", problemas)
        Next

        Assert.AreEqual(3, problemas.Count,
            "O verificador não acusou os três membros proibidos do controle negativo; " &
            "os testes de fronteira não teriam valor: " & String.Join(" | ", problemas))
    End Sub

    ''' <summary>
    ''' Existe SÓ para o controle negativo. Nunca use como modelo — cada
    ''' membro aqui é uma porta por onde um RCW passaria.
    ''' </summary>
    Public Class TipoRuimDePropósito
        Public Property Qualquer As Object
        Public Property Callback As Action
        Public Property ListaSolta As List(Of Object)
    End Class

    Private Shared Sub VerificarTipoProibido(tipo As Type, onde As String, problemas As List(Of String))
        For Each t In Expandir(tipo)
            If t Is GetType(Object) Then
                problemas.Add($"  {onde} é Object")
            ElseIf GetType([Delegate]).IsAssignableFrom(t) Then
                problemas.Add($"  {onde} é delegate ({t.Name})")
            End If
        Next
    End Sub

    ''' <summary>Tipo mais seus argumentos genéricos, recursivamente.</summary>
    Private Shared Iterator Function Expandir(tipo As Type) As IEnumerable(Of Type)
        If tipo Is Nothing Then Return
        Yield tipo
        If tipo.IsArray Then
            For Each t In Expandir(tipo.GetElementType())
                Yield t
            Next
        End If
        If tipo.IsGenericType Then
            For Each arg In tipo.GetGenericArguments()
                For Each t In Expandir(arg)
                    Yield t
                Next
            Next
        End If
    End Function

    ''' <summary>
    ''' Tipos que aparecem na superfície pública: propriedades, campos,
    ''' parâmetros e retornos.
    ''' </summary>
    Private Shared Iterator Function TiposUsadosNaSuperficie(t As Type) As IEnumerable(Of Type)
        For Each p In t.GetProperties()
            For Each x In Expandir(p.PropertyType) : Yield x : Next
        Next
        For Each f In t.GetFields()
            For Each x In Expandir(f.FieldType) : Yield x : Next
        Next
        For Each m In t.GetMethods()
            If m.DeclaringType IsNot t Then Continue For
            For Each x In Expandir(m.ReturnType) : Yield x : Next
            For Each prm In m.GetParameters()
                For Each x In Expandir(prm.ParameterType) : Yield x : Next
            Next
        Next
    End Function

End Class

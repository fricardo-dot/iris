Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' A fronteira do lado da apresentação.
'''
''' <para>
''' Iris.App PRECISA referenciar Iris.Outlook, porque o composition root
''' monta a implementação concreta. Logo, uma checagem no assembly inteiro
''' não distingue o uso permitido do proibido — e a versão anterior destes
''' testes afirmava, num comentário, fiscalizar ViewModels e Views sem
''' fiscalizar nada disso. Comentário que promete o que o código não faz é
''' pior que comentário nenhum.
''' </para>
''' <para>
''' LIMITE CONHECIDO: isto inspeciona a superfície declarada — campos,
''' propriedades, parâmetros, retornos, construtores — incluindo membros
''' privados. NÃO inspeciona corpos de método, então um tipo do Outlook
''' usado apenas como variável local dentro de um método escapa. A garantia
''' completa exigiria separar as Views e ViewModels num projeto próprio que
''' não referencia Iris.Outlook.
''' </para>
''' </summary>
<TestClass>
Public Class PresentationBoundaryTests

    Private Const NamespaceApresentacao As String = "Iris.App"
    Private Const NamespaceProibido As String = "Iris.Outlook"

    Private Shared Function AppAssembly() As Assembly
        Return GetType(Iris.App.ViewModels.ConnectionViewModel).Assembly
    End Function

    ''' <summary>
    ''' Nenhum tipo de ViewModel pode declarar membro do assembly do
    ''' Outlook, nem tipo de interop. O composition root — a classe
    ''' Application — é a única exceção, e por isso não entra na varredura.
    ''' </summary>
    <TestMethod>
    Public Sub ViewModels_nao_declaram_tipos_do_Outlook()
        Dim problemas As New List(Of String)()

        For Each t In AppAssembly().GetTypes()
            Dim ns = If(t.Namespace, "")
            If Not ns.StartsWith(NamespaceApresentacao & ".ViewModels", StringComparison.Ordinal) Then Continue For

            For Each tipo In TiposDeclarados(t)
                If EhProibido(tipo) Then
                    problemas.Add($"  {t.Name} declara {tipo.FullName}")
                End If
            Next
        Next

        Assert.AreEqual(0, problemas.Count,
            "ViewModel tocando a camada do Outlook:" & Environment.NewLine &
            String.Join(Environment.NewLine, problemas))
    End Sub

    ''' <summary>
    ''' CONTROLE NEGATIVO do detector: um tipo que declara justamente o que
    ''' é proibido precisa ser reprovado. Sem isto, um detector quebrado
    ''' aprovaria tudo para sempre.
    ''' </summary>
    <TestMethod>
    Public Sub Detector_reprova_tipo_que_declara_o_broker_concreto()
        Dim encontrados = TiposDeclarados(GetType(ViewModelProibidoDePropósito)).
                          Where(AddressOf EhProibido).ToList()

        Assert.IsTrue(encontrados.Count > 0,
            "O detector não reconheceu um campo do tipo concreto do Outlook; " &
            "o teste de fronteira da apresentação não teria valor.")
    End Sub

    ''' <summary>Existe só para o controle negativo. Nunca use como modelo.</summary>
    Public Class ViewModelProibidoDePropósito
        Private _broker As Iris.Outlook.OutlookBroker
    End Class

    Private Shared Function EhProibido(tipo As Type) As Boolean
        If tipo Is Nothing Then Return False
        Dim ns = If(tipo.Namespace, "")
        If ns.StartsWith(NamespaceProibido, StringComparison.Ordinal) Then Return True
        If ns.StartsWith("Microsoft.Office", StringComparison.OrdinalIgnoreCase) Then Return True
        Return tipo.GetCustomAttributes().Any(Function(a) a.GetType().Name = "TypeIdentifierAttribute")
    End Function

    ''' <summary>
    ''' Tudo que o tipo DECLARA, inclusive privado: campos, propriedades,
    ''' parâmetros e retornos de métodos e construtores.
    ''' </summary>
    Private Shared Iterator Function TiposDeclarados(t As Type) As IEnumerable(Of Type)
        Const Todos As BindingFlags =
            BindingFlags.Public Or BindingFlags.NonPublic Or
            BindingFlags.Instance Or BindingFlags.Static Or BindingFlags.DeclaredOnly

        For Each f In t.GetFields(Todos)
            Yield f.FieldType
        Next
        For Each p In t.GetProperties(Todos)
            Yield p.PropertyType
        Next
        For Each m In t.GetMethods(Todos)
            Yield m.ReturnType
            For Each prm In m.GetParameters()
                Yield prm.ParameterType
            Next
        Next
        For Each c In t.GetConstructors(Todos)
            For Each prm In c.GetParameters()
                Yield prm.ParameterType
            Next
        Next
    End Function

End Class

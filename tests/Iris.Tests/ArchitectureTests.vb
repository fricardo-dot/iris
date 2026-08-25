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

    ''' <summary>
    ''' O ciclo de sincronizacao. Ele recebe IOutlookBroker por INTERFACE e
    ''' nunca ve COM — o TFM sem -windows e o que garante, e este teste e o
    ''' que impede alguem acrescentar a referencia sem perceber.
    ''' </summary>
    Private Shared Function SyncAssembly() As Assembly
        Return GetType(Iris.Sync.SweepModel).Assembly
    End Function

    ' ===================================================================
    ' Fase da operação (F1-N)
    ' ===================================================================

    ''' <summary>
    ''' A fase de uma operação NÃO pode ser estado compartilhado do broker.
    '''
    ''' Isto não prova comportamento sob concorrência — para isso seria
    ''' preciso Outlook real e duas operações sobrepostas de verdade. Prova a
    ''' propriedade ESTRUTURAL que garante o comportamento: se cada
    ''' invocação instancia a sua fase e não existe campo guardando essa
    ''' fase, não há o que compartilhar.
    '''
    ''' Existe porque o defeito original era exatamente esse campo. Ele foi
    ''' removido, e alguém "simplificando" um dia pode trazê-lo de volta sem
    ''' perceber o que está desfazendo: uma operação concorrente zerando a
    ''' fase entre a falha de um Send e a classificação dela faz o envio que
    ''' talvez tenha saído voltar como retentável.
    ''' </summary>
    <TestMethod>
    Public Sub O_broker_nao_guarda_fase_de_operacao_em_campo()
        Dim broker = GetType(Iris.Outlook.OutlookBroker)

        Dim fase = broker.GetNestedTypes(BindingFlags.NonPublic).
                          FirstOrDefault(Function(t) t.Name.Contains("Fase"))

        Assert.IsNotNull(fase,
            "A classe de fase por invocação sumiu — a proteção pode ter voltado a ser campo.")

        Dim campos = broker.GetFields(BindingFlags.Instance Or BindingFlags.Static Or
                                      BindingFlags.NonPublic Or BindingFlags.Public).
                            Where(Function(f) f.FieldType Is fase).
                            Select(Function(f) f.Name).
                            ToList()

        Assert.AreEqual(0, campos.Count,
            "A fase da operação virou campo do broker de novo: " & String.Join(", ", campos) &
            ". Operações concorrentes passam a compartilhar a fase, e a defesa " &
            "contra reenviar uma mensagem que talvez tenha saido deixa de valer.")

        ' O campo removido chamava-se _effectStarted e era Integer. Procurar
        ' so por campo do TIPO da fase deixaria alguem reintroduzi-lo com
        ' outro tipo — que foi exatamente a forma original do defeito.
        Dim suspeitos = broker.GetFields(BindingFlags.Instance Or BindingFlags.Static Or
                                         BindingFlags.NonPublic Or BindingFlags.Public).
                               Where(Function(f) f.Name.IndexOf("effect", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                                 f.Name.IndexOf("mutation", StringComparison.OrdinalIgnoreCase) >= 0).
                               Select(Function(f) f.Name).
                               ToList()

        Assert.AreEqual(0, suspeitos.Count,
            "Campo com cara de fase de mutacao voltou ao broker: " & String.Join(", ", suspeitos))
    End Sub

    ''' <summary>
    ''' A classe de fase continua instanciavel e com a superficie que a
    ''' classificacao consome.
    '''
    ''' O que este teste NAO prova, e o nome anterior prometia: que RunAsync
    ''' chama o construtor, chama Marcar, ou le Iniciou. Reflexao sobre
    ''' metadados nao alcanca corpo de metodo. E um arame de tropeco
    ''' estrutural — pega quem apagar a classe ou renomear os membros — e nao
    ''' substitui o teste de concorrencia, que exigiria Outlook real e duas
    ''' operacoes sobrepostas de verdade.
    ''' </summary>
    <TestMethod>
    Public Sub A_classe_de_fase_continua_instanciavel_e_completa()
        Dim broker = GetType(Iris.Outlook.OutlookBroker)

        Dim fase = broker.GetNestedTypes(BindingFlags.NonPublic).
                          FirstOrDefault(Function(t) t.Name.Contains("Fase"))
        Assert.IsNotNull(fase)

        ' O construtor da fase precisa ser alcancavel: se ninguem instancia,
        ' a protecao nao existe, por mais que a classe esteja la.
        Dim ctor = fase.GetConstructors(BindingFlags.Instance Or BindingFlags.Public Or
                                        BindingFlags.NonPublic)
        Assert.IsTrue(ctor.Length > 0, "A classe de fase nao pode ser instanciada.")

        ' E ela precisa expor o que a classificacao consome.
        Dim marcar = fase.GetMethod("Marcar", BindingFlags.Instance Or BindingFlags.Public Or
                                              BindingFlags.NonPublic)
        Dim iniciou = fase.GetProperty("Iniciou", BindingFlags.Instance Or BindingFlags.Public Or
                                                  BindingFlags.NonPublic)

        Assert.IsNotNull(marcar, "Sem Marcar, a fase nunca e sinalizada.")
        Assert.IsNotNull(iniciou, "Sem Iniciou, a classificacao nao tem o que ler.")
    End Sub

    ''' <summary>
    ''' Controle negativo do teste acima: ele SABE achar campo do tipo que
    ''' procura. Sem esta conferência, um teste que nunca encontra nada
    ''' passaria para sempre, inclusive depois de o campo voltar.
    ''' </summary>
    <TestMethod>
    Public Sub A_busca_por_campo_de_fase_realmente_encontra_campos()
        Dim exemplo = GetType(ClasseComCampoDeFase)

        Dim achados = exemplo.GetFields(BindingFlags.Instance Or BindingFlags.NonPublic).
                              Where(Function(f) f.FieldType Is GetType(FaseDeMentira)).
                              Count()

        Assert.AreEqual(1, achados,
            "A técnica de reflexão do teste anterior não acha nada — ele passaria de graça.")
    End Sub

    Private NotInheritable Class FaseDeMentira
    End Class

    Private NotInheritable Class ClasseComCampoDeFase
        Private ReadOnly _fase As New FaseDeMentira()

        Public ReadOnly Property Fase As FaseDeMentira
            Get
                Return _fase
            End Get
        End Property
    End Class

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
        AssertSemInterop(SyncAssembly())
    End Sub

    ''' <summary>
    ''' Model e Core miram net10.0 SEM -windows. O interop do Outlook é
    ''' Windows-only, então este TFM torna a fronteira uma impossibilidade
    ''' de compilação, não uma convenção. Se alguém trocar o TFM, este teste
    ''' avisa antes de o problema aparecer.
    ''' </summary>
    <TestMethod>
    Public Sub Model_e_Core_nao_miram_windows()
        For Each asm In {ModelAssembly(), CoreAssembly(), SyncAssembly()}
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
        For Each asm In {ModelAssembly(), CoreAssembly(), SyncAssembly()}
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


    ' ==================================================================
    ' §26.2 — a UI nao pode contornar o dreno

    ''' <summary>
    ''' <b>A proibição da §26.2, executável.</b>
    '''
    ''' <blockquote>A UI não pode contornar o <c>PublicationDrain</c> usando
    ''' polling ou leitura direta como substituto silencioso da dívida
    ''' registrada.</blockquote>
    '''
    ''' Sem este teste a proibição seria um comentário, e este projeto já viu a
    ''' regra do <c>Permission</c> ficar escrita por três marcos enquanto o gate
    ''' falhava aberto.
    '''
    ''' O jeito de contornar é concreto e tentador: chamar
    ''' <c>ManifestReader</c> direto num timer. Funciona, parece igual, e a
    ''' dívida deixa de ser drenada — a fila enche, o <c>TravadoEm</c> nunca é
    ''' consultado, e ninguém descobre que a entrega parou.
    ''' </summary>
    <TestMethod>
    Public Sub A_UI_nao_instancia_ManifestReader_direto()
        Dim app = GetType(Iris.App.ViewModels.MainViewModel).Assembly
        Dim leitor = GetType(Iris.Integration.ManifestReader)

        Dim culpados =
            app.GetTypes().
                SelectMany(Function(t) t.GetMethods(Reflection.BindingFlags.Public Or
                                                    Reflection.BindingFlags.NonPublic Or
                                                    Reflection.BindingFlags.Instance Or
                                                    Reflection.BindingFlags.Static Or
                                                    Reflection.BindingFlags.DeclaredOnly)).
                Where(Function(m) UsaOTipo(m, leitor)).
                Select(Function(m) $"{m.DeclaringType.Name}.{m.Name}").
                ToList()

        Assert.AreEqual(0, culpados.Count,
            "a UI tem de receber o manifesto pelo AcervoService, que o dreno atualiza. " &
            "Ler direto e o contorno que a §26.2 proibe: " & String.Join(", ", culpados))
    End Sub

    ''' <summary>
    ''' Controle: a busca REALMENTE encontra uso do tipo.
    '''
    ''' Sem isto, um <c>UsaOTipo</c> que devolvesse sempre falso faria o teste
    ''' acima passar para sempre — e a proibição valeria zero.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_a_busca_por_uso_de_tipo_encontra_de_verdade()
        Dim eu = GetType(ArchitectureTests)
        Dim leitor = GetType(Iris.Integration.ManifestReader)
        Dim m = eu.GetMethod(NameOf(IscaQueUsaOManifestReader),
                             Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static)
        Assert.IsTrue(UsaOTipo(m, leitor),
            "a busca nao encontra nem um uso plantado de proposito — ela nao busca nada")
    End Sub

    Private Shared Function IscaQueUsaOManifestReader(db As Iris.Cache.CacheDatabase) As Object
        Return New Iris.Integration.ManifestReader(db)
    End Function

    ''' <summary>
    ''' Procura, no corpo compilado do método, referência ao tipo — construtor
    ''' ou chamada de membro.
    ''' </summary>
    Private Shared Function UsaOTipo(m As Reflection.MethodBase, alvo As Type) As Boolean
        Dim corpo = m.GetMethodBody()
        If corpo Is Nothing Then Return False

        Dim il = corpo.GetILAsByteArray()
        If il Is Nothing Then Return False

        Dim mod_ = m.Module
        Dim i = 0
        While i < il.Length - 4
            ' newobj (0x73) e call (0x28) / callvirt (0x6F) carregam um token
            ' de metodo nos quatro bytes seguintes.
            Dim op = il(i)
            If op = &H73 OrElse op = &H28 OrElse op = &H6F Then
                Dim token = BitConverter.ToInt32(il, i + 1)
                Try
                    Dim alvoMetodo = mod_.ResolveMethod(token,
                        If(m.DeclaringType?.GetGenericArguments(), Type.EmptyTypes),
                        If(TryCast(m, Reflection.MethodInfo)?.GetGenericArguments(), Type.EmptyTypes))
                    If alvoMetodo IsNot Nothing AndAlso alvoMetodo.DeclaringType Is alvo Then
                        Return True
                    End If
                Catch
                End Try
                i += 5
            Else
                i += 1
            End If
        End While
        Return False
    End Function

End Class

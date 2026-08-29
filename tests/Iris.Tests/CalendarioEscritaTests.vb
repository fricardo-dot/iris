Imports System.Linq
Imports Iris.Core
Imports Iris.Model
Imports Iris.Outlook
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' <b>ESCRITA NO CALENDÁRIO — e a invariante que a governa.</b>
'''
''' ------------------------------------------------------------------
''' <b>A REGRA QUE ESTES TESTES EXISTEM PARA PROTEGER</b>
'''
''' Salvar um <c>AppointmentItem</c> que seja <b>reunião</b> manda e-mail: o
''' Outlook envia convite ao criar, atualização ao editar e cancelamento ao
''' apagar. Isso não é efeito colateral que dê para desligar — é o
''' comportamento normal do objeto.
'''
''' O Iris tem uma regra sem exceção: <b>nada sai por e-mail sem o usuário
''' mandar</b>. A Fase 6 entregou a leitura em 28/08 justamente por isso, e a
''' escrita ficou esperando um desenho que não pudesse violar a regra por
''' descuido.
'''
''' <b>O desenho que passou:</b> a invariante é sustentada pelo <i>tipo</i>, e
''' não por alguém lembrar. <see cref="AppointmentDraft"/> não tem campo de
''' participante — então não existe caminho para o Iris criar uma reunião. E
''' editar ou apagar confere <c>MeetingStatus</c> <b>antes</b> de tocar no
''' item, recusando com <see cref="ErrorKind.Denied"/>.
'''
''' ------------------------------------------------------------------
''' <b>O QUE ESTES TESTES ALCANÇAM, E O QUE NÃO</b>
'''
''' <c>EhReuniao</c> precisa de um <c>AppointmentItem</c> real, então ele
''' <b>não</b> tem teste aqui — está declarado com esse nome, como o
''' <c>ContarAnexos</c> da paginação. O que dá para provar sem COM é a guarda
''' que decide se o Iris chega a tocar no Outlook, e é ela que está abaixo.
''' </summary>
<TestClass>
Public Class CalendarioEscritaTests

    Private Shared Function Rascunho(Optional assunto As String = "Reunião interna",
                                     Optional minutos As Integer = 60) As AppointmentDraft
        Dim inicio = New DateTimeOffset(2026, 9, 1, 14, 0, 0, TimeSpan.FromHours(-3))
        Return New AppointmentDraft With {
            .Subject = assunto,
            .De = inicio,
            .Ate = inicio.AddMinutes(minutos)
        }
    End Function

    ' ==================================================================

    ''' <summary>
    ''' <b>CONTROLE POSITIVO: um rascunho comum é aceito.</b>
    '''
    ''' Sem ele, uma guarda que recusasse tudo passaria em todos os testes
    ''' abaixo — que é o bloqueio sem controle negativo que o CLAUDE.md
    ''' descreve, e que esta sessão já cometeu quatro vezes.
    ''' </summary>
    <TestMethod>
    Public Sub Controle_um_rascunho_comum_e_aceito()
        Assert.IsNull(CalendarWriting.RecusarRascunho(Rascunho()),
                      "recusou um compromisso perfeitamente comum")
    End Sub

    ''' <summary>
    ''' <b>Compromisso sem assunto não entra.</b>
    '''
    ''' Um bloco anônimo na agenda é pior que nenhum bloco: quem olha depois
    ''' não sabe o que é, e o Iris não tem como explicar — ele não estava lá
    ''' quando aquilo foi criado.
    ''' </summary>
    <DataTestMethod>
    <DataRow("")>
    <DataRow("   ")>
    <DataRow(vbTab)>
    Public Sub Sem_assunto_RECUSA(assunto As String)
        Dim motivo = CalendarWriting.RecusarRascunho(Rascunho(assunto))

        Assert.IsNotNull(motivo, "aceitou compromisso sem assunto")
        StringAssert.Contains(motivo, "assunto")
    End Sub

    ''' <summary>
    ''' <b>Fim antes do início não entra — e o sintoma apareceria longe.</b>
    '''
    ''' O OOM <b>aceita</b> um compromisso invertido e a agenda mostra coisa
    ''' impossível. Pior: o <c>Restrict</c> da leitura passa a não achar o
    ''' item, então o defeito aparece como "sumiu um compromisso", longe de
    ''' quem o criou.
    ''' </summary>
    <TestMethod>
    Public Sub Fim_antes_do_inicio_RECUSA()
        Dim invertido = Rascunho(minutos:=-30)
        Dim motivo = CalendarWriting.RecusarRascunho(invertido)

        Assert.IsNotNull(motivo, "aceitou compromisso que termina antes de começar")
        StringAssert.Contains(motivo, "anterior ao início")
    End Sub

    ''' <summary>
    ''' <b>Duração zero é legítima.</b>
    '''
    ''' Um marcador de instante — "ligar para o fulano às 14h" — é compromisso
    ''' de verdade. A guarda recusa o <i>invertido</i>, e não o instantâneo:
    ''' sem esta linha, endurecer a comparação para <c>&lt;=</c> passaria
    ''' despercebido.
    ''' </summary>
    <TestMethod>
    Public Sub Duracao_zero_e_aceita()
        Assert.IsNull(CalendarWriting.RecusarRascunho(Rascunho(minutos:=0)),
                      "recusou um marcador de instante")
    End Sub

    ''' <summary>Rascunho nulo não explode: vira recusa.</summary>
    <TestMethod>
    Public Sub Rascunho_nulo_RECUSA()
        Assert.IsNotNull(CalendarWriting.RecusarRascunho(Nothing))
    End Sub

    ''' <summary>
    ''' <b>O TIPO NÃO TEM ONDE PÔR PARTICIPANTE — e é essa a garantia.</b>
    '''
    ''' Este teste não exercita comportamento: ele prende o <i>desenho</i>. A
    ''' invariante "o Iris não cria reunião" está sustentada por não existir
    ''' campo para preencher, e não por alguém lembrar de não preencher.
    '''
    ''' Se um dia alguém acrescentar <c>Recipients</c>, <c>Attendees</c> ou
    ''' <c>RequiredAttendees</c> ao rascunho, este teste cai — e quem o
    ''' acrescentar tem de ler o comentário do <c>CalendarWriting</c> antes de
    ''' apagá-lo.
    ''' </summary>
    <TestMethod>
    Public Sub O_rascunho_NAO_tem_campo_de_participante()
        Dim campos = GetType(AppointmentDraft).GetProperties().
                     Select(Function(p) p.Name.ToLowerInvariant()).ToList()

        For Each proibido In {"recipients", "attendees", "requiredattendees",
                              "optionalattendees", "resources", "to"}
            Assert.IsFalse(campos.Contains(proibido),
                $"AppointmentDraft ganhou '{proibido}': criar compromisso com " &
                "participante manda convite por e-mail, e o Iris não envia. " &
                "Leia o comentário do CalendarWriting antes de mexer nisto.")
        Next
    End Sub

End Class

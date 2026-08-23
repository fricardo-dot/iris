Imports System.IO

Namespace Global.Iris.App

    ''' <summary>
    ''' Perguntar ao usuário onde salvar.
    '''
    ''' É interface para que o ViewModel não conheça o diálogo do Windows —
    ''' senão ele deixaria de ser testável e passaria a exigir uma janela só
    ''' para verificar a lógica de salvar anexo.
    ''' </summary>
    Public Interface ISaveFileService
        ''' <returns>O caminho escolhido, ou Nothing se o usuário cancelou.</returns>
        Function AskWhereToSave(suggestedName As String) As String
    End Interface

    ''' <summary>
    ''' Perguntar ao usuário qual arquivo anexar.
    '''
    ''' Separada de <see cref="ISaveFileService"/> de propósito. São duas
    ''' direções opostas — uma tira dado de dentro da caixa, a outra põe
    ''' arquivo do disco numa mensagem que vai sair — e juntá-las numa
    ''' interface só faria qualquer teste do leitor ganhar, de graça, a
    ''' capacidade de anexar.
    ''' </summary>
    Public Interface IPickFileService
        ''' <returns>O caminho escolhido, ou Nothing se o usuário cancelou.</returns>
        Function AskWhichFileToAttach() As String
    End Interface

    Public NotInheritable Class WindowsPickFileService
        Implements IPickFileService

        Public Function AskWhichFileToAttach() As String _
            Implements IPickFileService.AskWhichFileToAttach

            Dim dialogo As New Microsoft.Win32.OpenFileDialog With {
                .Title = "Anexar arquivo",
                .CheckFileExists = True,
                .Multiselect = False,
                .Filter = "Todos os arquivos|*.*"
            }

            If dialogo.ShowDialog() = True Then Return dialogo.FileName
            Return Nothing
        End Function
    End Class

    Public NotInheritable Class WindowsSaveFileService
        Implements ISaveFileService

        Public Function AskWhereToSave(suggestedName As String) As String _
            Implements ISaveFileService.AskWhereToSave

            Dim seguro = Sanitizar(suggestedName)
            Dim extensao = Path.GetExtension(seguro)

            Dim dialogo As New Microsoft.Win32.SaveFileDialog With {
                .FileName = seguro,
                .Title = "Salvar anexo",
                .OverwritePrompt = True,
                .AddExtension = True,
                .Filter = If(String.IsNullOrEmpty(extensao),
                             "Todos os arquivos|*.*",
                             $"Arquivo {extensao}|*{extensao}|Todos os arquivos|*.*")
            }

            If dialogo.ShowDialog() = True Then Return dialogo.FileName
            Return Nothing
        End Function

        ''' <summary>
        ''' O nome vem de dentro de um e-mail, ou seja, de fora. Caracteres
        ''' inválidos e componentes de caminho não podem chegar ao diálogo:
        ''' um nome como "..\..\algo.exe" é entrada hostil, não nome de
        ''' arquivo.
        ''' </summary>
        Private Shared Function Sanitizar(nome As String) As String
            If String.IsNullOrWhiteSpace(nome) Then Return "anexo"

            ' Só o nome, nunca o caminho.
            nome = Path.GetFileName(nome)
            If String.IsNullOrWhiteSpace(nome) Then Return "anexo"

            For Each c In Path.GetInvalidFileNameChars()
                nome = nome.Replace(c, "_"c)
            Next

            Return If(String.IsNullOrWhiteSpace(nome), "anexo", nome)
        End Function
    End Class

End Namespace

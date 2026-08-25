Imports System.Globalization
Imports System.Windows
Imports System.Windows.Data

Namespace Global.Iris.App

    ''' <summary>
    ''' Booleano invertido para Visibility. Existe porque a mesma condição
    ''' precisa mostrar um elemento e esconder o outro, e duplicar a
    ''' propriedade no ViewModel só para inverter seria pior.
    ''' </summary>
    Public NotInheritable Class InverseBoolToVisibility
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object,
                                culture As CultureInfo) As Object Implements IValueConverter.Convert
            Dim verdadeiro = TypeOf value Is Boolean AndAlso CBool(value)
            Return If(verdadeiro, Visibility.Collapsed, Visibility.Visible)
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object,
                                    culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotSupportedException()
        End Function
    End Class

    ''' <summary>
    ''' Texto ausente ou vazio some; texto presente aparece.
    '''
    ''' Existe para a ressalva do acervo e para a falha de abertura do cache:
    ''' as duas só devem ocupar espaço quando têm o que dizer, e nenhuma pode
    ''' virar faixa vazia — faixa vazia é ruído que ensina o usuário a ignorar
    ''' aquele lugar da tela.
    ''' </summary>
    Public NotInheritable Class NullToVisibility
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object,
                                culture As CultureInfo) As Object Implements IValueConverter.Convert
            Dim texto = TryCast(value, String)
            If texto IsNot Nothing Then
                Return If(String.IsNullOrWhiteSpace(texto), Visibility.Collapsed, Visibility.Visible)
            End If
            Return If(value Is Nothing, Visibility.Collapsed, Visibility.Visible)
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object,
                                    culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotSupportedException()
        End Function
    End Class

End Namespace

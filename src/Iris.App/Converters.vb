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

End Namespace

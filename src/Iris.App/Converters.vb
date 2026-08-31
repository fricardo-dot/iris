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

    ''' <summary>
    ''' <b>Uma fração da altura da janela.</b>
    '''
    ''' A faixa da IA é linha <c>Auto</c> do MainWindow: ela cresce e a
    ''' lista de e-mails encolhe na mesma medida. Um teto em pixels fixos
    ''' seria certo numa janela e errado em todas as outras, e este
    ''' programa abre maximizado em telas que não são a minha.
    '''
    ''' O parâmetro é o fator, em texto: <c>ConverterParameter=0.42</c>.
    ''' Sem parâmetro legível, <b>metade</b>.
    '''
    ''' Altura não numérica, não finita ou zero devolve
    ''' <see cref="Double.PositiveInfinity"/> — que é o que "sem teto"
    ''' quer dizer para o WPF. Devolver zero esconderia o painel inteiro,
    ''' e um painel sumido por causa de uma medida ainda não feita é o
    ''' tipo de defeito que só aparece na máquina do outro.
    ''' </summary>
    Public NotInheritable Class FracaoDaAltura
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object,
                                culture As CultureInfo) As Object Implements IValueConverter.Convert
            Dim altura As Double
            If value Is Nothing OrElse Not Double.TryParse(
                    System.Convert.ToString(value, CultureInfo.InvariantCulture),
                    NumberStyles.Float, CultureInfo.InvariantCulture, altura) Then
                Return Double.PositiveInfinity
            End If
            If Double.IsNaN(altura) OrElse Double.IsInfinity(altura) OrElse altura <= 0 Then
                Return Double.PositiveInfinity
            End If

            Dim fator As Double = 0.5
            Dim lido As Double
            If parameter IsNot Nothing AndAlso Double.TryParse(
                    System.Convert.ToString(parameter, CultureInfo.InvariantCulture),
                    NumberStyles.Float, CultureInfo.InvariantCulture, lido) AndAlso
               lido > 0 AndAlso lido <= 1 Then
                fator = lido
            End If

            Return altura * fator
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object,
                                    culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotSupportedException()
        End Function
    End Class

End Namespace

'Useful generic tools
Namespace Toolkit

#Region " ComboBox/TickBox Item "
    'Custom content item for combo boxes or tickbox lists
    Public Class Item
        Property Name As String
        Property Value As Object

        Public Sub New(MyName As String, MyValue As Object)
            Name = MyName
            Value = MyValue
        End Sub

        Public Overrides Function ToString() As String
            'Generates the text shown in the control
            Return Name
        End Function
    End Class
#End Region

End Namespace
'Useful generic tools
Namespace Toolkit

#Region " ComboBox/TickBox Item "
    'Custom content item for combo boxes or tickbox lists
    Public Class Item
        Property Name As String
        Property Value As Object

        Public Sub New(MyName As String, MyValue As Object)
            If Not MyName Is Nothing Then Name = MyName
            If Not MyValue Is Nothing Then Value = MyValue
        End Sub

        Public Sub New(MyItem As Item) ' Cloning
            If Not MyItem Is Nothing Then
                Name = MyItem.Name
                Value = MyItem.Value
            End If
        End Sub

        Public Overrides Function ToString() As String
            'Generates the text shown in the control
            Return Name
        End Function
    End Class

    Public Class ReturnObject
        Property Success As Boolean
        Property ErrorMessage As String
        Property MyObject As Object

        Public Sub New(MySuccess As Boolean, MyErrorMessage As String, ReturnObject As Object)
            Success = MySuccess
            ErrorMessage = MyErrorMessage
            MyObject = ReturnObject
        End Sub

        Public Sub New(MySuccess As Boolean, MyErrorMessage As String)
            Success = MySuccess
            ErrorMessage = MyErrorMessage
            MyObject = Nothing
        End Sub
    End Class
#End Region

End Namespace
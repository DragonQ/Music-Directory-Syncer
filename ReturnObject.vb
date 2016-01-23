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
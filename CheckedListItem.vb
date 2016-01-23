Public Class CheckedListItem

    Property Data As Object
    Property MyName As String
    Property IsChecked As Boolean

    Public Sub New(NewData As Object, NewName As String, Checked As Boolean)
        Data = NewData
        MyName = NewName
        IsChecked = Checked
    End Sub

End Class

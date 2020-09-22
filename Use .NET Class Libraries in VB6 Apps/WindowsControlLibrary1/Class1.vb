Public Class Class1
    Private MyForm As Form1

    Public Sub ShowMessage()
        MsgBox("Hi!, Message From .Net", MsgBoxStyle.Information)
    End Sub

    Public Sub ShowForm()
        MyForm = New Form1
        MyForm.ShowDialog()
    End Sub

End Class

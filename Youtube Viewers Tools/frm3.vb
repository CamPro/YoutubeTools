Public Class frm3
    Private Sub Frm3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate(Form1.txtlink3.Text)
    End Sub
End Class
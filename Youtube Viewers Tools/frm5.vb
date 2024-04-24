Public Class frm5
    Private Sub Frm5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate(Form1.txtlink5.Text)
    End Sub
End Class
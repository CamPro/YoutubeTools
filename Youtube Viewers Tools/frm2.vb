Public Class frm2
    Private Sub Frm2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate(Form1.txtlink2.Text)
    End Sub
End Class
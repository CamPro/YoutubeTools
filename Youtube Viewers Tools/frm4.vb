Public Class frm4
    Private Sub Frm4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate(Form1.txtlink4.Text)
    End Sub
End Class
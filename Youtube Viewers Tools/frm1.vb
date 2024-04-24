Public Class frm1

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'WebBrowser1.Document.ParentWindow.scrollBy 0, 800
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

    End Sub

    Private Sub Frm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate(Form1.txtlink1.Text)
    End Sub
End Class
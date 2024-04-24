Imports System.IO
Imports System.Threading
Public Class Form1
    'OPENED SOURCE CODE BY NGUYỄN ĐẮC TÀI
    'THE DAYS OF OPEN SOURCE CODE : THỨ 2, NGÀY 27 THÁNG 5 NĂM 2019
    'GỬI LỜI CHÚC ĐẾN NGƯỜI TIẾP THEO PHÁT TRIỂN DỰ ÁN NÀY THÀNH CÔNG HƠN NHÉ !
    Private Declare Function UrlMkSetSessionOption Lib "urlmon.dll" (ByVal dwOption As Integer, ByVal pBuffer As String, ByVal dwBufferLength As Integer, ByVal dwReserved As Integer) As Integer

    Private Const URLMON_OPTION_USERAGENT As Integer = 268435457
    Public Sub ChangeUserAgent(ByVal Agent As String)
        UrlMkSetSessionOption(URLMON_OPTION_USERAGENT, Agent, Agent.Length, 0)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListBox1.Items.Clear()
        Dim openfile = New OpenFileDialog()
        openfile.Title = "CHỌN FILE CHỨA PROXY"
        openfile.Filter = "Text (*.txt)|*.txt"
        If (openfile.ShowDialog() = System.Windows.Forms.DialogResult.OK) Then
            Dim proxyfile As String = openfile.FileName
            Dim proxies As String() = File.ReadAllLines(proxyfile)
            For Each line As String In proxies
                ListBox1.Items.Add(line)
                lbproxy.Text = ListBox1.Items.Count
                ListBox1.SelectedIndex = 0
            Next
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lbug.Text = ListBox2.Items.Count
        btnmaxsize.Hide()
        ListBox2.SelectedIndex = 0

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox2.Items.Clear()
        Dim openfile = New OpenFileDialog()
        openfile.Title = "CHỌN FILE CHỨA USER AGENT"
        openfile.Filter = "Text (*.txt)|*.txt"
        If (openfile.ShowDialog() = System.Windows.Forms.DialogResult.OK) Then
            Dim ugfile As String = openfile.FileName
            Dim ug As String() = File.ReadAllLines(ugfile)
            For Each line As String In ug
                ListBox2.Items.Add(line)
                lbug.Text = ListBox2.Items.Count
                ListBox2.SelectedIndex = 0
            Next
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        tm_tuongtac.Start()
        Label12.Text = "runing..."
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Label6.Location = New Point(294, 6)
        tm_tuongtac.Stop()
        tm_nghi.Stop()
        Label12.Text = "stoped"
    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        If NumericUpDown1.Value = 1 Then
            'frm8.Close()
            'frm7.Close()
            'frm6.Close()
            'frm5.Close()
            'frm4.Close()
            'frm3.Close()
            'frm2.Close()
            'frm1.Show()
            clink1.Visible = True
            clink2.Visible = False
            clink3.Visible = False
            clink4.Visible = False
            clink5.Visible = False
            clink6.Visible = False
            clink7.Visible = False
            clink8.Visible = False
            Label10.Text = "1 luồng"
        ElseIf NumericUpDown1.Value = 2 Then
            'frm8.Close()
            'frm7.Close()
            'frm6.Close()
            'frm5.Close()
            'frm4.Close()
            'frm3.Close()
            'frm2.Show()
            'frm1.Show()
            clink1.Visible = True
            clink2.Visible = True
            clink3.Visible = False
            clink4.Visible = False
            clink5.Visible = False
            clink6.Visible = False
            clink7.Visible = False
            clink8.Visible = False
            Label10.Text = "2 luồng"
        ElseIf NumericUpDown1.Value = 3 Then
            'frm8.Close()
            'frm7.Close()
            'frm6.Close()
            'frm5.Close()
            'frm4.Close()
            'frm3.Show()
            'frm2.Show()
            'frm1.Show()
            clink1.Visible = True
            clink2.Visible = True
            clink3.Visible = True
            clink4.Visible = False
            clink5.Visible = False
            clink6.Visible = False
            clink7.Visible = False
            clink8.Visible = False
            Label10.Text = "3 luồng"
        ElseIf NumericUpDown1.Value = 4 Then
            'frm8.Close()
            'frm7.Close()
            'frm6.Close()
            'frm5.Close()
            'frm4.Show()
            'frm3.Show()
            'frm2.Show()
            'frm1.Show()
            clink1.Visible = True
            clink2.Visible = True
            clink3.Visible = True
            clink4.Visible = True
            clink5.Visible = False
            clink6.Visible = False
            clink7.Visible = False
            clink8.Visible = False
            Label10.Text = "4 luồng"
        ElseIf NumericUpDown1.Value = 5 Then
            'frm8.Close()
            'frm7.Close()
            'frm6.Close()
            'frm5.Show()
            'frm4.Show()
            'frm3.Show()
            'frm2.Show()
            'frm1.Show()
            clink1.Visible = True
            clink2.Visible = True
            clink3.Visible = True
            clink4.Visible = True
            clink5.Visible = True
            clink6.Visible = False
            clink7.Visible = False
            clink8.Visible = False
            Label10.Text = "5 luồng"
        ElseIf NumericUpDown1.Value = 6 Then
            'frm8.Close()
            'frm7.Close()
            'frm6.Show()
            'frm5.Show()
            'frm4.Show()
            'frm3.Show()
            'frm2.Show()
            'frm1.Show()
            clink1.Visible = True
            clink2.Visible = True
            clink3.Visible = True
            clink4.Visible = True
            clink5.Visible = True
            clink6.Visible = True
            clink7.Visible = False
            clink8.Visible = False
            Label10.Text = "6 luồng"
        ElseIf NumericUpDown1.Value = 7 Then
            'frm8.Close()
            'frm7.Show()
            'frm6.Show()
            'frm5.Show()
            'frm4.Show()
            'frm3.Show()
            'frm2.Show()
            'frm1.Show()
            clink1.Visible = True
            clink2.Visible = True
            clink3.Visible = True
            clink4.Visible = True
            clink5.Visible = True
            clink6.Visible = True
            clink7.Visible = True
            clink8.Visible = False
            Label10.Text = "7 luồng"
        ElseIf NumericUpDown1.Value = 8 Then
            'frm8.Show()
            'frm7.Show()
            'frm6.Show()
            'frm5.Show()
            'frm4.Show()
            'frm3.Show()
            'frm2.Show()
            'frm1.Show()
            clink1.Visible = True
            clink2.Visible = True
            clink3.Visible = True
            clink4.Visible = True
            clink5.Visible = True
            clink6.Visible = True
            clink7.Visible = True
            clink8.Visible = True
            Label10.Text = "8 luồng"
        End If
    End Sub

    Private Sub Btnmaxsize_Click(sender As Object, e As EventArgs) Handles btnmaxsize.Click
        Me.Width = 787
        Me.Height = 544
        btnmaxsize.Hide()
        btnthnhotool.Show()
    End Sub

    Private Sub Btnthnhotool_Click(sender As Object, e As EventArgs) Handles btnthnhotool.Click
        Me.Width = 399
        Me.Height = 544
        btnthnhotool.Hide()
        btnmaxsize.Show()
    End Sub

    Private Sub Btnlink1_Click(sender As Object, e As EventArgs) Handles btnlink1.Click
        frm1.Show()

    End Sub

    Private Sub Btnlink2_Click(sender As Object, e As EventArgs) Handles btnlink2.Click
        frm2.Show()
    End Sub

    Private Sub Btnlink3_Click(sender As Object, e As EventArgs) Handles btnlink3.Click
        frm3.Show()
    End Sub

    Private Sub Btnlink4_Click(sender As Object, e As EventArgs) Handles btnlink4.Click
        frm4.Show()
    End Sub

    Private Sub Btnlink5_Click(sender As Object, e As EventArgs) Handles btnlink5.Click
        frm5.Show()
    End Sub

    Private Sub Btnlink6_Click(sender As Object, e As EventArgs) Handles btnlink6.Click
        frm6.Show()
    End Sub

    Private Sub Btnlink7_Click(sender As Object, e As EventArgs) Handles btnlink7.Click
        frm7.Show()
    End Sub

    Private Sub Btnlink8_Click(sender As Object, e As EventArgs) Handles btnlink8.Click
        frm8.Show()
    End Sub
    Private Sub Tm_tuongtac_Tick(sender As Object, e As EventArgs) Handles tm_tuongtac.Tick
        Try
            ProgressBar2.Increment(70)
            If NumericUpDown1.Value = 1 Then
                frm1.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
            ElseIf NumericUpDown1.Value = 2 Then
                frm1.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm2.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
            ElseIf NumericUpDown1.Value = 3 Then
                frm1.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm2.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm3.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
            ElseIf NumericUpDown1.Value = 4 Then
                frm1.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm2.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm3.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm4.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
            ElseIf NumericUpDown1.Value = 5 Then
                frm1.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm2.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm3.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm4.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm5.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
            ElseIf NumericUpDown1.Value = 6 Then
                frm1.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm2.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm3.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm4.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm5.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm6.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
            ElseIf NumericUpDown1.Value = 7 Then
                frm1.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm2.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm3.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm4.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm5.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm6.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm7.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
            ElseIf NumericUpDown1.Value = 8 Then
                frm1.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm2.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm3.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm4.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm5.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm6.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm7.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
                frm8.WebBrowser1.Document.Window.ScrollTo(1, ProgressBar2.Value)
            End If


            If ProgressBar2.Value = ProgressBar2.Maximum Then
                ListBox2.SelectedIndex += 1
                ChangeUserAgent(ListBox2.SelectedItem)
                ListBox1.SelectedIndex += 1
                Use_Proxy.UseProxy(ListBox1.SelectedItem)
                tm_tuongtac.Stop()
                If NumericUpDown1.Value = 1 Then
                    frm1.WebBrowser1.Refresh()
                ElseIf NumericUpDown1.Value = 2 Then
                    frm1.WebBrowser1.Refresh()
                    frm2.WebBrowser1.Refresh()
                ElseIf NumericUpDown1.Value = 3 Then
                    frm1.WebBrowser1.Refresh()
                    frm2.WebBrowser1.Refresh()
                    frm3.WebBrowser1.Refresh()
                ElseIf NumericUpDown1.Value = 4 Then
                    frm1.WebBrowser1.Refresh()
                    frm2.WebBrowser1.Refresh()
                    frm3.WebBrowser1.Refresh()
                    frm4.WebBrowser1.Refresh()
                ElseIf NumericUpDown1.Value = 5 Then
                    frm1.WebBrowser1.Refresh()
                    frm2.WebBrowser1.Refresh()
                    frm3.WebBrowser1.Refresh()
                    frm4.WebBrowser1.Refresh()
                    frm5.WebBrowser1.Refresh()
                ElseIf NumericUpDown1.Value = 6 Then
                    frm1.WebBrowser1.Refresh()
                    frm2.WebBrowser1.Refresh()
                    frm3.WebBrowser1.Refresh()
                    frm4.WebBrowser1.Refresh()
                    frm5.WebBrowser1.Refresh()
                    frm6.WebBrowser1.Refresh()
                ElseIf NumericUpDown1.Value = 7 Then
                    frm1.WebBrowser1.Refresh()
                    frm2.WebBrowser1.Refresh()
                    frm3.WebBrowser1.Refresh()
                    frm4.WebBrowser1.Refresh()
                    frm5.WebBrowser1.Refresh()
                    frm6.WebBrowser1.Refresh()
                    frm7.WebBrowser1.Refresh()
                ElseIf NumericUpDown1.Value = 8 Then
                    frm1.WebBrowser1.Refresh()
                    frm2.WebBrowser1.Refresh()
                    frm3.WebBrowser1.Refresh()
                    frm4.WebBrowser1.Refresh()
                    frm5.WebBrowser1.Refresh()
                    frm6.WebBrowser1.Refresh()
                    frm7.WebBrowser1.Refresh()
                    frm8.WebBrowser1.Refresh()
                End If
                ProgressBar2.Value = 0
                tm_nghi.Start()
            End If
            'Dim rd As New Random
            'Dim toado As String
            'toado = (rd.Next(0, 1000))
            'frm1.WebBrowser1.Document.Window.ScrollTo(1, toado)
            'Thread.Sleep(TimeSpan.FromSeconds(5))
            'frm1.WebBrowser1.Document.Window.ScrollTo(1, toado)
            'Thread.Sleep(500)
            'tm_tuongtac.Stop()
            tm_nghi.Start()
        Catch ex As Exception
            tm_tuongtac.Stop()
        End Try

    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        'Dim tsTest As TimeSpan = TimeSpan.FromMinutes(NumericUpDown2.Value)
        'tm_tuongtac.Interval = CInt(tsTest.TotalMinutes)
        If NumericUpDown2.Value = 1 Then

        End If
    End Sub

    Private Sub Tm_nghi_Tick(sender As Object, e As EventArgs) Handles tm_nghi.Tick
        ProgressBar1.Increment(10)
        If ProgressBar1.Value = ProgressBar1.Maximum Then
            tm_tuongtac.Start()
            tm_nghi.Stop()
            ProgressBar1.Value = 0
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 2")
        Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 1")
        Label12.Text = "deleted track"
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Use_Proxy.UseProxy(ListBox1.SelectedItem)
    End Sub

    Private Sub Nbrchagerip_ValueChanged(sender As Object, e As EventArgs) Handles nbrchagerip.ValueChanged
        If nbrchagerip.Value = 1 Then
            ProgressBar2.Maximum = 1000
        ElseIf nbrchagerip.Value = 2 Then
            ProgressBar2.Maximum = 2000
        ElseIf nbrchagerip.Value = 3 Then
            ProgressBar2.Minimum = 3000
        ElseIf nbrchagerip.Value = 4 Then
            ProgressBar2.Maximum = 4000
        ElseIf nbrchagerip.Value = 5 Then
            ProgressBar2.Minimum = 5000
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        about.Show()
    End Sub
End Class

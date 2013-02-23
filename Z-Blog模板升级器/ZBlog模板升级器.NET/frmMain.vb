Public Class frmMain

    Dim objAero As clsAero
    Private Sub btnPlugin_Click(sender As Object, e As EventArgs) Handles btnPlugin.Click
        Me.Hide()
        frmUpdatePlugin.Show()
    End Sub

    Private Sub btnTheme_Click(sender As Object, e As EventArgs) Handles btnTheme.Click
        Me.Hide()
        frmUpdateTheme.Show()
    End Sub


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Environment.OSVersion.Version.Major >= 6 Then bolAero = True

        If bolAero Then
            objAero = New clsAero
            objAero.Form = Me
            objAero.Go()
        End If

    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start("https://code.google.com/p/zblog-1-9/source/browse/#svn%2Ftrunk%2FZ-Blog%E6%A8%A1%E6%9D%BF%E5%8D%87%E7%BA%A7%E5%99%A8")
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("http://www.zsxsoft.com/")
    End Sub
End Class
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
End Class
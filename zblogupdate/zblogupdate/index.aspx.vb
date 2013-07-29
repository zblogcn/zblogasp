Partial Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim dirname = System.Configuration.ConfigurationManager.AppSettings("dirname")

        Response.Clear()
        If Request.RawUrl.Contains("?install") = True Then
            Dim s As String = ""
            If Application("Release") = Nothing Then
                s = My.Computer.FileSystem.ReadAllText(System.Web.HttpContext.Current.Request.PhysicalApplicationPath & dirname & "\" & "Release.xml")
                Application("Release") = s
            Else
                s = Application("Release")
            End If
            System.Web.HttpContext.Current.Response.Filter = New IO.Compression.GZipStream(System.Web.HttpContext.Current.Response.Filter, IO.Compression.CompressionMode.Compress)
            System.Web.HttpContext.Current.Response.AppendHeader("Content-Encoding", "gzip")
            System.Web.HttpContext.Current.Response.ContentType = "text/xml"
            System.Web.HttpContext.Current.Response.AppendHeader("Last-Modified", System.DateTime.Parse(My.Computer.FileSystem.GetFileInfo(System.Web.HttpContext.Current.Request.PhysicalApplicationPath & dirname & "\" & "Release.xml").LastWriteTime).ToUniversalTime.ToString("r", System.Globalization.DateTimeFormatInfo.InvariantInfo))
            Response.Write(s)
        ElseIf Request.RawUrl.Contains("?beta") = True Then
            Dim s As String = System.Web.HttpContext.Current.Request.PhysicalApplicationPath & dirname & "\beta.html"
            Response.ContentEncoding = System.Text.Encoding.UTF8
            Response.ContentType = "text/plain"
            Response.WriteFile(s)
        ElseIf Request.QueryString.Count > 0 Then

            If System.Web.HttpContext.Current.Request.Headers("Accept-Encoding") IsNot Nothing Then
                If System.Web.HttpContext.Current.Request.Headers("Accept-Encoding").Contains("gzip") = True Then
                    System.Web.HttpContext.Current.Response.Filter = New IO.Compression.GZipStream(System.Web.HttpContext.Current.Response.Filter, IO.Compression.CompressionMode.Compress)
                    System.Web.HttpContext.Current.Response.AppendHeader("Content-Encoding", "gzip")
                ElseIf System.Web.HttpContext.Current.Request.Headers("Accept-Encoding").Contains("deflate") = True Then
                    System.Web.HttpContext.Current.Response.Filter = New IO.Compression.DeflateStream(System.Web.HttpContext.Current.Response.Filter, IO.Compression.CompressionMode.Compress)
                    System.Web.HttpContext.Current.Response.AppendHeader("Content-Encoding", "deflate")
                End If
            End If

            System.Web.HttpContext.Current.Response.ContentType = "application/octet-stream"
            Dim s As String = System.Web.HttpContext.Current.Request.PhysicalApplicationPath & dirname & "\" & Request.QueryString.Get(0)
            If s.Contains("../") = False And s.Contains("..\") = False Then
                Try
                    Response.WriteFile(s)
                Catch ex As Exception

                End Try

            End If

        Else
            Dim s As String = System.Web.HttpContext.Current.Request.PhysicalApplicationPath & dirname & "\now.html"
            Response.ContentEncoding = System.Text.Encoding.UTF8
            Response.ContentType = "text/plain"
            Response.WriteFile(s)
        End If

        Response.End()
    End Sub

End Class
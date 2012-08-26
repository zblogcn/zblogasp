<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 spirit 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备   注:     Aocool Studio Photo / Gallery Magic Show
'// 最后修改：   2009.12.30
'// 最后版本:    2.7.1
'///////////////////////////////////////////////////////////////////////////////
%>
<%Option Explicit%>
<%
Server.ScriptTimeout = 300
Response.Buffer = TRUE

On Error Resume Next

Function IsNullOrEmpty(ByVal String)
    IsNullOrEmpty = IsNull(String) Or String = ""
End Function

Function GetImage(ByVal URL)
    Dim oXmlHttp
    Set oXmlHttp = Server.CreateObject("Msxml2.XMLHTTP")

    If Err.Number <> 0 Then
        Response.Write("XMLHTTP Object not installed on this server, please go to Microsoft website download and install it.")
        Response.End()
    End If

    oXmlHttp.Open "GET", URL, FALSE
    oXmlHttp.setRequestHeader "Referer", URL
    oXmlHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)"
    oXmlHttp.Send()

    If oXmlHttp.readyState <> 4 Then
        GetImage = ""
    Else
        GetImage = oXmlHttp.responseBody
    End If

    Set oXmlHttp = Nothing
End Function

Function GetContentType(ByVal FileName)
    Dim FileExtension, ContentType
    FileExtension = Mid(FileName, InStrRev(FileName, ".") + 1)

    Select Case FileExtension
        Case "jpe"
            ContentType = "image/jpeg"
        Case "jpg"
            ContentType = "image/jpeg"
        Case "jpeg"
            ContentType = "image/jpeg"
        Case "gif"
            ContentType = "image/gif"
        Case "bmp"
            ContentType = "image/bmp"
        Case "png"
            ContentType = "image/png"
        Case "pnz"
            ContentType = "image/png"
        Case Else
            ContentType = "text/html"
    End Select

    GetContentType = ContentType
End Function

Dim URL, Bin
URL = Request.ServerVariables("QUERY_STRING")
Bin = GetImage(URL)

If IsNullOrEmpty(URL) = FALSE Then
    If Bin <> "" Then
        Response.ContentType = GetContentType(URL)
        Response.BinaryWrite Bin
        Response.Flush
    Else
        Response.ContentType = "text/html"
        Response.Write("Remote Server Error.")
    End If
Else
    Response.ContentType = "text/html"
    Response.Status = "400 Bad Request"
    Response.Write("400 Bad Request")
End If
%>
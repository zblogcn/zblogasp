Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Text.RegularExpressions
Module mdlBase
	Public objADO As Object
    Public objXML As Object
    Public bolAero As Boolean

    'Usage:升级
    'Param:strFilePath--文件名,intType--升级类型
    Function Update_Plugin(ByVal strFilePath As String, Optional ByRef intType As Short = 1, Optional ByVal OldPath As String = "", Optional ByVal NewPath As String = "") As Boolean
        Dim strFile As String = "", strTemp As String = ""
        Dim objExec As Object = Nothing
        If File.Exists(strFilePath) Then

            Log("Update: " & strFilePath & "  type:" & intType)
            strFile = File.ReadAllText(strFilePath)
            Select Case intType
                Case 1
                    '插件主体升级

                    OldPath = OldPath.Replace("..\", "..[/\\]")
                    '替换INCLUDE地址
                    For Each objExec In New Regex("<!-- +?#include +?file=""(" & OldPath & ").+?""", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Groups(1).Value, NewPath, 1, 1)
                        Log(objExec.Groups(1).Value & "-->" & NewPath)
                    Next objExec


                    '替换空行
                    'For Each objExec In New Regex("[" & vbTab & " ]+" & vbCrLf, RegexOptions.IgnoreCase).Matches(strFile)
                    'strFile = Replace(strFile, objExec.Value, "", 1, 1)
                    'Log(objExec.Value & "-->" & """""")
                    'Next objExec


                    '替换HEADER
                    For Each objExec In New Regex("<div class=[""']Header[""']>", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "<div class=""divHeader"">", 1, 1)
                        Log(objExec.Value & "-->" & "<div class=""divHeader"">")
                    Next objExec


                    '替换<head>
                    For Each objExec In New Regex("<!DOCTYPE html.+?>", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "<!--#include file=""" & NewPath & "zb_system\admin\admin_header.asp""-->", 1, 1)
                        Log(objExec.Value & "-->" & "<!--#include file=""" & NewPath & "zb_system\admin\admin_header.asp""-->")
                    Next objExec
                    strTemp = "<html.+?>" & vbCrLf & "|<title>.+?</title>|<head>|<meta.+?>" & vbCrLf & "|<body>|</head>|</html>" & _
                        "|<link rel=[""']stylesheet[""'] +?rev=[""']stylesheet[""'] +?href=[""'].+?CSS\/admin.css[""'] +?type=[""']text/css[""'] +?media=""screen"".+?>"

                    For Each objExec In New Regex(strTemp, RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & "")
                    Next objExec
                    For Each objExec In New Regex("<div id=[""']divMain[""']>", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "<!--#include file=""" & NewPath & "zb_system\admin\admin_top.asp""--><div id=""divMain"">", 1, 1)
                        Log(objExec.Value & "-->" & "<!--#include file=""" & NewPath & "zb_system\admin\admin_top.asp""--><div id=""divMain"">")
                    Next objExec
                    For Each objExec In New Regex("<div id=[""']divMain2[""']>", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "<div id=""divMain2""><script type=""text/javascript"">ActiveLeftMenu(""aPlugInMng"");</script>", 1, 1)
                        Log(objExec.Value & "-->" & "<div id=""divMain2""><script type=""text/javascript"">ActiveLeftMenu(""aPlugInMng"");</script>")
                    Next objExec

                    For Each objExec In New Regex("<% +?Call GetBlogHint\(\) +?%>", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "<div id=""ShowBlogHint""><%Call GetBlogHint()%></div>", 1, 1)
                        Log(objExec.Value & "-->" & "<div id=""ShowBlogHint""><%Call GetBlogHint()%></div>")
                    Next objExec
                    For Each objExec In New Regex("</div>[" & vbCrLf & vbTab & " ]+?</body>", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "<!--#include file=""" & NewPath & "zb_system\admin\admin_footer.asp""-->", 1, 1)
                        Log(objExec.Value & "-->" & "<!--#include file=""" & NewPath & "zb_system\admin\admin_footer.asp""-->")
                    Next objExec


                Case 2
                    For Each objExec In New Regex("BlogPath +?(\&|\+) +?""(\\|\/)?(PLUGIN|THEMES|CACHE|INCLUDE|c_option|c_custom|LANGUAGE|UPLOAD|FACE)", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "BlogPath & ""zb_users/" & objExec.Groups(3).Value, 1, 1)
                        Log(objExec.Value & "-->" & "BlogPath & ""zb_users/" & objExec.Groups(3).Value)
                    Next
                    For Each objExec In New Regex("BlogPath +?(\&|\+) +?""(\\|\/)?(FUNCTION|ADMIN|CSS|IMAGE|SCRIPT|XML-RPC)", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "BlogPath & ""zb_system/" & objExec.Groups(3).Value, 1, 1)
                        Log(objExec.Value & "-->" & "BlogPath & ""zb_system/" & objExec.Groups(3).Value)
                    Next
                    For Each objExec In New Regex("ZC_BLOG_HOST +?(\&|\+) +?""(\\|\/)?(PLUGIN|THEMES|CACHE|INCLUDE|c_option|c_custom|LANGUAGE|UPLOAD|FACE)", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "ZC_BLOG_HOST & ""zb_users\" & objExec.Groups(3).Value, 1, 1)
                        Log(objExec.Value & "-->" & "ZC_BLOG_HOST & ""zb_users\" & objExec.Groups(3).Value)
                    Next
                    For Each objExec In New Regex("ZC_BLOG_HOST +?(\&|\+) +?""(\\|\/)?(FUNCTION|ADMIN|CSS|IMAGE|SCRIPT|XML-RPC)", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "ZC_BLOG_HOST & ""zb_system\" & objExec.Groups(3).Value, 1, 1)
                        Log(objExec.Value & "-->" & "ZC_BLOG_HOST & ""zb_system\" & objExec.Groups(3).Value)
                    Next
                    For Each objExec In New Regex("\\..(\\|\/)(FUNCTION|ADMIN|CSS|IMAGE|SCRIPT|XML-RPC)", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "\..\zb_system\" & objExec.Groups(2).Value, 1, 1)
                        Log(objExec.Value & "-->" & "\..\zb_system\" & objExec.Groups(3).Value)
                    Next
                    For Each objExec In New Regex("\.\.(\\|\/)(PLUGIN|THEMES|CACHE|INCLUDE|c_option|c_custom|LANGUAGE|UPLOAD|FACE)", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "\..\zb_users\" & objExec.Groups(2).Value, 1, 1)
                        Log(objExec.Value & "-->" & "\..\zb_users\" & objExec.Groups(3).Value)
                    Next
                    strFile = Replace(strFile, "zb_users\themes", "zb_users\theme", 1, 1)
                    strFile = Replace(strFile, "zb_users\face", "zb_users\emotion", 1, 1)
                    strFile = Replace(strFile, "zb_users\c_custom", "zb_users\c_option", 1, 1)
            End Select
        Else
            Log(strFile & "找不到！")
        End If

        File.WriteAllText(strFilePath, strFile)

        Log("保存完毕")
    End Function

    'Usage:日志
    'Param:str--日志内容
    Sub Log(ByVal str_Renamed As String)
        frmUpdatePlugin.lstLog.Items.Add("【" & Now & "】" & str_Renamed)
    End Sub

End Module
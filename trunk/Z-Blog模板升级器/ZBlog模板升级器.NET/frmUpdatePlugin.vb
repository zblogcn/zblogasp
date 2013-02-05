Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Text.RegularExpressions

Friend Class frmUpdatePlugin
    Inherits System.Windows.Forms.Form

    Dim strSource, strTemplateFolder, strXMLPath As String
    Dim aryPluginFile() As String
    Dim aryOldPath() As String
    Dim aryNewPath() As String
    Dim objAero As clsAero


    Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click
        Dim strTemp As String
        fbdDialog.Description = "请选择插件文件夹"
        fbdDialog.ShowDialog()
        strTemp = fbdDialog.SelectedPath
        If Not strTemp = "" Then
            strTemplateFolder = strTemp
            txtPath.Text = strTemp
            Log("选择插件文件夹：" & strTemp)
        End If
    End Sub

    Private Sub cmdOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpen.Click
        Dim i As Short
        Log_Clear()
        strTemplateFolder = txtPath.Text
        If GetSubFolder(strTemplateFolder) Then
            Log("开始升级插件文件")
            For i = 0 To CShort(UBound(aryPluginFile))
                If Trim(aryPluginFile(i)) <> "" Then
                    Update_Plugin(aryPluginFile(i), 1, aryOldPath(i), aryNewPath(i))
                End If

            Next

            MsgBox(Replace("升级完毕！\n\n剩余以下部分没有升级，请自行修改：\n\n升级完成后，请在APP中心里编辑插件信息并保存，即可在2.0里激活插件。", "\n", vbCrLf), MsgBoxStyle.Information)
        End If

    End Sub

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        If Environment.OSVersion.Version.Major >= 6 Then bolAero = True

        If bolAero Then
            objAero = New clsAero
            objAero.Form = Me
            objAero.Go()
        End If


        strTemplateFolder = ""
        ReDim aryPluginFile(0)
        ReDim aryOldPath(0)
        ReDim aryNewPath(0)
        strSource = ""
        lblNote.Text = "说明：" & vbCrLf & "升级前必须备份。" & vbCrLf & _
            "您要升级的1.8插件必须符合以下要求：" & vbCrLf & _
            "以上条件有任意一点不符合，则本程序无法升级你的插件。"
    End Sub


    Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub




    'Usage:日志
    'Param:str--日志内容
    Sub Log(ByVal str_Renamed As String)
        lstLog.Items.Add("【" & Now & "】" & str_Renamed)
    End Sub

    'Usage:清除日志
    Sub Log_Clear()
        lstLog.Items.Clear()
    End Sub

    'Usage:扫描文件夹
    'Param:Folder--文件夹
    Function GetSubFolder(ByVal Folder As String, Optional ByVal OldPath As String = "..\..\", Optional ByVal NewPath As String = "..\..\..\", Optional ByVal SubFolder As Boolean = False) As Boolean
        GetSubFolder = False
        Dim objFor As Object
        If Directory.Exists(Folder) Then
            If Not SubFolder Then
                If File.Exists(Folder & "\plugin.xml") Then
                    strXMLPath = Folder & "\plugin.xml"
                    Log("找到插件XML信息")
                Else
                    Log("插件XML不存在")
                    Exit Function
                End If
            End If
            For Each objFor In Directory.GetFiles(Folder)
                If objFor Like "*.asp" Then
                    ReDim Preserve aryPluginFile(UBound(aryPluginFile) + 1)
                    ReDim Preserve aryOldPath(UBound(aryOldPath) + 1)
                    ReDim Preserve aryNewPath(UBound(aryNewPath) + 1)
                    aryPluginFile(UBound(aryPluginFile)) = objFor
                    aryOldPath(UBound(aryOldPath)) = OldPath
                    aryNewPath(UBound(aryNewPath)) = NewPath
                    Log("找到插件文件：" & objFor)
                End If
            Next objFor

            For Each objFor In Directory.GetDirectories(Folder)
                GetSubFolder(objFor, OldPath & "..\", NewPath & "..\", True)
            Next objFor

            GetSubFolder = True
        Else
            Log("文件夹不存在！")
        End If
    End Function


    'Usage:得到XML信息以判断是否Z-Blog
    'Param:XMLPath--XML地址
    Function LoadXMLInfo(ByVal XMLPath As String) As Boolean

    End Function




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
                    '模板主体和INCLUDE文件夹升级

                    OldPath = OldPath.Replace("..\", "..[/\\]")
                    '替换INCLUDE地址
                    For Each objExec In New Regex("<!-- +?#include +?file=""(" & OldPath & ").+?""", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Groups(1).Value, NewPath, 1, 1)
                        Log(objExec.Groups(1).Value & "-->" & NewPath)
                    Next objExec


                    '替换空行
                    For Each objExec In New Regex("[" & vbTab & " ]+" & vbCrLf, RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec


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

                    '保存
                    File.WriteAllText(strFilePath & "_update", strFile)
                    Log("保存完毕")

            End Select
        Else
            Log(strFile & "找不到！")
        End If

    End Function
End Class
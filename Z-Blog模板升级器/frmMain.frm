VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "1.8 模板升级器"
   ClientHeight    =   5370
   ClientLeft      =   7710
   ClientTop       =   4950
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   10725
   Begin VB.ListBox lstLog 
      Height          =   4200
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   10215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "打开(&O)"
      Height          =   375
      Left            =   9360
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "浏览(&B)"
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtPath 
      Height          =   270
      Left            =   1080
      TabIndex        =   1
      Top             =   280
      Width           =   7095
   End
   Begin VB.Label lblFolder 
      Caption         =   "模板路径"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   330
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strTemplateFolder As String, aryTemplateFile() As String, aryPluginFile() As String, strSource As String, strXMLPath As String


Private Sub cmdBrowse_Click()
    Dim strTemp As String
    strTemp = GetFolderPath("请选择模板文件夹", Me.hWnd)
    If Not strTemp = "False" Then
        strTemplateFolder = strTemp
        txtPath.Text = strTemp
        Log "选择模板文件夹：" & strTemp
    End If
End Sub

Private Sub cmdOpen_Click()
    Dim i As Integer
    Log_Clear
    strTemplateFolder = txtPath.Text
    If GetSubFolder(strTemplateFolder) Then
        Log "开始升级模板文件"
        For i = 0 To UBound(aryTemplateFile)
            If Trim(aryTemplateFile(i)) <> "" Then Update aryTemplateFile(i), 1
        Next
        Log "模板文件升级完毕"
        Log "开始升级source下asp"
        Update strSource, 2
        Log "source下asp升级完毕"
        Log "开始升级主题自带插件"
        For i = 0 To UBound(aryPluginFile)
            If Trim(aryPluginFile(i)) <> "" Then Update aryPluginFile(i), 3
        Next
        Log "主题自带插件升级完毕"
        '还差侧栏管理的升级
        'Log "升级XML信息"
        '升级XML是不是让APP升级好一点
        
        MsgBox Replace("升级完毕！\n\n剩余以下部分没有升级，请自行修改：\n\n侧栏部分（须符合2.0侧栏规范）\n主题插件\nXML信息\n\n升级完成后，请在APP中心里编辑主题信息并保存，即可在2.0里激活主题。", "\n", vbCrLf), vbInformation
    End If
    
End Sub

Private Sub Form_Load()
    Set objRegExp = New RegExp
    Set objFSO = New FileSystemObject
    Set objADO = CreateObject("ADODB.Stream")
    strTemplateFolder = ""
    objRegExp.Global = True
    objRegExp.IgnoreCase = True
    ReDim aryTemplateFile(0)
    ReDim aryPluginFile(0)
    strSource = ""
    txtPath.Text = "D:\Win8\Desktop\THEMES\Couple"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRegExp = Nothing
    Set objFSO = Nothing
    Set objADO = Nothing
End Sub




'Usage:日志
'Param:str--日志内容
Sub Log(ByVal str As String)
    lstLog.AddItem "【" & Now & "】" & str
End Sub

'Usage:清除日志
Sub Log_Clear()
    lstLog.Clear
End Sub

'Usage:扫描文件夹
'Param:Folder--文件夹
Function GetSubFolder(ByVal Folder As String) As Boolean
    GetSubFolder = False
    Dim objSub As Object, objFor
    If objFSO.FolderExists(Folder) Then
        If objFSO.FileExists(Folder & "/theme.xml") Then
            strXMLPath = objFSO.GetFile(Folder & "/theme.xml").Path
            Log "找到主题XML信息"
        Else
            Log "主题XML不存在"
            Exit Function
        End If
        If objFSO.FolderExists(Folder & "/template") Then
            For Each objFor In objFSO.GetFolder(Folder & "/template").Files
                ReDim Preserve aryTemplateFile(UBound(aryTemplateFile) + 1)
                aryTemplateFile(UBound(aryTemplateFile)) = objFor.Path
                Log "找到主题文件：" & objFor.Name
            Next
        End If
        If objFSO.FolderExists(Folder & "/plugin") Then
            For Each objFor In objFSO.GetFolder(Folder & "/plugin").Files
                ReDim Preserve aryPluginFile(UBound(aryPluginFile) + 1)
                aryPluginFile(UBound(aryPluginFile)) = objFor.Path
                Log "找到主题插件：" & objFor.Name
            Next
        End If
        If objFSO.FileExists(Folder & "/source/style.css.asp") Then
            strSource = objFSO.GetFile(Folder & "/source/style.css.asp").Path
            Log "找到STYLE.CSS.ASP"
        End If
        GetSubFolder = True
    Else
        Log "文件夹不存在！"
    End If
End Function


'Usage:得到XML信息以判断是否Z-Blog
'Param:XMLPath--XML地址
Function LoadXMLInfo(ByVal XMLPath As String) As Boolean

End Function




'Usage:升级
'Param:strFilePath--文件名,intType--升级类型
Function Update(ByVal strFilePath As String, Optional intType As Integer = 1) As Boolean
    Dim strFile As String, objExec As Object
    If objFSO.FileExists(strFilePath) Then
        Log "Update: " & strFilePath & "  type:" & intType
        strFile = LoadFromFile(strFilePath)
        Select Case intType
            Case 1
                '模板主体升级
                
                
                '替换zb_system下文件
                objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(admin|script|function|image|cmd.asp|login.asp)"
                
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_system/" & objExec.SubMatches(0), 1, 1)
                    Log objExec.SubMatches(0) & "-->" & "zb_system/" & objExec.SubMatches(0)
                Next
                
                '替换zb_users下文件
                objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(plugin|language|cache|upload)"
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_users/" & objExec.SubMatches(0), 1, 1)
                    Log objExec.SubMatches(0) & "-->" & "zb_users/" & objExec.SubMatches(0)
                Next
                
                '替换theme
                objRegExp.Pattern = "(\<\#ZC_BLOG_HOST\#\>themes)"
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.SubMatches(0), "<#ZC_BLOG_HOST#>zb_users/theme", 1, 1)
                    Log objExec.SubMatches(0) & "-->" & "<#ZC_BLOG_HOST#>zb_users/theme"
                Next
                
                '替换rss
                objRegExp.Pattern = "(\<\#ZC_BLOG_HOST\#\>rss\.xml)"
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.SubMatches(0), "<#ZC_BLOG_HOST#>feed.asp", 1, 1)
                    Log objExec.SubMatches(0) & "-->" & "<#ZC_BLOG_HOST#>feed.asp"
                Next
                
                
                '替换那些玩意
                objRegExp.Pattern = "var (str0[0-9]|intMaxLen|strBatchView|strBatchInculde|strBatchCount)=.+?;"
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.Value, "", 1, 1)
                    Log objExec.Value & "-->" & """"""
                Next
                
                '强插c_html_js_add.asp
                If InStr(LCase(strFile), "c_html_js_add.asp") = 0 And InStr(LCase(strFile), "</head>") > 0 Then
                    strFile = Replace(strFile, "</head>", "<script src=""<#ZC_BLOG_HOST#>zb_system/function/c_html_js_add.asp"" type=""text/javascript""></script>" & vbCrLf & "</head>")
                    Log "强制插入c_html_js_add.asp"
                End If
                
                '删除无用UBB部分
                objRegExp.Pattern = "InsertQuote.+?\;|ExportUbbFrame\(\)\;?"
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.Value, "", 1, 1)
                    Log objExec.Value & "-->" & """"""
                Next
                
                '替换空行
                objRegExp.Pattern = "[" & vbTab & vbSpace & "]+" & vbCrLf
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.Value, "", 1, 1)
                    Log objExec.Value & "-->" & """"""
                Next
                '保存
                SaveToFile strFilePath, strFile
                Log "保存完毕"
        Case 2
                'SOURCE\STYLE.CSS.ASP升级
                
                '替换<%
                strFile = Replace(strFile, "<%", "<!-- #include file=""../../../../zb_system/function/c_function.asp"" -->" & vbCrLf & "<%")
                Log "引用c_function.asp"
                
                '替换路径
                strFile = Replace(strFile, """themes""", """zb_users/theme""")
                Log """themes"" --> ""zb_users/theme"""
                
                '替换HOST
                strFile = Replace(strFile, "ZC_BLOG_HOST", "GetCurrentHost()")
                Log "ZC_BLOG_HOST --> GetCurrentHost()"
                
                
                SaveToFile strFilePath, strFile
                Log "保存完毕"
        
        Case 3
                '插件\主题插件升级
        Case 4
                '侧栏管理升级
        Case 5
                'XML升级
                
        End Select
    Else
        Log strFile & "找不到！"
    End If
End Function




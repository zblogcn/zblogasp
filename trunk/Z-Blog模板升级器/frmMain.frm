VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "1.8 ģ��������"
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
      Caption         =   "��(&O)"
      Height          =   375
      Left            =   9360
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "���(&B)"
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
      Caption         =   "ģ��·��"
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
    strTemp = GetFolderPath("��ѡ��ģ���ļ���", Me.hWnd)
    If Not strTemp = "False" Then
        strTemplateFolder = strTemp
        txtPath.Text = strTemp
        Log "ѡ��ģ���ļ��У�" & strTemp
    End If
End Sub

Private Sub cmdOpen_Click()
    Dim i As Integer
    Log_Clear
    strTemplateFolder = txtPath.Text
    If GetSubFolder(strTemplateFolder) Then
        Log "��ʼ����ģ���ļ�"
        For i = 0 To UBound(aryTemplateFile)
            If Trim(aryTemplateFile(i)) <> "" Then Update aryTemplateFile(i), 1
        Next
        Log "ģ���ļ��������"
        Log "��ʼ����source��asp"
        
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
    txtPath.Text = "D:\Win8\Desktop\THEMES\Qeeke"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRegExp = Nothing
    Set objFSO = Nothing
    Set objADO = Nothing
End Sub




'Usage:��־
'Param:str--��־����
Sub Log(ByVal str As String)
    lstLog.AddItem "��" & Now & "��" & str
End Sub

'Usage:�����־
Sub Log_Clear()
    lstLog.Clear
End Sub

'Usage:ɨ���ļ���
'Param:Folder--�ļ���
Function GetSubFolder(ByVal Folder As String) As Boolean
    GetSubFolder = False
    Dim objSub As Object, objFor
    If objFSO.FolderExists(Folder) Then
        If objFSO.FileExists(Folder & "/theme.xml") Then
            strXMLPath = objFSO.GetFile(Folder & "/theme.xml").Path
            Log "�ҵ�����XML��Ϣ"
        Else
            Log "����XML������"
            Exit Function
        End If
        If objFSO.FolderExists(Folder & "/template") Then
            For Each objFor In objFSO.GetFolder(Folder & "/template").Files
                ReDim Preserve aryTemplateFile(UBound(aryTemplateFile) + 1)
                aryTemplateFile(UBound(aryTemplateFile)) = objFor.Path
                Log "�ҵ������ļ���" & objFor.Name
            Next
        End If
        If objFSO.FolderExists(Folder & "/plugin") Then
            For Each objFor In objFSO.GetFolder(Folder & "/plugin").Files
                ReDim Preserve aryPluginFile(UBound(aryPluginFile) + 1)
                aryPluginFile(UBound(aryPluginFile)) = objFor.Path
                Log "�ҵ���������" & objFor.Name
            Next
        End If
        If objFSO.FileExists(Folder & "/source/style.css.asp") Then
            strSource = objFSO.GetFile(Folder & "/source/style.css.asp").Path
            Log "�ҵ�STYLE.CSS.ASP"
        End If
        GetSubFolder = True
    Else
        Log "�ļ��в����ڣ�"
    End If
End Function


'Usage:�õ�XML��Ϣ���ж��Ƿ�Z-Blog
'Param:XMLPath--XML��ַ
Function LoadXMLInfo(ByVal XMLPath As String) As Boolean

End Function




'Usage:����
'Param:strFilePath--�ļ���,intType--��������
Function Update(ByVal strFilePath As String, Optional intType As Integer = 1) As Boolean
    Dim strFile As String, objExec As Object
    If objFSO.FileExists(strFilePath) Then
        Log "Update: " & strFilePath & "  type:" & intType
        strFile = LoadFromFile(strFilePath)
        Select Case intType
            Case 1
                '�滻zb_system���ļ�
                objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(admin|script|function|image|cmd.asp|login.asp)"
                
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_system/" & objExec.SubMatches(0), 1, 1)
                    Log objExec.SubMatches(0) & "-->" & "zb_system/" & objExec.SubMatches(0)
                Next
                
                '�滻zb_users���ļ�
                objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(plugin|language|cache|upload)"
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_users/" & objExec.SubMatches(0), 1, 1)
                    Log objExec.SubMatches(0) & "-->" & "zb_users/" & objExec.SubMatches(0)
                Next
                
                '�滻theme
                objRegExp.Pattern = "(\<\#ZC_BLOG_HOST\#\>themes)"
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.SubMatches(0), "<#ZC_BLOG_HOST#>zb_users/theme", 1, 1)
                    Log objExec.SubMatches(0) & "-->" & "<#ZC_BLOG_HOST#>zb_users/theme"
                Next
                
                '�滻rss
                objRegExp.Pattern = "(\<\#ZC_BLOG_HOST\#\>rss\.xml)"
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.SubMatches(0), "<#ZC_BLOG_HOST#>feed.asp", 1, 1)
                    Log objExec.SubMatches(0) & "-->" & "<#ZC_BLOG_HOST#>feed.asp"
                Next
                
                
                '�滻��Щ����
                objRegExp.Pattern = "var (str0[0-9]|intMaxLen|strBatchView|strBatchInculde|strBatchCount)=.+?;"
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.Value, "", 1, 1)
                    Log objExec.Value & "-->" & """"""
                Next
                
                'ǿ��c_html_js_add.asp
                If InStr(LCase(strFile), "c_html_js_add.asp") = 0 And InStr(LCase(strFile), "</head>") > 0 Then
                    strFile = Replace(strFile, "</head>", "<script src=""<#ZC_BLOG_HOST#>zb_system/function/c_html_js_add.asp"" type=""text/javascript""></script>" & vbCrLf & "</head>")
                    Log "ǿ�Ʋ���c_html_js_add.asp"
                End If
                
                'ɾ������UBB����
                objRegExp.Pattern = "InsertQuote.+?\;|ExportUbbFrame\(\)\;?"
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.Value, "", 1, 1)
                    Log objExec.Value & "-->" & """"""
                Next
                
                '�滻����
                objRegExp.Pattern = "[" & vbTab & vbSpace & "]+" & vbCrLf
                For Each objExec In objRegExp.Execute(strFile)
                    strFile = Replace(strFile, objExec.Value, "", 1, 1)
                    Log objExec.Value & "-->" & """"""
                Next
                '����
                SaveToFile strFilePath, strFile
                Log "�������"
        End Select
    Else
        Log strFile & "�Ҳ�����"
    End If
End Function




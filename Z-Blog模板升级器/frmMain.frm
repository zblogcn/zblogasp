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
Public strTemplateFolder As String, objFSO As FileSystemObject, objRegExp As RegExp, objADO As Object, objXML As New DOMDocument
Dim aryTemplateFile() As String, aryPluginFile() As String, strSource As String, strXMLPath As String


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
            Update aryTemplateFile(i), 1
        Next
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
'Param:strFile--文件名,intType--升级类型
Function Update(ByVal strFile As String, Optional intType As Integer = 1) As Boolean

End Function


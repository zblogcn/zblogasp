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
Public strTemplateFolder As String, objFSO As FileSystemObject, objRegExp As RegExp, objADO As Object


Private Sub cmdBrowse_Click()
    Dim strTemp As String
    strTemp = GetFolderPath("请选择模板文件夹", Me.hWnd)
    If Not strTemp = "False" Then
        strTemplateFolder = strTemp
        txtPath.Text = strTemp
        Log "选择模板文件夹：" & strTemp
    End If
End Sub

Private Sub Form_Load()
    Set objRegExp = New RegExp
    Set objFSO = New FileSystemObject
    Set objADO = CreateObject("ADODB.Stream")
    strTemplateFolder = ""
    objRegExp.Global = True
    objRegExp.IgnoreCase = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRegExp = Nothing
    Set objFSO = Nothing
    Set objADO = Nothing
End Sub




'Usage:日志
'Param:str--日志内容
Function Log(ByVal str As String)
    lstLog.AddItem "【" & Now & "】" & str
End Function

'Usage:清除日志
Function Log_Clear()
    lstLog.Clear
End Function

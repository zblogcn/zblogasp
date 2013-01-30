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
Public strTemplateFolder As String, objFSO As FileSystemObject, objRegExp As RegExp, objADO As Object, objXML As New DOMDocument
Dim aryTemplateFile() As String, aryPluginFile() As String, strSource As String, strXMLPath As String


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
'Param:strFile--�ļ���,intType--��������
Function Update(ByVal strFile As String, Optional intType As Integer = 1) As Boolean

End Function


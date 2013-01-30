Attribute VB_Name = "mdlBase"

'获得文件夹路径
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
               Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
                                             ByVal pszPath As String) As Long
'显示文件夹列表框
Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
               Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Const BIF_RETURNONLYFSDIRS = &H1


'Usage:打开一个文件夹选取窗口
'Param:strMsg--标题，hWnd--Form句柄

Function GetFolderPath(ByVal strMsg As String, ByRef hWnd As Long)

    Dim broInfo As BROWSEINFO
    Dim lngGet As Long
    Dim lngPID As Long
    Dim strPath As String
    broInfo.hOwner = hWnd
    broInfo.pidlRoot = 0&
    broInfo.lpszTitle = strMsg
    broInfo.ulFlags = &H1  'BIF_RETURNONLYFSDIRS
    lngPID = SHBrowseForFolder(broInfo)
    strPath = Space$(512)
    lngGet = SHGetPathFromIDList(lngPID&, strPath)
    If lngGet Then
        'API获取到的有Space，比较坑爹
        GetFolderPath = Left(strPath, InStr(strPath, Chr$(0)) - 1)
    Else
        GetFolderPath = False
    End If
    
End Function

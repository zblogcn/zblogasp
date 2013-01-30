Attribute VB_Name = "mdlBase"
Public objFSO As FileSystemObject, objRegExp As RegExp, objADO As Object, objXML As New DOMDocument
Public bolAero As Boolean

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type


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


'获得文件夹路径
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
               Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
                                             ByVal pszPath As String) As Long
'显示文件夹列表框
Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
               Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

'得到操作系统版本
Public Declare Function GetVersionEx Lib "kernel32" _
               Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Const BIF_RETURNONLYFSDIRS = &H1
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Const adTypeBinary = 1
Const adTypeText = 2

Const adModeRead = 1
Const adModeReadWrite = 3

Const adSaveCreateNotExist = 1
Const adSaveCreateOverWrite = 2




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

'Usage:得到系统版本
Function GetSystemVersion()
    Dim objOS As OSVERSIONINFO
    objOS.dwOSVersionInfoSize = 148
    objOS.szCSDVersion = Space$(128)
    Call GetVersionEx(objOS)
    GetSystemVersion = CDbl(objOS.dwMajorVersion & "." & objOS.dwMinorVersion)
    If GetSystemVersion >= 6 Then
        bolAero = True
    End If
End Function

Function LoadFromFile(ByVal strPath As String, Optional strCharset As String = "UTF-8")

    With objADO
        .Type = adTypeText
        .Mode = adModeReadWrite
        .Open
        .Charset = strCharset
        .Position = .Size
        .LoadFromFile strPath
        LoadFromFile = .ReadText
        .Close
    End With

    Err.Clear

End Function

Function SaveToFile(strFullName As String, strContent As String, Optional strCharset As String = "UTF-8")


    With objADO
        .Type = adTypeText
        .Mode = adModeReadWrite
        .Open
        .Charset = strCharset
        .Position = .Size
        .WriteText = strContent
        .SaveToFile strFullName, adSaveCreateOverWrite
        .Close
    End With
    

End Function


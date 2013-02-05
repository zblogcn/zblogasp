Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Text.RegularExpressions
Module mdlBase
	Public objADO As Object
    Public objXML As Object
    Public bolAero As Boolean

    Function Update_Plugin(ByVal strFilePath As String, Optional ByRef intType As Short = 1) As Boolean

        Dim strFile As String = ""
        Dim objExec As Object = Nothing
        If File.Exists(strFilePath) Then

            Log("Update: " & strFilePath & "  type:" & intType)
            strFile = File.ReadAllText(strFilePath)
            Select Case intType
                Case 1
                    '替换zb_system下文件
                    For Each objExec In New Regex("\<\#ZC_BLOG_HOST\#\>(admin|script|function|image|cmd.asp|login.asp)").Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_system/" & objExec.Groups(1).Value, 1, 1)
                        Log(objExec.Groups(1).Value & "-->" & "zb_system/" & objExec.Groups(1).Value)
                    Next objExec



                    '保存
                    File.WriteAllText(strFilePath, strFile)
                    Log("保存完毕")
                Case 2

                Case 3
                Case 4
                Case 5

            End Select
        Else
            Log(strFile & "找不到！")
        End If

    End Function
    Sub Log(str As String)
        'test
    End Sub

End Module
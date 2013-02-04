Option Strict Off
Option Explicit On
Imports System.IO
Module mdlBase
	Public objRegExp As VBScript_RegExp_55.RegExp
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
                    objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(admin|script|function|image|cmd.asp|login.asp)"

                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_system/" & objExec.SubMatches(0), 1, 1)
                        Log(objExec.SubMatches(0) & "-->" & "zb_system/" & objExec.SubMatches(0))
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
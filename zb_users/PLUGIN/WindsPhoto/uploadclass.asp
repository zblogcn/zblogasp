<%
Dim Data_5xsoft

Class upload_5xsoft

    Dim objForm, objFile, Version

    Public Function Form(strForm)
        strForm = LCase(strForm)
        If Not objForm.Exists(strForm) Then
            Form = ""
        Else
            Form = objForm(strForm)
        End If
    End Function

    Public Function File(strFile)
        strFile = LCase(strFile)
        If Not objFile.Exists(strFile) Then
            Set File = New FileInfo
        Else
            Set File = objFile(strFile)
        End If
    End Function


    Private Sub Class_Initialize
        Dim RequestData, sStart, VBCRLF, sInfo, iInfoStart, iInfoEnd, tStream, iStart, theFile
        Dim iFileSize, sFilePath, sFileType, sFormValue, sFileName
        Dim iFindStart, iFindEnd
        Dim iFormStart, iFormEnd, sFormName
        Version = "HTTPVersion   2.0"
        Set objForm = Server.CreateObject("Scripting.Dictionary")
        Set objFile = Server.CreateObject("Scripting.Dictionary")
        If Request.TotalBytes<1 Then Exit Sub
        Set tStream = Server.CreateObject("adodb.stream")
        Set Data_5xsoft = Server.CreateObject("adodb.stream")
        Data_5xsoft.Type = 1
        Data_5xsoft.Mode = 3
        Data_5xsoft.Open
        Data_5xsoft.Write Request.BinaryRead(Request.TotalBytes)
        Data_5xsoft.Position = 0
        RequestData = Data_5xsoft.Read

        iFormStart = 1
        iFormEnd = LenB(RequestData)
        VBCRLF = chrB(13) & chrB(10)
        sStart = MidB(RequestData, 1, InStrB(iFormStart, RequestData, VBCRLF) -1)
        iStart = LenB (sStart)
        iFormStart = iFormStart + iStart + 1

        While (iFormStart + 10) < iFormEnd
            iInfoEnd = InStrB(iFormStart, RequestData, VBCRLF & VBCRLF) + 3
            tStream.Type = 1
            tStream.Mode = 3
            tStream.Open
            Data_5xsoft.Position = iFormStart
            Data_5xsoft.CopyTo tStream, iInfoEnd - iFormStart
            tStream.Position = 0
            tStream.Type = 2
            tStream.Charset = "utf-8" '********************************这里原来是gb2312
            sInfo = tStream.ReadText
            tStream.Close

            iFormStart = InStrB(iInfoEnd, RequestData, sStart)
            iFindStart = InStr(22, sInfo, "name=""", 1) + 6
            iFindEnd = InStr(iFindStart, sInfo, """", 1)
            sFormName = LCase(Mid (sinfo, iFindStart, iFindEnd - iFindStart))

            If InStr (45, sInfo, "filename=""", 1) > 0 Then
                Set theFile = New FileInfo

                iFindStart = InStr(iFindEnd, sInfo, "filename=""", 1) + 10
                iFindEnd = InStr(iFindStart, sInfo, """", 1)
                sFileName = Mid (sinfo, iFindStart, iFindEnd - iFindStart)
                theFile.FileName = GetFileName(sFileName)
                theFile.FilePath = getFilePath(sFileName)

                iFindStart = InStr(iFindEnd, sInfo, "Content-Type:   ", 1) + 14
                iFindEnd = InStr(iFindStart, sInfo, VBCR)
                theFile.FileType = Mid (sinfo, iFindStart, iFindEnd - iFindStart)
                theFile.FileStart = iInfoEnd
                theFile.FileSize = iFormStart - iInfoEnd -3
                theFile.FormName = sFormName
                If Not objFile.Exists(sFormName) Then
                    objFile.Add sFormName, theFile
                End If
            Else

                tStream.Type = 1
                tStream.Mode = 3
                tStream.Open
                Data_5xsoft.Position = iInfoEnd
                Data_5xsoft.CopyTo tStream, iFormStart - iInfoEnd -3
                tStream.Position = 0
                tStream.Type = 2
                tStream.Charset = "utf-8" '*********************************这里原来是gb2312

                sFormValue = tStream.ReadText
                tStream.Close
                If objForm.Exists(sFormName) Then
                    objForm(sFormName) = objForm(sFormName)&",   "&sFormValue
                Else
                    objForm.Add sFormName, sFormValue
                End If
            End If
            iFormStart = iFormStart + iStart + 1
        Wend
        RequestData = ""
        Set tStream = Nothing
    End Sub

    Private Sub Class_Terminate
        If Request.TotalBytes>0 Then
            objForm.RemoveAll
            objFile.RemoveAll
            Set objForm = Nothing
            Set objFile = Nothing
            Data_5xsoft.Close
            Set Data_5xsoft = Nothing
        End If
    End Sub


    Private Function GetFilePath(FullPath)
        If FullPath <> "" Then
            GetFilePath = Left(FullPath, InStrRev(FullPath, "\"))
        Else
            GetFilePath = ""
        End If
    End Function

    Private Function GetFileName(FullPath)
        If FullPath <> "" Then
            GetFileName = Mid(FullPath, InStrRev(FullPath, "\") + 1)
        Else
            GetFileName = ""
        End If
    End Function

End Class

Class FileInfo
    Dim FormName, FileName, FilePath, FileSize, FileType, FileStart

    Private Sub Class_Initialize
        FileName = ""
        FilePath = ""
        FileSize = 0
        FileStart = 0
        FormName = ""
        FileType = ""
    End Sub

    Public Function SaveAs(FullPath)
        Dim dr, ErrorChar, i
        SaveAs = TRUE
        If Trim(fullpath) = "" Or FileStart = 0 Or FileName = "" Or Right(fullpath, 1) = "/" Then Exit Function
        Set dr = CreateObject("Adodb.Stream")
        dr.Mode = 3
        dr.Type = 1
        dr.Open
        Data_5xsoft.position = FileStart
        Data_5xsoft.copyto dr, FileSize
        dr.SaveToFile FullPath, 2
        dr.Close
        Set dr = Nothing
        SaveAs = FALSE
    End Function

End Class
%>

﻿<%
    dim upfile_5xSoft_Stream

    Class upload_5xSoft

    dim Form,File,Version

    Private Sub Class_Initialize
    dim iStart,iFileNameStart,iFileNameEnd,iEnd,vbEnter,iFormStart,iFormEnd,theFile
    dim strDiv,mFormName,mFormValue,mFileName,mFileSize,mFilePath,iDivLen,mStr
    Version=""
    if Request.TotalBytes<1 then Exit Sub
    set Form=CreateObject("Scripting.Dictionary")
    set File=CreateObject("Scripting.Dictionary")
    set upfile_5xSoft_Stream=CreateObject("Adodb.Stream")
    upfile_5xSoft_Stream.mode=3
    upfile_5xSoft_Stream.type=1
    upfile_5xSoft_Stream.open
    upfile_5xSoft_Stream.write Request.BinaryRead(Request.TotalBytes)

    vbEnter=Chr(13)&Chr(10)
    iDivLen=inString(1,vbEnter)+1
    strDiv=subString(1,iDivLen)
    iFormStart=iDivLen
    iFormEnd=inString(iformStart,strDiv)-1
    while iFormStart < iFormEnd
    iStart=inString(iFormStart,"name=""")
    iEnd=inString(iStart+6,"""")
    mFormName=subString(iStart+6,iEnd-iStart-6)
    iFileNameStart=inString(iEnd+1,"filename=""")
    if iFileNameStart>0 and iFileNameStart<iFormEnd then
    iFileNameEnd=inString(iFileNameStart+10,"""")
    mFileName=subString(iFileNameStart+10,iFileNameEnd-iFileNameStart-10)
    iStart=inString(iFileNameEnd+1,vbEnter&vbEnter)
    iEnd=inString(iStart+4,vbEnter&strDiv)
    if iEnd>iStart then
    mFileSize=iEnd-iStart-4
    else
    mFileSize=0
    end if
    set theFile=new FileInfo
    theFile.FileName=getFileName(mFileName)
    theFile.FilePath=getFilePath(mFileName)
    theFile.FileSize=mFileSize
    theFile.FileStart=iStart+4
    theFile.FormName=FormName
    file.add mFormName,theFile
    else
    iStart=inString(iEnd+1,vbEnter&vbEnter)
    iEnd=inString(iStart+4,vbEnter&strDiv)

    if iEnd>iStart then
    mFormValue=subString(iStart+4,iEnd-iStart-4)
    else
    mFormValue=""
    end if
    form.Add mFormName,mFormValue
    end if

    iFormStart=iformEnd+iDivLen
    iFormEnd=inString(iformStart,strDiv)-1
    wend
    End Sub

    Private Function subString(theStart,theLen)
    dim i,c,stemp
    upfile_5xSoft_Stream.Position=theStart-1
    stemp=""
    for i=1 to theLen
    if upfile_5xSoft_Stream.EOS then Exit for
    c=ascB(upfile_5xSoft_Stream.Read(1))
    If c > 127 Then
    if upfile_5xSoft_Stream.EOS then Exit for
    stemp=stemp&Chr(AscW(ChrB(AscB(upfile_5xSoft_Stream.Read(1)))&ChrB(c)))
    i=i+1
    else
    stemp=stemp&Chr(c)
    End If
    Next
    subString=stemp
    End function

    Private Function inString(theStart,varStr)
    dim i,j,bt,theLen,str
    InString=0
    Str=toByte(varStr)
    theLen=LenB(Str)
    for i=theStart to upfile_5xSoft_Stream.Size-theLen
    if i>upfile_5xSoft_Stream.size then exit Function
    upfile_5xSoft_Stream.Position=i-1
    if AscB(upfile_5xSoft_Stream.Read(1))=AscB(midB(Str,1)) then
    InString=i
    for j=2 to theLen
    if upfile_5xSoft_Stream.EOS then
    inString=0
    Exit for
    end if
    if AscB(upfile_5xSoft_Stream.Read(1))<>AscB(MidB(Str,j,1)) then
    InString=0
    Exit For
    end if
    next
    if InString<>0 then Exit Function
    end if
    next
    End Function

    Private Sub Class_Terminate
    form.RemoveAll
    file.RemoveAll
    set form=nothing
    set file=nothing
    upfile_5xSoft_Stream.close
    set upfile_5xSoft_Stream=nothing
    End Sub


    Private function GetFilePath(FullPath)
    If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, "\"))
    Else
    GetFilePath = ""
    End If
    End function

    Private function GetFileName(FullPath)
    If FullPath <> "" Then
    GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
    Else
    GetFileName = ""
    End If
    End function

    Private function toByte(Str)
    dim i,iCode,c,iLow,iHigh
    toByte=""
    For i=1 To Len(Str)
    c=mid(Str,i,1)
    iCode =Asc(c)
    If iCode<0 Then iCode = iCode + 65535
    If iCode>255 Then
    iLow = Left(Hex(Asc(c)),2)
    iHigh =Right(Hex(Asc(c)),2)
    toByte = toByte & chrB("&H"&iLow) & chrB("&H"&iHigh)
    Else
    toByte = toByte & chrB(AscB(c))
    End If
    Next
    End function
    End Class


    Class FileInfo
    dim FormName,FileName,FilePath,FileSize,FileStart
    Private Sub Class_Initialize
    FileName = ""
    FilePath = ""
    FileSize = 0
    FileStart= 0
    FormName = ""
    End Sub

    Public function SaveAs(FullPath)
    dim dr,ErrorChar,i
    SaveAs=1
    if trim(fullpath)="" or FileSize=0 or FileStart=0 or FileName="" then exit function
    if FileStart=0 or right(fullpath,1)="/" then exit function
    set dr=CreateObject("Adodb.Stream")
    dr.Mode=3
    dr.Type=1
    dr.Open
    upfile_5xSoft_Stream.position=FileStart-1
    upfile_5xSoft_Stream.copyto dr,FileSize
	
    dr.SaveToFile FullPath,2
    dr.Close
    set dr=nothing
    SaveAs=0
    end function
    End Class 
%>
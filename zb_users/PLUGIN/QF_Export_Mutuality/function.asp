<%
'******************************************************************************
Function QF_GetImgSrc(str)
	Dim tmp, objRegExp, Matches, Match
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global = false
	objRegExp.Pattern = "<img [^>]*?\bsrc=(['""]?)([^'""\x20]+)\1 [^>]*?>"
	Set Matches = objRegExp.Execute(str)
	For Each Match In Matches
	tmp = tmp & Match.SubMatches(1)
	Next
	'If tmp = "" Then tmp = ZC_BLOG_HOST & "THEMES/product/STYLE/images/blank.jpg" '图片不存在时的默认图片
	QF_GetImgSrc = tmp
End Function 
'******************************************************************************
'Function MiniImgSrc(ByVal strUrl,ByVal miniwidth,ByVal miniheight)
'	Dim strOriginalPath,strFileName
'	If strUrl="" Then
'		MiniImgSrc=ZC_BLOG_HOST & "plugin/zblog_alipay/noimg.jpg"
'	Else
'		If InStr(LCase(strUrl),LCase(ZC_BLOG_HOST & ZC_UPLOAD_DIRECTORY))>0 Then
'			strOriginalPath=Replace(LCase(strUrl),LCase(ZC_BLOG_HOST),BlogPath)
'			strOriginalPath=QF_URLDecode(strOriginalPath)
'			strFileName=LCase(Mid(strOriginalPath,InStrRev(Replace(strOriginalPath,"/","\"),"\")+1))
'			MiniImgSrc=CreatMini(strOriginalPath,miniwidth,miniheight)
'		Else
'			MiniImgSrc=strUrl
'		End If
'	End If
'End Function
'******************************************************************************
Function QF_URLDecode(ByVal strUrl)
	'On Error Resume Next

	If Not InStr(strUrl,"%")>0 Then
		QF_URLDecode=strUrl
		Exit Function
	End If

	Dim ObjUrlCode
	Set ObjURLCode = New QF_UTFURLCode
		QF_URLDecode=ObjURLCode.UrlDecode(strUrl)
	Set objURLCode = Nothing

	'Err.Clear
End Function
'******************************************************************************
Function QF_CreatMini(ByVal strUrl,ByVal miniwidth,ByVal miniheight)
'On Error Resume Next
	Dim strOriginalPath,strFileName
	If strUrl="" Then
		QF_CreatMini=ZC_BLOG_HOST & "zb_users/plugin/zblog_alipay/noimg.jpg"
	Else
		If InStr(LCase(strUrl),LCase(ZC_BLOG_HOST & ZC_UPLOAD_DIRECTORY))>0 Then
			strOriginalPath=Replace(LCase(strUrl),LCase(ZC_BLOG_HOST),BlogPath)
			strOriginalPath=QF_URLDecode(strOriginalPath)
			'strFileName=LCase(Mid(strOriginalPath,InStrRev(Replace(strOriginalPath,"/","\"),"\")+1))

     

	Dim Jpeg,h,w,m,n,strMiniPath
	strMiniPath=Left(strOriginalPath,InStrRev(strOriginalPath,".")-1)&"_"& miniwidth&"_"&miniheight& "_miniimg.jpg"
QF_CreatMini=Replace(strMiniPath,BlogPath,ZC_BLOG_HOST)
	Set Jpeg = Server.CreateObject("Persits.Jpeg")

		If QF_FileExists(strMiniPath) Then
			Jpeg.Open strMiniPath
				If Jpeg.OriginalWidth=miniwidth And Jpeg.OriginalHeight=miniheight Then
					Jpeg.Close
					Exit Function
				End If
			Jpeg.Close
		End If

		Jpeg.Open strOriginalPath

			m= miniwidth / miniheight
			n= Jpeg.OriginalWidth / Jpeg.OriginalHeight

			If n < m Then
				w = miniwidth
				h = CInt(w / n)
				Jpeg.Width = w
				Jpeg.Height = h
				jpeg.Crop 0, 0, w, CInt(w / m)
			ElseIf n > m Then
				h = miniheight
				w = CInt(h * n)
				Jpeg.Width = w
				Jpeg.Height = h
				jpeg.Crop CInt((w - m * h)/2), 0, CInt(m * h + (w - m * h)/2), h
			Else
				w = miniwidth
				h= CInt(w / n)
				Jpeg.Width = w
				Jpeg.Height = h
			End If

			Jpeg.Save strMiniPath

		Jpeg.Close

	Set Jpeg=Nothing
			
		Else
			QF_CreatMini=strUrl
		End If
	End If

'Err.Clear
End Function

'******************************************************************************
Class QF_UTFURLCode

	Public Function UrlEncode(ByVal UTFStr)
		UrlEncode = Server.UrlEncode(UTFStr)
	End Function


	Public Function UrlDecode(ByVal UTFStr)
	Dim Dig,GBStr
		For Dig=1 To len(UTFStr)
			If mid(UTFStr,Dig,1)="%" Then
				If len(UTFStr) >= Dig+8 Then
					GBStr=GBStr & ConvChinese(mid(UTFStr,Dig,9))
					Dig=Dig+8
				Else
					GBStr=GBStr & mid(UTFStr,Dig,1)
				End If
			Else
				GBStr=GBStr & mid(UTFStr,Dig,1)
			End If
		Next
		UrlDecode=GBStr
	End Function 

	Private Function ConvChinese(ByVal x) 
	Dim a,i,j,DigS,Unicode
		A=split(mid(x,2),"%")
		i=0
		j=0
		
		For i=0 To ubound(A) 
			A(i)=c16To2(A(i))
		Next
			
		For i=0 To ubound(A)-1
			DigS=instr(A(i),"0")
			Unicode=""
			For j=1 To DigS-1
				If j=1 Then 
					A(i)=right(A(i),len(A(i))-DigS)
					Unicode=Unicode & A(i)
				Else
					i=i+1
					A(i)=right(A(i),len(A(i))-2)
					Unicode=Unicode & A(i) 
				End If 
			Next
			
			If len(c2To16(Unicode))=4 Then
				ConvChinese=ConvChinese & chrw(int("&H" & c2To16(Unicode)))
			Else
				ConvChinese=ConvChinese & chr(int("&H" & c2To16(Unicode)))
			End If
		Next
	End Function 

	Private Function c16To2(ByVal x)
	'这个函数是用来转换16进制到2进制的，可以是任何长度的，一般转换UTF-8的时候是两个长度，比如A9
	'比如：输入“C2”，转化成“11000010”,其中1100是"c"是10进制的12（1100），那么2（10）不足4位要补齐成（0010）。
	Dim tempstr
	Dim i:i=0'临时的指针 
	For i=1 To len(trim(x))
	  tempstr= c10To2(cint(int("&h" & mid(x,i,1))))
	  Do While len(tempstr)<4
	   tempstr="0" & tempstr'如果不足4位那么补齐4位数
	  Loop
	  c16To2=c16To2 & tempstr
	Next
	End Function 

	Private Function c2To16(ByVal x)
	  '2进制到16进制的转换，每4个0或1转换成一个16进制字母，输入长度当然不可能不是4的倍数了 
	  Dim i:i=1'临时的指针
	  For i=1 To len(x)  Step 4
	   c2To16=c2To16 & hex(c2To10(mid(x,i,4)))
	  Next
	End Function

	Private Function c2To10(ByVal x)
	  '单纯的2进制到10进制的转换，不考虑转16进制所需要的4位前零补齐。
	  '因为这个函数很有用！以后也会用到，做过通讯和硬件的人应该知道。
	  '这里用字符串代表二进制
	   c2To10=0
	   If x="0" Then Exit Function'如果是0的话直接得0就完事
	   Dim i:i=0'临时的指针
	   For i= 0 To len(x) -1'否则利用8421码计算，这个从我最开始学计算机的时候就会，好怀念当初教我们的谢道建老先生啊！
		If mid(x,len(x)-i,1)="1" Then c2To10=c2To10+2^(i)
	   Next
	End Function

	Private Function c10To2(ByVal x)
	'10进制到2进制的转换
	  Dim sign, result
	  result = ""
	  '符号
	  sign = sgn(x)
	  x = abs(x)
	  If x = 0 Then
		c10To2 = 0
		Exit Function
	  End If
	  Do until x = "0"
		result = result & (x mod 2)
		x = x \ 2
	  Loop
	  result = strReverse(result)
	  If sign = -1 Then
		c10To2 = "-" & result
	  Else
		c10To2 = result
	  End If
	End Function

End Class

'******************************************************************************
Function QF_FileExists(ByVal strFilePath)
	'On Error Resume Next

	Dim FileExists
	FileExists=False

	Dim objFSO
	Set objFSO=CreateObject("Scripting.FileSystemObject")
		If objFSO.FileExists(strFilePath) Then
			FileExists=True
		End If
	Set objFSO=Nothing

	QF_FileExists=FileExists

	'Err.Clear
End Function
'******************************************************************************
Function Miniimg_Del(byval ID,byval AuthorID,byval FileSize,byval FileName,byval PostTime,byval DirByTime)
	'On Error Resume Next

	Call CheckParameter(ID,"int",0)

	Dim objRS,strFilePath

	Set objRS=objConn.Execute("SELECT * FROM [blog_UpLoad] WHERE [ul_ID] = " & ID)

	If (Not objRS.bof) And (Not objRS.eof) Then

		If objRS("ul_DownNum")=-1 Then
			objConn.Execute("UPDATE [blog_Upload] SET [ul_DownNum]=(0) WHERE [ul_DownNum]=("& ID &")")
		End If

		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")

		strFilePath = BlogPath & "/"& ZC_UPLOAD_DIRECTORY &"/" & objRS("ul_FileName")
		strFilePath=Left(strFilePath,InStrRev(strFilePath,".")-1) & "_miniimg.jpg"
		If fso.FileExists( strFilePath ) Then
			fso.DeleteFile( strFilePath )
		End If

		strFilePath = BlogPath & "/"& ZC_UPLOAD_DIRECTORY & "/" & Year(objRS("ul_PostTime")) & "/" & Month(objRS("ul_PostTime")) &"/" & objRS("ul_FileName")
		strFilePath=Left(strFilePath,InStrRev(strFilePath,".")-1) & "_miniimg.jpg"
		If fso.FileExists( strFilePath ) Then
			fso.DeleteFile( strFilePath )
		End If

		Set fso = Nothing

	Else

		Exit Function

	End If

	objRS.Close
	Set objRS=Nothing

	'Err.Clear
End Function
%>
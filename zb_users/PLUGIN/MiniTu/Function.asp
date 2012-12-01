<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8及以上的版本
'// 插件制作:  zblog管理员之家(www.zbadmin.com)
'// 备    注:   Mini缩略图插件代码
'// 最后修改：   2012/2/20
'// 最后版本:    0.1
'///////////////////////////////////////////////////////////////////////////////
%>
<%
'*********************************************************
' 目的：    生成指定大小的缩略图
'*********************************************************
Function MiniTu_CreatMini(ByVal strOriginalPath,ByVal strMiniPath)
On Error Resume Next

	Dim Jpeg,h,w,m,n

	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	If Err.Number<>0 Then Exit Function
		If MiniTu_FileExists(strMiniPath) Then
			Jpeg.Open strMiniPath
				If Jpeg.OriginalWidth=MiniTu_MiniImgWidth And Jpeg.OriginalHeight=MiniTu_MiniImgHeight Then
					Jpeg.Close
					Exit Function
				End If
			Jpeg.Close
		End If

		Jpeg.Open strOriginalPath
			n= Jpeg.OriginalWidth / Jpeg.OriginalHeight
			If MiniTu_MiniImgHeight=0 Then
				w = MiniTu_MiniImgWidth
				h = CInt(w / n)
				Jpeg.Width = w
				Jpeg.Height = h
				jpeg.Crop 0, 0, w, h
			else
				m= MiniTu_MiniImgWidth / MiniTu_MiniImgHeight
				

				If n < m Then
					w = MiniTu_MiniImgWidth
					h = CInt(w / n)
					Jpeg.Width = w
					Jpeg.Height = h
					jpeg.Crop 0, 0, w, CInt(w / m)
				ElseIf n > m Then
					h = MiniTu_MiniImgHeight
					w = CInt(h * n)
					Jpeg.Width = w
					Jpeg.Height = h
					jpeg.Crop CInt((w - m * h)/2), 0, CInt(m * h + (w - m * h)/2), h
				Else
					w = MiniTu_MiniImgWidth
					h= CInt(w / n)
					Jpeg.Width = w
					Jpeg.Height = h
				End If
			End If

			Jpeg.Save strMiniPath

		Jpeg.Close

	Set Jpeg=Nothing

'Err.Clear
End Function

'*********************************************************
' 目的：    检查某目录下的某文件是否存在
'*********************************************************
Function MiniTu_FileExists(ByVal strFilePath)
	'On Error Resume Next

	Dim FileExists
	FileExists=False

	Dim objFSO
	Set objFSO=CreateObject("Scripting.FileSystemObject")
		If objFSO.FileExists(strFilePath) Then
			FileExists=True
		End If
	Set objFSO=Nothing

	MiniTu_FileExists=FileExists

	'Err.Clear
End Function

'*********************************************************
' 目的：    URL解码
'*********************************************************
Function MiniTu_URLDecode(ByVal strUrl)
	'On Error Resume Next

	If Not InStr(strUrl,"%")>0 Then
		MiniTu_URLDecode=strUrl
		Exit Function
	End If

	Dim ObjUrlCode
	Set ObjURLCode = New MiniTu_UTFURLCode
		MiniTu_URLDecode=ObjURLCode.UrlDecode(strUrl)
	Set objURLCode = Nothing

	'Err.Clear
End Function

'=======================================================
'类名：    UrlDecode
'功能：      将utf-8编码解码为中文
'UTFStr:     需要编码的字符，在utf-8和gb2312两种页面编码中均可使用
'备注:       此类改编自网络
'=======================================================
Class MiniTu_UTFURLCode

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
'=======================================================
%>
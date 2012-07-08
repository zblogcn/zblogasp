<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_function.asp
'// 开始时间:    2004.07.28
'// 最后修改:    
'// 备    注:    函数模块
'///////////////////////////////////////////////////////////////////////////////




'*********************************************************
' 目的：    显示错误页面
' 输入：    id
' 返回：    无
'*********************************************************
Dim ShowError_Custom
Sub ShowError(id)
	If IsEmpty(ShowError_Custom)=False Then
		Execute(ShowError_Custom)
		Exit Sub
	End If
	Response.Redirect ZC_BLOG_HOST & "zb_system/function/c_error.asp?errorid=" & id & "&number=" & Err.Number & "&description=" & Server.URLEncode(Err.Description) & "&source=" & Server.URLEncode(Err.Source) & "&sourceurl="  &Server.URLEncode(Request.ServerVariables("Http_Referer")) 
End Sub
'*********************************************************




'*********************************************************
' 目的：    XML-RPC显示错误页面
'*********************************************************
Function RespondError(faultCode,faultString)

	Dim strXML
	Dim strError

	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><fault><value><struct><member><name>faultCode</name><value><int>$1</int></value></member><member><name>faultString</name><value><string>$2</string></value></member></struct></value></fault></methodResponse>"

	strError=strXML
	strError=Replace(strError,"$1",TransferHTML(faultCode,"[html-format]"))
	strError=Replace(strError,"$2",TransferHTML(faultString,"[html-format]"))

	Response.Clear
	Response.BinaryWrite ChrB(&HEF) & ChrB(&HBB) & ChrB(&HBF)
	Response.Write strError
	Response.End

End Function
'*********************************************************




'*********************************************************
' 目的：    检查正则式
' 输入：    id
' 返回：    成功为True
'*********************************************************
Function CheckRegExp(source,para)

	If para="[username]" Then
		para="^[.A-Za-z0-9\u4e00-\u9fa5]+$"
	End If
	If para="[password]" Then
		para="^[a-z0-9]+$"
	End If
	If para="[email]" Then
		para="^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*\.)+[a-zA-Z]*)$"
	End If
	If para="[homepage]" Then
		para="^[a-zA-Z]+://[a-zA-Z0-9\_\-\.\&\?/:=#\u4e00-\u9fa5]+?/*$"
	End If
	If para="[nojapan]" Then
		para="[\u3040-\u30ff]+"
	End If
	If para="[guid]" Then
		para="^\w{8}\-\w{4}\-\w{4}\-\w{4}\-\w{12}$"
	End If

	Dim re
	Set re = New RegExp
	re.Global = True
	re.Pattern = para
	re.IgnoreCase = False
	CheckRegExp = re.Test(source)

End Function
'*********************************************************




'*********************************************************
' 目的：    检查参数
' 返回：    出错则转到ShowError(3)
'*********************************************************
Function CheckParameter(byRef source,strType,default)

	On Error Resume Next

	If strType="int" Then

		'数值
		If IsNull(source) Then
			source=default
		ElseIf IsEmpty(source) Then
			source=default
		ElseIf IsNumeric(source) Then
			source=CLng(source)
		ElseIf source="" Then
			source=default
		Else
			Call ShowError(3)
		End if
		If Err.Number<>0 Then Call ShowError(3)

		CheckParameter=True

	ElseIf  strType="dtm" Then

		'日期
		If IsNull(source) Then
			source=default
		ElseIf IsEmpty(source) Then
			source=default
		ElseIf IsDate(source) Then
			source=source
			Call FormatDateTime(source,vbLongDate)
			Call FormatDateTime(source,vbShortDate)
		ElseIf source="" Then
			source=default
		Else
			Call ShowError(3)
		End if
		If Err.Number<>0 Then Call ShowError(3)

		CheckParameter=True

	ElseIf strType="sql" Then

		'SQL
		If IsNull(source) Or Trim(source)="" Or IsEmpty(source) Then
			source=default
		Else
			source=CStr(Replace(source,Chr(39),Chr(39)&Chr(39)))
		End If

	ElseIf strType="bool" Then

		'Boolean
		source=CBool(source)

		If Err.Number<>0 Then
			Err.Clear
			If IsEmpty(source)=True Then
				source=True
			Else
				source=False
			End If
		End If

	Else
		Call ShowError(0)
	End If

End Function
'*********************************************************




'*********************************************************
' 目的：    检查引用
' 返回：    无
'*********************************************************
Sub CheckReference(strDestination)

	Exit Sub

	Dim strReferer
	strReferer=CStr(Request.ServerVariables("HTTP_REFERER"))

	If Instr(strReferer,ZC_BLOG_HOST)=0 Then 
		ShowError(5)
	End If

End Sub
'*********************************************************




'*********************************************************
' 目的：    搜索字符串
' 返回：    
' 备注:     不区分大小写
'*********************************************************
Function Search(strText,strQuestion)

	Dim s
	Dim i
	Dim j

	s=strText
	i=Instr(1,s,strQuestion,vbTextCompare)
	If i>0 Then
		s=Left(s,i+Len(strQuestion)+100)
		s=Right(s,Len(strQuestion)+200)
	Else
		s=""
	End If

	If s<>"" Then
		i=1
		Do While InStr(i,s,strQuestion,vbTextCompare)>0
			j=InStr(i,s,strQuestion,vbTextCompare)
			If Len(s)-j-Len(strQuestion)<0 Then
				s=Left(s,j-1) & "<b style='color:#FF6347'>" & strQuestion & "</b>"
				Exit Do
			Else
				s=Left(s,j-1) & "<b style='color:#FF6347'>" & strQuestion & "</b>" & Right(s,Len(s)-j-Len(strQuestion)+1)
			End If
			i=j+Len("<b style='color:#FF6347'>" & strQuestion & "</b>")-1
			If i>=Len(s) Then Exit Do
		Loop

	End If

	If s="" Then
		Search=strText
	Else
		Search=s
	End If

End Function
'*********************************************************




'*********************************************************
' 目的：    检查引用
' 输入：    SQL值（引用）
' 返回：    
'*********************************************************
Function FilterSQL(strSQL)

	FilterSQL=CStr(Replace(strSQL,chr(39),chr(39)&chr(39)))

End Function
'*********************************************************




'*********************************************************
' 目的：    检查引用
' 输入：    
' 输入：    要替换的字符代号
' 返回：    
'*********************************************************
Function TransferHTML(ByVal source,para)

	Dim objRegExp

	'先换"&"
	If Instr(para,"[&]")>0 Then  source=Replace(source,"&","&amp;")
	If Instr(para,"[<]")>0 Then  source=Replace(source,"<","&lt;")
	If Instr(para,"[>]")>0 Then  source=Replace(source,">","&gt;")
	If Instr(para,"[""]")>0 Then source=Replace(source,"""","&quot;")
	If Instr(para,"[space]")>0 Then source=Replace(source," ","&nbsp;")
	If Instr(para,"[enter]")>0 Then
		source=Replace(source,vbCrLf,"<br/>")
		source=Replace(source,vbLf,"<br/>")
	End If
	If Instr(para,"[vbCrlf]")>0 And ZC_AUTO_NEWLINE Then 

		Set objRegExp=New RegExp
		objRegExp.IgnoreCase =True
		objRegExp.Global=True

		objRegExp.Pattern="((</?form[^\n<]*>)|(<select[^\n<]*>)|(<textarea[^\n<]*>)|(</?option[^\n<]*>)|(</?dl[^\n<]*>)|(</?dt[^\n<]*>)|(</?dd[^\n<]*>)|(</?th[^\n<]*>)|(</?tr[^\n<]*>)|(</?td[^\n<]*>)|(</?tbody[^\n<]*>)|(</?table[^\n<]*>)|(</?hr[^\n<]*>)|(</?div[^\n<]*>)|(</?ul[^\n<]*>)|(</?li[^\n<]*>)|(</?ol[^\n<]*>)|(</?h1[^\n<]*>)|(</?h2[^\n<]*>)|(</?h3[^\n<]*>)|(</?h4[^\n<]*>)|(</?h5[^\n<]*>)|(</?h6[^\n<]*>)|(</?p[^\n<]*>))(\x20*(\r\n|\n))"

		source= objRegExp.Replace(source,"$1")

		objRegExp.Pattern="(\r\n|\n)"
		source= objRegExp.Replace(source,"<br/>")

		source=Replace(source,"<html>","")
		source=Replace(source,"</html>","")
		source=Replace(source,"<body>","")
		source=Replace(source,"</body>","")

	End If
	If Instr(para,"[vbTab]")>0 Then source=Replace(source,vbTab,"&nbsp;&nbsp;")
	If Instr(para,"[upload]")>0 Then
		source=Replace(source,"src=""upload/","src="""& ZC_BLOG_HOST & "zb_users/" & ZC_UPLOAD_DIRECTORY & "/")
		source=Replace(source,"href=""upload/","href="""& ZC_BLOG_HOST & "zb_users/" &  ZC_UPLOAD_DIRECTORY & "/")
		source=Replace(source,"value=""upload/","value="""& ZC_BLOG_HOST & "zb_users/" &  ZC_UPLOAD_DIRECTORY & "/")
		source=Replace(source,"href=""http://upload/","href="""& ZC_BLOG_HOST & "zb_users/" &  ZC_UPLOAD_DIRECTORY & "/")
		source=Replace(source,"(this.nextSibling,'upload/","(this.nextSibling,'"& ZC_BLOG_HOST & "zb_users/" &  ZC_UPLOAD_DIRECTORY & "/")

		source=Replace(source,"src=""image/face/","src="""& ZC_BLOG_HOST & "zb_system/" &  "image/face/")
	End If
	If Instr(para,"[anti-upload]")>0 Then
		source=Replace(source,"src="""& ZC_BLOG_HOST & "zb_users/" &  ZC_UPLOAD_DIRECTORY & "/","src=""upload/")
		source=Replace(source,"href="""& ZC_BLOG_HOST & "zb_users/" &  ZC_UPLOAD_DIRECTORY & "/","href=""upload/")
		source=Replace(source,"value="""& ZC_BLOG_HOST & "zb_users/" &  ZC_UPLOAD_DIRECTORY & "/","value=""upload/")
		source=Replace(source,"href="""& ZC_BLOG_HOST & "zb_users/" &  ZC_UPLOAD_DIRECTORY & "/","href=""http://upload/")
		source=Replace(source,"(this.nextSibling,'"& ZC_BLOG_HOST & "zb_users/" &  ZC_UPLOAD_DIRECTORY & "/","(this.nextSibling,'upload/")

		source=Replace(source,"src="""& ZC_BLOG_HOST & "zb_system/" & "image/face/","src=""image/face/")
	End If
	If Instr(para,"[zc_blog_host]")>0 Then
		source=Replace(source,"<#ZC_BLOG_HOST#>",ZC_BLOG_HOST)
	End If
	If Instr(para,"[anti-zc_blog_host]")>0 Then
		source=Replace(source,ZC_BLOG_HOST,"<#ZC_BLOG_HOST#>")
	End If
	If Instr(para,"[no-asp]")>0 Then
		source=Replace(source,"<"&"%","&lt;"&"%")
		source=Replace(source,"%"&">","%"&"&gt;")
	End If
	If ZC_COMMENT_NOFOLLOW_ENABLE And Instr(para,"[nofollow]")>0 Then
		source=Replace(source,"<a ","<a rel=""nofollow"" ")
	End If
	If Instr(para,"[nohtml]")>0  Then

		Set objRegExp=New RegExp
		objRegExp.IgnoreCase =True
		objRegExp.Global=True
		objRegExp.Pattern="<[^>]*>"
		source= objRegExp.Replace(source,"")

	End If
	If Instr(para,"[filename]")>0  Then
		source=Replace(source,"/","")
		source=Replace(source,"\","")
		source=Replace(source,":","")
		source=Replace(source,"?","")
		source=Replace(source,"*","")
		source=Replace(source,"""","")
		source=Replace(source,"<","")
		source=Replace(source,">","")
		source=Replace(source,"|","")
		source=Replace(source," ","")
	End If
	If Instr(para,"[normalname]")>0  Then
		source=Replace(source,"$","")
		source=Replace(source,"(","")
		source=Replace(source,")","")
		source=Replace(source,"*","")
		source=Replace(source,"+","")
		source=Replace(source,",","")
		source=Replace(source,"[","")
		source=Replace(source,"]","")
		source=Replace(source,"{","")
		source=Replace(source,"}","")
		source=Replace(source,"?","")
		source=Replace(source,"\","")
		source=Replace(source,"^","")
		source=Replace(source,"|","")
		source=Replace(source,":","")
		source=Replace(source,"""","")
		source=Replace(source,"'","")
	End If
	If Instr(para,"[textarea]")>0 Then
		'Set objRegExp=New RegExp
		'objRegExp.IgnoreCase =True
		'objRegExp.Global=True
		'objRegExp.Pattern="(&)([#a-z0-9]{2,10})(;)"
		'source= objRegExp.Replace(source,"&amp;$2$3")
		source=Replace(source,"&","&amp;")
		source=Replace(source,"%","&#037;")
		source=Replace(source,"<","&lt;")
		source=Replace(source,">","&gt;")
	End If
	If ZC_JAPAN_TO_HTML And Instr(para,"[japan-html]")>0 Then
		source=Replace(source,"ガ","&#12460;")
		source=Replace(source,"ギ","&#12462;")
		source=Replace(source,"ア","&#12450;")
		source=Replace(source,"ゲ","&#12466;")
		source=Replace(source,"ゴ","&#12468;")
		source=Replace(source,"ザ","&#12470;")
		source=Replace(source,"ジ","&#12472;")
		source=Replace(source,"ズ","&#12474;")
		source=Replace(source,"ゼ","&#12476;")
		source=Replace(source,"ゾ","&#12478;")
		source=Replace(source,"ダ","&#12480;")
		source=Replace(source,"ヂ","&#12482;")
		source=Replace(source,"ヅ","&#12485;")
		source=Replace(source,"デ","&#12487;")
		source=Replace(source,"ド","&#12489;")
		source=Replace(source,"バ","&#12496;")
		source=Replace(source,"パ","&#12497;")
		source=Replace(source,"ビ","&#12499;")
		source=Replace(source,"ピ","&#12500;")
		source=Replace(source,"ブ","&#12502;")
		source=Replace(source,"ブ","&#12502;")
		source=Replace(source,"プ","&#12503;")
		source=Replace(source,"ベ","&#12505;")
		source=Replace(source,"ペ","&#12506;")
		source=Replace(source,"ボ","&#12508;")
		source=Replace(source,"ポ","&#12509;")
		source=Replace(source,"ヴ","&#12532;")
	End If
	If ZC_JAPAN_TO_HTML And Instr(para,"[html-japan]")>0 Then
		source=Replace(source,"&#12460;","ガ")
		source=Replace(source,"&#12462;","ギ")
		source=Replace(source,"&#12450;","ア")
		source=Replace(source,"&#12466;","ゲ")
		source=Replace(source,"&#12468;","ゴ")
		source=Replace(source,"&#12470;","ザ")
		source=Replace(source,"&#12472;","ジ")
		source=Replace(source,"&#12474;","ズ")
		source=Replace(source,"&#12476;","ゼ")
		source=Replace(source,"&#12478;","ゾ")
		source=Replace(source,"&#12480;","ダ")
		source=Replace(source,"&#12482;","ヂ")
		source=Replace(source,"&#12485;","ヅ")
		source=Replace(source,"&#12487;","デ")
		source=Replace(source,"&#12489;","ド")
		source=Replace(source,"&#12496;","バ")
		source=Replace(source,"&#12497;","パ")
		source=Replace(source,"&#12499;","ビ")
		source=Replace(source,"&#12500;","ピ")
		source=Replace(source,"&#12502;","ブ")
		source=Replace(source,"&#12502;","ブ")
		source=Replace(source,"&#12503;","プ")
		source=Replace(source,"&#12505;","ベ")
		source=Replace(source,"&#12506;","ペ")
		source=Replace(source,"&#12508;","ボ")
		source=Replace(source,"&#12509;","ポ")
		source=Replace(source,"&#12532;","ヴ")
	End If
	If Instr(para,"[html-format]")>0 Then
		source=Replace(source,"&","&amp;")
		source=Replace(source,"<","&lt;")
		source=Replace(source,">","&gt;")
		source=Replace(source,"""","&quot;")
	End If
	If Instr(para,"[anti-html-format]")>0 Then
		source=Replace(source,"&quot;","""")
		source=Replace(source,"&lt;","<")
		source=Replace(source,"&gt;",">")
		source=Replace(source,"&amp;","&")
	End If
	If Instr(para,"[wapnohtml]")>0 Then
		source=Replace(source,"<br/>",vbCrLf)
		source=Replace(source,"<br>",vbCrLf)
		Set objRegExp=New RegExp
		objRegExp.IgnoreCase =True
		objRegExp.Global=True
		objRegExp.Pattern="(<[^>]*)|([^<]*>)"
		source= objRegExp.Replace(source,"")
		objRegExp.Pattern="(\r\n|\n)"
		source= objRegExp.Replace(source,"<br/>")
	End If

	If Instr(para,"[nbsp-br]")>0 Then
		Set objRegExp=New RegExp
		objRegExp.IgnoreCase =True
		objRegExp.Global=True
		objRegExp.Pattern="&lt;$|&lt;b$|&lt;br$|&lt;br/$"
		source= objRegExp.Replace(source,"")
		objRegExp.Pattern="^br/&gt;|^r/&gt;|^/&gt;|^&gt;"
		source= objRegExp.Replace(source,"")
		objRegExp.Pattern="&lt;br/&gt;"
		source= objRegExp.Replace(source,"<br/>")
		objRegExp.Pattern="&amp;nbsp;"
		source= objRegExp.Replace(source," ")
	End If

	If Instr(para,"[closehtml]")>0 Then
		source=closeHTML(source)
	End If


	TransferHTML=source

End Function
'*********************************************************




'*********************************************************
' 目的：   301 Moved
' 输入：    
' 返回：    
'*********************************************************
Sub RedirectBy301(strURL)

	Response.Clear
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location",strURL
	Response.End

End Sub
'*********************************************************




'*********************************************************
' 目的：   Random Number Create
' 输入：    
' 返回：    
'*********************************************************
Sub CreateVerifyNumber()

	Dim i,j,s,t
	Randomize

	Dim aryVerifyNumber(10000)
	For j=0 To 10000
		s=""
		For i = 0 To 4
			t = Int(Rnd * Len(ZC_VERIFYCODE_STRING))
			s= s & Mid(ZC_VERIFYCODE_STRING,t + 1, 1)
		Next
		aryVerifyNumber(j)=s
	Next

	Application.Lock
	Application(ZC_BLOG_CLSID & "VERIFY_NUMBER")=aryVerifyNumber
	Application.UnLock

End Sub
'*********************************************************




'*********************************************************
' 目的：   Random Number Get
' 输入：    
' 返回：    
'*********************************************************
Function GetVerifyNumber()

	Randomize
	Dim i,j,s,t
	Dim aryVerifyNumber

	Application.Lock
	aryVerifyNumber=Application(ZC_BLOG_CLSID & "VERIFY_NUMBER")
	Application.UnLock

	If IsEmpty(aryVerifyNumber)=True Or IsArray(aryVerifyNumber)=False Then
		Call CreateVerifyNumber()
		Application.Lock
		aryVerifyNumber=Application(ZC_BLOG_CLSID & "VERIFY_NUMBER")
		Application.UnLock
	End If


	For i=0 To 10000
		If (aryVerifyNumber(i)<>"") And (Len(aryVerifyNumber(i))=5) Then 
			GetVerifyNumber=aryVerifyNumber(i)
			Exit For
		End If
	Next

	aryVerifyNumber(i)=aryVerifyNumber(i)&"-"

	If i=5000 Then
		For j=5001 To 10000
			s=""
			For i = 0 To 4
				t = Int(Rnd * Len(ZC_VERIFYCODE_STRING))
				s= s & Mid(ZC_VERIFYCODE_STRING,t + 1, 1)
			Next
			aryVerifyNumber(j)=s
		Next
	End If

	If i=10000 Then
		For j=0 To 5000
			s=""
			For i = 0 To 4
				t = Int(Rnd * Len(ZC_VERIFYCODE_STRING))
				s= s & Mid(ZC_VERIFYCODE_STRING,t + 1, 1)
			Next
			aryVerifyNumber(j)=s
		Next
	End If

	Application.Lock
	Application(ZC_BLOG_CLSID & "VERIFY_NUMBER")=aryVerifyNumber
	Application.UnLock

End Function
'*********************************************************




'*********************************************************
' 目的：   Random Number Check
' 输入：    
' 返回：    
'*********************************************************
Function CheckVerifyNumber(ByVal strNumber)

	Dim i,j,s,t
	Dim aryVerifyNumber

	Application.Lock
	aryVerifyNumber=Application(ZC_BLOG_CLSID & "VERIFY_NUMBER")
	Application.UnLock

	If IsEmpty(aryVerifyNumber) Then Exit Function

	strNumber=UCase(strNumber)

	For j=0 To 10000

		If aryVerifyNumber(j)=strNumber & "-" Then

			Randomize
			s=""
			For i = 0 To 4
				t = Int(Rnd * Len(ZC_VERIFYCODE_STRING))
				s= s & Mid(ZC_VERIFYCODE_STRING,t + 1, 1)
			Next
			aryVerifyNumber(j)=s

			Application.Lock
			Application(ZC_BLOG_CLSID & "VERIFY_NUMBER")=aryVerifyNumber
			Application.UnLock

			CheckVerifyNumber=True

			Exit Function

		End If

	Next

End Function
'*********************************************************




'*********************************************************
' 目的：    UBB
' 输入：    
' 输入：    
' 返回：    
'*********************************************************
Function UBBCode(ByVal strContent,strType)

	Dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True

	If ZC_UBB_LINK_ENABLE And Instr(strType,"[link]")>0 Then

		objRegExp.Pattern="(\[URL\])(([a-zA-Z0-9]+?):\/\/[^ :\(\)\f\n\r\t\v]+?)(\[\/URL\])"
		strContent= objRegExp.Replace(strContent,"<a href=""$2"" target=""_blank"">$2</a>")

		objRegExp.Pattern="(\[URL\])([^ :\(\)\f\n\r\t\v]+?)(\[\/URL\])"
		strContent= objRegExp.Replace(strContent,"<a href=""http://$2"" target=""_blank"">$2</a>")

		objRegExp.Pattern="(\[URL=)(([a-zA-Z0-9]+?):\/\/[^ :\(\)\f\n\r\t\v]+?)(\])(.+?)(\[\/URL\])"
		strContent= objRegExp.Replace(strContent,"<a href=""$2"" target=""_blank"">$5</a>")

		objRegExp.Pattern="(\[URL=)([^ :\(\)\f\n\r\t\v]+?)(\])(\S+?)(\[\/URL\])"
		strContent= objRegExp.Replace(strContent,"<a href=""http://$2"" target=""_blank"">$4</a>")

	End If


	If ZC_UBB_LINK_ENABLE And Instr(strType,"[email]")>0 Then

		objRegExp.Pattern="(\[EMAIL\])(\S+\@\S+?)(\[\/EMAIL\])"
		strContent= objRegExp.Replace(strContent,"<a href=""mailto:$2"" >$2</a>")

		objRegExp.Pattern="(\[EMAIL=)(\S+\@\S+?)(\])(.+?)(\[\/EMAIL\])"
		strContent= objRegExp.Replace(strContent,"<a href=""mailto:$2"">$4</a>")

	End If


	If ZC_UBB_FONT_ENABLE And Instr(strType,"[font]")>0 Then

		objRegExp.Pattern="(\[I\])([\u0000-\uffff]+?)(\[\/I\])"
		strContent=objRegExp.Replace(strContent,"<i>$2</i>")

		objRegExp.Pattern="(\[B\])([\u0000-\uffff]+?)(\[\/B\])"
		strContent=objRegExp.Replace(strContent,"<b>$2</b>")

		objRegExp.Pattern="(\[U\])([\u0000-\uffff]+?)(\[\/U\])"
		strContent=objRegExp.Replace(strContent,"<u>$2</u>")

		objRegExp.Pattern="(\[S\])([\u0000-\uffff]+?)(\[\/S\])"
		strContent=objRegExp.Replace(strContent,"<s>$2</s>")

		objRegExp.Pattern="(\[QUOTE\])([\u0000-\uffff]+?)(\[\/QUOTE\])"
		strContent=objRegExp.Replace(strContent,"<blockquote><div class=""quote"">$2"&"</div></blockquote>")

		objRegExp.Pattern="(\[QUOTE=)(.+?)(\])([\u0000-\uffff]+?)(\[\/QUOTE\])"
		strContent= objRegExp.Replace(strContent,"<blockquote><div class=""quote quote2""><div class=""quote-title"">"&ZC_MSG153&" $2</div>$4"&"</div></blockquote>")

		objRegExp.Pattern="(\[REVERT=)(.+?)(\])([\u0000-\uffff]+?)(\[\/REVERT\])"
		strContent= objRegExp.Replace(strContent,"<blockquote><div class=""quote quote3""><div class=""quote-title"">$2</div>$4</div></blockquote>")

	End If


	If ZC_UBB_CODE_ENABLE And Instr(strType,"[code]")>0 Then

		Dim strCode
		Dim Match, Matches

		strContent =Replace(strContent,vbLf,"")

		'[CODELITE]
		objRegExp.Pattern="(\[CODE_LITE\])(.+?)(\[\/CODE_LITE\])"
		Set Matches = objRegExp.Execute(strContent)

		For Each Match in Matches

			strCode=Match
			strCode = TransferHTML(strCode,"[<][>][space][vbTab]")
			strCode=Replace(strCode,vbCr,"<br/>")
			strContent =Replace(strContent,Match,strCode)

			objRegExp.Global=False

			objRegExp.Pattern="(\[CODE_LITE\](<br\/>)?)(.+?)(\[\/CODE_LITE\])"
			strContent=objRegExp.Replace(strContent,"<p class=""code""><code>$3</code></p>")

			objRegExp.Global=True

		Next
		Set Matches = Nothing

		'[CODE]
		objRegExp.Pattern="(\[CODE\])(.+?)(\[\/CODE\])"
		Set Matches = objRegExp.Execute(strContent)

		For Each Match in Matches

			strCode=Match
			strCode = TransferHTML(strCode,"[<][>][space][vbTab]")
			strCode = Replace(strCode,vbCr,Chr(8)&Chr(11)&Chr(9)&Chr(12))
			strContent =Replace(strContent,Match,strCode)


			objRegExp.Global=False

			objRegExp.Pattern="(\[CODE\])(.+?)(\[\/CODE\])"
			strContent=objRegExp.Replace(strContent,"<textarea class=""code"" rows=""10"" cols=""50"">$2</textarea>")

			objRegExp.Global=True

		Next
		Set Matches = Nothing

		strContent =Replace(strContent,vbCr,vbCrLf)
		strContent =Replace(strContent,Chr(8)&Chr(11)&Chr(9)&Chr(12),vbCr)

	End If


	If ZC_UBB_FACE_ENABLE And Instr(strType,"[face]")>0 Then

		objRegExp.Pattern="(\[F\])(.+?)(\[\/F\])"
		strContent= objRegExp.Replace(strContent,"<img src="""& ZC_BLOG_HOST &"ZB_SYSTEM/image/face/$2.gif"" style=""padding:2px;border:0;"" width="""&ZC_EMOTICONS_FILESIZE&""" title=""$2"" alt=""$2"" />")

	End If


	If ZC_UBB_IMAGE_ENABLE And Instr(strType,"[image]")>0 Then
	'[img]

		objRegExp.Pattern="(\[IMG=)([0-9]*),([0-9]*),([^\n\[]*)(\])(.+?)(\[\/IMG\])"
		strContent= objRegExp.Replace(strContent,"<img src=""$6"" alt=""$4"" title=""$4"" width=""$2"" height=""$3""/>")

		objRegExp.Pattern="(\[IMG=)([0-9]*),([^\n\[]*)(\])(.+?)(\[\/IMG\])"
		strContent= objRegExp.Replace(strContent,"<img src=""$5"" alt=""$3"" title=""$3"" width=""$2""/>")

		objRegExp.Pattern="(\[IMG=)([0-9]*)(\])(.+?)(\[\/IMG\])"
		strContent= objRegExp.Replace(strContent,"<img src=""$4"" alt="""" title="""" width=""$2""/>")

		objRegExp.Pattern="(\[IMG\])(.+?)(\[\/IMG\])"
		strContent= objRegExp.Replace(strContent,"<img onload=""ResizeImage(this,"&ZC_IMAGE_WIDTH&")"" src=""$2"" alt="""" title=""""/>")


		objRegExp.Pattern="(\[IMG_LEFT=)([0-9]*),([0-9]*),([^\n\[]*)(\])(.+?)(\[\/IMG_LEFT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-left"" style=""float:left"" src=""$6"" alt=""$4"" title=""$4"" width=""$2"" height=""$3""/>")

		objRegExp.Pattern="(\[IMG_LEFT=)([0-9]*),([^\n\[]*)(\])(.+?)(\[\/IMG_LEFT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-left"" style=""float:left"" src=""$5"" alt=""$3"" title=""$3"" width=""$2""/>")

		objRegExp.Pattern="(\[IMG_LEFT=)([0-9]*)(\])(.+?)(\[\/IMG_LEFT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-left"" style=""float:left"" src=""$4"" alt="""" title="""" width=""$2""/>")

		objRegExp.Pattern="(\[IMG_LEFT\])(.+?)(\[\/IMG_LEFT\])"
		strContent= objRegExp.Replace(strContent,"<img onload=""ResizeImage(this,"&ZC_IMAGE_WIDTH&")"" class=""float-left"" style=""float:left"" src=""$2"" alt="""" title=""""/>")


		objRegExp.Pattern="(\[IMG_RIGHT=)([0-9]*),([0-9]*),(.*)(\])(.+?)(\[\/IMG_RIGHT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-right"" style=""float:right"" src=""$6"" alt=""$4"" title=""$4"" width=""$2"" height=""$3""/>")

		objRegExp.Pattern="(\[IMG_RIGHT=)([0-9]*),(.*)(\])(.+?)(\[\/IMG_RIGHT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-right"" style=""float:right"" src=""$5"" alt=""$3"" title=""$3"" width=""$2""/>")

		objRegExp.Pattern="(\[IMG_RIGHT=)([0-9]*)(\])(.+?)(\[\/IMG_RIGHT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-right"" style=""float:right"" src=""$4"" alt="""" title="""" width=""$2""/>")

		objRegExp.Pattern="(\[IMG_RIGHT\])(.+?)(\[\/IMG_RIGHT\])"
		strContent= objRegExp.Replace(strContent,"<img onload=""ResizeImage(this,"&ZC_IMAGE_WIDTH&")"" class=""float-right"" style=""float:right"" src=""$2"" alt="""" title=""""/>")



	End If


	If ZC_UBB_FLASH_ENABLE And Instr(strType,"[flash]")>0 Then
	'[flash]

		objRegExp.Pattern="(\[FLASH=)([0-9]*),([0-9]*),([a-z]*)(\])(.+?)(\[\/FLASH\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"" width=""$2"" height=""$3""><param name=""movie"" value=""$6""><param name=""quality"" value=""high""><param name=""play"" value=""$4""><embed src=""$6"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""$2"" height=""$3"" play=""$4""></embed></object>")

	End If



	If ZC_UBB_TYPESET_ENABLE And Instr(strType,"[typeset]")>0 Then

		objRegExp.Pattern="(\[ALIGN-CENTER\])([\u0000-\uffff]+?)(\[\/ALIGN-CENTER\])"
		strContent=objRegExp.Replace(strContent,"<div style=""margin:10px 0 10px 0;text-align:center;"">$2</div>")

		objRegExp.Pattern="(\[ALIGN-LELT\])([\u0000-\uffff]+?)(\[\/ALIGN-LELT\])"
		strContent=objRegExp.Replace(strContent,"<div style=""margin:10px 0 10px 0;text-align:left;"">$2</div>")

		objRegExp.Pattern="(\[ALIGN-RIGHT\])([\u0000-\uffff]+?)(\[\/ALIGN-RIGHT\])"
		strContent=objRegExp.Replace(strContent,"<div style=""margin:10px 0 10px 0;text-align:right;"">$2</div>")

		objRegExp.Pattern="(\[HR\])([\u0000-\uffff]?)(\[\/HR\])"
		strContent=objRegExp.Replace(strContent,"<hr/>")

		objRegExp.Pattern="(\[FONT-FACE=)([a-z\x20]*)(\])([\u0000-\uffff_]+?)(\[\/FONT-FACE\])"
		strContent=objRegExp.Replace(strContent,"<font face=""$2"">$4</font>")

		objRegExp.Pattern="(\[FACE=)([a-z\x20]*)(\])([\u0000-\uffff_]+?)(\[\/FACE\])"
		strContent=objRegExp.Replace(strContent,"<font face=""$2"">$4</font>")

		objRegExp.Pattern="(\[FONT-SIZE=)([1-7]*)(\])([\u0000-\uffff]+?)(\[\/FONT-SIZE\])"
		strContent=objRegExp.Replace(strContent,"<font size=""$2"">$4</font>")

		objRegExp.Pattern="(\[SIZE=)([1-7]*)(\])([\u0000-\uffff]+?)(\[\/SIZE\])"
		strContent=objRegExp.Replace(strContent,"<font size=""$2"">$4</font>")

		objRegExp.Pattern="(\[FONT-COLOR=)([#0-9a-z]*)(\])([\u0000-\uffff]+?)(\[\/FONT-COLOR\])"
		strContent=objRegExp.Replace(strContent,"<font color=""$2"">$4</font>")

		objRegExp.Pattern="(\[COLOR=)([#0-9a-z]*)(\])([\u0000-\uffff]+?)(\[\/COLOR\])"
		strContent=objRegExp.Replace(strContent,"<font color=""$2"">$4</font>")

	End If



	If ZC_UBB_MEDIA_ENABLE And Instr(strType,"[media]")>0 Then

		'[WMA]
		objRegExp.Pattern="(\[WMA=)([a-z]*)(\])(.+?)(\[\/WMA\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95"" height=""68"" width=""350""><param name=""ShowStatusBar"" value=""-1""><param name=""AutoStart"" value=""$2""><param name=""Filename"" value=""$4""><embed type=""application/x-mplayer2"" pluginspage=""http://www.microsoft.com/Windows/MediaPlayer/"" src=""$4"" autostart=""$2"" width=""350"" height=""45""></embed></object>")

		objRegExp.Pattern="(\[WMA\])(.+?)(\[\/WMA\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95"" height=""68"" width=""350""><param name=""ShowStatusBar"" value=""-1""><param name=""AutoStart"" value=""true""><param name=""Filename"" value=""$2""><embed type=""application/x-mplayer2"" pluginspage=""http://www.microsoft.com/Windows/MediaPlayer/"" src=""$2"" autostart=""true"" width=""350"" height=""45""></embed></object>")

		'[WMV]
		objRegExp.Pattern="(\[WMV=)([0-9]*),([0-9]*),([a-z]*)(\])(.+?)(\[\/WMV\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95""  width=""$2"" height=""$3""><param name=""ShowStatusBar"" value=""-1""><param name=""AutoStart"" value=""$4""><param name=""Filename"" value=""$6""><embed type=""application/x-mplayer2"" pluginspage=""http://www.microsoft.com/Windows/MediaPlayer/"" src=""$6"" autostart=""$4""></embed></object>")

		objRegExp.Pattern="(\[WMV\])(.+?)(\[\/WMV\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95""><param name=""ShowStatusBar"" value=""-1""><param name=""AutoStart"" value=""true""><param name=""Filename"" value=""$2""><embed type=""application/x-mplayer2"" pluginspage=""http://www.microsoft.com/Windows/MediaPlayer/"" src=""$2"" autostart=""true""></embed></object>")

		'[RMV]
		objRegExp.Pattern="(\[RM=)([0-9]*),([0-9]*),([a-z]*)(\])(.+?)(\[\/RM\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" width=""$2"" height=""$3""><param name=""src"" value=""$6""><param name=""controls"" value=""imagewindow""><param name=""console"" value=""one""><param name=""AutoStart"" value=""$4""><embed src=""$6"" type=""audio/x-pn-realaudio-plugin"" width=""$2"" height=""$3"" nojava=""true"" controls=""imagewindow,ControlPanel,StatusBar"" console=""one"" autostart=""$4""></object>")

		objRegExp.Pattern="(\[RM\])(.+?)(\[\/RM\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA""><param name=""src"" value=""$2""><param name=""controls"" value=""imagewindow""><param name=""console"" value=""one""><param name=""AutoStart"" value=""true""><embed src=""$2"" type=""audio/x-pn-realaudio-plugin"" nojava=""true"" controls=""imagewindow,ControlPanel,StatusBar"" console=""one"" autostart=""true""></embed></object>")

		'[RA]
		objRegExp.Pattern="(\[RA=)([a-z]*)(\])(.+?)(\[\/RA\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" width=""350"" height=""36""><param name=""src"" value=""$4""><param name=""controls"" value=""ControlPanel""><param name=""console"" value=""one""><param name=""AutoStart"" value=""$2""><embed src=""$4"" type=""audio/x-pn-realaudio-plugin"" nojava=""true"" controls=""ControlPanel,StatusBar"" console=""one"" autostart=""$2"" width=""350"" height=""36""></embed></object>")

		objRegExp.Pattern="(\[RA\])(.+?)(\[\/RA\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" width=""350"" height=""36""><param name=""src"" value=""$2""><param name=""controls"" value=""ControlPanel""><param name=""console"" value=""one""><param name=""AutoStart"" value=""true""><embed src=""$2"" type=""audio/x-pn-realaudio-plugin"" nojava=""true"" controls=""ControlPanel,StatusBar"" console=""one"" autostart=""true"" width=""350"" height=""36""></embed></object>")

		'[QT]
		objRegExp.Pattern="(\[QT=)([0-9]*),([0-9]*),([a-z]*)(\])(.+?)(\[\/QT\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:02BF25D5-8C17-4B23-BC80-D3488ABDDC6B"" codebase=""http://www.apple.com/qtactivex/qtplugin.cab"" width=""$2"" height=""$3"" ><param name=""src"" value=""$6"" ><param name=""autoplay"" value=""$4"" ><embed  src=""qtmimetype.pntg"" type=""image/x-macpaint"" pluginspage=""http://www.apple.com/quicktime/download"" qtsrc=""$6"" width=""$2"" height=""$3"" autoplay=""$4"" ></embed></object>")

		objRegExp.Pattern="(\[QT\])(.+?)(\[\/QT\])"
		strContent= objRegExp.Replace(strContent,"<object classid=""clsid:02BF25D5-8C17-4B23-BC80-D3488ABDDC6B"" codebase=""http://www.apple.com/qtactivex/qtplugin.cab"" ><param name=""src"" value=""$2"" ><param name=""autoplay"" value=""true"" ><embed  src=""qtmimetype.pntg"" type=""image/x-macpaint"" pluginspage=""http://www.apple.com/quicktime/download"" qtsrc=""$2"" autoplay=""true"" ></embed></object>")

		'[MEDIA]
		objRegExp.Pattern="(\[MEDIA=)([a-z]*),([0-9]*),([0-9]*)(\])(.+?)(\[\/MEDIA\])"
		strContent= objRegExp.Replace(strContent,"<div class=""media""><a href="""" onclick=""javascript:ShowMedia(this.nextSibling,'$6','$2',$3,$4);return(false);"">"& ZC_MSG103 &"</a><div class=""media-object""></div></div>")

		objRegExp.Pattern="(\[MEDIA=)([0-9]*),([0-9]*)(\])(.+?)(\[\/MEDIA\])"
		strContent= objRegExp.Replace(strContent,"<div class=""media""><a href="""" onclick=""javascript:ShowMedia(this.nextSibling,'$5','AUTO',$2,$3);return(false);"">"& ZC_MSG103 &"</a><div class=""media-object""></div></div>")

		objRegExp.Pattern="(\[MEDIA\])(.+?)(\[\/MEDIA\])"
		strContent= objRegExp.Replace(strContent,"<div class=""media""><a href="""" onclick=""javascript:ShowMedia(this.nextSibling,'$2','AUTO',400,300);return(false);"">"& ZC_MSG103 &"</a><div class=""media-object""></div></div>")


	End If



	If ZC_UBB_AUTOLINK_ENABLE And Instr(strType,"[autolink]")>0 Then

		objRegExp.Pattern="(^|\r\n|\n)((http|https|ftp|mailto|gopher|news|telnet|mms|rtsp|ed2k|tencent|nfcall|dic|pig2pig|callto|exeem|ymsgr|thunder|p4p|pplive|synacast|ppstream|feed|wangwang|qqtv|rssfeed|msnim|chrome|file|ppg|thunder):{1}\/{0,2}[^<>\f\n\r\t\v]+?)(\r\n|\n|$)"
		strContent=objRegExp.Replace(strContent,vbCrlf & "<a href=""$2""  target=""_blank"">$2</a>" & vbCrlf)

		objRegExp.Pattern="(^|\r\n|\n)((http|https|ftp|mailto|gopher|news|telnet|mms|rtsp|ed2k|tencent|nfcall|dic|pig2pig|callto|exeem|ymsgr|thunder|p4p|pplive|synacast|ppstream|feed|wangwang|qqtv|rssfeed|msnim|chrome|file|ppg|thunder):{1}\/{0,2}[^<>\f\n\r\t\v]+?)(\r\n|\n|$)"
		strContent=objRegExp.Replace(strContent,vbCrlf & "<a href=""$2""  target=""_blank"">$2</a>" & vbCrlf)

	End If


	If ZC_UBB_AUTOKEY_ENABLE And Instr(strType,"[key]")>0 Then

		Dim i,j

		If IsArray(KeyWords) Then
			For i=Lbound(KeyWords,2) To Ubound(KeyWords,2)

				objRegExp.Pattern="((<.*)("&KeyWords(1,i)&")(.*>))|((<a.*)("&KeyWords(1,i)&")(\/a>))"

				Set Matches = objRegExp.Execute(strContent)
				For Each Match in Matches
					strContent=Replace(strContent,Match,vbVerticalTab & vbTab & vbVerticalTab)
				Next

				strContent=Replace(strContent,KeyWords(1,i),"<a href="""& KeyWords(2,i) &""" target=""_blank"">"& KeyWords(1,i) &"</a>")


				For Each Match in Matches
					strContent=Replace(strContent,vbVerticalTab & vbTab & vbVerticalTab,Match,1,1)
				Next
				Set Matches = Nothing

			Next
		End If

	End If


	If ZC_UBB_LINK_ENABLE And Instr(strType,"[link-antispam]")>0 Then

		Dim Match2, Matches2 ,strCode2

		objRegExp.Pattern="(href="".+?"")"
		Set Matches2 = objRegExp.Execute(strContent)

		For Each Match2 in Matches2
			strCode2=Match2
			strCode2=Left(strCode2,Len(strCode2)-1)
			strCode2=Right(strCode2,Len(strCode2)-6)
			strCode2=URLEncodeForAntiSpam(strCode2)
			strContent =Replace(strContent,Match2,"href=""" & strCode2 & """")
		Next
		Set Matches2 = Nothing

	End If


	Set objRegExp=Nothing
	UBBCode=strContent

End Function
'*********************************************************




'*********************************************************
' 目的：    Save Text to File
' 输入：    
' 输入：    
' 返回：    
'*********************************************************
Function SaveToFile(strFullName,strContent,strCharset,bolRemoveBOM)

	On Error Resume Next

	Dim objStream

	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = adTypeText
	.Mode = adModeReadWrite
	.Open
	.Charset = strCharset
	.Position = objStream.Size
	.WriteText = strContent
	.SaveToFile strFullName,adSaveCreateOverWrite
	.Close
	End With
	Set objStream = Nothing

	If bolRemoveBOM Then
		If strContent<>"" And ZC_STATIC_TYPE="shtml" Then
			Call RemoveBOM(strFullName)
		End If
	End If

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    Load Text form File
' 输入：    
' 输入：    
' 返回：    
'*********************************************************
Function LoadFromFile(strFullName,strCharset)

	On Error Resume Next

	Dim objStream

	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = adTypeText
	.Mode = adModeReadWrite
	.Open
	.Charset = strCharset
	.Position = objStream.Size
	.LoadFromFile strFullName
	LoadFromFile=.ReadText
	.Close
	End With
	Set objStream = Nothing

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    Remove BOM from UTF-8
'*********************************************************
Function RemoveBOM(strFullName)

	On Error Resume Next

	Dim objStream
	Dim strContent

	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = adTypeBinary
	.Mode = adModeReadWrite
	.Open
	.Position = objStream.Size
	.LoadFromFile strFullName
	.Position = 3
	strContent=.Read
	.Close
	End With
	Set objStream = NoThing

	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = adTypeBinary
	.Mode = adModeReadWrite
	.Open
	.Position = objStream.Size
	.Write = strContent
	.SaveToFile strFullName,adSaveCreateOverWrite
	.Close
	End With
	Set objStream = Nothing

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    Save Value For Setting
'*********************************************************
Function SaveValueForSetting(ByRef strContent,bolConst,strTypeVar,strItem,strValue)

	Dim i,j,s,t
	Dim strConst
	Dim objRegExp

	If bolConst=True Then strConst="Const"

	Set objRegExp=New RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True

	strValue=TransferHTML(strValue,"[no-asp]")

	If strTypeVar="String" Then

		strValue=Replace(strValue,"""","""""")
		strValue=""""& strValue &""""

		objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))(.+?)(\r\n|\n|$)"
		If objRegExp.Test(strContent)=True Then
			strContent=objRegExp.Replace(strContent,"$1$2"& strValue &"$8")
			SaveValueForSetting=True
			Exit Function
		End If

	End If

	If strTypeVar="Boolean" Then

		strValue=Trim(strValue)
		If LCase(strValue)="true" Then
			strValue="True"
		Else
			strValue="False"
		End If

		If objRegExp.Test(strContent)=True Then
			objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([a-z]+)( *)(\r\n|\n|$)"
			strContent=objRegExp.Replace(strContent,"$1$2"& strValue &"$9")
			SaveValueForSetting=True
			Exit Function
		End If


	End If

	If strTypeVar="Numeric" Then

		strValue=Trim(strValue)
		If IsNumeric(strValue)=False Then
			strValue=0
		End If

		If objRegExp.Test(strContent)=True Then
			objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([0-9.]+)( *)(\r\n|\n|$)"
			strContent=objRegExp.Replace(strContent,"$1$2"& strValue &"$9")
			SaveValueForSetting=True
			Exit Function
		End If

	End If

	SaveValueForSetting=False

End Function
'*********************************************************




'*********************************************************
' 目的：    Load Value For Setting
'*********************************************************
Function LoadValueForSetting(strContent,bolConst,strTypeVar,strItem,ByRef strValue)

	Dim i,j,s,t
	Dim strConst
	Dim objRegExp
	Dim Matches,Match

	If bolConst=True Then strConst="Const"

	Set objRegExp=New RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True


	If strTypeVar="String" Then

		objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))(.+?)(\r\n|\n|$)"
		Set Matches = objRegExp.Execute(strContent)
		If Matches.Count=1 Then

			t=Matches(0).Value
			t=Replace(t,VbCrlf,"")
			t=Replace(t,Vblf,"")
			objRegExp.Pattern="( *)""(.*)""( *)($)"
			Set Matches = objRegExp.Execute(t)

			If Matches.Count>0 Then

				s=Trim(Matches(0).Value)
				s=Mid(s,2,Len(s)-2)
				s=Replace(s,"""""","""")
				strValue=s

				LoadValueForSetting=True
				Exit Function

			End If
		End If

	End If

	If strTypeVar="Boolean" Then

		objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([a-z]+)( *)(\r\n|\n|$)"
		Set Matches = objRegExp.Execute(strContent)
		If Matches.Count=1 Then
			t=Matches(0).Value
			t=Replace(t,VbCrlf,"")
			t=Replace(t,Vblf,"")
			objRegExp.Pattern="( *)((True)|(False))( *)($)"
			Set Matches = objRegExp.Execute(t)

			If Matches.Count>0 Then

				s=Trim(Matches(0).Value)
				s=LCase(Matches(0).Value)
				If InStr(s,"true")>0 Then
					strValue=True
				ElseIf InStr(s,"false")>0 Then
					strValue=False
				End If

				LoadValueForSetting=True
				Exit Function

			End If
		End If

	End If

	If strTypeVar="Numeric" Then

		objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([0-9.]+)( *)(\r\n|\n|$)"
		Set Matches = objRegExp.Execute(strContent)
		If Matches.Count=1 Then
			t=Matches(0).Value
			t=Replace(t,VbCrlf,"")
			t=Replace(t,Vblf,"")
			objRegExp.Pattern="( *)([0-9.]+)( *)($)"
			Set Matches = objRegExp.Execute(t)

			If Matches.Count>0 Then

				s=Trim(Matches(0).Value)
				If IsNumeric(s)=True Then

					strValue=s

					LoadValueForSetting=True
					Exit Function

				End If

			End If
		End If

	End If

	LoadValueForSetting=False

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function URLEncodeForAntiSpam(strUrl)

	Dim i,s
	For i =1 To Len(strUrl)
		s=s & Mid(strUrl,i,1) & CStr(Int((10 * Rnd)))
	Next
	URLEncodeForAntiSpam=ZC_BLOG_HOST & "../zb_system/function/c_urlredirect.asp?url=" & Server.URLEncode(s)

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function URLDecodeForAntiSpam(strUrl)

	Dim i,s
	For i =1 To Len(strUrl) Step 2
		s=s & Mid(strUrl,i,1)
	Next
	s=TransferHTML(s,"[anti-html-format]")
	If CheckRegExp(s,"[homepage]")=False Then s=""

	URLDecodeForAntiSpam=s

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function GetTime(t)

        GetTime=DateAdd("h", -(ZC_HOST_TIME_ZONE / 100) + (ZC_TIME_ZONE / 100) , t)

End Function
'*********************************************************




'*********************************************************
'目的：自动闭合HTML
'*********************************************************
Function closeHTML(strContent)

  Dim arrTags,i,OpenPos,ClosePos,re,strMatchs,j,Match
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
    arrTags=array("strong","em","strike","b","u","i","font","span","a", "h1","h2","h3","h4","h5","h6","p","li","ol","ul","td","tr","tbody","table","blockquote","pre","cite","div")
  For i=0 To ubound(arrTags)
   OpenPos=0
   ClosePos=0
   
   re.Pattern="\<"+arrTags(i)+"( [^\<\>]+|)\>"
   Set strMatchs=re.Execute(strContent)
   For Each Match In strMatchs
    OpenPos=OpenPos+1
   Next
   re.Pattern="\</"+arrTags(i)+"\>"
   Set strMatchs=re.Execute(strContent)
   For Each Match In strMatchs
    ClosePos=ClosePos+1
   Next
   For j=1 To OpenPos-ClosePos
      strContent=strContent+"</"+arrTags(i)+">"
   Next
  Next
  closeHTML=strContent

End Function 
'*********************************************************




'*********************************************************
' 目的：三态
'*********************************************************
Function IIf(ByVal expr,ByVal  truepart,ByVal  falsepart)
	If expr=True Then
		IIf=truepart
		Exit Function
	Else
		IIf=falsepart
	End If
End Function
'*********************************************************




'*********************************************************
' 目的：   检查是否手机端访问
'*********************************************************
Function CheckMobile()

	'是否由wap转入电脑版
	If  Not IsEmpty(Request.ServerVariables("HTTP_REFERER"))  And  InStr(LCase(Request.ServerVariables("HTTP_REFERER")),ZC_FILENAME_WAP) Then 
			CheckMobile=False:Exit Function  
	End If 

	'是否专用wap浏览器
	If InStr(LCase(Request.ServerVariables("HTTP_ACCEPT")), "application/vnd.wap.xhtml+xml") Or Not IsEmpty(Request.ServerVariables("HTTP_X_PROFILE")) Or Not IsEmpty(Request.ServerVariables("HTTP_PROFILE")) Then
			CheckMobile=True:Exit Function
	End If 

	'是否（智能）手机浏览器
	Dim MobileBrowser_List,PCBrowser_List,UserAgent
	MobileBrowser_List ="up.browser|up.link|mmp|iphone|android|wap|netfront|java|opera\smini|ucweb|windows\sce|symbian|series|webos|sonyericsson|sony|blackberry|cellphone|dopod|nokia|samsung|palmsource|palmos|pda|xphone|xda|smartphone|pieplus|meizu|midp|cldc|brew|tear"
	PCBrowser_List="mozilla|chrome|safari|opera|m3gate|winwap|openwave"
	UserAgent = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
	If CheckRegExp(UserAgent,MobileBrowser_List) Then 
		CheckMobile=True:Exit Function
	ElseIf CheckRegExp(UserAgent,PCBrowser_List) Then '未知手机浏览器，其UA标识为常见浏览器，不跳转
		CheckMobile=False:Exit Function
	Else 
		CheckMobile=False 
	End If 

End Function 
'*********************************************************

'*********************************************************
' 目的：    unescape
' 输入：    
' 输入：    要替换的字符
' 返回：    
'*********************************************************
%>
<script language="javascript" runat="server">

	function vbsunescape(source){
		return unescape(source);
	}

</script>
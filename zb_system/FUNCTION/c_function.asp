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




Dim PublicObjAdo
Dim PublicObjFSO

Set PublicObjAdo=Server.CreateObject("ADODB.Stream")
Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")


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
	'Response.Redirect GetCurrentHost() & "zb_system/function/c_error.asp?errorid=" & id & "&number=" & Err.Number & "&description=" & Server.URLEncode(Err.Description) & "&source=" & Server.URLEncode(Err.Source) & "&sourceurl="  &Server.URLEncode(Request.ServerVariables("Http_Referer"))
	Response.Clear
	
	If id=2 Then
		Response.Status="404 Not Found"
	Else
		Response.Status="500 Internal Server Error"
	End If
	%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<link rel="stylesheet" rev="stylesheet" href="<%=ZC_BLOG_HOST%>zb_system/css/admin.css" type="text/css" media="screen" />
	<title><%=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG045%></title>
</head>
<body class="short">

<div class="bg">
<div id="wrapper">
  <div class="logo"><img src="<%=ZC_BLOG_HOST%>zb_system/image/admin/none.gif" title="Z-Blog" alt="Z-Blog"/></div>
  <div class="login">
	<form id="frmLogin" method="post" action="">
	  <div class="divHeader"><%=ZC_MSG045%></div>

<%
	Response.Write "<p>" & ZC_MSG098 & ":" & ZVA_ErrorMsg(id) & "</p>"

	If Err.Number<>0 Then
		Response.Write "<p>" & ZC_MSG076 & ":" & "" & Err.Number & "</p>"
		Response.Write "<p>" & ZC_MSG016 & ":" & "<br/>" & TransferHTML(Err.Description,"[html-format]") & "</p>"
		Response.Write "<p>" & TransferHTML(Err.Source,"[html-format]") & "</p>"
	End If
		Response.Write "<p><br/></p>"
	If CheckRegExp(Request.ServerVariables("Http_Referer"),"[homepage]")=True Then
		Response.Write "<p style='text-align:right;'><a href=""" & TransferHTML(Request.ServerVariables("Http_Referer"),"[html-format]") & """>" & ZC_MSG207 & "</a></p>"
	Else
		Response.Write "<p style='text-align:right;'><a href=""" & GetCurrentHost() & """>" & ZC_MSG207 & "</a></p>"
	End If

	If id=6 Then
		Response.Write "<p style='text-align:right;'><a href=""../cmd.asp?act=login"" target=""_top"">"& ZC_MSG009 & "</a></p>"
	End If
%>

    </form>
  </div>
</div>
</div>
</body>
</html>
	<%
	Response.End
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
		para="^[\.\_A-Za-z0-9\u4e00-\u9fa5]+$"
	End If
	If para="[password]" Then
		para="^[A-Za-z0-9`~!@#\$%\^&\*\-_]+$"
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

	If Instr(strReferer,GetCurrentHost())=0 Then 
		ShowError(5)
	End If

End Sub
'*********************************************************


'*********************************************************
' 目的：    得到真实IP
' 返回：    IP
'*********************************************************
Function GetReallyIP()

	Dim strIP
	strIP=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If strIP="" Or InStr(strIP,"unknown") Then
		strIP=Request.ServerVariables("REMOTE_ADDR")
	ElseIf InStr(strIP,",") Then
		strIP=Split(strIP,",")(0)
	ElseIf InStr(strIP,";") Then
		strIP=Split(strIP,";")(0)
	End If
	
	GetReallyIP=Trim(strIP)

End Function
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

	If Len(strQuestion)=0 Then Exit Function

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
				s=Left(s,j-1) & "<b style='color:red'>" & strQuestion & "</b>"
				Exit Do
			Else
				s=Left(s,j-1) & "<b style='color:red'>" & strQuestion & "</b>" & Right(s,Len(s)-j-Len(strQuestion)+1)
			End If
			i=j+Len("<b style='color:red'>" & strQuestion & "</b>")-1
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

	If IsNull(strSQL) Then strSQL=""
	FilterSQL=CStr(Replace(strSQL,chr(39),chr(39)&chr(39)))

End Function
'*********************************************************




'*********************************************************
' 目的：    检查引用
' 输入：    
' 输入：    要替换的字符代号
' 返回：    
'*********************************************************
Function TransferHTML(ByVal source,ByVal para)

	Dim objRegExp

	If IsNull(source)=True Then Exit Function


	If InStr(para,"[mobilerequest]") Then
		para=para&"[enter][closehtml]"	
		'如何判断HTML标签和用户输入的类似0<1这种数据，还真是个大麻烦	
	End If

	'先换"&"
	If Instr(para,"[&]")>0 Then  source=Replace(source,"&","&amp;")
	If Instr(para,"[<]")>0 Then  source=Replace(source,"<","&lt;")
	If Instr(para,"[>]")>0 Then  source=Replace(source,">","&gt;")
	If Instr(para,"[""]")>0 Then source=Replace(source,"""","&quot;")
	If Instr(para,"[space]")>0 Then source=Replace(source," ","&nbsp;")
	If Instr(para,"[delspace]")>0 Then
		Source=Replace(source," ","")
		Source=Replace(source,"　","")
	End If
	If Instr(para,"[enter]")>0 Then
		source=Replace(source,vbCrLf,"<br/>")
		source=Replace(source,vbLf,"<br/>")
	End If
	If Instr(para,"[vbCrlf]")>0 Then 

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
		source=Replace(source,"src=""upload/","src=""<#ZC_BLOG_HOST#>" & ZC_UPLOAD_DIRECTORY & "/")
		source=Replace(source,"href=""upload/","href=""<#ZC_BLOG_HOST#>" & ZC_UPLOAD_DIRECTORY & "/")
		source=Replace(source,"value=""upload/","value=""<#ZC_BLOG_HOST#>" & ZC_UPLOAD_DIRECTORY & "/")
		source=Replace(source,"href=""http://upload/","href=""<#ZC_BLOG_HOST#>" & ZC_UPLOAD_DIRECTORY & "/")
		source=Replace(source,"(this.nextSibling,'upload/","(this.nextSibling,'<#ZC_BLOG_HOST#>" & ZC_UPLOAD_DIRECTORY & "/")

		source=Replace(source,"src=""image/face/","src=""<#ZC_BLOG_HOST#>zb_users/emotion/face/")
	End If
	If Instr(para,"[anti-upload]")>0 Then
		source=Replace(source,"src="""& GetCurrentHost() & ZC_UPLOAD_DIRECTORY & "/","src=""upload/")
		source=Replace(source,"href="""& GetCurrentHost() & ZC_UPLOAD_DIRECTORY & "/","href=""upload/")
		source=Replace(source,"value="""& GetCurrentHost() & ZC_UPLOAD_DIRECTORY & "/","value=""upload/")
		source=Replace(source,"href="""& GetCurrentHost() & ZC_UPLOAD_DIRECTORY & "/","href=""http://upload/")
		source=Replace(source,"(this.nextSibling,'"& GetCurrentHost() & ZC_UPLOAD_DIRECTORY & "/","(this.nextSibling,'upload/")

		source=Replace(source,"src="""& GetCurrentHost() & "zb_users/emotion/face/","src=""<#ZC_BLOG_HOST#>zb_users/emotion/face/")
	End If
	If Instr(para,"[zc_blog_host]")>0 Then
		source=Replace(source,"<#ZC_BLOG_HOST#>",GetCurrentHost())
	End If
	If Instr(para,"[no-asp]")>0 Then
		source=Replace(source,"<"&"%","&lt;"&"%")
		source=Replace(source,"%"&">","%"&"&gt;")
	End If
	If Instr(para,"[nofollow]")>0 Then
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
	If Instr(para,"[directory&file]")>0  Then
		source=Replace(source,"/","/")
		source=Replace(source,"\","/")
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
	If Instr(para,"[normaltag]")>0  Then
		source=Replace(source,"$","")
		source=Replace(source,"(","")
		source=Replace(source,")","")
		source=Replace(source,"*","")
		source=Replace(source,"+","")
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
		Set objRegExp=New RegExp
		objRegExp.IgnoreCase =True
		objRegExp.Global=True
		objRegExp.Pattern=",+"
		source= objRegExp.Replace(source,",")
		objRegExp.Pattern="(^,|,$)"
		source= objRegExp.Replace(source,"")
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
	If Instr(para,"[japan-html]")>0 Then
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
	If Instr(para,"[html-japan]")>0 Then
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
	If Instr(para,"[anti-zc_blog_host]")>0 Then
		source=Replace(source,GetCurrentHost(),"<#ZC_BLOG_HOST#>")
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

	If ZC_UBB_ENABLE=False Then
		UBBCode=strContent
		Exit Function
	End If

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
		strContent= objRegExp.Replace(strContent,"<img src=""<#ZC_BLOG_HOST#>zb_users/emotion/face/$2.gif"" style=""padding:2px;border:0;"" width="""&ZC_EMOTICONS_FILESIZE&""" title=""$2"" alt=""$2"" />")

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
		strContent= objRegExp.Replace(strContent,"<img src=""$2"" alt="""" title=""""/>")


		objRegExp.Pattern="(\[IMG_LEFT=)([0-9]*),([0-9]*),([^\n\[]*)(\])(.+?)(\[\/IMG_LEFT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-left"" style=""float:left"" src=""$6"" alt=""$4"" title=""$4"" width=""$2"" height=""$3""/>")

		objRegExp.Pattern="(\[IMG_LEFT=)([0-9]*),([^\n\[]*)(\])(.+?)(\[\/IMG_LEFT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-left"" style=""float:left"" src=""$5"" alt=""$3"" title=""$3"" width=""$2""/>")

		objRegExp.Pattern="(\[IMG_LEFT=)([0-9]*)(\])(.+?)(\[\/IMG_LEFT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-left"" style=""float:left"" src=""$4"" alt="""" title="""" width=""$2""/>")

		objRegExp.Pattern="(\[IMG_LEFT\])(.+?)(\[\/IMG_LEFT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-left"" style=""float:left"" src=""$2"" alt="""" title=""""/>")


		objRegExp.Pattern="(\[IMG_RIGHT=)([0-9]*),([0-9]*),(.*)(\])(.+?)(\[\/IMG_RIGHT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-right"" style=""float:right"" src=""$6"" alt=""$4"" title=""$4"" width=""$2"" height=""$3""/>")

		objRegExp.Pattern="(\[IMG_RIGHT=)([0-9]*),(.*)(\])(.+?)(\[\/IMG_RIGHT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-right"" style=""float:right"" src=""$5"" alt=""$3"" title=""$3"" width=""$2""/>")

		objRegExp.Pattern="(\[IMG_RIGHT=)([0-9]*)(\])(.+?)(\[\/IMG_RIGHT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-right"" style=""float:right"" src=""$4"" alt="""" title="""" width=""$2""/>")

		objRegExp.Pattern="(\[IMG_RIGHT\])(.+?)(\[\/IMG_RIGHT\])"
		strContent= objRegExp.Replace(strContent,"<img class=""float-right"" style=""float:right"" src=""$2"" alt="""" title=""""/>")



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
' 目的：    Del to File
' 输入：    
' 输入：    
' 返回：    
'*********************************************************
Function DelToFile(strFullName)

	On Error Resume Next
	DelToFile=False

	Dim TxtFile
	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")

	If PublicObjFSO.FileExists(strFullName) Then
		Set TxtFile = PublicObjFSO.GetFile(strFullName)
		TxtFile.Delete
		If Err.Number=0 Then DelToFile=True
	End If

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

	If Not IsObject(PublicObjAdo) Then Set PublicObjAdo = Server.CreateObject("ADODB.Stream")
	With PublicObjAdo
	.Type = adTypeText
	.Mode = adModeReadWrite
	.Open
	.Charset = strCharset
	.Position = .Size
	.WriteText = strContent
	.SaveToFile strFullName,adSaveCreateOverWrite
	.Close
	End With

	If bolRemoveBOM Then
		If strContent<>"" And ZC_STATIC_TYPE="shtml" Then
			Call RemoveBOM(strFullName)
		End If
	End If

	Err.Clear

End Function
'*********************************************************

'*********************************************************
' 目的：    Save Binary to File
' 输入：    
' 输入：    
' 返回：    
'*********************************************************
Function SaveBinary(BinaryData,FilePath)
	On Error Resume Next

	If Not IsObject(PublicObjAdo) Then Set PublicObjAdo = Server.CreateObject("ADODB.Stream")

	With PublicObjAdo
		.Type = adTypeBinary
		.Open
		.Write BinaryData
		.SaveToFile FilePath, adSaveCreateOverWrite
		.Close
	End With
	
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

	If Not IsObject(PublicObjAdo) Then Set PublicObjAdo = Server.CreateObject("ADODB.Stream")

	With PublicObjAdo
	.Type = adTypeText
	.Mode = adModeReadWrite
	.Open
	.Charset = strCharset
	.Position = .Size
	.LoadFromFile strFullName
	LoadFromFile=.ReadText
	.Close
	End With

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    Remove BOM from UTF-8
'*********************************************************
Function RemoveBOM(strFullName)

	On Error Resume Next

	Dim strContent
	
	If Not IsObject(PublicObjAdo) Then Set PublicObjAdo = Server.CreateObject("ADODB.Stream")

	With PublicObjAdo
	.Type = adTypeBinary
	.Mode = adModeReadWrite
	.Open
	.Position = .Size
	.LoadFromFile strFullName
	.Position = 3
	strContent=.Read
	.Close
	End With

	With PublicObjAdo
	.Type = adTypeBinary
	.Mode = adModeReadWrite
	.Open
	.Position = .Size
	.Write = strContent
	.SaveToFile strFullName,adSaveCreateOverWrite
	.Close
	End With

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
		
		objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([a-z]+)( *)(\r\n|\n|$)"
		If objRegExp.Test(strContent)=True Then
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
' 目的：    测试某个object是否已经安装
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
' 目的：    加密URL
'*********************************************************
Function URLEncodeForAntiSpam(strUrl)

	If InStr(strUrl,"c_urlredirect.asp")>0 Then
		URLEncodeForAntiSpam=strUrl
		Exit Function
	End If

	Dim i,s
	For i =1 To Len(strUrl)
		s=s & Mid(strUrl,i,1) & CStr(Int((10 * Rnd)))
	Next
	URLEncodeForAntiSpam=GetCurrentHost() & "zb_system/function/c_urlredirect.asp?url=" & Server.URLEncode(s)

End Function
'*********************************************************




'*********************************************************
' 目的：    解密URL
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
' 目的：    根据t格式化时区得到时间
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
' 目的：Remove Li by Url'''''
'*********************************************************
Function RemoveLibyUrl(ByVal s,ByVal url)

	Dim i,b
	b=Split(s,"</li>")
	If UBound(b)>0 then
		For i=0 To UBound(b)-1
			b(i)=b(i) & "</li>"
			If InStr(b(i),"href="""&url)>0 Then
				b(i)=""
			End If
			If InStr(b(i),"href='"&url)>0 Then
				b(i)=""
			End If
		Next
		RemoveLibyUrl=Join(b)
		Exit Function
	End if

	RemoveLibyUrl=s

End Function
'*********************************************************




'*********************************************************
' 目的：    根据文件路径得到最后更改时间
'*********************************************************
Function GetFileModified(Path)
	On Error Resume Next
	
	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")
	
	If PublicObjFSO.FileExists(Path) Then
		GetFileModified=PublicObjFSO.GetFile(Path).DateLastModified
	Else
		GetFileModified=Now
	End If

End Function
'*********************************************************




'*********************************************************
' 目的：   生成随机Guid
'*********************************************************
Function RndGuid()

	Dim i,s

	Const c="0123456789ABCDEF"

	Randomize

	For i=1 To 32 

		s=s & Mid(c,Int(Rnd*16)+1,1)
		If i=1 And s="0" Then s="F"

		If i=8 Then s=s & "-"
		If i=12 Then s=s & "-"
		If i=16 Then s=s & "-"
		If i=20 Then s=s & "-"

	Next

	RndGuid=s

End Function 
'*********************************************************




'*********************************************************
' 目的：    得到当前地址
'*********************************************************
Dim CurrentHostUrl
Dim CurrentReallyDirectory
Function GetCurrentHost()

	On Error Resume Next

	If CurrentHostUrl<>"" Then
		GetCurrentHost=CurrentHostUrl
		Exit Function
	End If

	Dim PhysicsPath

	PhysicsPath=Server.MapPath(".") & "\"

	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")
	
	If PublicObjFSO.FolderExists(PhysicsPath & "ZB_SYSTEM\") Then
		PhysicsPath=PhysicsPath
	ElseIf PublicObjFSO.FolderExists(PhysicsPath & "..\ZB_SYSTEM\") Then
		PhysicsPath=PhysicsPath & "..\"
	ElseIf PublicObjFSO.FolderExists(PhysicsPath & "..\..\ZB_SYSTEM\") Then
		PhysicsPath=PhysicsPath & "..\..\"
	ElseIf PublicObjFSO.FolderExists(PhysicsPath & "..\..\..\ZB_SYSTEM\") Then
		PhysicsPath=PhysicsPath & "..\..\..\"
	ElseIf PublicObjFSO.FolderExists(PhysicsPath & "..\..\..\..\ZB_SYSTEM\") Then
		PhysicsPath=PhysicsPath & "..\..\..\..\"
	ElseIf PublicObjFSO.FolderExists(PhysicsPath & "..\..\..\..\..\ZB_SYSTEM\") Then
		PhysicsPath=PhysicsPath & "..\..\..\..\..\"
	ElseIf PublicObjFSO.FolderExists(PhysicsPath & "..\..\..\..\..\..\ZB_SYSTEM\") Then
		PhysicsPath=PhysicsPath & "..\..\..\..\..\..\"
	ElseIf PublicObjFSO.FolderExists(PhysicsPath & "..\..\..\..\..\..\..\ZB_SYSTEM\") Then
		PhysicsPath=PhysicsPath & "..\..\..\..\..\..\..\"
	End If
	Set fso=Nothing

	PhysicsPath=PublicObjFSO.GetFolder(PhysicsPath).Path
	If Right(PhysicsPath,1)<>"\" Then PhysicsPath=PhysicsPath & "\"
	CurrentReallyDirectory=PhysicsPath
	
	Err.Clear
	If ZC_PERMANENT_DOMAIN_ENABLE=True Then
		CurrentHostUrl=ZC_BLOG_HOST
		GetCurrentHost=CurrentHostUrl
		If Err.Number=0 Then
			Exit Function
		End If
	End If

	Dim s,t,u,i,w,x

	s=LCase(Replace(Request.ServerVariables("PATH_TRANSLATED"),"\","/"))

	t=LCase(Request.ServerVariables("HTTP_HOST") & Split(Request.ServerVariables("URL"),"?")(0))
	'Kangle 3.0下Request.ServerVariables("URL")含有QueryString

	w=LCase(Replace(PhysicsPath,"\","/"))

	x=Right(s,Len(s)-Len(w))

	u=Replace(t,x,"")

	if Request.ServerVariables("HTTPS")<>"on" Then
	'Kangle的返回值为True\False..
		If Request.ServerVariables("HTTPS")=True Then
			u= "https://" & u
		Else 
			u= "http://" & u
		End If
		
	else
		u= "https://" & u
	end If

	If Right(u,1)<>"/" Then u=u & "/"

	CurrentHostUrl=u

	GetCurrentHost=CurrentHostUrl

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function SetValueByNameInArrays(ByRef arrayname,ByRef arrayvalue,name,value)

	Dim IsFind
	IsFind=False

	Dim i,j
	j=UBound(arrayname)
	For i=1 To j
		If LCase(arrayname(i))=LCase(name) Then
			arrayvalue(i)=value
			IsFind=True
			Exit For
		End If 
	Next


	If IsFind=True Then
		SetValueByNameInArrays=True
		Exit Function
	End If

	j=j+1
	ReDim Preserve arrayname(j)
	ReDim Preserve arrayvalue(j)
	arrayname(j)=name
	arrayvalue(j)=value

	SetValueByNameInArrays=True

End Function
'*********************************************************




'*********************************************************
Function SearchInArrays(a,arrays)

	Dim c

	If IsArray(arrays)=True Then
		For Each c In arrays
			If LCase(a)=LCase(c) Then
				SearchInArrays=True
				Exit Function
			End If
		Next
	End If

	SearchInArrays=False

End Function
'*********************************************************




'*********************************************************
' 目的：   
'*********************************************************
Function TagCloud(Count)
	Dim i
	If Count<=5 Then
		i=0
	ElseIf Count>5 And Count<=10 Then
		i=1
	ElseIf Count>10 And Count<=20 Then
		i=2
	ElseIf Count>20 And Count<=35 Then
		i=3
	ElseIf Count>35 And Count<=70 Then
		i=4
	ElseIf Count>70 And Count<=130 Then
		i=5
	ElseIf Count>130 And Count<=200 Then
		i=6
	ElseIf Count>200 Then
		i=7
	End If
	TagCloud=i
End Function
'*********************************************************




'*********************************************************
' 目的：   检查是否手机端访问
'*********************************************************
Function CheckMobile()

	CheckMobile=False

	'是否（智能）手机浏览器
	Dim MobileBrowser_List,UserAgent
	MobileBrowser_List ="android|iphone|ipad|windows\sphone|kindle|rim\stablet|meego|netfront|java|opera\smini|opera\smobi|ucweb|windows\sce|symbian|series|webos|sonyericsson|sony|blackberry|cellphone|dopod|nokia|samsung|palmsource|palmos|xphone|xda|smartphone|meizu|up.browser|up.link|pieplus|midp|cldc|motorola|foma|docomo|huawei|coolpad|alcatel|amoi|ktouch|philips|benq|haier|bird|zte|wap|mobile"

	UserAgent = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
	If CheckRegExp(UserAgent,MobileBrowser_List) Then 
		CheckMobile=True
		Exit Function
	End If 

	'是否专用wap浏览器
	If InStr(LCase(Request.ServerVariables("HTTP_ACCEPT")), "application/vnd.wap.xhtml+xml") Then
		CheckMobile=True
		Exit Function
	End If
	If InStr(LCase(Request.ServerVariables("HTTP_VIA")), "wap")>0 Then
		CheckMobile=True
		Exit Function
	End If
	If Not IsEmpty(Request.ServerVariables("HTTP_X_WAP_PROFILE")) Then
		CheckMobile=True
		Exit Function
	End If
	If Not IsEmpty(Request.ServerVariables("HTTP_PROFILE")) Then
		CheckMobile=True
		Exit Function
	End If

End Function 
'*********************************************************




'*********************************************************
Function ParseDateForRFC822GMT(dtmDate)

	dtmDate=DateAdd("h", 0-(CInt(ZC_HOST_TIME_ZONE)/100), dtmDate)

	Dim dtmDay, dtmWeekDay, dtmMonth, dtmYear
	Dim dtmHours, dtmMinutes, dtmSeconds

	Select Case WeekDay(dtmDate)
		Case 1:dtmWeekDay="Sun"
		Case 2:dtmWeekDay="Mon"
		Case 3:dtmWeekDay="Tue"
		Case 4:dtmWeekDay="Wed"
		Case 5:dtmWeekDay="Thu"
		Case 6:dtmWeekDay="Fri"
		Case 7:dtmWeekDay="Sat"
	End Select

	Select Case Month(dtmDate)
		Case 1:dtmMonth="Jan"
		Case 2:dtmMonth="Feb"
		Case 3:dtmMonth="Mar"
		Case 4:dtmMonth="Apr"
		Case 5:dtmMonth="May"
		Case 6:dtmMonth="Jun"
		Case 7:dtmMonth="Jul"
		Case 8:dtmMonth="Aug"
		Case 9:dtmMonth="Sep"
		Case 10:dtmMonth="Oct"
		Case 11:dtmMonth="Nov"
		Case 12:dtmMonth="Dec"
	End Select

	dtmYear = Year(dtmDate)
	dtmDay = Right("00" & Day(dtmDate),2)

	dtmHours = Right("00" & Hour(dtmDate),2)
	dtmMinutes = Right("00" & Minute(dtmDate),2)
	dtmSeconds = Right("00" & Second(dtmDate),2)

	ParseDateForRFC822GMT = dtmWeekDay & ", " & dtmDay &" " & dtmMonth & " " & dtmYear & " " & dtmHours & ":" & dtmMinutes & ":" & dtmSeconds & " GMT"

End Function
'*********************************************************




'*********************************************************
' Derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm,
' as set out in the memo RFC1321.
'
' See the VB6 project that accompanies this sample for full code comments on how
' it works.
'
' ASP VBScript code for generating an MD5 'digest' or 'signature' of a string. The
' MD5 algorithm is one of the industry standard methods for generating digital
' signatures. It is generically known as a digest, digital signature, one-way
' encryption, hash or checksum algorithm. A common use for MD5 is for password
' encryption as it is one-way in nature, that does not mean that your passwords
' are not free from a dictionary attack.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on this code provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk

Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32

Private m_lOnBits(30)
Private m_l2Power(30)
 
    m_lOnBits(0) = CLng(1)
    m_lOnBits(1) = CLng(3)
    m_lOnBits(2) = CLng(7)
    m_lOnBits(3) = CLng(15)
    m_lOnBits(4) = CLng(31)
    m_lOnBits(5) = CLng(63)
    m_lOnBits(6) = CLng(127)
    m_lOnBits(7) = CLng(255)
    m_lOnBits(8) = CLng(511)
    m_lOnBits(9) = CLng(1023)
    m_lOnBits(10) = CLng(2047)
    m_lOnBits(11) = CLng(4095)
    m_lOnBits(12) = CLng(8191)
    m_lOnBits(13) = CLng(16383)
    m_lOnBits(14) = CLng(32767)
    m_lOnBits(15) = CLng(65535)
    m_lOnBits(16) = CLng(131071)
    m_lOnBits(17) = CLng(262143)
    m_lOnBits(18) = CLng(524287)
    m_lOnBits(19) = CLng(1048575)
    m_lOnBits(20) = CLng(2097151)
    m_lOnBits(21) = CLng(4194303)
    m_lOnBits(22) = CLng(8388607)
    m_lOnBits(23) = CLng(16777215)
    m_lOnBits(24) = CLng(33554431)
    m_lOnBits(25) = CLng(67108863)
    m_lOnBits(26) = CLng(134217727)
    m_lOnBits(27) = CLng(268435455)
    m_lOnBits(28) = CLng(536870911)
    m_lOnBits(29) = CLng(1073741823)
    m_lOnBits(30) = CLng(2147483647)
    
    m_l2Power(0) = CLng(1)
    m_l2Power(1) = CLng(2)
    m_l2Power(2) = CLng(4)
    m_l2Power(3) = CLng(8)
    m_l2Power(4) = CLng(16)
    m_l2Power(5) = CLng(32)
    m_l2Power(6) = CLng(64)
    m_l2Power(7) = CLng(128)
    m_l2Power(8) = CLng(256)
    m_l2Power(9) = CLng(512)
    m_l2Power(10) = CLng(1024)
    m_l2Power(11) = CLng(2048)
    m_l2Power(12) = CLng(4096)
    m_l2Power(13) = CLng(8192)
    m_l2Power(14) = CLng(16384)
    m_l2Power(15) = CLng(32768)
    m_l2Power(16) = CLng(65536)
    m_l2Power(17) = CLng(131072)
    m_l2Power(18) = CLng(262144)
    m_l2Power(19) = CLng(524288)
    m_l2Power(20) = CLng(1048576)
    m_l2Power(21) = CLng(2097152)
    m_l2Power(22) = CLng(4194304)
    m_l2Power(23) = CLng(8388608)
    m_l2Power(24) = CLng(16777216)
    m_l2Power(25) = CLng(33554432)
    m_l2Power(26) = CLng(67108864)
    m_l2Power(27) = CLng(134217728)
    m_l2Power(28) = CLng(268435456)
    m_l2Power(29) = CLng(536870912)
    m_l2Power(30) = CLng(1073741824)

Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
    Dim lX4
    Dim lY4
    Dim lX8
    Dim lY8
    Dim lResult
 
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
 
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
 
    AddUnsigned = lResult
End Function

Private Function FFFF(x, y, z)
    FFFF = (x And y) Or ((Not x) And z)
End Function

Private Function GGGG(x, y, z)
    GGGG = (x And z) Or (y And (Not z))
End Function

Private Function HHHH(x, y, z)
    HHHH = (x Xor y Xor z)
End Function

Private Function IIII(x, y, z)
    IIII = (y Xor (x Or (Not z)))
End Function

Private Sub FF(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(FFFF(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub GG(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(GGGG(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub HH(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(HHHH(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub II(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(IIII(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(sMessage)
    Dim lMessageLength
    Dim lNumberOfWords
    Dim lWordArray()
    Dim lBytePosition
    Dim lByteCount
    Dim lWordCount
    
    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448
    
    lMessageLength = Len(sMessage)
    
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
    
    ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(lValue)
    Dim lByte
    Dim lCount
    
    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
    Next
End Function

Public Function MD5(sMessage)
    Dim x
    Dim k
    Dim AA
    Dim BB
    Dim CC
    Dim DD
    Dim a
    Dim b
    Dim c
    Dim d
    
    Const S11 = 7
    Const S12 = 12
    Const S13 = 17
    Const S14 = 22
    Const S21 = 5
    Const S22 = 9
    Const S23 = 14
    Const S24 = 20
    Const S31 = 4
    Const S32 = 11
    Const S33 = 16
    Const S34 = 23
    Const S41 = 6
    Const S42 = 10
    Const S43 = 15
    Const S44 = 21

    x = ConvertToWordArray(sMessage)
    
    a = &H67452301
    b = &HEFCDAB89
    c = &H98BADCFE
    d = &H10325476

    For k = 0 To UBound(x) Step 16
        AA = a
        BB = b
        CC = c
        DD = d
    
        FF a, b, c, d, x(k + 0), S11, &HD76AA478
        FF d, a, b, c, x(k + 1), S12, &HE8C7B756
        FF c, d, a, b, x(k + 2), S13, &H242070DB
        FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
        FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
        FF d, a, b, c, x(k + 5), S12, &H4787C62A
        FF c, d, a, b, x(k + 6), S13, &HA8304613
        FF b, c, d, a, x(k + 7), S14, &HFD469501
        FF a, b, c, d, x(k + 8), S11, &H698098D8
        FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
        FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
        FF b, c, d, a, x(k + 11), S14, &H895CD7BE
        FF a, b, c, d, x(k + 12), S11, &H6B901122
        FF d, a, b, c, x(k + 13), S12, &HFD987193
        FF c, d, a, b, x(k + 14), S13, &HA679438E
        FF b, c, d, a, x(k + 15), S14, &H49B40821
    
        GG a, b, c, d, x(k + 1), S21, &HF61E2562
        GG d, a, b, c, x(k + 6), S22, &HC040B340
        GG c, d, a, b, x(k + 11), S23, &H265E5A51
        GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
        GG a, b, c, d, x(k + 5), S21, &HD62F105D
        GG d, a, b, c, x(k + 10), S22, &H2441453
        GG c, d, a, b, x(k + 15), S23, &HD8A1E681
        GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
        GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
        GG d, a, b, c, x(k + 14), S22, &HC33707D6
        GG c, d, a, b, x(k + 3), S23, &HF4D50D87
        GG b, c, d, a, x(k + 8), S24, &H455A14ED
        GG a, b, c, d, x(k + 13), S21, &HA9E3E905
        GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
        GG c, d, a, b, x(k + 7), S23, &H676F02D9
        GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
            
        HH a, b, c, d, x(k + 5), S31, &HFFFA3942
        HH d, a, b, c, x(k + 8), S32, &H8771F681
        HH c, d, a, b, x(k + 11), S33, &H6D9D6122
        HH b, c, d, a, x(k + 14), S34, &HFDE5380C
        HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
        HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
        HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
        HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
        HH a, b, c, d, x(k + 13), S31, &H289B7EC6
        HH d, a, b, c, x(k + 0), S32, &HEAA127FA
        HH c, d, a, b, x(k + 3), S33, &HD4EF3085
        HH b, c, d, a, x(k + 6), S34, &H4881D05
        HH a, b, c, d, x(k + 9), S31, &HD9D4D039
        HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
        HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
        HH b, c, d, a, x(k + 2), S34, &HC4AC5665
    
        II a, b, c, d, x(k + 0), S41, &HF4292244
        II d, a, b, c, x(k + 7), S42, &H432AFF97
        II c, d, a, b, x(k + 14), S43, &HAB9423A7
        II b, c, d, a, x(k + 5), S44, &HFC93A039
        II a, b, c, d, x(k + 12), S41, &H655B59C3
        II d, a, b, c, x(k + 3), S42, &H8F0CCC92
        II c, d, a, b, x(k + 10), S43, &HFFEFF47D
        II b, c, d, a, x(k + 1), S44, &H85845DD1
        II a, b, c, d, x(k + 8), S41, &H6FA87E4F
        II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
        II c, d, a, b, x(k + 6), S43, &HA3014314
        II b, c, d, a, x(k + 13), S44, &H4E0811A1
        II a, b, c, d, x(k + 4), S41, &HF7537E82
        II d, a, b, c, x(k + 11), S42, &HBD3AF235
        II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
        II b, c, d, a, x(k + 9), S44, &HEB86D391
    
        a = AddUnsigned(a, AA)
        b = AddUnsigned(b, BB)
        c = AddUnsigned(c, CC)
        d = AddUnsigned(d, DD)
    Next
    
    MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
End Function
'*********************************************************




'*********************************************************
' 目的：    newClass
' 输入：   
' 输入：   类名 
' 返回：    一个初始化的类，给JS用
'*********************************************************
Function newClass(ClassName)
	Execute "Set newClass=New " & ClassName
End Function
'*********************************************************
' 目的：    vbsArray
' 输入：   
' 输入：   数组名,i
' 返回：    用于读取VBS数组，给JS用
'*********************************************************
Function vbsArray(arr,i)
	vbsArray=arr(i)
End Function
'*********************************************************
' 目的：    vbsArrayEdit
' 输入：   
' 输入：   数组名,i,内容
' 返回：    用于修改VBS数组，给JS用
'*********************************************************
Function vbsArrayEdit(arr,i,c)
	arr(i)=c
	vbsArrayEdit=arr
End Function
'*********************************************************
' 目的：    vbsArrayRedim
' 输入：   
' 输入：   数组名,下标,保留数据
' 返回：    用于修改VBS数组下标，给JS用
'*********************************************************
Function vbsArrayRedim(arr,i,pre)
	Execute "Redim " & IIf(pre=True,"Preserve ","") & "arr(i)" 
	vbsArrayRedim=arr
End Function
'*********************************************************
' 目的：    vbsReplace
' 输入：   
' 输入：   
' 返回：    
'*********************************************************
Function vbsReplace(s1,s2,s3)
	vbsReplace=Replace(s1,s2,s3)
End Function

'*********************************************************
' 目的：    unescape
' 输入：    
' 输入：    要替换的字符
' 返回：    
'*********************************************************
%>
<script language="javascript" runat="server">

	function vbsunescape(source){
		if(typeof(source)=="undefined"){return ""};
		if(source===""){return ""};
		var a;
		a=unescape(source);
		return (a==("undefined"||undefined) ? "" : a)
	}
	function vbsescape(source){
		if(typeof(source)=="undefined"){return ""};
		if(source===""){return ""}
		var a;
		a=escape(source);
		return (a==("undefined"||undefined) ? "" : a)
	}
	String.prototype.vbsreplace=function(s1,s2){return vbsReplace(this,s1,s2)}
</script>
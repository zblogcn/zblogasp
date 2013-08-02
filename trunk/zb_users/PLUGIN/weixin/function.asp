<%
'****************************************
' weixin 子菜单
'****************************************
Function weixin_SubMenu(id)
	Dim aryName,aryPath,aryFloat,aryInNewWindow,i
	aryName=Array("基本设置","微信连接设置")
	aryPath=Array("main.asp","tokenst.asp")
	aryFloat=Array("m-left","m-left")
	aryInNewWindow=Array(False,False)
	For i=0 To Ubound(aryName)
		weixin_SubMenu=weixin_SubMenu & MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i))
	Next
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'首次关注
Function wx_Welcome(blogtitle,num,welcomestr)
	wx_Welcome = Replace(welcomestr,"{%title%}",blogtitle)
	wx_Welcome = Replace(wx_Welcome,"{%num%}",num)
	wx_Welcome = Replace(wx_Welcome,"<br/>",vbCrLf)
End Function

'帮助
Function wx_Help()
	wx_Help="您可以输入“最新文章”来查看博客的最新图文文章；或者输入关键词来搜索博客中的文章并在微信中查看。"
End Function


'查询文章
Function wx_Search(strQuestion,show_num,shou_meta)
	Dim LTRS,InserNewHtml:InserNewHtml = ""
	
	If ZC_MSSQL_ENABLE=False Then
		Set LTRS=objConn.Execute("SELECT TOP "&show_num&"  [log_ID],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_PostTime],[log_FullUrl] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_ID]>0) AND( (InStr(1,LCase([log_Title]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('"&strQuestion&"'),0)<>0) )")
	Else
		Set LTRS=objConn.Execute("SELECT TOP "&show_num&"  [log_ID],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_PostTime],[log_FullUrl] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_ID]>0) AND( (CHARINDEX('"&strQuestion&"',[log_Title])<>0) OR (CHARINDEX('"&strQuestion&"',[log_Intro])<>0) OR (CHARINDEX('"&strQuestion&"',[log_Content])<>0) )")
	End If	
	
	Do Until LTRS.Eof
		If shou_meta=1 Then
			InserNewHtml = InserNewHtml & "<a href=""" & ZC_BLOG_HOST & "ZB_USERS/plugin/weixin/viewwx.asp?wid=" & LTRS("log_ID") & """>" & LTRS("log_ID") & "、" & LTRS("log_Title") & "</a>" & VBCrLf & VBCrLf
		ElseIf shou_meta=3 Then
			InserNewHtml = InserNewHtml & "<a href=""" & ZC_BLOG_HOST & "view.asp?nav=" & LTRS("log_ID") & """>" & LTRS("log_ID") & "、" & LTRS("log_Title") & "</a>" & VBCrLf & VBCrLf 
		End If
		'InserNewHtml = InserNewHtml & TransferHTML(LTRS("log_Content"),"[nohtml]")
		'Exit Do
		LTRS.MoveNext
	Loop
	Set LTRS=Nothing

	InserNewHtml = Replace(InserNewHtml,"&nbsp;"," ")
	InserNewHtml = Replace(InserNewHtml,"<#ZC_BLOG_HOST#>",BlogHost)
	
	wx_Search = "“" & strQuestion & "”搜索结果：" & VBCrLf
	wx_Search = wx_Search & InserNewHtml & "  提示：请直接点击文章标题查看博客文章。"
End Function

'最新文章
Function wx_LastPost(number)
	Dim LTRS,InserNewHtml:InserNewHtml = ""
	Set LTRS=objConn.Execute("SELECT TOP "&number&" [log_ID], [log_Title], [log_Intro], [log_Content], [log_PostTime], [log_Type] FROM blog_Article WHERE ((([log_Type])=0)) ORDER BY [log_PostTime] DESC")
	Do Until LTRS.Eof
		InserNewHtml = InserNewHtml & "<item><Title><![CDATA[" & LTRS("log_Title") & "]]></Title><Description><![CDATA[" & TransferHTML(LTRS("log_Intro"),"[nohtml]") & "]]></Description><PicUrl><![CDATA["

		if wx_GetFirstUrl(LTRS("log_Content"))="" then
			InserNewHtml = InserNewHtml & BlogHost &"ZB_USERS/plugin/weixin/defaultpic.jpg"
		else
			InserNewHtml = InserNewHtml & wx_GetFirstUrl(LTRS("log_Content"))
		End if

		InserNewHtml = InserNewHtml & "]]></PicUrl><Url><![CDATA[" & ZC_BLOG_HOST & "ZB_USERS/plugin/weixin/viewwx.asp?wid=" & LTRS("log_ID") & "]]></Url></item>"
		LTRS.MoveNext
	Loop
	Set LTRS=Nothing

	InserNewHtml = Replace(InserNewHtml,"&nbsp;"," ")
	InserNewHtml = Replace(InserNewHtml,"<#ZC_BLOG_HOST#>",BlogHost)

	wx_LastPost = InserNewHtml
End Function

'=======================================================
'函数: 从正文中提取图片路径.
'输入: 文章全文.
'返回: 有图则返回图片路径, 无图返回空.
'=======================================================
Function wx_GetFirstUrl(ByVal strContent)
	'On Error Resume Next
	Dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase=True
	objRegExp.Global=False

	objRegExp.Pattern="(<img[^>]+(src|data-original)[^""]+"")([^""]+)([^>]+>)"

	Dim Match, Matches, Value
	Set Matches=objRegExp.Execute(strContent)
		For Each Match in Matches
			Value=objRegExp.Replace(Match.value,"$3")
		Next
	Set Matches=Nothing

	Set objRegExp=Nothing

	wx_GetFirstUrl=Value

	'Err.Clear
End Function
%>
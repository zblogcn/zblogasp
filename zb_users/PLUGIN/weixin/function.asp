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
Function Welcome(title,num)
	Welcome = Content & "欢迎关注《"&title&"》！！！"& VBCrLf
	Welcome = Welcome & "您可发送“最新文章”来查看博客最新的"&num&"篇文章，或者直接发送关键词来搜索博客中已发表的文章。更多使用帮助请输入英文“help”或者数字“0”来查看。"
End Function

'查询文章
Function Search(Content)
	Dim LTRS,InserNewHtml:InserNewHtml = ""
	Set LTRS=objConn.Execute("SELECT TOP 15  [log_ID],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_PostTime],[log_FullUrl] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_ID]>0) AND( (InStr(1,LCase([log_Title]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('"&strQuestion&"'),0)<>0) )")
	Do Until LTRS.Eof
		InserNewHtml = InserNewHtml & "<a href=""" & ZC_BLOG_HOST & "ZB_USERS/plugin/weixin/view.asp?wid=" & LTRS("log_ID") & """>" & LTRS("log_ID") & "、" & LTRS("log_Title") & "</a>" & VBCrLf & VBCrLf 'LTRS("log_PostTime") & 
		'InserNewHtml = InserNewHtml & TransferHTML(LTRS("log_Content"),"[nohtml]")
		'Exit Do
		LTRS.MoveNext
	Loop
	Set LTRS=Nothing

	InserNewHtml = Replace(InserNewHtml,"&nbsp;"," ")
	InserNewHtml = Replace(InserNewHtml,"<#ZC_BLOG_HOST#>",BlogHost)
	
	Content = "“" & Content & "”搜索结果：" & VBCrLf
	Search = Content & InserNewHtml & VBCrLf & "  提示：请直接点击文章标题查看博客文章，或者回复标题前的编号直接在微信中查看文字版。"
End Function

'最新文章
Function LastPost()
	Dim LTRS,InserNewHtml:InserNewHtml = ""
	Set LTRS=objConn.Execute("SELECT TOP 5 [log_ID], [log_Title], [log_Intro], [log_Content], [log_PostTime], [log_Type] FROM blog_Article WHERE ((([log_Type])=0)) ORDER BY [log_PostTime] DESC")
	Do Until LTRS.Eof
		InserNewHtml = InserNewHtml & "<item><Title><![CDATA[" & LTRS("log_Title") & "]]></Title><Description><![CDATA[" & TransferHTML(LTRS("log_Intro"),"[nohtml]") & "]]></Description><PicUrl><![CDATA["

		if GetFirstUrl(LTRS("log_Content"))="" then
			InserNewHtml = InserNewHtml & "http://imzhou.com/zb_system/image/logo/zblog.gif"
		else
			InserNewHtml = InserNewHtml & GetFirstUrl(LTRS("log_Content"))
		End if

		InserNewHtml = InserNewHtml & "]]></PicUrl><Url><![CDATA[" & ZC_BLOG_HOST & "ZB_USERS/plugin/weixin/view.asp?wid=" & LTRS("log_ID") & "]]></Url></item>"
		LTRS.MoveNext
	Loop
	Set LTRS=Nothing

	InserNewHtml = Replace(InserNewHtml,"&nbsp;"," ")
	InserNewHtml = Replace(InserNewHtml,"<#ZC_BLOG_HOST#>",BlogHost)

	LastPost = InserNewHtml
End Function

'=======================================================
'函数: 从正文中提取图片路径.
'输入: 文章全文.
'返回: 有图则返回图片路径, 无图返回空.
'=======================================================
Function GetFirstUrl(ByVal strContent)
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

	GetFirstUrl=Value

	'Err.Clear
End Function

%>
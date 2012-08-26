<%
'///////////////////////////////////////////////////////////////////////////////
'// 月上之木 2012.8.25
'///////////////////////////////////////////////////////////////////////////////


'注册插件
Call RegisterPlugin("SweetTitles","ActivePlugin_SweetTitles")

Function ActivePlugin_SweetTitles()

	Call Add_Filter_Plugin("Filter_Plugin_TArticleList_Build_Template","SweetTitles_InnerCode_List")

	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Export_Template","SweetTitles_InnerCode_Single")

End Function

Function SweetTitles_InnerCode_Single(ByRef html,ByRef subhtml)

	Call SweetTitles_InnerCode(html)

End Function


Function SweetTitles_InnerCode_List(ByRef html)

	Call SweetTitles_InnerCode(html)

End Function



Function SweetTitles_InnerCode(ByRef Ftemplate)

	If IsEmpty(Ftemplate) Then Exit Function

	Dim objRegExp
	Set objRegExp = new RegExp
	objRegExp.IgnoreCase = True
	objRegExp.Global = False

	objRegExp.Pattern = "(</head>)"
	Ftemplate = objRegExp.Replace(Ftemplate,SweetTitles_Code&"$1")

	Set objRegExp = Nothing

End Function


Function SweetTitles_Code

	Dim innerHtml : innerHtml = vbTab
	innerHtml = innerHtml & "<script type=""text/javascript"" src=""" & BlogHost & "zb_users/PLUGIN/SweetTitles/jquery.sweetTitles.js""></script>" & vbCrLf & vbTab
	innerHtml = innerHtml & "<link rel=""stylesheet"" href=""" & BlogHost & "zb_users/PLUGIN/SweetTitles/sweetTitles.css"" />" & vbCrLf

	
	SweetTitles_Code = innerHtml

End Function


'安装插件
Function InstallPlugin_SweetTitles

	Call SetBlogHint(Empty,True,True)

End Function

'卸载插件
Function UnInstallPlugin_SweetTitles

	Call SetBlogHint(Empty,True,True)

End Function
%>
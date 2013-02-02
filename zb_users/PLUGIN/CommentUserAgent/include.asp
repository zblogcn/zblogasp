<!-- #include file="function.asp" -->
<%
Dim CommentUserAgent_UserAgent
'注册插件
Call RegisterPlugin("CommentUserAgent","ActivePlugin_CommentUserAgent")
'挂口部分
Function ActivePlugin_CommentUserAgent()

	Call Add_Filter_Plugin("Filter_Plugin_TComment_MakeTemplate_TemplateTags","CommentUserAgent_Comment_MakeTemplate_TemplateTags")
	
End Function


Function CommentUserAgent_Comment_MakeTemplate_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)
	On Error Resume Next
	
	If UBound(aryTemplateTagsValue)>26 Then
		If aryTemplateTagsName(27)="article/comment/agent" Then
			CommentUserAgent_UserAgent=aryTemplateTagsValue(27)
		Else
			CommentUserAgent_UserAgent=objConn.Execute("SELECT [comm_Agent] FROM [blog_Comment] WHERE [comm_ID]="&aryTemplateTagsValue(1))(0)
		End If
	Else
		CommentUserAgent_UserAgent=objConn.Execute("SELECT [comm_Agent] FROM [blog_Comment] WHERE [comm_ID]="&aryTemplateTagsValue(1))(0)
	End if

	Dim strFull
	Dim img
	Set img=detect_platform(CommentUserAgent_UserAgent)
	Dim aryTag(14),aryValue(14)

	aryTag(1)="zsxsoft/cmtua/useragent"
	aryValue(1)=Replace(TransferHTML(CommentUserAgent_UserAgent,"[html-format]"),"(Unreal_CommentUserAgentAutoGenerate)","")
	aryTag(2)="zsxsoft/cmtua/platform/src"
	aryValue(2)=BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename
	aryTag(3)="zsxsoft/cmtua/platform/system"
	aryValue(3)=img.text
	aryTag(4)="zsxsoft/cmtua/platform/version"
	aryValue(4)=img.ver
	aryTag(5)="zsxsoft/cmtua/platform/link"
	aryValue(5)=img.link
	aryTag(6)="zsxsoft/cmtua/platform/img"
	aryValue(6)="<img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename&"' width='16 height='16' alt='"&img.text&"系统' />"
	aryTag(7)="zsxsoft/cmtua/platform"
	aryValue(7)="<span class='cmtua_platform'><img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename&"' width='16 height='16' alt='"&img.text&"系统' title='"&img.text&"'/>"&img.text&"</span>"
	
	Set img=detect_webbrowser(CommentUserAgent_UserAgent)
	aryTag(8)="zsxsoft/cmtua/browser/src"
	aryValue(8)=BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename
	aryTag(9)="zsxsoft/cmtua/browser/browser"
	aryValue(9)=img.text
	aryTag(10)="zsxsoft/cmtua/browser/version"
	aryValue(10)=img.ver
	aryTag(11)="zsxsoft/cmtua/browser/link"
	aryValue(11)=img.link
	aryTag(12)="zsxsoft/cmtua/browser/img"
	aryValue(12)="<img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename&"' width='16 height='16' alt='"&img.text&"浏览器' />"
	aryTag(13)="zsxsoft/cmtua/browser"
	aryValue(13)="<span class='cmtua_browser'><img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename&"' width='16 height='16' alt='"&img.text&"浏览器' title='"&img.text&"' />"&img.text&"</span>"

	aryTag(14)="zsxsoft/cmtua/all"
	aryValue(14)="<div class='cmtua'>"&aryValue(7)&"&nbsp;&nbsp;"&aryValue(13)&"</div>"

	Dim jj,s
	s=Ubound(aryTemplateTagsName)
	Redim Preserve aryTemplateTagsName(s+Ubound(aryTag)+1)
	Redim Preserve aryTemplateTagsValue(s+Ubound(aryTag)+1)
	For jj=0 To Ubound(aryTag)
		aryTemplateTagsName(s+jj+1)=aryTag(jj)
		aryTemplateTagsValue(s+jj+1)=aryValue(jj)
	Next
	
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load "CommentUserAgent"
	Dim j_,k_,m_
	j_=Split(objConfig.Read("Item"),",")
	
	For k_=0 To Ubound(j_)
		m_=CInt(j_(k_))
		strFull=strFull & aryValue(m_+1)
	Next
	

	aryTemplateTagsValue(7)=aryTemplateTagsValue(7) & strFull


End Function

Function InstallPlugin_CommentUserAgent()
	Call SetBlogHint(Empty,Empty,True)
End Function

Function UninstallPlugin_CommentUserAgent()
	Call SetBlogHint(Empty,Empty,True)
End Function
%>
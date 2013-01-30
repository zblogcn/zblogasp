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

	CommentUserAgent_UserAgent=aryTemplateTagsValue(27)

	Dim strFull
	Dim img
	Set img=detect_platform(CommentUserAgent_UserAgent)
	Dim s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13,s14
	Dim t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12,t13,t14

	s1="<#zsxsoft/cmtua/useragent#>"
	t1=Replace(TransferHTML(CommentUserAgent_UserAgent,"[html-format]"),"(Unreal_CommentUserAgentAutoGenerate)","")
	s2="<#zsxsoft/cmtua/platform/src#>"
	t2=BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename
	s3="<#zsxsoft/cmtua/platform/system#>"
	t3=img.text
	s4="<#zsxsoft/cmtua/platform/version#>"
	t4=img.ver
	s5="<#zsxsoft/cmtua/platform/link#>"
	t5=img.link
	s6="<#zsxsoft/cmtua/platform/img#>"
	t6="<img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename&"' width='16 height='16' alt='"&img.text&"系统' />"
	s7="<#zsxsoft/cmtua/platform#>"
	t7="<span class='cmtua_platform'><img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename&"' width='16 height='16' alt='"&img.text&"系统' />"&img.text&"</span>"
	
	Set img=detect_webbrowser(CommentUserAgent_UserAgent)
	s8="<#zsxsoft/cmtua/browser/src#>"
	t8=BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename
	s9="<#zsxsoft/cmtua/browser/browser#>"
	t9=img.text
	s10="<#zsxsoft/cmtua/browser/version#>"
	t10=img.ver
	s11="<#zsxsoft/cmtua/browser/link#>"
	t11=img.link
	s12="<#zsxsoft/cmtua/browser/img#>"
	t12="<img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename&"' width='16 height='16' alt='"&img.text&"浏览器' />"
	s13="<#zsxsoft/cmtua/browser#>"
	t13="<span class='cmtua_browser'><img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & img.fullfilename&"' width='16 height='16' alt='"&img.text&"浏览器' />"&img.text&"</span>"

	s14="<#zsxsoft/cmtua/all#>"
	t14="<div class='cmtua'>"&t7&"&nbsp;&nbsp;"&t13&"</div>"

	Call Execute("strFull=t14")'在这里把设置代入吧,骚年

	Call SetValueByNameInArrays(aryTemplateTagsName,aryTemplateTagsValue,"article/comment/content",aryTemplateTagsValue(7) & strFull)


End Function
%>
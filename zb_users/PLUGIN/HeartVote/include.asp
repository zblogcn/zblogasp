<%


'注册插件
Call RegisterPlugin("HeartVote","ActivePlugin_HeartVote")

Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type='text/javascript' src='" & BlogHost & "zb_users/plugin/heartvote/js/vote.js'></script>"");")

Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<link rel='stylesheet' rev='stylesheet'  href='" & BlogHost & "zb_users/plugin/heartvote/css/stars.css' type='text/css' media='screen' />"");")

Function InstallPlugin_HeartVote()

	On Error Resume Next

	Err.Clear

End Function


Function UninstallPlugin_HeartVote()

	On Error Resume Next

	Err.Clear

End Function


'具体的接口挂接
Function ActivePlugin_HeartVote() 

	'Action_Plugin_TArticle_Export_Begin
	Call Add_Action_Plugin("Action_Plugin_TArticle_Export_Begin","Call Add_Filter_Plugin(""Filter_Plugin_TArticle_Export_TemplateTags"",""HeartVote_Filter_Plugin_TArticle_Export_TemplateTags"")")

End Function

Function HeartVote_Filter_Plugin_TArticle_Export_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue) 

	On Error Resume Next

	Dim i,s
	i=UBound(aryTemplateTagsName)

ReDim Preserve aryTemplateTagsName(i+1)

ReDim Preserve aryTemplateTagsValue(i+1)

aryTemplateTagsName(i)="HeartVoteForm"

s=s&"<div class=""heart-vote"" id=""HeartVote_"&aryTemplateTagsValue(1)&""">"
s=s&"<ul class=""unit-rating"">"
s=s&"<li class='current-rating' style=""width:0px;""></li>"
s=s&"<li><a href=""javascript:heartVote('1','"&aryTemplateTagsValue(1)&"')"" title=""打1分"" class=""r1-unit"">1</a></li>"
s=s&"<li><a href=""javascript:heartVote('2','"&aryTemplateTagsValue(1)&"')"" title=""打2分"" class=""r2-unit"">2</a></li>"
s=s&"<li><a href=""javascript:heartVote('3','"&aryTemplateTagsValue(1)&"')"" title=""打3分"" class=""r3-unit"">3</a></li>"
s=s&"<li><a href=""javascript:heartVote('4','"&aryTemplateTagsValue(1)&"')"" title=""打4分"" class=""r4-unit"">4</a></li>"
s=s&"<li><a href=""javascript:heartVote('5','"&aryTemplateTagsValue(1)&"')"" title=""打5分"" class=""r5-unit"">5</a></li>"
s=s&"<li><a href=""javascript:heartVote('6','"&aryTemplateTagsValue(1)&"')"" title=""打6分"" class=""r6-unit"">6</a></li>"
s=s&"<li><a href=""javascript:heartVote('7','"&aryTemplateTagsValue(1)&"')"" title=""打7分"" class=""r7-unit"">7</a></li>"
s=s&"<li><a href=""javascript:heartVote('8','"&aryTemplateTagsValue(1)&"')"" title=""打8分"" class=""r8-unit"">8</a></li>"
s=s&"<li><a href=""javascript:heartVote('9','"&aryTemplateTagsValue(1)&"')"" title=""打9分"" class=""r9-unit"">9</a></li>"
s=s&"<li><a href=""javascript:heartVote('10','"&aryTemplateTagsValue(1)&"')"" title=""打10分"" class=""r10-unit"">10</a></li>"
s=s&"</ul><p><b>0</b><i>分/0个投票</i></p></div>"

s=s&"<script src=""<#ZC_BLOG_HOST#>zb_users/plugin/heartvote/getvote.asp?id="&aryTemplateTagsValue(1)&""" type=""text/javascript""></script>"

aryTemplateTagsValue(i)=s

	Err.Clear

End Function

%>
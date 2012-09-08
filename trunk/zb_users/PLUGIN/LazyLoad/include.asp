<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog 1.8
'// 作    者:    ZSXSOFT
'//	http://zsxsoft.com
'///////////////////////////////////////////////////////////////////////////////

	Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"" src=\""" & BlogHost & "zb_users/PLUGIN/LazyLoad/jquery.lazyload.min.js\""></script>"");")
	Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","$(document).ready(function($){$('img').lazyload({effect: 'fadeIn'});});")


'注册插件
Call RegisterPlugin("LazyLoad","ActivePlugin_LazyLoad")

Function ActivePlugin_LazyLoad()

	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Export_TemplateTags","LazyLoad_Add_Code1")
	Call Add_Filter_Plugin("Filter_Plugin_TComment_MakeTemplate_Template","LazyLoad_Add_Code3")
	Call Add_Filter_Plugin("Filter_Plugin_TComment_MakeTemplate_TemplateTags","LazyLoad_Add_Code4")
End Function

Function LazyLoad_Add_Code3(ByRef CommentHTML)
	Call LazyLoad_InnerCode2(CommentHTML)
End Function

Function LazyLoad_Add_Code4(ByRef aryTemplateSubName2,ByRef aryTemplateSubValue2)
	Call LazyLoad_InnerCode2(aryTemplateSubValue2(7))

End Function
Function LazyLoad_Add_Code1(ByRef aryTemplateSubName2,ByRef aryTemplateSubValue2)
	Call LazyLoad_InnerCode2(aryTemplateSubValue2(5))
	Call LazyLoad_InnerCode2(aryTemplateSubValue2(4))
End Function
'插入代码



Function LazyLoad_InnerCode2(ByRef HtmlContent)
	Dim NewRegExp,ForVar1,ForVar2
	Set NewRegExp=New RegExp
	NewRegExp.Pattern="<img([\D\d]+?)src=([""'])([\D\d]+?)([""'])([\D\d]+?)>"
	NewRegExp.Global=True
	HtmlContent=NewRegExp.Replace(HtmlContent,"<img$1 data-original=$2$3$4 $5>")
End Function


'安装插件
Function InstallPlugin_LazyLoad

	Call SetBlogHint(Empty,Empty,True)

End Function

'卸载插件
Function UnInstallPlugin_LazyLoad

	Call SetBlogHint(Empty,Empty,True)

End Function



%>
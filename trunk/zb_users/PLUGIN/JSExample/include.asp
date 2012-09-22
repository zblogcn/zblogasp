<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog 1.8
'// 作    者:    ZSXSOFT
'//	http://zsxsoft.com
'///////////////////////////////////////////////////////////////////////////////
Const JSEMPTY=Empty
JSExample()
%>
<script language="javascript" runat="server">
function JSExample(){
	RegisterPlugin("JSExample","ActivePlugin_JSExample");
	Add_Response_Plugin("Response_Plugin_Html_Js_Add","$(document).ready(function(){$(\"#BlogSubTitle\").after(\"JSExample\")})");
}
function ActivePlugin_JSExample(){
	Add_Action_Plugin("Action_Plugin_Edit_Form","JSExample_ResponseWrite(EditArticle)");
	Add_Filter_Plugin("Filter_Plugin_PostArticle_Succeed","JSExample_ShowMsg");
}
function JSExample_ResponseWrite(o){Add_Response_Plugin("Response_Plugin_Edit_Form","JSExample插件强力路过ID为"+o.ID+"的文章！")}
function JSExample_ShowMsg(o){SetBlogHint_Custom("JSExample提示：\n<br/>ArticleID:"+o.ID+"\n<br/>FormData:"+Request.Form)}
function UnInstallPlugin_jsExample(){
	SetBlogHint(JSEMPTY,JSEMPTY,true)
}
</script>
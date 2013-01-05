<script language="javascript" src="Config.js" runat="server" type="text/javascript"></script>
<!--#include file="FUNCTION/YT.Function.asp" -->
<%
Call RegisterPlugin("YTCMS","ActivePlugin_YT_CMS")
Sub ActivePlugin_YT_CMS()
	Call Add_Filter_Plugin("Filter_Plugin_TArticleList_Build_Template_Succeed","YT_CMS_Filter_Plugin_TArticleList_Build_Template_Succeed")
	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Build_Template_Succeed","YT_CMS_Filter_Plugin_TArticle_Build_Template_Succeed")
	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Del","YT_CMS_Filter_Plugin_TArticle_Del")
	Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Succeed","YT_CMS_Filter_Plugin_PostArticle_Succeed")
	Call Add_Action_Plugin("Action_Plugin_MakeBlogReBuild_Core_Begin","YT_CMS_Action_Plugin_MakeBlogReBuild_Core_Begin")
	Call Add_Action_Plugin("Action_Plugin_TArticle_Export_End","Call YT_CMS_Action_Plugin_TArticle_Export_End(html,subhtml,aryTemplateTagsName,aryTemplateTagsValue)")
	Call Add_Response_Plugin("Response_Plugin_Edit_Form",YT_Model_Analysis)
	Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(1,"Content Manage System",BlogHost&"zb_users/plugin/YTCMS/YT.Panel.asp","","_top"))
End Sub
Function UnInstallPlugin_YTCMS()
	'卸载模型
'	Dim t,i,j
'	Set t=new YT_Model_XML
'		i=t.Length
'		For j=0 To i-1
'			Call t.Model("UnInstall",j)
'		Next
'	Set t=Nothing
End Function
Function InstallPlugin_YTCMS()
	'安装模型
	Dim t,i,j
	Set t=new YT_Model_XML
		i=t.Length
		For j=0 To i-1
			Call t.Model("Install",j)
		Next
	Set t=Nothing
	Call new YT_Block_XML
End Function
%>
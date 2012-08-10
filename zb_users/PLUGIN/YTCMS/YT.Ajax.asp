<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷(YT.Single)
'// 技术支持:    13120003225@qq.com
'// 程序名称:    	Content Manage System
'// 开始时间:    	2011.03.26
'// 最后修改:    2012-08-08
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #INCLUDE FILE="../../C_OPTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_FUNCTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_LIB.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_BASE.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_EVENT.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_PLUGIN.ASP" -->
<!-- #INCLUDE FILE="../../PLUGIN/P_CONFIG.ASP" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 
If CheckPluginState("YTCMS") = False Then Call ShowError(48)
Dim Action
	Action = Request("Action")
	Select Case Action
		Case "tplList":Response.Write(new YT_Template.List())
		Case "GetFile":Response.Write(new YT_Template.GetFile(Request.QueryString("File")))
		Case "SaveFile":Response.Write(SaveToFile(BlogPath & "ZB_USERS/THEME/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY & "/" & Request.Form("Name"),Request.Form("Content"),"utf-8",False))
		Case "SaveModel":Response.Write(LCase(new YT_Model_XML.Add(jsonToObject(Request.Form("Json")),-1)))
		Case "UpdateModel":Response.Write(LCase(new YT_Model_XML.Add(jsonToObject(Request.Form("Json")),Request.Form("Index"))))
		Case "DelModel":Call new YT_Model_XML.Del(Request.Form("Index"))
		Case "SaveBlock":Response.Write(LCase(new YT_Block_XML.Add(jsonToObject(Request.Form("Json")),-1)))
		Case "UpdateBlock":Response.Write(LCase(new YT_Block_XML.Add(jsonToObject(Request.Form("Json")),Request.Form("Index"))))
		Case "DelTPL":Call new YT_TPL_XML.Del(Request.Form("Index"))
		Case "SaveTPL":Response.Write(LCase(new YT_TPL_XML.Add(jsonToObject(Request.Form("Json")),-1)))
		Case "UpdateTPL":Response.Write(LCase(new YT_TPL_XML.Add(jsonToObject(Request.Form("Json")),Request.Form("Index"))))
		Case "DelBlock":Call new YT_Block_XML.Del(Request.Form("Index"))
		Case "Exist":Response.Write(LCase(new YT_Table.Exist(Request.Form("Name"))))
		Case "Install":Call new YT_Model_XML.Model("Install",Request.Form("Index"))
		Case "UnInstall":Call new YT_Model_XML.Model("UnInstall",Request.Form("Index"))
		Case "GetData":Response.Write(YT_Data_GetRow(Request.Form("Name"),Request.Form("ID")))
	End Select
Call System_Terminate()
If Err.Number<>0 then
  Call ShowError(0)
End If
%>
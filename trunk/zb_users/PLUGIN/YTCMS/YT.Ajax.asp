<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷
'// 技术支持:    33195@qq.com
'// 程序名称:    	Content Manage System
'// 开始时间:    	2011-03-26
'// 最后修改:    2012-11-04
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
Dim Action,code
	Action = Request("Action")
	Response.Clear
	Select Case Action
		Case "tplList":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(YT_FileJsonList())
		Case "GetFile":
			Call Response.AddHeader("Content-Type","text/html")
			Response.Write(YT_GetFile(Request.QueryString("File")))
		Case "SaveFile":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(LCase(YT_SaveFile(Request.Form("Name"),Request.Form("Content"))))
		Case "DelFile":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(LCase(YT_DelFile(Request.Form("Name"))))
		Case "SaveModel":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(LCase(new YT_Model_XML.Add(Request.Form("Json"),-1)))
		Case "UpdateModel":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(LCase(new YT_Model_XML.Add(Request.Form("Json"),Request.Form("Index"))))
		Case "DelModel":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(LCase(new YT_Model_XML.Del(Request.Form("Index"))))
		Case "SaveBlock":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(LCase(new YT_Block_XML.Add(Request.Form("Json"),-1)))
		Case "UpdateBlock":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(LCase(new YT_Block_XML.Add(Request.Form("Json"),Request.Form("Index"))))
		Case "DelBlock":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(LCase(new YT_Block_XML.Del(Request.Form("Index"))))
		Case "Exist":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(LCase(new YT_Table.Exist(Request.Form("Name"))))
		Case "Install":
			Call Response.AddHeader("Content-Type","text/html")
			Call new YT_Model_XML.Model("Install",Request.Form("Index"))
			Response.Write(LCase("Install"))
		Case "UnInstall":
			Call Response.AddHeader("Content-Type","text/html")
			Call new YT_Model_XML.Model("UnInstall",Request.Form("Index"))
			Response.Write(LCase("UnInstall"))
		Case "GetData":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(YT_Data_GetRow(Request.Form("Name"),Request.Form("ID")))
		Case "ImportList":
			Call Response.AddHeader("Content-Type","application/json")
			code = new YT_Table.List()
			dim xl,sxl
			for each xl in code
				sxl = sxl & CHR(34) & xl & CHR(34) & ","
			next
			if right(sxl,1) = "," then sxl = left(sxl,len(sxl)-1)
			Response.Write("["&sxl&"]")
		Case "Import":
			Call Response.AddHeader("Content-Type","application/json")
			Response.Write(LCase(new YT_Table.Import(Request("Name"))))
		Case "Demo":
			Call Response.AddHeader("Content-Type","text/html")
			code = LoadFromFile(Server.MapPath(".")&"\DEMO.TPL","utf-8")
			If Len(code) > 0 Then Response.Write(YT_TPL_display(array(code,"DEMO")))
	End Select
Call System_Terminate()
'If Err.Number<>0 then
''  Call ShowError(0)
'End If
%>
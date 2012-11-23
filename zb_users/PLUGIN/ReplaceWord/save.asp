<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<!-- #include file="function.asp" -->

<%
Dim XmlDom
ShowError_Custom="Response.Write ""{'success':false,'error':""&id&""}"":Response.End"
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ReplaceWord")=False Then Call ShowError(48)
BlogTitle="敏感词替换器"
replaceword.init()
Select Case Request.QueryString("act")
	Case "delete"
		Set XmlDom=replaceword.words(id)
		replaceword.xmldom.documentElement.removeChild xmlDom
	Case Else
		Dim Frm,id
		For Each Frm In Request.Form
			id=Split(Frm,"_")(1)
			Select Case Left(Frm,3)
				Case "exp"
					replaceword.words(id).attributes.getNamedItem("regexp").value=Request.Form(Frm).Item
				Case "str"
					replaceword.words(id).selectSingleNode("str").text=Request.Form(Frm).Item
				Case "rep"
					replaceword.words(id).selectSingleNode("replace").text=Request.Form(Frm).Item
				Case "des"
					replaceword.words(id).selectSingleNode("description").text=Request.Form(Frm).Item
			End Select
		Next
End Select
replaceword.xmldom.Save(Server.MapPath("config.xml"))
Response.Write "{'success':true}"
'Response.Redirect "main.asp"
%>

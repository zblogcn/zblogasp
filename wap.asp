<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    (zx.asd)&(sipo)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    wap.asp
'// 开始时间:    2006-3-19
'// 最后修改:    
'// 备    注:    WAP模块
'///////////////////////////////////////////////////////////////////////////////

Option Explicit
On Error Resume Next
Response.Charset="UTF-8"
Response.Buffer=True
Response.Expires = "0"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-cache, must-revalidate"
%>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_function_md5.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_wap.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%
'If ZC_IE_DISPLAY_WAP Then
'	If InStr(LCase(Request.ServerVariables("HTTP_ACCEPT")),"text/vnd.wap.wml") > 0 Then Response.ContentType = "text/vnd.wap.wml"
'Else
'	Response.ContentType = "text/vnd.wap.wml"
'End If

Response.ContentType = "text/vnd.wap.wml"

ShowError_Custom="Call ShowError_WAP(id)"

%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head><meta forua="true" http-equiv="Cache-Control" content="max-age=0" /></head>
<%
Call System_Initialize()

'plugin node
For Each sAction_Plugin_Wap_Begin in Action_Plugin_Wap_Begin
	If Not IsEmpty(sAction_Plugin_Wap_Begin) Then Call Execute(sAction_Plugin_Wap_Begin)
Next

PubLic intPageCount
	Select Case ReQuest("act")
		Case "View"
			Call WapView()
		Case "Com"
			Call WapCom()
		Case "Main"
			Call WapMain()
		Case "Login"
			Call WapLogin()
		Case "Err"
			Call WapError()
		Case "Cate"
			Call WapCate()
		Case "Stat"
			Call WapStat()
		Case "AddCom"
			Call WapAddCom(0)
		Case "PostCom"
			Call WapPostCom()
		Case "DelCom"
			Call WapDelCom()
		Case "AddArt"
		    Call WapEdtArt()
		Case "PostArt"
		    Call WapPostArt()
		Case "DelArt"
			Call WapDelArt()
		Case "Logout"
			Call WapLogout()
		Case Else
			Call WapMenu()
	End Select

'plugin node
For Each sAction_Plugin_Wap_End in Action_Plugin_Wap_End
	If Not IsEmpty(sAction_Plugin_Wap_End) Then Call Execute(sAction_Plugin_Wap_End)
Next

Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
<br/>
<a href="<%=ZC_BLOG_HOST&ZC_FILENAME_WAP%>"><%=ZC_MSG213%></a>
</p>
</card>
</wml>
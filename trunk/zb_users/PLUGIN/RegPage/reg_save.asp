<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../p_config.asp" -->
<!-- #include file="include_plugin.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")

If CheckPluginState("RegPage")=False Then Call ShowError(48)
For Each sAction_Plugin_RegSave_Begin in Action_Plugin_RegSave_Begin
	If Not IsEmpty(sAction_Plugin_RegSave_Begin) Then Call Execute(sAction_Plugin_RegSave_Begin)
Next
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<title>注册验证</title>
</head>
<%
Sub ExportErr(str)
	Response.Write "<script type=""text/javascript"">alert('"&replace(str,"'","\'")&"');history.go(-1)</script>"
	Response.End
End Sub
If CheckVerifyNumber(Request.Form("edtCheckOut"))=False Then
	ExportErr ZVA_ErrorMsg(38)
End If

dim Username,UserPassword,UserMail,UserHomePage
Username=TransferHTML(filtersql(request.form("username")),"[nohtml]")
UserPassword=TransferHTML(filtersql(request.form("password")),"[nohtml]")
UserMail=TransferHTML(filtersql(request.form("email")),"[nohtml]")
UserHomePage=TransferHTML(filtersql(request.Form("site")),"[nohtml]")
Dim chkUserName,chkPassWord,chkUserMail,chkHomePage


chkUserName=CheckRegExp(Username,"[username]")
If   chkUserName=False or len(username)>ZC_USERNAME_MAX Then
	ExportErr "用户名格式不正确！\n\n请确认：\n用户名小于"&ZC_USERNAME_MAX&"个字符\n用户名没有非法字符"
End If


chkPassWord=CheckRegExp(UserPassWord,"[password]")
If  len(UserPassWord)<8  or len(UserPassWord)>14 Or chkPassWord=False Then
	ExportErr "密码格式不正确！"
End If
UserPassWord=MD5(UserPassWord)


chkHomePage=CheckRegExp(UserHomePage,"[homepage]")
If  chkHomePage=False Then
	ExportErr "网站格式不正确！"
End If

chkUserMail=CheckRegExp(UserMail,"[email]")
If  chkUserMail=False Then
	ExportErr "电子邮箱格式不正确！"
End If

dim objRs
set objRs = objConn.execute ("SELECT * FROM [blog_Member] where mem_Name= '" & Username & "' ")
if not (objRs.Bof or objRs.eof) then
	ExportErr Username & "已被他人注册，请选用其它用户名！"
End If
objRs.close
set objRs = nothing       

For Each sAction_Plugin_RegSave_VerifyOK in Action_Plugin_RegSave_VerifyOK
	If Not IsEmpty(sAction_Plugin_RegSave_VerifyOK) Then Call Execute(sAction_Plugin_RegSave_VerifyOK)
Next

Dim RegUser
Set RegUser=New TUser
'RegUser.LoadInfoById 18
RegUser.Level=4
RegUser.Name=UserName
RegUser.Email=UserMail
RegUser.HomePage=UserHomePage
RegUser.Password=UserPassword
RegUser.Register
'RegUser.LoadInfoById RegUser.ID
Response.Cookies("password")=RegUser.PassWord
Response.Cookies("password").Expires = DateAdd("d", 1, now)
Response.Cookies("password").Path = "/"
Response.Cookies("username")=RegUser.Name
Response.Cookies("username").Expires = DateAdd("d", 1, now)
Response.Cookies("username").Path = "/"


Dim strResponse
strResponse="<script language='javascript' type='text/javascript'>alert('恭喜，注册成功。\n欢迎您成为本站一员。\n\n单击确定登陆本站。');location.href="""&GetCurrentHost&"""</script>"

For Each sAction_Plugin_RegSave_End in Action_Plugin_RegSave_End

	If Not IsEmpty(sAction_Plugin_RegSave_End) Then Call Execute(sAction_Plugin_RegSave_End)
Next

response.write strResponse


Set RegUser=Nothing
%>
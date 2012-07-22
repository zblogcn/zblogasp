<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="c_option.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_function_md5.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="plugin/p_config.asp" -->
<%
Call System_Initialize()
Call CheckReference("")
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
Username=TransferHTML(filtersql(request.form("username")),"[no-html]")
UserPassword=TransferHTML(filtersql(request.form("password")),"[no-html]")
UserMail=TransferHTML(filtersql(request.form("email")),"[no-html]")
UserHomePage=TransferHTML(filtersql(request.Form("site")),"[no-html]")
Dim chkUserName,chkPassWord,chkUserMail,chkHomePage


chkUserName=CheckRegExp(Username,"[username]")
If  len(username)<4  Or chkUserName=False or len(username)>ZC_USERNAME_MAX Then
	ExportErr "用户名格式不正确！\n\n请确认：\n用户名大于4个字符\n用户名小于"&ZC_USERNAME_MAX&"个字符\n用户名没有非法字符"
End If


chkPassWord=CheckRegExp(UserPassWord,"[password]")
If  len(UserPassWord)<6  or len(UserPassWord)>14 Or chkPassWord=False Then
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
	ExportErr Username & "已被他人注册，请选用其它用户名！'"
End If
objRs.close
set objRs = nothing       

Dim RegUser
Set RegUser=New TUser
RegUser.Level=4
RegUser.Name=UserName
RegUser.Email=UserMail
RegUser.HomePage=UserHomePage
RegUser.Password=UserPassword
RegUser.Register("")
response.write "<script language='javascript' type='text/javascript'>"
response.write "alert('恭喜，注册成功。\n欢迎您成为本站一员。\n\n单击确定登陆本站。');location.href="""&ZC_BLOG_HOST&"/zb_system/cmd.asp?act=login"""
response.write "</script>"
%>
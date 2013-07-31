<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    default.asp
'// 开始时间:    2004.07.25
'// 最后修改:    
'// 备    注:    主页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="zh-CN" />
<meta name="robots" content="none" />
<title>页面转移</title>
</head>
<body>
<%
'兼容性策略
Dim s
s=TransferHTML(Request.QueryString,"[html-format]")

If CheckPluginState("Wap") Then
	Response.Status="301 Moved Permanently"
	s=BlogHost & "?mod=wap&" & s
	Response.AddHeader "Location",s
%>
<p>您好，由于系统升级，您访问的Wap页面将重定向。</p>
<p>如果您没有被重定向成功，请点击以下网址：</p>
<p><a href="<%=s%>"><%=s%></a></p>
<%
Else
	Response.Redirect "default.asp"
%>
<p>您好，本站关闭了WAP访问功能。</p>
<p>您可以<a href="<%=BlogHost%>">点击这里返回首页</a>。</p>
<%
End If

%>

</body>
</html>
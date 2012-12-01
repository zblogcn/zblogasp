<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!-- #include file="function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("AppCentre")=False Then Call ShowError(48)
BlogTitle="应用中心"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-CN" lang="zh-CN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="zh-CN" />
<title>Z-Blog 应用中心</title>
<style type="text/css">
.aaa{color:white;text-decoration:none;font-size:36px;font-family:微软雅黑,宋体}
.aaa2{color:white;text-decoration:none;font-size:30px;font-family:微软雅黑,宋体}

</style>
</head>
<body style="background:url(appcentre.png)">
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="100%" border="0">
  <tr>
    <td style="text-align: center"><span class="aaa" style="text-align:center;">Coming Soon</span></td>
  </tr>
  <tr>
    <td style="text-align: center"><span class="aaa" style="text-align:center;">全新的应用中心即将诞生</span></td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr>
    <td style="text-align: center"><span style="text-align:center;"><a href="javascript:void(0)" class="aaa2">返回</a> &nbsp;&nbsp;<a href="plugin_edit.asp" class="aaa2">创建新插件</a></span></td>
  </tr>
</table>

</body>
</html>
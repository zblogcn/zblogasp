<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Devo Or Newer
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    批量管理文章插件 - 跳转页
'// 最后修改：   2008-10-24
'// 最后版本:    1.4
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

Dim ShowWarning
Dim objConfig
Set objConfig=New TConfig
objConfig.Load "BatchArticles"
ShowWarning=CBool(objConfig.Read("ShowWarning"))
if ShowWarning then
	Response.Redirect "warning.asp"
else
	Response.Redirect "articlelist.asp"
end if

Call System_Terminate()

If Err.Number<>0 then
'  Call ShowError(0)
End If
%>
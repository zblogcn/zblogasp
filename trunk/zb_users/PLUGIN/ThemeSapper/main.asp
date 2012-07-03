<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Pre Terminator 及以上版本, 其它版本的Z-blog未知
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    主题管理插件
'// 最后修改：   2008-6-28
'// 最后版本:    1.2
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../function/c_function.asp" -->
<!-- #include file="../../function/c_system_lib.asp" -->
<!-- #include file="../../function/c_system_base.asp" -->
<!-- #include file="../../function/c_system_plugin.asp" -->
<!-- #include file="c_sapper.asp" -->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then
	Response.Write "<div style=""float:right;height:15px;width:200px;padding:5px 10px;background:#8B0000;color:#FFFFFF;font-size:12px;"">您无权使用此插件, 正在退出...</div>"
	Response.Write "<script>setTimeout(""self.history.back(1)"",2000);</script>"
	Response.End
End If

If CheckPluginState("ThemeSapper")=False Then
	Response.Write "<div style=""float:right;height:15px;width:200px;padding:5px 10px;background:#8B0000;color:#FFFFFF;font-size:12px;"">此插件已停用, 正在退出...</div>"
	Response.Write "<script>setTimeout(""self.history.back(1)"",2000);</script>"
	Response.End
End If

Response.Write "<script>self.location.href=""ThemeList.asp"";</script>"


Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
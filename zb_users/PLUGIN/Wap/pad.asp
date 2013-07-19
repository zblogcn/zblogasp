<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    (zx.asd)&(sipo)&(月上之木)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    wap.asp
'// 开始时间:    2006-3-19
'// 最后修改:    2011-7-23
'// 备    注:    WAP模块
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Response.CacheControl="no-cache" %>
<!-- #include file="../../../zb_users/c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../../zb_users/plugin/wap/c_system_pad.asp" -->
<!-- #include file="../../../zb_users/plugin/p_config.asp" -->
<%
Call System_Initialize()

Wap_Type="pad"

Dim Pad
Set Pad=New TPad

Pad.Run

Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%
Call System_Initialize()
Call Totoro_Initialize
If Request.QueryString("act")="cleanall" And BlogUser.Level=1 Then
	Totoro_Config.Write "TOTORO_THROWCOUNT",0
	Totoro_Config.Write "TOTORO_CHECKCOUNT",0
	Totoro_Config.Save
End If
  %><%If BlogUser.Level=1 Then %>
<p style="-ms-text-size-adjust: 100%; ">Totoro已经为您拦截<%=Totoro_THROWCOUNT%>条评论，加入审核<%=Totoro_CHECKCOUNT%>条评论
<a href="javascript:" onclick="$(document).ready(function(){$.get('<%=GetCurrentHost()%>zb_users/plugin/totoro/getcount.asp',{'rnd':Math.random(),'act':'cleanall'},function(txt){$('#totoro').html(txt)})})">[清空计数器]</a><%End If%>。</p>
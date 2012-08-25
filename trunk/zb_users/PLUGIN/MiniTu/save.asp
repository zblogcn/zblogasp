<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8及以上的版本
'// 插件制作:  zblog管理员之家(www.zbadmin.com)
'// 备    注:   Mini缩略图插件代码
'// 最后修改：   2012/2/20
'// 最后版本:    0.1
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 

Call MiniTu_Initialize
MiniTu_Config.Write "MiniImgWidth",Request.Form("MiniTu_MiniImgWidth")
MiniTu_Config.Write "MiniImgHeight",Request.Form("MiniTu_MiniImgHeight")
MiniTu_Config.Save

Call SetBlogHint(Empty,True,Empty)

Call System_Terminate()


If Err.Number<>0 then
  Call ShowError(0)
End If
%>
<script type="text/javascript">window.location="setting.asp"</script>

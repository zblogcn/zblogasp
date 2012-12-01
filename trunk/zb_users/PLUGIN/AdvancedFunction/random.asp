<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="functions.asp"-->
<%
Response.ContentType="application/x-javascript"
System_Initialize
init()
%>

<script language="javascript" runat="server">
function init(){
	advancedfunction.init();
	Response.Write("$(\"#divRandomArticle ul\").html('"+advancedfunction.随机文章(false).replace(/<#ZC_BLOG_HOST#>/g,BlogHost)+"');")
}
</script>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->

<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize
Call ZBQQConnect_Initialize
'ZBQQConnect_class.callbackurl=ZBQQConnect_class.callbackurl&"?act=login"
Response.Redirect ZBQQConnect_class.Authorize


%>
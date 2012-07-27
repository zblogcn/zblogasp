<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->

<%
Call System_Initialize
Call ZBQQConnect_Class.OpenId(ZBQQConnect_class.CallBack)
Select Case LCase(Request.QueryString("act"))
	Case "login"
		ZBQQConnect_DB.OpenID=Session(strSessionClsID&"ZBQQConnect_strOpenID")
		ZBQQConnect_DB.AccessToken=Session(strSessionClsID&"ZBQQConnect_strAccessToken")
		Response.Redirect GetCurrentHost&"/ZB_SYSTEM/ADMIN/ADMIN.ASP?ACT=SiteInfo"
	Case "reg"
		ZBQQConnect_DB.OpenID=Session(strSessionClsID&"ZBQQConnect_strOpenID")
		ZBQQConnect_DB.AccessToken=Session(strSessionClsID&"ZBQQConnect_strAccessToken")
		Set ZBQQConnect_DB.objUser=BlogUser
		ZBQQConnect_DB.Email=MD5(BlogUser.EMail)
		ZBQQConnect_DB.Bind
	Case Else
		'Response.Write "a"
		'Response.write LCase(Request.QueryString("act"))
	End Select
%>


<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->

<%
Call System_Initialize
Call ZBQQConnect_Initialize()
Call ZBQQConnect_Class.GetOpenId(ZBQQConnect_class.CallBack)
'Select Case LCase(Request.QueryString("act"))
'	Case "login"
'Response.Write BlogUser.ID
'Response.end

ZBQQConnect_DB.OpenID=ZBQQConnect_Class.OpenID
ZBQQConnect_DB.AccessToken=ZBQQConnect_Class.AccessToken

If ZBQQConnect_DB.LoadInfo(4) Then
	If CInt(ZBQQConnect_DB.objUser.ID)<>0 Then
		If ZBQQConnect_DB.Login=True Then
			Response.Redirect GetCurrentHost&"/ZB_SYSTEM/ADMIN/ADMIN.ASP?ACT=SiteInfo"
		Else
			Response.Write ZBQQConnect_DB.objUser.ID
		End If
	Else
		a
	End If
Else
	a
End If
	'Case "admin"
		
'	Case Else
		'Response.Write "a"
'		'Response.write LCase(Request.QueryString("act"))
'	End Select
Function a
		Dim b
		b=ZBQQConnect_class.API("https://graph.qq.com/user/get_info","{'format':'json'}","GET&")
		Set b=ZBQQConnect_toobject(b)
		ZBQQConnect_DB.tHead=b.data.head
		b=ZBQQConnect_class.API("https://graph.qq.com/user/get_user_info","{'format':'json'}","GET&")
		Set b=ZBQQConnect_toobject(b)
		ZBQQConnect_DB.QZoneHead=b.figureurl_2
		Set ZBQQConnect_DB.objUser=BlogUser
		ZBQQConnect_DB.Email=MD5(BlogUser.EMail)
		ZBQQConnect_DB.Bind
		If BlogUser.Level=5 Then
			Response.Redirect "select.asp?QQOPENID="&ZBQQConnect_Class.OpenID
		Else
			Response.write "<script>opener.location.href=opener.location.href.replace(""act=logout"","""");window.close()</script>"
		End If
		Response.Cookies("inpName")=b.nickname
		Response.Cookies("inpName").Expires = DateAdd("d", 365, now)
		Response.Cookies("inpName").Path="/"
End Function
%>

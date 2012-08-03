<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
'
Call System_Initialize()
Call ZBQQConnect_Initialize()

Call CheckReference("")
If CheckPluginState("ZBQQConnect")=False Then Call ShowError(48)
Dim tmpa
Dim get_user_info
dim tmpbl
dim for1,for2,obj1
'判断是否注销
if request.QueryString("act")="qqlogout" then
		Set ZBQQConnect_DB.objUser=BlogUser
		ZBQQConnect_DB.LoadInfo 2
		ZBQQConnect_class.OpenID=ZBQQConnect_DB.OpenID
 		ZBQQConnect_class.logout
		response.Redirect("main.asp")
end if 
if request.QueryString("act")="wblogout" then
		ZBQQConnect_Config.Write "WBToken",""
		ZBQQConnect_Config.Write "WBName",""
		ZBQQConnect_Config.Write "WBSecret",""
		ZBQQConnect_Config.Save
		'response.end
		response.Redirect("main.asp")
end if 

BlogTitle="ZBQQConnect-首页"
%>
    
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader">ZBQQConnect</div>
<div class="SubMenu"><%=ZBQQConnect_SBar(1)%></div>
<div id="divMain2">

      <%
	Dim ZBQQConnect_get_authorize_url
	Set ZBQQConnect_DB.objUser=BlogUser
	Dim ZBQQConnect_A

	If ZBQQConnect_Config.Exists("AppID")=True Then
		If ZBQQConnect_DB.LoadInfo(2)=False Or BlogUser.Level=5 Then
			
			ZBQQConnect_class.callbackurl=IIf(BlogUser.Level=5,GetCurrentHOst&"/ZB_USERS/PLUGIN/ZBQQConnect/callback.asp?act=login",GetCurrentHOst&"/ZB_USERS/PLUGIN/ZBQQConnect/callback.asp?act=admin")
			Response.Write "<a onclick='window.open(""" & ZBQQConnect_class.Authorize & """);$(""#fff"").show();' href='javascript:void(0);'><img src='logo_230_48.png'/></a></div><div id='fff' style='display:none'>如果您无法正常获取到授权码，请<a href='javascript:location.href=""main.asp?""+Math.random()'>点击刷新本页</a></div>"
		Else
			
			ZBQQConnect_class.OpenID=ZBQQConnect_DB.OpenID
			ZBQQConnect_class.AccessToken=ZBQQConnect_DB.AccessToken
			ZBQQConnect_get_authorize_url = "main.asp?act=qqlogout"
			Response.Write "<a href=""" & ZBQQConnect_get_authorize_url & """>解除QQ与该ID的绑定</a>"	
			
			ZBQQConnect_A=ZBQQConnect_class.API("https://graph.qq.com/user/get_user_info","{'format':'json'}","GET&")
			Set ZBQQConnect_A=ZBQQConnect_ToObject(ZBQQConnect_A)
			Response.Write "<br/>空间信息：<br/>姓名"&ZBQQConnect_A.nickname&"<br/>性别"&ZBQQConnect_A.Gender
		End If
	Else
		Response.Write "您还没有配置APPID，无法使用QQ登录功能。<br/>"
	End If
	response.write "<br/><br/>"
	If BlogUser.Level=1 Then
		Dim wbToken,wbSecret,wbName
		If ZBQQConnect_Config.Exists("WBToken")=True And CStr(ZBQQConnect_Config.Read("WBToken"))<>"" Then
			ZBQQConnect_get_authorize_url = "main.asp?act=wblogout"
			Response.Write "<a href=""" & ZBQQConnect_get_authorize_url & """>注销微博</a>"	
			ZBQQConnect_A=ZBQQConnect_class.fakeQQConnect.API("http://open.t.qq.com/api/user/info","{'format':'json'}","GET&")
			Set ZBQQConnect_A=ZBQQConnect_ToObject(ZBQQConnect_A)
			Response.Write "<br/>微博信息：<br/>帐号"&ZBQQConnect_A.data.name&"<br/>性别"&IIf(ZBQQConnect_A.data.sex=2,"女","男")

		Else
			ZBQQConnect_class.fakeQQConnect.callbackurl=GetCurrentHOst&"/ZB_USERS/PLUGIN/ZBQQConnect/callback.asp?act=login&tp=wb"
			Response.Write "<a onclick='window.open(""" & ZBQQConnect_class.fakeQQConnect.Run(1,"","","","") & """);$(""#fff"").show();' href='javascript:void(0);'><img src='wblogin.gif'/></a></div><div id='fff' style='display:none'>如果您无法正常获取到授权码，请<a href='javascript:location.href=""main.asp?""+Math.random()'>点击刷新本页</a></div>"
		End If
	End If
%>


</div>
</div>

<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%

function pdkz(text)
	if text=null or text=empty or text="" then pdkz="空转" else pdkz=text
end function
set ZBQQConnect_class=nothing
%>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="include.asp" -->
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
'
Call System_Initialize()
Call CheckReference("")

Dim tmpa
Dim get_user_info
dim tmpbl
dim for1,for2,obj1
'判断是否注销
if request.QueryString("act")="logout" then
 		ZBQQConnect_class.logout
		response.Redirect("example_index.asp")
end if 
'判断AJAX拉取时用户有无权限，若有则添加codepage
If ZBQQConnect_class.logined=false and request.QueryString("typ")<>"" Then
	response.write "error"
	response.end
else
	session.CodePage=65001
end if

%>
    
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">
<div class="Header">ZBQQConnect</div>
<div class="SubMenu"><%=ZBQQConnect_SBar(1)%></div>
<div id="divMain2">

      <%
	Dim ZBQQConnect_get_authorize_url
	If ZBQQConnect_class.logined=false Then
		Dim ZBQQConnect_A
		
		Response.Write "<div align='center' id='main'><a onclick='window.open(""" & ZBQQConnect_class.Authorize & """);$(""#fff"").show();' href='javascript:void(0);'><img src='logo_230_48.png'/></a></div><div id='fff' style='display:none'>如果您无法正常获取到授权码，请<a href='javascript:location.href=""example_index.asp?""+Math.random()'>点击刷新本页</a></div>"
	Else
		ZBQQConnect_get_authorize_url = "example_index.asp?act=logout"
		Response.Write "<a href=""" & ZBQQConnect_get_authorize_url & """>注销</a>"	
		Set ZBQQConnect_A=ZBQQConnect_ToObject(ZBQQConnect_class.API("https://graph.qq.com/user/get_user_info","{'format':'json'}","GET&"))
		Response.Write "<br/>你好，"&ZBQQConnect_A.nickname&IIf(ZBQQConnect_A.Gender="男","先生","女士")&"<br/><img src="""&ZBQQConnect_A.figureurl_2&"""/>"
	End If

%>


</div>
</div>

<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
'以下为example页面代码，对SDK开发无用
'导航栏生成 
Function ZBQQConnect_SBar(Btype)
	dim b(1,3),i,j,k
	b(1,1)="m-left"
	b(1,2)="example_index.asp"
	b(1,3)="首页"
	
	For i=1 to 1
		if btype=i then
			k=k&"<span class=""" & b(i,1) & " m-now""><a href=""" & b(i,2) & """>" & b(i,3) & "</a></span>"
		else
			k=k&"<span class=""" & b(i,1) & """><a href=""" & b(i,2) & """>" & b(i,3) & "</a></span>"
		end if
	Next
	ZBQQConnect_SBar=k
End Function
'空转判断
function pdkz(text)
	if text=null or text=empty or text="" then pdkz="空转" else pdkz=text
end function
set ZBQQConnect_class=nothing
%>
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!-- #include file="function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("AppCentre")=False Then Call ShowError(48)

BlogTitle="应用中心"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<link rel="stylesheet" href="images/style.css" type="text/css" media="screen" />
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
	<%Call SubMenu(1)%>
  </div>
  <div id="divMain2">
   <script type="text/javascript">ActiveLeftMenu("aAppcentre");$("#leftmenu #nav_appcentre.on span").css("background-image","url('<%=GetCurrentHost%>zb_users/plugin/appcentre/images/web2.png')")</script>
   <%
Response.Write "<div style='box-shadow: 0 0 10px #aaa;margin:5px 20px 40px 5px;height:500px;text-align:center;'><br/><br/><br/><br/><br/><p><img src='"&BlogHost&"zb_system/image/admin/logo.png'/></p><br/><p style='font-size:20px'>应用中心建设中，敬请期待！</p><p>&nbsp;</p><p><a href='http://bbs.rainbowsoft.org/forum-121-1.html' target='_blank' style='font-size:20px'>论坛2.0应用发布区</a></p></div>"
Response.End
Response.Flush

Dim strTemp,strFunc
strFunc="ListTheme"
strTemp="?findapp=1"
Select Case Request.QueryString("act")
	Case "detail"
		strTemp=strTemp&"&page=1&id=" & Request("id")
		strFunc="DetailTheme"
	Case Else
		strTemp=strTemp&"&page=" & IIf(IsEmpty(Request("page")),1,Request("page"))
End Select
strTemp=strTemp & "&count=" & IIf(IsEmpty(Request("count")),10,Request("count"))
strTemp=GetHTTPPage(APPCENTRE_URL & strTemp)
Execute "Response.Write "&strFunc&"(strTemp)"
   %>
  </div>
</div>
<script type="text/javascript">

</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
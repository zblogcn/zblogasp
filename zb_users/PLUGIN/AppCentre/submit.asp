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

Call AppCentre_InitConfig

BlogTitle="应用中心-上传应用到官方网站"

Dim ZipPathDir,ZipPathFile,PackFile,ShortDir,ID,s,t

ID=Request.QueryString("id")

ID=Request.QueryString("id")

If Request.Form.Count=0 Then

	If Request.QueryString("type")="plugin" Then

		PackFile=MD5(ZC_BLOG_CLSID & ID) & ".zba"
		ZipPathDir = BlogPath & "zb_users\plugin\" & ID & "\"
		ZipPathFile = BlogPath & "zb_users\cache\" & PackFile
		ShortDir = ID & "\"


		Call CreatePluginXml(ZipPathFile)
		Call LoadAppFiles(ZipPathDir,ZipPathFile,ShortDir)

	End If

	If Request.QueryString("type")="theme" Then

		PackFile=MD5(ZC_BLOG_CLSID & ID) & ".zba"
		ZipPathDir = BlogPath & "zb_users\theme\" & ID & "\"
		ZipPathFile = BlogPath & "zb_users\cache\" & PackFile
		ShortDir = ID & "\"


		Call CreateThemeXml(ZipPathFile)
		Call LoadAppFiles(ZipPathDir,ZipPathFile,ShortDir)

	End If


	Dim objPing
	Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP")

	objPing.open "GET", APPCENTRE_SUBMITBEFORE_URL & Server.URLEncode(id),False
	objPing.send ""

	If objPing.ReadyState=4 Then
		If objPing.Status=200 Then
			s=objPing.responseText
		End If
	End If

	Set objPing = Nothing

Else

	Dim objXmlHttp
	Set objXmlHttp=Server.CreateObject("MSXML2.ServerXMLHTTP")

	objXmlHttp.Open "POST",APPCENTRE_SUBMIT_URL
	objXmlhttp.SetRequestHeader "Content-Type","application/x-www-form-urlencoded"
	objXmlhttp.SetRequestHeader "Cookie","username="&vbsescape(login_un)&"; password="&vbsescape(login_pw)
	objXmlHttp.Send "file="'zsx帮我加这个zba输出到server

	If objXmlHttp.ReadyState=4 Then
		If objXmlhttp.Status=200 Then
			t=objXmlhttp.ResponseText
		End If
	End If

	Set objXmlHttp = Nothing

End If

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
	<%If ID="" Then Call AppCentre_SubMenu(1) Else Call AppCentre_SubMenu(-1):Response.Write "<a href=""plugin_pack.asp?id="&ID&""" target=""_blank""><span class=""m-right"">导出当前插件</span></a>" End If%>
  </div>
  <div id="divMain2">
<form method="post" action="">
<table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
<tr><th colspan="2" width='28%'>&nbsp;拟提交发布或更新的应用信息</th></tr>

<tr><td><p><b>· 应用ID</b></p></td><td><p>&nbsp;<input id="app_id" name="app_id" style="width:550px;"  type="text" value="<%=id%>" readonly="readonly" /></p></td></tr>
<tr><td><p><b>· 应用文件名</b></p></td><td><p>&nbsp;<input id="app_name" name="app_name" style="width:550px;"  type="text" value="<%=id & ".zba"%>" readonly="readonly" /></p></td></tr>
<tr><td><p><b>· 最后更新日期</b></p></td><td><p>&nbsp;<input id="app_name" name="app_name" style="width:550px;"  type="text" value="<%=app_modified%>" readonly="readonly" /></p></td></tr>


</table>

<%If s<>"" Then%>
<table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
<tr><th colspan="2" width='28%'>&nbsp;应用中心目标应用的相关信息</th></tr>
<tr><td><p><b>· 应用提交用户</b></p></td><td><p>&nbsp;<input id="zblog_app_id" name="zblog_app_id" style="width:550px;"  type="text" value="" readonly="readonly" /></p></td></tr>
<tr><td><p><b>· 最后更新日期</b></p></td><td><p>&nbsp;<input id="zblog_app_name" name="zblog_app_name" style="width:550px;"  type="text" value="" readonly="readonly" /></p></td></tr>
</table>
<script type="text/javascript">
var jsoninfo=eval(<%=s%>);
$("#zblog_app_id").val(jsoninfo.username);
$("#zblog_app_name").val(jsoninfo.lastmodified);
</script>
<%End If%>

<%If t<>"" Then%>
<script type="text/javascript">alert('<%=t%>')</script>
<%End If%>

<p> 提示:金牌开发者和银牌开发者和铜牌开发者只能更新和提交自己的应用,管理员可以更新和提交所有应用.</p>
<p><br/><input type="submit" class="button" value="提交" id="btnPost" onclick='' /></p><p>&nbsp;</p>


</form>
  </div>
</div>
   <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script>


<%
	If login_pw<>"" Then
		Response.Write "<script type='text/javascript'>$('div.SubMenu a[href=\'login.asp\']').hide();$('div.footer_nav p').html('&nbsp;&nbsp;&nbsp;<b>"&login_un&"</b>您好,欢迎来到APP应用中心!').css('visibility','inherit');</script>"
	End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
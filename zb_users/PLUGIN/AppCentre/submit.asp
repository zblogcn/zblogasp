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

	Call Server_Open("submitpre")
	s=strResponse

	If Request.QueryString("type")="plugin" Then Call LoadPluginXmlInfo(ID) Else Call LoadThemeXmlInfo(ID) 

Else

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

	Call Server_Open("submit")
	If left(strResponse,4)="http" Then
		Response.Redirect strResponse
	Else
		Response.Write "<script type=""text/javascript"">alert('" & strResponse & "')</script>"
	End If
	
	Call DelToFile(ZipPathFile)

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
<tr><th colspan="2" width='28%'>&nbsp;拟提交发布或更新的应用信息&nbsp;&nbsp;<a href="<%=Request.QueryString("type")%>_edit.asp?id=<%=Request.QueryString("id")%>">[点击修改]</a></th></tr>

<tr><td><p><b>· 应用ID</b></p></td><td><p>&nbsp;<input id="app_id" name="app_id" style="width:550px;"  type="text" value="<%=id%>" readonly /></p></td></tr>
<tr><td><p><b>· 应用文件名</b></p></td><td><p>&nbsp;<input id="app_name" name="app_name" style="width:550px;"  type="text" value="<%=id & ".zba"%>" readonly /></p></td></tr>
<tr><td><p><b>· 最后更新日期</b></p></td><td><p>&nbsp;<input id="app_name" name="app_name" style="width:550px;"  type="text" value="<%=app_modified%>" readonly /></p></td></tr>


</table>

<%If s<>"" Then%>
<table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
<tr><th colspan="2" width='28%'>&nbsp;“Z-Blog应用中心”目标应用的相关信息</th></tr>
<tr><td><p><b>· 应用发布ID</b></p></td><td><p>&nbsp;<input id="zblog_app_id" name="zblog_app_id" style="width:550px;"  type="text" value="" readonly /></p></td></tr>
<tr><td><p><b>· 应用提交用户</b></p></td><td><p>&nbsp;<input id="zblog_app_user" name="zblog_app_user" style="width:550px;"  type="text" value="" readonly /></p></td></tr>
<tr><td><p><b>· 最后更新日期</b></p></td><td><p>&nbsp;<input id="zblog_app_name" name="zblog_app_name" style="width:550px;"  type="text" value="" readonly /></p></td></tr>
</table>
<script type="text/javascript">
var jsoninfo=eval(<%=s%>);
$("#zblog_app_id").val(jsoninfo.id);
$("#zblog_app_user").val(jsoninfo.author=="null"?"未提交":jsoninfo.author);
$("#zblog_app_name").val(jsoninfo.modified=="null"?"未提交":jsoninfo.modified);
</script>
<%End If%>

<%If t<>"" Then%>
<script type="text/javascript">alert('<%=t%>')</script>
<%End If%>

<p> 提示:金牌开发者、银牌开发者、铜牌开发者和铁牌开发者只能更新和提交自己的应用,白金开发者可以更新和提交所有应用.</p>
<p><br/><input type="submit" class="button" value="提交" id="btnPost" onclick='return confirm("您确定要提交吗？")' /></p><p>&nbsp;</p>


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
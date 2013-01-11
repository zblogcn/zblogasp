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

Dim ID

ID=Request.QueryString("id")

BlogTitle="应用中心-主题编辑"

If Request.Form.Count>0 Then

	If ID="" Then
		Call CreateNewTheme(Request.Form("app_id"))
	End If

	Call SaveThemeXmlInfo(Request.Form("app_id"))

End If


If ID="" Then
	app_pubdate=FormatDateTime(Now,vbShortDate)
	app_modified=FormatDateTime(Now,vbShortDate)
	app_author_name=BlogUser.FirstName
	app_author_email=BlogUser.EMail
	app_author_url=BlogUser.HomePage
	app_price=0
	app_version="1.0"
Else

	Call LoadThemeXmlInfo(ID)
	If app_price="" Then app_price=0
	app_modified=AppCentre_GetLastModifiTime(app_path)

End If 


%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
	<%If ID="" Then Call AppCentre_SubMenu(2) Else Call AppCentre_SubMenu(-1):Response.Write "<a href=""theme_pack.asp?id="&ID&""" target=""_blank""><span class=""m-right"">导出当前主题</span></a>" End If%>
  </div>
  <div id="divMain2">
<form method="post" action="">
<table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
<tr><th width='28%'>&nbsp;</th><th>&nbsp;</th></tr>

<tr><td><p><b>· 主题ID</b><br/><span class="note">&nbsp;&nbsp;主题ID为主题的目录名,且不能重复.</span></p></td>
<td><p>&nbsp;<input id="app_id" name="app_id" style="width:550px;"  type="text" value="<%=app_id%>" <%=IIF(ID="","","readonly=""readonly""")%> /></p></td></tr>
<tr><td><p><b>· 主题名称</b></p></td><td><p>&nbsp;<input id="app_name" name="app_name" style="width:550px;"  type="text" value="<%=app_name%>" /></p></td></tr>
<tr><td><p><b>· 主题发布页面</b></p></td><td><p>&nbsp;<input id="app_url" name="app_url" style="width:550px;"  type="text" value="<%=app_url%>" /></p></td></tr>
<tr><td><p><b>· 主题简介</b></p></td><td><p>&nbsp;<input id="app_note" name="app_note" style="width:550px;"  type="text" value="<%=app_note%>" /></p></td></tr>
<tr><td><p><b>· 适用的最低要求 Z-Blog 版本</b></p></td><td>
<p>&nbsp;<select name="app_adapted" id="app_adapted" style="width:400px;">
    <option value="121221" selected="selected">Z-Blog 2.0 Doomsday Build 121221</option>
    <option value="121028">Z-Blog 2.0 Beta2 Build 121028</option>
    <option value="121001">Z-Blog 2.0 Beta1 Build 121001</option>
  </select></p>
</td></tr>
<tr><td><p><b>· 主题版本号</b></p></td><td><p>&nbsp;<input id="app_version" name="app_version" style="width:550px;" type="number" step="0.1" value="<%=app_version%>" /></p></td></tr>
<tr><td><p><b>· 主题首发时间</b><br/><span class="note">&nbsp;&nbsp;日期格式为2012-12-12</span></p></td><td><p>&nbsp;<input id="app_pubdate" name="app_pubdate" style="width:550px;"  type="text" value="<%=app_pubdate%>" /></p></td></tr>
<tr><td><p><b>· 主题最后修改时间</b><br/><span class="note">&nbsp;&nbsp;系统自动检查目录内文件的最后修改日期</span></p></td><td><p>&nbsp;<input id="app_modified" name="app_modified" style="width:550px;"  type="text" value="<%=app_modified%>" readonly="readonly" /></p></td></tr>

<tr><td><p><b>· 作者名称</b></p></td><td><p>&nbsp;<input id="app_author_name" name="app_author_name" style="width:550px;"  type="text" value="<%=app_author_name%>" /></p></td></tr>
<tr><td><p><b>· 作者邮箱</b></p></td><td><p>&nbsp;<input id="app_author_email" name="app_author_email" style="width:550px;"  type="text" value="<%=app_author_email%>" /></p></td></tr>
<tr><td><p><b>· 作者网站</b></p></td><td><p>&nbsp;<input id="app_author_url" name="app_author_url" style="width:550px;"  type="text" value="<%=app_author_url%>" /></p></td></tr>
<tr><td><p><b>· 源作者名称</b> (可选)</p></td><td><p>&nbsp;<input id="app_source_name" name="app_source_name" style="width:550px;"  type="text" value="<%=app_source_name%>" /></p></td></tr>
<tr><td><p><b>· 源作者邮箱</b> (可选)</p></td><td><p>&nbsp;<input id="app_source_email" name="app_source_email" style="width:550px;"  type="text" value="<%=app_source_email%>" /></p></td></tr>
<tr><td><p><b>· 源作者网站</b> (可选)</p></td><td><p>&nbsp;<input id="app_source_url" name="app_source_url" style="width:550px;"  type="text" value="<%=app_source_url%>" /></p></td></tr>
<tr><td><p><b>· 依赖插件（以|分隔）</b>(可选)</p></td><td><p>&nbsp;<input id="app_dependency" name="app_dependency" style="width:550px;"  type="text" value="<%=app_dependency%>" /></p></td></tr>


<tr><td><p><b>· 内置插件管理页</b> (可选)<br/><span class="note">&nbsp;&nbsp;默认为main.asp</span></p></td><td><p>&nbsp;<input id="app_plugin_path" name="app_plugin_path" style="width:550px;"  type="text" value="<%=app_plugin_path%>" /></p></td></tr>
<tr><td><p><b>· 内置插件嵌入页</b> (可选)<br/><span class="note">&nbsp;&nbsp;默认为include.asp</span></p></td><td><p>&nbsp;<input id="app_plugin_include" name="app_plugin_include" style="width:550px;"  type="text" value="<%=app_plugin_include%>" /></p></td></tr>
<tr><td><p><b>· 内置插件管理权限</b> (可选)</p></td>
<td>
<p>&nbsp;<select name="app_plugin_level" id="app_plugin_level" style="width:200px;">
    <option value="1" <%=IIF(app_plugin_level="1","selected='selected'","")%>><%=ZVA_User_Level_Name(1)%></option>
    <option value="2" <%=IIF(app_plugin_level="2","selected='selected'","")%>><%=ZVA_User_Level_Name(2)%></option>
    <option value="3" <%=IIF(app_plugin_level="3","selected='selected'","")%>><%=ZVA_User_Level_Name(3)%></option>
    <option value="4" <%=IIF(app_plugin_level="4","selected='selected'","")%>><%=ZVA_User_Level_Name(4)%></option>
    <option value="5" <%=IIF(app_plugin_level="5","selected='selected'","")%>><%=ZVA_User_Level_Name(5)%></option>
  </select></p>
</td></tr>

<tr><td><p><b>· 【高级】内置插件重写系统函数列表（以|分隔）</b>(可选)</p></td><td><p>&nbsp;<input id="app_rewritefunctions" name="app_rewritefunctions" style="width:550px;"  type="text" value="<%=app_rewritefunctions%>" /></p></td></tr>
<tr><td><p><b>· 【高级】内置插件冲突插件列表（以|分隔）</b>(可选)</p></td><td><p>&nbsp;<input id="app_conflict" name="app_conflict" style="width:550px;"  type="text" value="<%=app_conflict%>" /></p></td></tr>
<tr><td><p><b>· 主题定价</b></p></td><td><p>&nbsp;<input id="app_price" name="app_price" style="width:550px;"  type="text" value="<%=app_price%>" /></p></td></tr>
<tr><td><p><b>· 详细说明</b> (可选)</p></td><td><p>&nbsp;<textarea cols="3" rows="6" id="app_description" name="app_description" style="width:550px;"><%=TransferHTML(app_description,"[html-format]")%></textarea></p></td></tr>


</table>
<p> 提示:主题的缩略图是名为ScreenShot.png的<b>300x240px</b>大小的png文件,放在插件的目录下.</p>
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
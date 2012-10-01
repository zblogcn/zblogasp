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

Dim ID

ID=Request.QueryString("id")

BlogTitle="应用中心-插件编辑"


If Request.Form.Count>0 Then

	If ID="" Then
		Call CreateNewPlugin(Request.Form("app_id"))
	End If

	Call SavePluginXmlInfo(Request.Form("app_id"))

End If


If ID="" Then

Else

	Call LoadPluginXmlInfo(ID)

End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
	<%If ID="" Then Call SubMenu(2) Else Call SubMenu(-1) End If%>
  </div>
  <div id="divMain2">
<form method="post" action="">
<table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
<tr><th width='28%'>&nbsp;</th><th>&nbsp;</th></tr>

<tr><td><p><b>· 插件ID</b><br/><span class="note">&nbsp;&nbsp;插件ID为插件的目录名,且不能重复.</span></p></td>
<td><p>&nbsp;<input id="app_id" name="app_id" style="width:550px;"  type="text" value="<%=app_id%>" <%=IIF(ID="","","readonly=""readonly""")%> /></p></td></tr>
<tr><td><p><b>· 插件名称</b></p></td><td><p>&nbsp;<input id="app_name" name="app_name" style="width:550px;"  type="text" value="<%=app_name%>" /></p></td></tr>
<tr><td><p><b>· 插件发布页面</b></p></td><td><p>&nbsp;<input id="app_url" name="app_url" style="width:550px;"  type="text" value="<%=app_url%>" /></p></td></tr>
<tr><td><p><b>· 插件简介</b></p></td><td><p>&nbsp;<input id="app_note" name="app_note" style="width:550px;"  type="text" value="<%=app_note%>" /></p></td></tr>
<tr><td><p><b>· 适用的 Z-Blog 版本</b></p></td><td>
<p>&nbsp;<select name="app_adapted" id="app_adapted" style="width:200px;">
    <option value="Z-Blog 2.0">Z-Blog 2.0</option>
  </select></p>
</td></tr>
<tr><td><p><b>· 插件版本号</b></p></td><td><p>&nbsp;<input id="app_version" name="app_version" style="width:550px;"  type="text" value="<%=app_version%>" /></p></td></tr>
<tr><td><p><b>· 插件首发时间</b><br/><span class="note">&nbsp;&nbsp;日期格式为2012-12-12</span></p></td><td><p>&nbsp;<input id="app_pubdate" name="app_pubdate" style="width:550px;"  type="text" value="<%=app_pubdate%>" /></p></td></tr>
<tr><td><p><b>· 插件最后修改时间</b></p></td><td><p>&nbsp;<input id="app_modified" name="app_modified" style="width:550px;"  type="text" value="<%=app_modified%>" /></p></td></tr>

<tr><td><p><b>· 作者名称</b></p></td><td><p>&nbsp;<input id="app_author_name" name="app_author_name" style="width:550px;"  type="text" value="<%=app_author_name%>" /></p></td></tr>
<tr><td><p><b>· 作者邮箱</b></p></td><td><p>&nbsp;<input id="app_author_email" name="app_author_email" style="width:550px;"  type="text" value="<%=app_author_email%>" /></p></td></tr>
<tr><td><p><b>· 作者网站</b></p></td><td><p>&nbsp;<input id="app_author_url" name="app_author_url" style="width:550px;"  type="text" value="<%=app_author_url%>" /></p></td></tr>

<tr><td><p><b>· 插件管理页</b> <br/><span class="note">&nbsp;&nbsp;默认为main.asp</span></p></td><td><p>&nbsp;<input id="app_plugin_path" name="app_plugin_path" style="width:550px;"  type="text" value="<%=app_plugin_path%>" /></p></td></tr>
<tr><td><p><b>· 插件嵌入页</b> <br/><span class="note">&nbsp;&nbsp;默认为include.asp</span></p></td><td><p>&nbsp;<input id="app_plugin_include" name="app_plugin_include" style="width:550px;"  type="text" value="<%=app_plugin_include%>" /></p></td></tr>
<tr><td><p><b>· 插件管理权限</b> </p></td><td>
<p>&nbsp;<select name="app_plugin_level" id="app_plugin_level" style="width:200px;">
    <option value="1" <%=IIF(app_plugin_level="1","selected='selected'","")%>><%=ZVA_User_Level_Name(1)%></option>
    <option value="2" <%=IIF(app_plugin_level="2","selected='selected'","")%>><%=ZVA_User_Level_Name(2)%></option>
    <option value="3" <%=IIF(app_plugin_level="3","selected='selected'","")%>><%=ZVA_User_Level_Name(3)%></option>
    <option value="4" <%=IIF(app_plugin_level="4","selected='selected'","")%>><%=ZVA_User_Level_Name(4)%></option>
    <option value="5" <%=IIF(app_plugin_level="5","selected='selected'","")%>><%=ZVA_User_Level_Name(5)%></option>
  </select></p>
</td></tr>
<tr><td><p><b>· 插件定价</b></p></td><td><p>&nbsp;<input id="app_price" name="app_price" style="width:550px;"  type="text" value="<%=app_price%>" /></p></td></tr>
<tr><td><p><b>· 详细说明</b> (可选)</p></td><td><p>&nbsp;<textarea cols="3" rows="6" id="app_description" name="app_description" style="width:550px;"><%=TransferHTML(app_description,"[html-format]")%></textarea></p></td></tr>


</table>

<p><br/><input type="submit" class="button" value="提交" id="btnPost" onclick='' /></p><p>&nbsp;</p>


</form>
  </div>
</div>
   <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
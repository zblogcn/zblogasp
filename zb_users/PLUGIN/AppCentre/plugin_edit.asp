<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
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

Dim app_id
Dim app_name
Dim app_url
Dim app_note

Dim app_author_name
Dim app_author_email
Dim app_author_url

Dim app_source_name
Dim app_source_email
Dim app_source_url

Dim app_plugin_path
Dim app_plugin_include
Dim app_plugin_level

Dim app_adapted
Dim app_version
Dim app_pubdate
Dim app_modified
Dim app_description
Dim app_price

If ID="" Then

Else

	Call LoadThemeXmlInfo(ID)

End If 


%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
	<%Call SubMenu(2)%>
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
    <option value="1"><%=ZVA_User_Level_Name(1)%></option>
    <option value="2"><%=ZVA_User_Level_Name(2)%></option>
    <option value="3"><%=ZVA_User_Level_Name(3)%></option>
    <option value="4"><%=ZVA_User_Level_Name(4)%></option>
    <option value="5"><%=ZVA_User_Level_Name(5)%></option>
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
<%
Function LoadThemeXmlInfo(id)

	On Error Resume Next

	Dim objXmlFile,strXmlFile
	Dim fso, s
	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FileExists(BlogPath & "zb_users/plugin" & "/" & id & "/" & "plugin.xml") Then

		strXmlFile =BlogPath & "zb_users/plugin" & "/" & id & "/" & "plugin.xml"

		Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		objXmlFile.load(strXmlFile)
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode <> 0 Then
			Else

				app_id=id
				app_name=objXmlFile.documentElement.selectSingleNode("name").text
				app_url=objXmlFile.documentElement.selectSingleNode("url").text

				app_adapted=objXmlFile.documentElement.selectSingleNode("adapted").text
				app_version=objXmlFile.documentElement.selectSingleNode("version").text
				app_pubdate=objXmlFile.documentElement.selectSingleNode("pubdate").text
				app_modified=objXmlFile.documentElement.selectSingleNode("modified").text

				app_note=objXmlFile.documentElement.selectSingleNode("note").text
				app_description=objXmlFile.documentElement.selectSingleNode("description").text

				app_author_name=objXmlFile.documentElement.selectSingleNode("author/name").text
				app_author_email=objXmlFile.documentElement.selectSingleNode("author/email").text
				app_author_url=objXmlFile.documentElement.selectSingleNode("author/url").text

				'app_source_name=objXmlFile.documentElement.selectSingleNode("source/name").text
				'app_source_email=objXmlFile.documentElement.selectSingleNode("source/email").text
				'app_source_url=objXmlFile.documentElement.selectSingleNode("source/url").text


				app_plugin_path=objXmlFile.documentElement.selectSingleNode("path").text
				app_plugin_include=objXmlFile.documentElement.selectSingleNode("include").text
				app_plugin_level=objXmlFile.documentElement.selectSingleNode("level").text

				app_price=objXmlFile.documentElement.selectSingleNode("app_price").text

			End If
		End If
	End If

End Function
%>
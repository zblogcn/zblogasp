<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../../c_option.asp" -->
<!-- #include file="../../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../../plugin/p_config.asp" -->

<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("X2013")=False Then Call ShowError(48)
BlogTitle="X2013主题设置"

Dim objConfig
Set objConfig=New TConfig
objConfig.Load("X2013")
If objConfig.Exists("Version")=False Then
	objConfig.Write "Version","1.0"
	objConfig.Write "SetWeiboSina","http://weibo.com/810888188"
	objConfig.Write "SetWeiboQQ","http://t.qq.com/involvements"
	objConfig.Write "DisplayFeed","True"
	objConfig.Write "SetMailKey","4e54e0008863773ff0f44e54eb9c1805cf165e63a0601789"
	objConfig.Write "PostAdHeader",""
	objConfig.Write "PostAdFooter",""
	objConfig.Save
End If

Dim strAct,PostAdHeader,PostAdFooter
strAct=Request.QueryString("act")
If strAct="Save" Then
	PostAdHeader=Request.Form("PostAdHeader")
	PostAdFooter=Request.Form("PostAdFooter")
	
	objConfig.Write "PostAdHeader",PostAdHeader
	objConfig.Write "PostAdFooter",PostAdFooter
	objConfig.Save
	
	Call SetBlogHint(True,Empty,True)
End If

%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<style>
input.text{background:#FFF;border:1px double #aaa;font-size:1em;padding:0.25em;}
p{line-height:1.5em;padding:0.5em 0;}
</style>
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain">
	<div id="ShowBlogHint"><%Call GetBlogHint()%></div>
	<div class="divHeader"><%=BlogTitle%></div>
  	<div class="SubMenu"><%=X2013_SubMenu(3)%></div>
	<div id="divMain2">
	<script type="text/javascript">ActiveTopMenu("aX2013");</script> 
	<!--SetCon Star.-->
	<form id="form1" name="form1" method="post">
	
    <table width="100%" style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
  <tr>
    <th width='20%'><p align="center">广告位</p></th>
    <th width='70%'><p align="center">内容</p></th>
  </tr>
  <tr>
    <td  rowspan="2" colspan="1"><b><label for="PostAdHeader"><p align="center">文章开始广告位(建议宽度880px)</p></label></b></td>
    <td><p align="left"><textarea name="PostAdHeader" type="text" id="PostAdHeader" style="width: 80%;"><%=objConfig.Read("PostAdHeader")%></textarea></p></td>
    
  </tr>
   <tr><td><%=objConfig.Read("PostAdHeader")%></td></tr>
  <tr>
    <td rowspan="2" colspan="1"><b><label for="PostAdFooter"><p align="center">文章结束广告位(建议宽度880px)</p></label></b></td>
    <td><p align="left"><textarea name="PostAdFooter" type="text" id="PostAdFooter" style="width: 80%;"><%=objConfig.Read("PostAdFooter")%></textarea></p></td>
  </tr>
   <tr><td><%=objConfig.Read("PostAdFooter")%></tr>
</table>
 <br />
   <input name="" type="Submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=Save";'/>

    </form>
    <!--SetCon End.-->
<br />

	</div>
</div>
<!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
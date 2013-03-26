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

Dim strAct
strAct=Request.QueryString("act")
If strAct="Save" Then
	ZC_MSG266=""
	Dim SetWeiboSina
	SetWeiboSina=Request.Form("SetWeiboSina")
	If SetWeiboSina<>"" Then
		If SetWeiboSina<>objConfig.Read("SetWeiboSina") Then
			objConfig.Write "SetWeiboSina",SetWeiboSina
			objConfig.Save
			ZC_MSG266 = "新浪微博地址设置成功；"
		Else
			ZC_MSG266 = "新浪微博地址未更改；"
		End If
	Else
		objConfig.Write "SetWeiboSina",SetWeiboSina
		objConfig.Save
		Call SetBlogHint_Custom("新浪微博地址为空，前台将不显示此图标.")
	End If
	
	Dim SetWeiboQQ
	SetWeiboQQ=Request.Form("SetWeiboQQ")
	If SetWeiboQQ<>"" Then
		If SetWeiboQQ<>objConfig.Read("SetWeiboQQ") Then
			objConfig.Write "SetWeiboQQ",SetWeiboQQ
			objConfig.Save
			ZC_MSG266 = ZC_MSG266 + "腾讯微博地址设置成功；"
		Else
			ZC_MSG266 = ZC_MSG266 + "腾讯微博地址未更改；"
		End If
	Else
		objConfig.Write "SetWeiboQQ",SetWeiboQQ
		objConfig.Save
		Call SetBlogHint_Custom("腾讯微博地址为空，前台将不显示此图标.")
	End If
	
	Dim DisplayFeed,SetMailKey
	DisplayFeed = Request.Form("DisplayFeed")
	SetMailKey  = Request.Form("SetMailKey")
	objConfig.Write "DisplayFeed",DisplayFeed
	objConfig.Write "SetMailKey",SetMailKey
	objConfig.Save
	
	Call SetBlogHint(True,Empty,True)
	'ZC_MSG266=" 操作成功."
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
  	<div class="SubMenu"><%=X2013_SubMenu(0)%></div>
	<div id="divMain2">
	<script type="text/javascript">ActiveTopMenu("aX2013");</script> 
	<!--SetCon Star.-->
	<form id="form1" name="form1" method="post">
	
    <table width="100%" style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
  <tr>
    <th width='20%'><p align="center">设置</p></th>
    <th width='70%'><p align="center">内容</p></th>
    
  </tr>
  <tr>
    <td><b><label for="SetWeiboSina"><p align="center">新浪微博</p></label></b></td>
    <td><p align="left"><input name="SetWeiboSina" type="text" id="SetWeiboSina" size="100%" value="<%=objConfig.Read("SetWeiboSina")%>" /></p></td>
    
  </tr>
  <tr>
    <td><b><label for="SetWeiboQQ"><p align="center">腾讯微博</p></label></b></td>
    <td><p align="left"><input name="SetWeiboQQ" type="text" id="SetWeiboQQ" size="100%" value="<%=objConfig.Read("SetWeiboQQ")%>" /></p></td>
  </tr>
  <tr>
    <td><b><label for="DisplayFeed"><p align="center">是否显示邮件订阅</p></label></b></td>
    <td><p align="left"><input id="DisplayFeed" name="DisplayFeed" style="display: none; " type="text" value="<%=CBool(objConfig.Read("DisplayFeed"))%>" class="checkbox"></p></td>
  </tr>
  <tr>
    <td><b><label for="SetMailKey"><p align="center">QQMail邮件订阅key</p></label></b></td>
    <td><p align="left"><input name="SetMailKey" type="text" id="SetMailKey" size="100%" value="<%=objConfig.Read("SetMailKey")%>" /></p></td>
  </tr>  
</table>
 <br />
   <input name="" type="Submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=Save";'/>
  
    </form>
    <!--SetCon End.-->
<br />

	</div>
</div>
<!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
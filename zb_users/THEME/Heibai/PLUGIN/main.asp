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
<!-- #include file="Function.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("Heibai")=False Then Call ShowError(48)
BlogTitle="Heibai主题设置"

Dim objConfig
Set objConfig=New TConfig
objConfig.Load("Heibai")
If objConfig.Exists("Version")=False Then
	objConfig.Write "Version","0.1"
	objConfig.Write "SetNewArt","10"
	objConfig.Write "SetCommArt","10"
	objConfig.Write "SetRandomArt","10"
	objConfig.Write "SetNewComm","10"
	objConfig.Write "SetHotCommer","10"
	objConfig.Write "SetTags","30"
	objConfig.Write "SetWeiboSina","http://weibo.com/810888188"
	objConfig.Write "SetWeiboQQ","http://t.qq.com/involvements"
	objConfig.Save
End If

ZC_MSG266=""
Select Case Request.QueryString("act")
	Case "SetNewArt"
	objConfig.Write "SetNewArt",Request.Form("SetNewArt")
	objConfig.Save
	ZC_MSG266 = "最新文章"
	
	Case "SetCommArt"
	objConfig.Write "SetCommArt",Request.Form("SetCommArt")
	objConfig.Save
	ZC_MSG266 = "热评文章"

	Case "SetRandomArt"
	objConfig.Write "SetRandomArt",Request.Form("SetRandomArt")
	objConfig.Save
	ZC_MSG266 = "随机文章"
	
	Case "SetNewComm"
	objConfig.Write "SetNewComm",Request.Form("SetNewComm")
	objConfig.Save
	ZC_MSG266 = "最新评论"
	

	Case "SetHotCommer"
	objConfig.Write "SetHotCommer",Request.Form("SetHotCommer")
	objConfig.Save
	ZC_MSG266 = "读者墙"
	
	Case "SetTags"
	objConfig.Write "SetTags",Request.Form("SetTags")
	objConfig.Save
	ZC_MSG266 = "标签列表"
End Select

If Request.QueryString("act")<>"" Then
	ZC_MSG266 = "<spam style='color:#ff0000'>"+ZC_MSG266 + "</spam>设置成功"
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
  	<div class="SubMenu"><a href="main.asp"><span class="m-left m-now">主题显示调用数量设置</span></a><a href="weibo.asp"><span class="m-left">作者微博设置</span></a></div>
	<div id="divMain2">
	<script type="text/javascript">ActiveTopMenu("aHeibai");</script> 
	<!--SetCon Star.-->
	<form id="form1" name="form1" method="post">
	
    <table width="100%" style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
  <tr>
    <th width='20%'><p align="center">排序方式</p></th>
    <th width='20%'><p align="center">设置调用条数</p></th>
    <th width='60%'>&nbsp;</th>
  </tr>
  <tr>
    <td><b><label for="SetNewArt">最新文章</label></b></td>
    <td><p align="center"><input name="SetNewArt" type="text" id="SetNewArt" size="10" value="<%=objConfig.Read("SetNewArt")%>" /></p></td>
    <td><input name="" type="submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=SetNewArt";'/></td>
  </tr>
  <tr>
    <td><b><label for="SetCommArt">热评文章</label></b></td>
    <td><p align="center"><input name="SetCommArt" type="text" id="SetCommArt" size="10" value="<%=objConfig.Read("SetCommArt")%>" /></p></td>
    <td><input name="" type="submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=SetCommArt";'/></td>
  </tr>
  <tr>
    <td><b><label for="SetRandomArt">随机文章</label></b></td>
    <td><p align="center"><input name="SetRandomArt" type="text" id="SetRandomArt" size="10" value="<%=objConfig.Read("SetRandomArt")%>" /></p></td>
    <td><input name="" type="submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=SetRandomArt";'/></td>
  </tr>
  <tr>
    <td><b><label for="SetNewComm">最新评论</label></b></td>
    <td><p align="center"><input name="SetNewComm" type="text" id="SetNewComm" size="10" value="<%=objConfig.Read("SetNewComm")%>" /></p></td>
    <td><input name="" type="submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=SetNewComm";'/></td>
  </tr>
   <tr>
    <td><b><label for="SetHotCommer">读者墙</label></b></td>
    <td><p align="center"><input name="SetHotCommer" type="text" id="SetHotCommer" size="10" value="<%=objConfig.Read("SetHotCommer")%>" /></p></td>
    <td><input name="" type="submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=SetHotCommer";'/></td>
  </tr>
   <tr>
    <td><b><label for="SetTags">标签列表</label></b></td>
    <td><p align="center"><input name="SetTags" type="text" id="SetTags" size="10" value="<%=objConfig.Read("SetTags")%>" /></p></td>
    <td><input name="" type="submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=SetTags";'/></td>
  </tr>
</table>
  </form>
<br />

	</div>
</div>
<!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->

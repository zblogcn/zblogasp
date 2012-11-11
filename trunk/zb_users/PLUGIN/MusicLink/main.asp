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
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("MusicLink")=False Then Call ShowError(48)
BlogTitle="MusicLink音乐外链设置"

Dim objConfig
Set objConfig=New TConfig
objConfig.Load("MusicLink")
If objConfig.Exists("Version")=False Then
	objConfig.Write "Version","0.1"
	objConfig.Write "AutoPlay","True"
	objConfig.Write "Player","baidu"
	objConfig.Save
End If

Dim strAct
strAct=Request.QueryString("act")
If strAct="Save" Then
	
	Dim AutoPlay
	AutoPlay=Request.Form("AutoPlay")
	If AutoPlay<>"" Then
		objConfig.Write "AutoPlay",AutoPlay
		objConfig.Save
	End If
	
	Dim Player
	Player=Request.Form("Player")
	If Player<>"" Then
		objConfig.Write "Player",Player
		objConfig.Save
	End If
	
	Call SetBlogHint(True,Empty,Empty)
	
End If

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style>
input.text{background:#FFF;border:1px double #aaa;font-size:1em;padding:0.25em;}
p{line-height:1.5em;padding:0.5em 0;}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain">
	<div id="ShowBlogHint"><%Call GetBlogHint()%></div>
	<div class="divHeader"><%=BlogTitle%></div>
  	<div class="SubMenu"><a href="main.asp"><span class="m-left m-now">插件设置</span></a></div>
	<div id="divMain2">

	<!--SetCon Star.-->
	<form id="form1" name="form1" method="post">
	
    <table width="100%" style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
  <tr>
    <th width='20%'><p align="center">选项</p></th>
    <th width='70%'><p align="left">内容</p></th>
    
  </tr>
  <tr>
    <td><b><label for="AutoPlay"><p align="center">是否自动播放</p></label></b></td>
    <td><p align="left"><input id="AutoPlay" name="AutoPlay" style="display: none; " type="text" value="<%=objConfig.Read("AutoPlay")%>" class="checkbox"></p></td>
    
  </tr>
  <tr>
    <td><b><label for="Player"><p align="center">默认播放器</p></label></b></td>
    <td><p align="left"><select name="Player" id="Player" style="width:200px;">
		<%
			Dim strContent
			If objConfig.Read("Player")="baidu" Then
				strContent = "<option value='baidu' selected='selected'>百度音乐播放器</option>"
				strContent = strContent & "<option value='yige'>亦歌音乐播放器</option>"
			ElseIf objConfig.Read("Player")="yige" Then
				strContent = "<option value='baidu'>百度音乐播放器</option>"
				strContent = strContent & "<option value='yige' selected='selected'>亦歌音乐播放器</option>"
			End If
			Response.Write strContent
		%>
	</select></p></td>
    
  </tr>
</table>
 <br />
   <input name="" type="submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=Save";'/>
  
    </form>
    <!--SetCon End.-->
<br />

	</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
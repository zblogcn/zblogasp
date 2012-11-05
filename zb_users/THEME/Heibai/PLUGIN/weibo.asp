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

Dim strAct
strAct=Request.QueryString("act")
If strAct="SaveWeibo" Then
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
  	<div class="SubMenu"><a href="main.asp"><span class="m-left">主题显示调用数量设置</span></a><a href="weibo.asp"><span class="m-left m-now">作者微博设置</span></a></div>
	<div id="divMain2">
	<script type="text/javascript">ActiveTopMenu("aHeibai");</script> 
	<!--SetCon Star.-->
	<form id="form1" name="form1" method="post">
	
    <table width="100%" style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
  <tr>
    <th width='20%'><p align="center">微博</p></th>
    <th width='70%'><p align="center">微博地址</p></th>
    
  </tr>
  <tr>
    <td><b><label for="SetWeiboSina"><p align="center">新浪微博</p></label></b></td>
    <td><p align="left"><input name="SetWeiboSina" type="text" id="SetWeiboSina" size="100%" value="<%=objConfig.Read("SetWeiboSina")%>" /></p></td>
    
  </tr>
  <tr>
    <td><b><label for="SetWeiboQQ"><p align="center">腾讯微博</p></label></b></td>
    <td><p align="left"><input name="SetWeiboQQ" type="text" id="SetWeiboQQ" size="100%" value="<%=objConfig.Read("SetWeiboQQ")%>" /></p></td>
    
  </tr>
</table>
 <br />
   <input name="" type="submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=SaveWeibo";'/>
  
    </form>
    <!--SetCon End.-->
<br />

	</div>
</div>
<!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->

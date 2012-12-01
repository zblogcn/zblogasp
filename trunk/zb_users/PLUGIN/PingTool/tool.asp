<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog(http://www.rainbowsoft.org)
'// 插件制作:    zx.asd
'// 备    注:    Ping中心通知程序
'// 最后修改：   2005-12-8
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_manage.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 
If CheckPluginState("PingTool")=False Then Call ShowError(48)
BlogTitle="Ping中心和引用通告发送器"
Dim objConfig
Set objConfig=New TConfig
objConfig.Load "PingTool"
If Request.QueryString("ok")="1" Then
	objConfig.Write "Content",Request.Form("txaContent")
	objConfig.Save
	Call SetBlogHint(True,Empty,Empty)
End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain">
<div id="ShowBlogHint">
<%
Call GetBlogHint()
%>
</div>
<div class="divHeader">Ping中心和引用通告发送器</div>
<div id="divMain2">
<form border="1" name="edit" id="edit" method="post" action="tool.asp?ok=1">
<p><b>设置Ping中心</b></p>
<p>
<textarea style="height:300px;width:100%" name="txaContent" id="txaContent"><%=TransferHTML(objConfig.Read("Content"),"[textarea]")%></textarea>
</p>
<p>&nbsp;</p>
<p>
<input type="submit" class="button" value="提交" name="B1"/>
</p>
几个常用地址：<div id="ccdz">
<p>谷歌Ping地址：<a href="http://blogsearch.google.com/ping/RPC2" onclick="return false;">http://blogsearch.google.com/ping/RPC2</a> </p>
<p>百度Ping地址：<a href="http://ping.baidu.com/ping/RPC2" onclick="return false;">http://ping.baidu.com/ping/RPC2</a> </p>
<p>PingOmatic:<a href="http://rpc.pingomatic.com/" onclick="return false">http://rpc.pingomatic.com/</a></p>
</div>
</form>
</div>
<script type="text/javascript">$(document).ready(function(e) {$("#ccdz a").each(function(i){$(this).bind("click",function(){$("#txaContent").html($("#txaContent").html()+$(this).attr("href")+"\n")})})});</script>
<%Set objConfig=Nothing%>

<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
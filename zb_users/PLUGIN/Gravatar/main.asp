<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
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
If CheckPluginState("Gravatar")=False Then Call ShowError(48)
BlogTitle="Gravatar"
Dim c
Set c=New TConfig
c.Load "Gravatar"
If Request.QueryString("act")="save" Then
	c.Write "c",Request.Form("a")
	c.Save
	Call SetBlogHint(True,Empty,Empty)
EnD iF
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->


<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
   <a href="main.asp"> <span class="m-left m-now">[插件后台管理页] </span></a>
  </div>
  <div id="divMain2">
   <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>

<form id="form1" name="form1" method="post" action="?act=save">
<label for="a">请输入Gravatar头像地址</label><br/><br/>
<input type="text" name="a" id="a" style="width:100%" value="<%=c.read("c")%>" /><br/><br/>
<input name="" type="submit" class="button" value="提交" />
</form>

<a href="LoadHeader.asp?act=refresh">刷新后台头像</a>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->


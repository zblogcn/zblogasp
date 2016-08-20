<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
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
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("Gravatar")=False Then Call ShowError(48)
BlogTitle="Gravatar头像"
Dim c
Set c=New TConfig
c.Load "Gravatar"
If Request.QueryString("act")="save" Then
	Gravatar_Enable=Request.Form("Gravatar_Enable")
	Gravatar_EmailMD5=Request.Form("Gravatar_EmailMD5")
	c.Write "c",Gravatar_EmailMD5
	c.Write "e",Gravatar_Enable
	c.Save
	Call SetBlogHint(True,Empty,Empty)
	If Request.Form("Gravatar_Refresh")="True" Then
		Dim objRS
		Set objRS=objConn.Execute("SELECT [mem_ID],[mem_Name] FROM [blog_Member] ORDER BY [mem_ID] ASC")
		If (Not objRS.bof) And (Not objRS.eof) Then
			Do While Not objRS.eof
				Call AddBatch("缓存用户"& objRS("mem_Name")&"的Gravatar头像","Gravatar_GetImage "& objRS("mem_ID"))
				objRS.MoveNext
			Loop
		End If
	End If
EnD iF
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->


<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div id="divMain2">
   <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>

<form id="form" name="form" method="post" action="?act=save">

<input type="hidden" name="edtZC_STATIC_MODE" id="edtZC_STATIC_MODE" value="<%=ZC_STATIC_MODE%>" />
<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>
<tr><td><p  align='left'><b>·启用Gravatar头像</b></p></td><td><p><input id="Gravatar_Enable" name="Gravatar_Enable" style="" type="text" value="<%=Gravatar_Enable%>" class="checkbox"/></p></td></tr>


<tr><td width='30%'><p align='left'><b>·Gravatar URL</b><br/><span class='note'>推荐设置一般无需改动</span></p></td><td><p><input id='Gravatar_EmailMD5' name='Gravatar_EmailMD5' style='width:90%;' type='text' value='<%=Gravatar_EmailMD5%>' /></p></td></tr>
<tr>
<td><span class='note'>可选值: </span></td>
<td>
	<p><b>V2EX</b>：<a href="javascript:void(0)" class="enterGravatar">https://cdn.v2ex.com/gravatar/{%emailmd5%}.png?s=60&d=mm&r=G</a></p>
	<p><b>MoeNet公共库</b>：<a href="javascript:void(0)" class="enterGravatar">https://gravatar.moefont.com/avatar/{%emailmd5%}?s=60&amp;d=mm&amp;r=G</a></p>
	<p><b>多说CDN</b>：<a href="javascript:void(0)" class="enterGravatar">http://gravatar.duoshuo.com/avatar/{%emailmd5%}?s=60&amp;d=mm&amp;r=G</a></p>
	<p><b>官方加密</b>：<a href="javascript:void(0)" class="enterGravatar">https://secure.gravatar.com/avatar/{%emailmd5%}?s=60&amp;d=mm&amp;r=G</a></p>
	<p><b>官方</b>：<a href="javascript:void(0)" class="enterGravatar">http://cn.gravatar.com/avatar/{%emailmd5%}?s=60&amp;d=mm&amp;r=G</a></p>
	</td>
</tr>
<tr><td width='30%'><p align='left'><b>·刷新注册用户Gravatar头像的缓存</b><br/><span class='note'>如果用户数多会比较耗费时间和占用AVATAR目录空间</span></p></td><td><p><input id="Gravatar_Refresh" name="Gravatar_Refresh" style="" type="text" value="<%=Gravatar_Refresh%>" class="checkbox"/></p></td></tr>
</table>

<br/>
<input name="" type="submit" class="button" value="保存"/>
</form>
<script>
$(function() {
	$(".enterGravatar").click(function() {
		var $this = $(this);
		$("#Gravatar_EmailMD5").val($this.text());
	});
});
</script>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->


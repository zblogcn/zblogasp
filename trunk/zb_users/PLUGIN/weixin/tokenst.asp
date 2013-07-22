<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("weixin")=False Then Call ShowError(48)
BlogTitle="微信搜索"

	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load("weixin")
	
	If Request.QueryString("act")="Save" Then
		objConfig.Write "token",Request.Form("token")
		objConfig.Save
	End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=weixin_SubMenu(1)%></div>
          <div id="divMain2"> 
		  您的微信接口URL为：<a target="_blank" href="<%=ZC_BLOG_HOST%>zb_users/plugin/weixin/api.asp"><%=ZC_BLOG_HOST%>zb_users/plugin/weixin/api.asp</a>
			<form id="form1" name="form1" method="post" >
			<table width="100%" style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
		  <tr>
			<th width='10%'><p align="center">设置项目</p></th>
			<th width='5%'><p align="center">设置内容</p></th>
			<th width='50%'><p align="left">设置说明</p></th>
			
		  </tr>
		  <tr>
			<td><b><label for="token"><p align="center">微信验证token</p></label></b></td>
			<td><p align="center"><input name="token" type="text" id="token"  style="width:80px;" value="<%=objConfig.Read("token")%>" /></p></td>
			<td><b><label for="token"><p align="left">&nbsp;&nbsp;设置token，该值在微信后台验证博客时会用到。</p></label></b></td>
		  </tr>
		</table>
		 <br />
		   <input name="" type="submit" class="button" value="保存" onclick="document.getElementById('form1').action='?act=Save';"/>
		  
			</form>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>

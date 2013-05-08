<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<!-- #include file="encode\sha1.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("api")=False Then Call ShowError(48)
BlogTitle="API"
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load("api")
	If (objConfig.Read("id")="0") Then
		objConfig.Write "id",md5(Base64Encode(ZC_BLOG_HOST))
		objConfig.Write "secret",hex_sha1(md5(Base64Encode(ZC_BLOG_HOST&allRand(10))))
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
          <div class="SubMenu"><%=api_SubMenu(0)%></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("aPlugInMng");</script> 
            Access Key ID :===>  <%=objConfig.Read("id")%>
			<br>
			Access Key Secret :===>  <%=objConfig.Read("secret")%>
			<br>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>

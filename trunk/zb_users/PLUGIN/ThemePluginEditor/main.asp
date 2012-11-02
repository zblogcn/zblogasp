<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
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
If CheckPluginState("ThemePluginEditor")=False Then Call ShowError(48)
BlogTitle="主题插件生成器"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <p>这个插件，只是为了帮助主题作者制作可以进行管理的主题而已，嗯。</p>
            <p>你需要给主题INCLUDE文件夹下添加需要引用的文件，然后这里就会自动出现。</p>
            <form action="save.asp" method="post">
            <table width="100%" border="1" width="100%" class="tableBorder">
            <tr>
              <th scope="col" height="32" width="100px">文件名</th>
              <th scope="col" width="50px">文件类型</th>
              <th scope="col">文件注释</th>
            </tr>
            <%
			Dim oFso,oF
			Set oFso=Server.CreateObject("scripting.filesystemobject")
			Set oF=oFso.GetFolder(BlogPath & "\zb_users\theme\" & ZC_BLOG_THEME & "\include").Files
			Dim oS,s
			For Each oS In oF
			s=TransferHTML(oS.Name,"[html-format]")
			%>
            <tr>
            <td><%=oS.Name%></td>
            <td><select name="type_<%=s%>"><option value="1">HTML</option><option value="2">文件</option></select></td><td>
            <input type="text" id="<%=s%>" name="include_<%=s%>" value="" style="width:98%"/>
            </td></tr>
            <%
			Next
			%>
            
            </table>
            <input type="submit" value="提交"/>
            </form>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>

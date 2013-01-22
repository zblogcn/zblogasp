<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="function.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ZBDK")=False Then Call ShowError(48)
BlogTitle=zbdk_title
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"> <%=ZBDK.submenu.export("default")%> </div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("zbdk");</script> 
            <p>ZBDK，全称Z-Blog Plugin Development Kit，是为插件开发人员开发的一套工具包。它集合了许多插件开发中常用的工具，可以帮助插件开发者更好地进行插件开发。<span style="color:red">但不支持IE6.</span></p>
            <p>&nbsp;</p>
            <p>该插件有一定的危险性，一旦进行了误操作可能导致博客崩溃，请谨慎使用。</p>
            <p>&nbsp;</p>
            <p>工具列表：</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            
            <table width='100%'>
            <tr height='40'><td width='50'>ID</td><td width='120'>工具名</td><td>信息</td></tr>
            <%=ZBDK.mainpage.export()%>
            </table>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'************************************
' Powered by ThemePluginEditor<%=版本号%>
' zsx http://www.zsxsoft.com
'************************************
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\..\c_option.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\..\..\plugin\p_config.asp" -->

<%

Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

BlogTitle="主题管理"
%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<script type="text/javascript">
function copydata(This){
	if(navigator.userAgent.search(/ie/i)>-1){
		window.clipboardData.setData("text",$(This).attr("bindtag"));
		alert("复制成功！");
	}
	else{
		prompt("您的浏览器不支持复制，请按下CTRL+C手动复制！",$(This).attr("bindtag"));
	}	
}
</script>
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"></div>
          <div id="divMain2"> 
          <form action="save.asp" method="post">
          <table width="100%" border="1" width="100%" class="tableBorder">
            <tr>
              <th scope="col" height="32" width="100px">配置项</th>
              <th scope="col">配置内容</th>
              <th scope="col" width="500px">调用代码</th>
            </tr>
			<%=表格%>
          </table>
          <input name="ok" type="submit" class="button" value="提交"/>
          </form>
          <script type="text/javascript">ActiveLeftMenu("aThemeMng");</script> 
          </div>
        </div>
        <!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>

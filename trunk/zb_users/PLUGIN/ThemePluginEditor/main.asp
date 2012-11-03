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
<script type="text/javascript">
function newtr(This){
	var m=window.prompt("请输入新文件名","xxxx.html");
	if(m==null) return false;
	$(This).parent().parent().before("<tr><td>"+m+"</td><td>"+"<select name=\"type_"+m+"\"><option value=\"1\" selected=\"selected\">HTML</option><option value=\"2\">文件</option></select></td><td><input type=\"text\" id=\""+m+"\" name=\"include_"+m+"\" value=\"\" style=\"width:98%\"/><input type=\"hidden\" id=\""+m+"_2\" name=\"new_"+m+"\"/></td></tr>");bmx2table();
}
</script>
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
            <p>如果需要删除，删除INCLUDE下的文件，这里也会相应删除。</p>
            <form action="save.asp" method="post">
            <table width="100%" border="1" width="100%" class="tableBorder">
            <tr>
              <th scope="col" height="32" width="100px">文件名</th>
              <th scope="col" width="100px">文件类型</th>
              <th scope="col">文件注释</th>
            </tr>
            <%
			Dim oFso,oF
			Set oFso=Server.CreateObject("scripting.filesystemobject")
			If oFSO.FolderExists(BlogPath & "\zb_users\theme\" & ZC_BLOG_THEME & "\include")=False Then
				oFSO.CreateFolder BlogPath & "\zb_users\theme\" & ZC_BLOG_THEME & "\INCLUDE"
			End If
			Set oF=oFso.GetFolder(BlogPath & "\zb_users\theme\" & ZC_BLOG_THEME & "\include").Files
			Dim oS,s
			For Each oS In oF
			s=TransferHTML(oS.Name,"[html-format]")
			%>
            <tr>
            <td><%=oS.Name%></td>
            <td><select name="type_<%=s%>">
            <option value="1">HTML</option>
            <option value="2"<%=IIf(isHTML(s),""," selected=""selected"")")%>>文件</option></select></td><td>
            <input type="text" id="<%=s%>" name="include_<%=s%>" value="" style="width:98%"/>
            </td></tr>
            
            <%
			Next
			%>
            <tr id="new"><td><a href='javascript:;' onclick='newtr(this)'>新建..</a></td><td></td><td></td></tr>
            </table>
            <input type="submit" value="提交"/>
            </form>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>
<%
Function isHTML(s)
	Dim j
	If Instr(s,".") Then
		j=Mid(s,InstrRev(s,".")+1)
	Else
		j=""
	End If
	Select Case LCase(j)
		Case "html","htm","css","js","txt","xml","asp" isHTML=True
	End Select
End Function
%>
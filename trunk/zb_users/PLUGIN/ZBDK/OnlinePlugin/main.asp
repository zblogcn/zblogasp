<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ZBDK")=False Then Call ShowError(48)
BlogTitle=zbdk_title
If Request.QueryString("act")="save" Then
	Call SaveToFile(Server.MapPath("aspcode.asp"),Request.Form("asp"),"utf-8",false)
	Response.Redirect "main.asp"
	Response.End
End If
%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->

<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"> <%=ZBDK.submenu.export("OnlinePlugin")%> </div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("zbdk");</script>
            <form id="form1" name="form1" method="post" action="main.asp?act=save">
              <input type="submit" value="保存" class="button"/><label for="asp">在这里直接输入ASP代码，只要启用ZBDK插件这段代码就有效。不能有语法错误。</label>
              <p>&nbsp;</p>
              <textarea name="asp" id="asp" cols="45" rows="5" style="width:100%;height:500px" ><%=TransferHTML(LoadFromFile(Server.MapPath("aspcode.asp"),"utf-8"),"[textarea]")%></textarea>
            </form>
<div id="result"></div>
          </div>
        </div>


        <!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
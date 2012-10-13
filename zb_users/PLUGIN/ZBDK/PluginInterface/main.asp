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
BlogTitle=title
If Request.QueryString("act")="interface" Then
	Dim s,n,j,i
	s=LCase(Request.Form("interface"))
	
	Select Case Left(s,6)
		Case "action"
		'我X不能用Join
			Response.Write "<table width='100%'><tr><td height='40'>代码（共"
			Execute "Response.Write Ubound("&s&")"
			Response.Write "行）</td></tr>"
			Execute "j="&s
			For i=1 To Ubound(j)
				n=n & "<tr><td height='40'>" & TransferHTML(j(i),"[html-format]") & "</td></tr>"
			Next
			Response.Write n
			
		Case "filter"
			Response.Write "<table width='100%'><tr><td height='40'>函数（共"
			Execute "j=Split(s"&s&",""|"")"
			Response.Write Ubound(j)&"个）</td></tr>"
			For i=0 To Ubound(j)-1
				n=n & "<tr><td height='40'>" & TransferHTML(j(i),"[html-format]") & "</td></tr>"
			Next
			Response.Write n
		Case "respon"
			Execute "Response.Write TransferHTML("&s&",""[html-format]"")"
	End Select
	Response.Write "</table>"
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
          <div class="SubMenu"> <%=ZBDK.submenu(4)%> </div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <form id="form1" onsubmit="return false">
            <label for="interface">输入接口名</label>
            <input type="text" name="interface" id="interface" style="width:80%"/>
            <input type="submit" name="ok" id="ok" value="查看" onclick=""/>
            </form>
            <div id="result"></div>
          </div>
        </div>
        <script type="text/javascript">
		$(document).ready(function() {
            $("#form1").bind("submit",function(){
				$.post("main.asp?act=interface",{"interface":$("#interface").val()},function(data){
					$("#result").html(data);
					bmx2table();
				}
				)
			}
			)
        });
		</script>
        <!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
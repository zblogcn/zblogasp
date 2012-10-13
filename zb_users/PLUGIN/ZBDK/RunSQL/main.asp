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
If Request.QueryString("act")="sql" Then
	Dim s,objRs,i
	s=Request.Form("sql")
	On Error Resume Next
	Set objRs=objConn.Execute(s)
	Response.Write "查询用时" & RunTime & "ms<br/>"
	If Err.Number=0 Then
		If Not objRs.Eof Then
			Response.Write "<table width='100%'><tr>"
			For i=0 To objRs.fields.count-1 
				response.write "<td height='40'>"&objRs(i).Name& "</td>" 
			Next 
			Response.Write "</tr>"
			Do Until objRs.Eof
				Response.Write "<tr>"
				For i=0 To objRs.fields.count-1 
					response.write "<td>"&TransferHTML(objRs(i),"[html-format]")& "</td>" 
				Next 	
				Response.Write "</tr>"
				objRs.MoveNext
			Loop
			Response.Write "</table>"
		End If
	Else
		Response.Write "出现错误：" & Err.Number & "<br/>错误信息：" & Err.Description 
	End If
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
          <div class="SubMenu"> <%=ZBDK.submenu(3)%> </div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <form id="form1" onSubmit="return false">
            <label for="sql">输入SQL代码</label>
            <input type="text" name="sql" id="sql" style="width:80%"/>
            <input type="submit" name="ok" id="ok" value="提交" onClick=""/>
            </form>
            <div id="result"></div>
          </div>
        </div>
        <script type="text/javascript">
		$(document).ready(function() {
            $("#form1").bind("submit",function(){
				$("#result").html("Waiting...");
				$.post("main.asp?act=sql",{"sql":$("#sql").val()},function(data){
					$("#result").html(data);
					bmx2table();
				}
				)
			}
			)
        });
		</script>
        <!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
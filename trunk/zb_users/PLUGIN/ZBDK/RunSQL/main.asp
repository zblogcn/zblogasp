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
	Response.Write "<div class=""DIVBlogConfigtop""><span id=""name"">查询用时" & RunTime & "ms</span><a href='javascript:;' onclick='$(""#tree"").toggleClass(""hide"");$(""#result"").css(""margin-left"",$(""#result"").css(""margin-left"")==""200px""?""0px"":""200px"")'>[显示/隐藏左侧表]</a></div>"
	If Err.Number=0 Then
		If objRs.fields.count=0 Then Response.End
		Response.Write "<table width='100%' class='tablesorter'><tr>"
		For i=0 To objRs.fields.count-1 
			response.write "<th height='40'>"&objRs(i).Name& "</th>" 
		Next 
		Response.Write "</tr>"
		
		Do Until objRs.Eof
			Response.Write "<tr height='32'>"
			For i=0 To objRs.fields.count-1 
				response.write "<td>"&TransferHTML(objRs(i),"[html-format]")& "</td>" 
			Next 	
			Response.Write "</tr>"
			objRs.MoveNext
		Loop
		Response.Write "</table>"
	'End If
	Else
		Response.Write "<br/>出现错误：" & Err.Number & "<br/>错误信息：" & Err.Description 
	End If
	Response.End
End If
%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<link rel="stylesheet" href="../css/BlogConfig.css" type="text/css" media="screen"/>
<script type="text/javascript" src="../script/colResizable-1.3.min.js"></script>
<style type="text/css">
td{text-align: center}
</style>
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"> <%=ZBDK.submenu(3)%> </div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("zbdk");</script>
            <form id="form1" onSubmit="return false">
              <label for="sql">输入SQL代码</label>
              <input type="text" name="sql" id="sql" style="width:80%"/>
              <input type="submit" name="ok" id="ok" value="提交" onClick=""/>
            </form>
            <div class="DIVBlogConfig">
              <div class="DIVBlogConfignav" name="tree" id="tree">
                <ul>
                  <%=ReadTables%>
                </ul>
              </div>
              <div id="result" class="DIVBlogConfigcontent"> 
               
              </div>
              <div class="clear"></div>
            </div>
          </div>
        </div>
      </div>
    </div>
    <script type="text/javascript">
		$(document).ready(function() {
            $("#form1").bind("submit",function(){
				$("#result").html("Waiting...");
				$.post("main.asp?act=sql",{"sql":$("#sql").val()},function(data){
					$("#result").html(data);
					bmx2table();
					 $("#result table").colResizable({
						liveDrag:true,
//						gripInnerHtml:"<div class='grip'>ceshi</div>", 
						draggingClass:"dragging", 
						onResize:function(e){  
    						var table = $(e.currentTarget); //reference to the resized table
  						}
					  });  
				}
				)
			}
			);
			$("a[sql]").click(function(){
				var h=$(this);
				$("#sql").val('SELECT TOP 100 * FROM '+h.attr("table"));
				$("#form1").submit();
			});
        });
		</script> 
    <!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
	<%

	Function ReadTables
		Dim objRs
		If ZC_MSSQL_ENABLE Then
			Set objRs=objConn.Execute("SELECT [name] As [table_name] FROM [dbo].[sysobjects] WHERE TYPE='u'")
		Else
			Set objRs=objConn.OpenSchema(20)
			objRs.Filter="table_type='table'"
		End If
		Do Until objRs.Eof
			ReadTables=ReadTables&"<li><a table='"&Server.HTMLEncode(objRs("table_name"))&"' sql='sql' id='a" & Server.HTMLEncode(objRs("table_name")) & "' href='javascript:;'>"&objRs("table_name")&"</a></li>"
			objRs.MoveNext
		Loop
		Set objRs=Nothing
	End Function
	%>
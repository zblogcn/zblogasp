<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 

If CheckPluginState("Totoro")=False Then Call ShowError(48)

BlogTitle="TotoroⅢ（基于TotoroⅡ的Z-Blog的评论管理审核系统增强版）"

%><!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

			<div id="divMain">
<div class="divHeader"><%=BlogTitle%></div>
<div class="SubMenu"><a href="setting.asp"><span class="m-left">TotoroⅢ设置</span></a><a href="setting1.asp"><span class="m-left m-now">审核评论<%
	Dim objRS1
	Set objRS1=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [comm_isCheck]=-1")
	If (Not objRS1.bof) And (Not objRS1.eof) Then
		Response.Write "("&objRS1(0)&"条未审核的评论)"
	End If
%></span></a></div>
<div id="divMain2">
<%

	Dim intPage,intContent

	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage

	intPage=Request.QueryString("page")

	Call CheckParameter(intPage,"int",1)
	intContent=FilterSQL(intContent)

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	strSQL=strSQL&" WHERE  ([comm_isCheck]=-1) "

	If CheckRights("Root")=False Then strSQL=strSQL & "AND( ([comm_AuthorID] = " & BlogUser.ID & " ) OR ((SELECT [log_AuthorID] FROM [blog_Article] WHERE [blog_Article].[log_ID]=[blog_Comment].[log_ID])=" & BlogUser.ID & " )) "

	If Trim(intContent)<>"" Then strSQL=strSQL & " AND ( ([comm_Author] LIKE '%" & intContent & "%') OR ([comm_Content] LIKE '%" & intContent & "%') OR ([comm_HomePage] LIKE '%" & intContent & "%') ) "

	Call GetBlogHint()

	Response.Write "<form id=""frmBatch"" method=""post"" action=""""><p><input class=""button button2"" type=""submit"" onclick='this.form.action="""&ZC_BLOG_HOST&"zb_users/plugin/totoro/commentdel.asp?act=delALL"&""";return window.confirm("""& ZC_MSG058 &""");' value=""删除数据库中所有未审核的评论"" id=""btnPost""/></p></form><br/>"

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""1"" cellpadding=""1"">"
	Response.Write "<tr><td width='5%'>"& ZC_MSG076 &"</td><td width='14%'>"& ZC_MSG001 &"</td><td>"& ZC_MSG055 &"</td><td width='12%'>"& ZC_MSG080 &"</td><td width='15%'>"& ZC_MSG075 &"</td><td width='5%'  align='center'><a href='' onclick='BatchSelectAll();return false'>"& ZC_MSG229 &"</a></td></tr>"'

	objRS.Open("SELECT * FROM [blog_Comment] "& strSQL &" ORDER BY [comm_ID] DESC")


	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			Response.Write "<tr>"
			Response.Write "<td>" & objRS("comm_ID") & "</td>"
			If Trim(objRS("comm_Email"))="" Then
			Response.Write "<td>"& objRS("comm_Author") & "</td>"
			Else
			Response.Write "<td><a href=""mailto:"& objRS("comm_Email") &""">" & objRS("comm_Author") & "</a></td>"
			End If

			Response.Write "<td><a href="""" onclick='javascript:$(this).parent().html(""" & TransferHTML(objRS("comm_Content"),"[html-format][enter][""]") & """);return false;' title="""&TransferHTML(TransferHTML(UBBCode(objRS("comm_Content"),"[face][link][autolink][font][code][image][typeset][media][flash][key][upload]"),"[html-japan][upload]"),"[nohtml]")&""">" & Left(objRS("comm_Content"),40) & "...</a></td>"
			Response.Write "<td>" & objRS("comm_IP") & "</td>"
			Response.Write "<td>" & objRS("comm_PostTime") & "</td>"
			Response.Write "<td align=""center"" ><input type=""checkbox"" name=""edtDel"" id=""edtDel"" value="""&objRS("comm_ID")&"""/></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	Response.Write "</table>"

	For i=1 to objRS.PageCount
		strPage=strPage &"<a href='"&ZC_BLOG_HOST&"/zb_users/plugin/totoro/setting1.asp?page="& i &"'>["& Replace(ZC_MSG036,"%s",i) &"]</a> "
	Next
	Response.Write "<br/><form id=""frmBatch"" method=""post"" action=""""><p><input type=""hidden"" id=""edtBatch"" name=""edtBatch"" value=""""/><input class=""button button2"" type=""submit"" onclick='BatchDeleteAll(""edtBatch"");if(document.getElementById(""edtBatch"").value){this.form.action="""&ZC_BLOG_HOST&"/zb_users/plugin/totoro/commentdel.asp"&""";return window.confirm("""& ZC_MSG058 &""");}else{return false}' value=""删除所选择的评论"" id=""btnPost""/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input class=""button button2"" type=""submit"" onclick='BatchDeleteAll(""edtBatch"");if(document.getElementById(""edtBatch"").value){this.form.action="""&ZC_BLOG_HOST&"/zb_users/plugin/totoro/commentpass.asp"&""";return window.confirm("""& ZC_MSG058 &""");}else{return false}' value=""通过所选择的评论"" id=""btnPost""/></p><form><br/>" & vbCrlf

	Response.Write "<hr/>" & ZC_MSG042 & ": " & strPage

	objRS.Close
	Set objRS=Nothing


%>

</div>
<script language="javascript">

	//斑马线
	var tables=document.getElementsByTagName("table");
	var b=false;
	for (var j = 0; j < tables.length; j++){

		var cells = tables[j].getElementsByTagName("tr");

		cells[0].className="color1";
		for (var i = 1; i < cells.length; i++){
			if(b){
				cells[i].className="color2";
				b=false;
			}
			else{
				cells[i].className="color3";
				b=true;
			};
		};
	}

function ChangeValue(obj){

	if (obj.value=="True")
	{
	obj.value="False";
	return true;
	}

	if (obj.value=="False")
	{
	obj.value="True";
	return true;
	}
}
</script>
</div><!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>


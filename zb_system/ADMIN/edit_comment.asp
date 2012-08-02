<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog 彩虹网志个人版
'// 作    者:    朱煊(zx.asd) & Sipo
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    edit_comment.asp
'// 开始时间:    2006.12.30
'// 最后修改:    
'// 备    注:    评论编辑页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

'plugin node
For Each sAction_Plugin_Edit_Comment_Begin in Action_Plugin_Edit_Comment_Begin
	If Not IsEmpty(sAction_Plugin_Edit_Comment_Begin) Then Call Execute(sAction_Plugin_Edit_Comment_Begin)
Next

'检查非法链接
Call CheckReference("")

'检查权限
If Not CheckRights("CommentEdt") Then Call ShowError(6)

GetCategory()
GetUser()

Dim EditComment
Set EditComment=New TComment

Dim EditArticle
Set EditArticle=New TArticle

Dim IsRev
IsRev=False

If Request.QueryString("id")<>0 Then
	If EditComment.LoadInfoByID(Request.QueryString("id"))=False Then
		Call ShowError(12)
	Else
		EditArticle.LoadInfoByID EditComment.log_ID
	End If
End If

If Request.QueryString("revid")<>0 Then
	Set EditComment=New TComment
	EditComment.ParentID=Trim(Request.QueryString("revid"))
	EditComment.log_ID=Trim(Request.QueryString("log_id"))
	EditComment.ID=0
	EditComment.Author=BlogUser.Name
	EditComment.EMail=BlogUser.Email
	EditComment.HomePage=BlogUser.HomePage
	EditComment.Content=""
	EditArticle.LoadInfoByID Trim(Request.QueryString("log_id"))
	IsRev=True
Else
	IsRev=False
End If

BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG272

%>
<!--#include file="admin_header.asp"-->
<!--#include file="admin_top.asp"-->
<div id="divMain">
<%	Call GetBlogHint()	%>
<div class="divHeader2"><%=ZC_MSG272%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_CommentEdt_SubMenu & "</div>"
%>
<div id="divMain2">
<form id="edit" name="edit" method="post" action="">
<%
	Response.Write "<input id=""inpID"" name=""inpID""  type=""hidden"" value="""& EditComment.ID &""" />"
'If IsRev=False Then
	Response.Write "<p><span class='title'>"& ZC_MSG270 &":</span><br/><input type=""hidden"" id=""inpLogID"" name=""inpLogID"" value="""& EditComment.log_ID &""" />"& EditArticle.HtmlTitle &"</p>"
'Else
'	Response.Write "<input type=""hidden"" id=""inpLogID"" name=""inpLogID"" value="""& EditComment.log_ID &""" />"
'End If

	Response.Write "<p><span class='title'>"& ZC_MSG095 &":</span><br/><input type=""text"" readonly=""readonly"" id=""intRevID"" name=""intRevID"" value="""& TransferHTML(EditComment.ParentID,"[html-format]") &""" size=""40""  /></p>"
If IsRev=False Then
	Response.Write "<p><span class='title'>"& ZC_MSG001 &":</span><span class='star'>(*)</span><br/><input type=""text"" id=""inpName"" name=""inpName"" value="""& TransferHTML(EditComment.Author,"[html-format]") &""" size=""40"" /></p>"
	Response.Write "<p><span class='title'>"& ZC_MSG053 &":</span><br/><input type=""text"" name=""inpEmail"" value="""& TransferHTML(EditComment.Email,"[html-format]") &""" size=""40""  /></p>"
	Response.Write "<p><span class='title'>"& ZC_MSG054 &":</span><br/><input type=""text"" name=""inpHomePage"" value="""& TransferHTML(EditComment.HomePage,"[html-format]") &""" size=""40""  /></p>"
Else
	Response.Write "<input type=""hidden"" id=""inpName"" name=""inpName"" value="""& TransferHTML(EditComment.Author,"[html-format]") &""" />"
	Response.Write "<input type=""hidden"" name=""inpEmail"" value="""& TransferHTML(EditComment.Email,"[html-format]") &""" />"
	Response.Write "<input type=""hidden"" name=""inpHomePage"" value="""& TransferHTML(EditComment.HomePage,"[html-format]") &""" />"
End If
	Response.Write "<p><span class='title'>"& ZC_MSG090 &":</span><span class='star'>(*)</span><br/><textarea name=""txaContent"" id=""txaContent"" onchange=""GetActiveText(this.id);"" onclick=""GetActiveText(this.id);"" onfocus=""GetActiveText(this.id);"" cols=""80"" rows=""12"">"&EditComment.Content&"</textarea>(*)</p>"

	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='return checkCateInfo();' /><br/><script language=""JavaScript"" type=""text/javascript"">objActive=""txaArticle"";ExportUbbFrame();</script></p>"
%>
</form>
</div>

			</div>
<script type="text/javascript">
// <![CDATA[
	var str17="<%=ZC_MSG118%>";
	var str18="<%=ZC_MSG035%>";
	var str19="<%=ZVA_ErrorMsg(9)%>";

	function checkCateInfo(){
		document.getElementById("edit").action="../cmd.asp?act=CommentSav&revid=<%=Request.QueryString("revid")%>";

		if(!document.getElementById("inpID").value){
			alert(str19);
			return false
		}

		if(!document.getElementById("inpName").value){
			alert(str17);
			return false
		}

		if(!document.getElementById("txaArticle").value){
			alert(str18);
			return false
		}

	}
// ]]>
</script>
<script type="text/javascript">ActiveLeftMenu("aCommentMng");</script>
<!--#include file="admin_footer.asp"-->
<% 
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
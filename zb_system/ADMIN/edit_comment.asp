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
<% On Error Resume Next %>
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

Dim EditComment
Set EditComment=New TComment

If Not IsEmpty(Request.QueryString("id")) Then

	If EditComment.LoadInfoByID(Request.QueryString("id"))=False Then Call ShowError(12)

End If


BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG066

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<link rel="stylesheet" rev="stylesheet" href="../CSS/admin.css" type="text/css" media="screen" />
	<script language="JavaScript" src="../script/common.js" type="text/javascript"></script>
	<link rel="stylesheet" href="../CSS/jquery.bettertip.css" type="text/css" media="screen">
	<script language="JavaScript" src="../script/jquery.bettertip.pack.js" type="text/javascript"></script>
	<title><%=BlogTitle%></title>
</head>
<body>
<script language="JavaScript" type="text/javascript">
	var str00="<%=ZC_BLOG_HOST%>";
	var str01="<%=ZC_MSG033%>";
	var str02="<%=ZC_MSG034%>";
	var str03="<%=ZC_MSG035%>";
	var str06="<%=ZC_MSG057%>";
	var intMaxLen="<%=ZC_CONTENT_MAX%>";
	var strFaceName="<%=ZC_EMOTICONS_FILENAME%>";
	var strFaceSize="<%=ZC_EMOTICONS_FILESIZE%>";
</script>
<div id="divMain">
<div class="Header"><%=ZC_MSG272%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_CommentEdt_SubMenu & "</div>"
%>
<div id="divMain2">
<% Call GetBlogHint() %>
<form id="edit" name="edit" method="post">
<%
	Response.Write "<input id=""edtID"" name=""edtID""  type=""hidden"" value="""& EditComment.ID &""" />"
	Response.Write "<p>"& ZC_MSG277 &":</p><p><input type=""text"" id=""inpID"" name=""inpID"" value="""& EditComment.log_ID &""" size=""40"" />(*)</p>"
	Response.Write "<p>"& "回复评论ID(设置为0则单独为一个评论，同时该ID不能为其他文章评论的ID)" &":</p><p><input type=""text"" name=""intRepComment"" value="""& TransferHTML(EditComment.ParentID,"[html-format]") &""" size=""40""  /></p>"

	Response.Write "<p>"& ZC_MSG001 &":</p><p><input type=""text"" id=""inpName"" name=""inpName"" value="""& TransferHTML(EditComment.Author,"[html-format]") &""" size=""40"" />(*)</p>"
	Response.Write "<p>"& ZC_MSG053 &":</p><p><input type=""text"" name=""inpEmail"" value="""& TransferHTML(EditComment.Email,"[html-format]") &""" size=""40""  /></p>"
	Response.Write "<p>"& ZC_MSG054 &":</p><p><input type=""text"" name=""inpHomePage"" value="""& TransferHTML(EditComment.HomePage,"[html-format]") &""" size=""40""  /></p>"
	
	Response.Write "<p>"& ZC_MSG055 &":</p><p><textarea name=""txaArticle"" id=""txaArticle"" onchange=""GetActiveText(this.id);"" onclick=""GetActiveText(this.id);"" onfocus=""GetActiveText(this.id);"" cols=""80"" rows=""12"">"&EditComment.Content&"</textarea>(*)</p>"

	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='return checkCateInfo();' /></p><p><script language=""JavaScript"" type=""text/javascript"">objActive=""txaArticle"";ExportUbbFrame();</script></p>"
%>
</form>
</div>

			</div>

</body>
<script>

	var str17="<%=ZC_MSG118%>";
	var str18="<%=ZC_MSG035%>";
	var str19="<%=ZVA_ErrorMsg(9)%>";

	function checkCateInfo(){
		document.getElementById("edit").action="../cmd.asp?act=CommentSav";

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
</script>
</html>
<% 
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
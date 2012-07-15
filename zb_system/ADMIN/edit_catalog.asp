<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog 彩虹网志个人版
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:
'// 程序版本:
'// 单元名称:    edit_catalog.asp
'// 开始时间:    2005.03.03
'// 最后修改:
'// 备    注:    编辑页
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
For Each sAction_Plugin_Edit_Catalog_Begin in Action_Plugin_Edit_Catalog_Begin
	If Not IsEmpty(sAction_Plugin_Edit_Catalog_Begin) Then Call Execute(sAction_Plugin_Edit_Catalog_Begin)
Next


'检查非法链接
Call CheckReference("")

'检查权限
If Not CheckRights("CategoryEdt") Then Call ShowError(6)

GetCategory()
GetUser()

Dim EditCategory
Set EditCategory=New TCategory

If Not IsEmpty(Request.QueryString("id")) Then

	If EditCategory.LoadInfoByID(Request.QueryString("id"))=False Then Call ShowError(12)

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

<div id="divMain">
<div class="Header"><%=ZC_MSG243%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_CategoryEdt_SubMenu & "</div>"
%>
<div id="divMain2">
<% Call GetBlogHint() %>
<form id="edit" name="edit" method="post">
<%
	Response.Write "<input id=""edtID"" name=""edtID""  type=""hidden"" value="""& EditCategory.ID &""" />"
	Response.Write "<p>"& ZC_MSG001 &":</p><p><input id=""edtName"" style='width:300px;' size=""40"" name=""edtName""  type=""text"" value="""& TransferHTML(EditCategory.Name,"[html-format]") &""" />(*)</p><p></p>"
	Response.Write "<p>"& ZC_MSG079 &":</p><p><input id=""edtOrder"" style='width:300px;' size=""40"" name=""edtOrder""  type=""text"" value="""& EditCategory.Order &""" /></p><p></p>"
	Response.Write "<p>"& ZC_MSG320 &":</p>"
	Response.Write "<p><select style='width:310px;' id=""edtPareID"" name=""edtPareID"" class=""edit"" size=""1"">"
	Response.Write "<option value=""0"" "
	If EditCategory.ParentID=0 Then Response.Write "selected=""selected"" "
	Response.Write ">"& ZC_MSG322 &"</option>"
	Dim Category,bolHasSubCate
	bolHasSubCate=False
	For Each Category in Categorys
		If IsObject(Category) Then
			If Category.ParentID=EditCategory.ID Then bolHasSubCate=True
		End If
	Next
	If EditCategory.ID=Empty Then bolHasSubCate=False
	If bolHasSubCate=False Then
		Dim aryCateInOrder,i
		aryCateInOrder=GetCategoryOrder()
		If IsArray(aryCateInOrder) Then
			For i=LBound(aryCateInOrder)+1 To Ubound(aryCateInOrder)
				If Categorys(aryCateInOrder(i)).ParentID=0 And Categorys(aryCateInOrder(i)).ID<>EditCategory.ID Then
					Response.Write "<option value="""&Categorys(aryCateInOrder(i)).ID&""" "
					If Categorys(aryCateInOrder(i)).ID=EditCategory.ParentID Then Response.Write "selected=""selected"" "
					Response.Write ">"&TransferHTML(Categorys(aryCateInOrder(i)).Name,"[html-format]")&"</option>"
				End If
			Next
		End If
	End If
	Response.Write "</select></p><p></p>"


	Response.Write "<p>"&ZC_MSG324&":</p><p><select style='width:310px;' class='edit' size='1' id='cmbTemplate' onchange='edtTemplate.value=this.options[this.selectedIndex].value'>"

	'Response.Write "<option value="""">"&ZC_MSG325&"</option>"

	Dim aryFileList

	aryFileList=LoadIncludeFiles("zb_users\theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY)

	If IsArray(aryFileList) Then
		Dim j,t
		j=UBound(aryFileList)
		For i=1 to j
			t=UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
			If EditCategory.TemplateName=t Then
				Response.Write "<option value="""&t&""" selected=""selected"">"&t&"</option>"
			Else
				Response.Write "<option value="""&t&""">"&t&"</option>"
			End If
		Next
	End If

	If EditCategory.TemplateName="" Then
	Response.Write "<option value='' selected='selected'>"&ZC_MSG325&"(CATALOG)</option>"
	Else
	Response.Write "<option value=''>"&ZC_MSG325&"(CATALOG)</option>"
	End If

	Response.Write "</select><input type='hidden' name='edtTemplate' id='edtTemplate' value="&EditCategory.TemplateName&" />"





	Response.Write "</p><p></p>"
	Response.Write "<p>"& ZC_MSG147 &":</p><p><input id=""edtAlias"" style='width:300px;' size=""40"" name=""edtAlias""  type=""text"" value="""& TransferHTML(EditCategory.Alias,"[html-format]") &""" /></p><p></p>"
	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='return checkCateInfo();' /></p><p></p>"
%>
</form>
</div>

			</div>

</body>
<script>

	var str17="<%=ZC_MSG118%>";

	function checkCateInfo(){
		document.getElementById("edit").action="../cmd.asp?act=CategoryPst";

		if(!document.getElementById("edtName").value){
			alert(str17);
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
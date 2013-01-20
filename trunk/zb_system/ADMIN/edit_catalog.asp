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

Dim EditCategory
Set EditCategory=New TCategory

EditCategory.Name=""

If Not IsEmpty(Request.QueryString("id")) Then

	If EditCategory.LoadInfoByID(Request.QueryString("id"))=False Then Call ShowError(12)

End If

BlogTitle=ZC_MSG066

'为1号输出输口准备的Action接口
'plugin node
For Each sAction_Plugin_EditCatalog_Form in Action_Plugin_EditCatalog_Form
	If Not IsEmpty(sAction_Plugin_EditCatalog_Form) Then Call Execute(sAction_Plugin_EditCatalog_Form)
Next

%>
<!--#include file="admin_header.asp"-->
<!--#include file="admin_top.asp"-->
<div id="divMain">
<%	Call GetBlogHint()	%>
<div class="divHeader2"><%=ZC_MSG243%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_CategoryEdt_SubMenu & "</div>"
%>
<div id="divMain2">
<form id="edit" name="edit" method="post" action="">
<%
	Response.Write "<input id=""edtID"" name=""edtID""  type=""hidden"" value="""& EditCategory.ID &""" />"
	Response.Write "<p><span class='title'>"& ZC_MSG001 &":</span><span class='star'>(*)</span><br/><input id=""edtName"" style='width:300px;' size=""40"" name=""edtName"" maxlength=""50"" type=""text"" value="""& TransferHTML(EditCategory.Name,"[html-format]") &""" /></p>"
	Response.Write "<p><span class='title'>"& ZC_MSG147 &":</span><br/><input id=""edtAlias"" style='width:300px;' size=""40"" name=""edtAlias""  type=""text"" value="""& TransferHTML(EditCategory.Alias,"[html-format]") &""" /></p>"

If Request.QueryString("id")<>"0" Then

	Response.Write "<p><span class='title'>"& ZC_MSG079 &":</span><br/><input id=""edtOrder"" style='width:300px;' size=""40"" name=""edtOrder""  type=""text"" value="""& EditCategory.Order &""" /></p>"

	Response.Write "<p><span class='title'>"& ZC_MSG195 &":</span><br/><select style='width:310px;' id=""edtPareID"" name=""edtPareID"" class=""edit"" size=""1"">"
	Response.Write "<option value=""0"" "
	If EditCategory.ParentID=0 Then Response.Write "selected=""selected"" "
	Response.Write ">"& ZC_MSG180 &"</option>"
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
	Response.Write "</select></p>"


	Response.Write "<p><span class='title'>"&ZC_MSG188&":</span><br/><select style='width:310px;' class='edit' size='1' id='cmbTemplate' onchange='edtTemplate.value=this.options[this.selectedIndex].value'>"

	Dim aryFileList
	aryFileList=LoadIncludeFilesOnlyType("zb_users\theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY)

	If IsArray(aryFileList) Then
		Dim j,t
		j=UBound(aryFileList)
		For i=1 to j
			t=UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
			If Left(t,2)<>"B_" AND t<>"FOOTER" And t<>"HEADER" And t<>"SINGLE" And t<>"PAGE" Then
				If EditCategory.GetDefaultTemplateName=t Then
					Response.Write "<option value="""&t&""" selected=""selected"">"&t&IIF(EditCategory.TemplateName=""," ("&ZC_MSG187&")","")&"</option>"
				Else
					Response.Write "<option value="""&t&""">"&t&"</option>"
				End If
			End If
		Next
	End If

	'If EditCategory.TemplateName="" Then
	'Response.Write "<option value='' selected='selected'>"&ZC_MSG187&"(CATALOG)</option>"
	'Else
	'Response.Write "<option value=''>"&ZC_MSG187&"(CATALOG)</option>"
	'End If

	Response.Write "</select><input type='hidden' name='edtTemplate' id='edtTemplate' value='"&EditCategory.TemplateName&"' />"
	Response.Write "</p>"


	Response.Write "<p><span class='title'>"&ZC_MSG179&":</span><br/><select style='width:310px;' class='edit' size='1' id='cmbLogTemplate' onchange='edtLogTemplate.value=this.options[this.selectedIndex].value'>"

	aryFileList=LoadIncludeFilesOnlyType("zb_users\theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY)

	If IsArray(aryFileList) Then
		j=UBound(aryFileList)
		For i=1 to j
			t=UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
			If Left(t,2)<>"B_" AND t<>"FOOTER" And t<>"HEADER" And  t<>"CATALOG" And t<>"DEFAULT" Then
				If EditCategory.GetDefaultLogTemplateName=t Then
					Response.Write "<option value="""&t&""" selected=""selected"">"&t&IIF(EditCategory.LogTemplate=""," ("&ZC_MSG187&")","")&"</option>"
				Else
					Response.Write "<option value="""&t&""">"&t&"</option>"
				End If
			End If
		Next
	End If

	'If EditCategory.LogTemplate="" Then
	'Response.Write "<option value='' selected='selected'>"&ZC_MSG187&"(SINGLE)</option>"
	'Else
	'Response.Write "<option value=''>"&ZC_MSG187&"(SINGLE)</option>"
	'End If

	Response.Write "</select><input type='hidden' name='edtLogTemplate' id='edtLogTemplate' value='"&EditCategory.LogTemplate&"' />"

	Response.Write "</p>"

	Response.Write "<p><label><input type='checkbox' name='edtAddNavbar' id='edtAddNavbar' value='True'' />&nbsp;&nbsp;<span class='title'>"& ZC_MSG046 &"</span></label></p>"
Else
	Response.Write "<p>"& ZC_MSG261 &"</p>"
End If

	'<!-- 1号输出接口 -->
	If Response_Plugin_EditCatalog_Form<>"" Then Response.Write "<div id=""divEditForm1"">"&Response_Plugin_EditCatalog_Form&"</div>"

	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='return checkCateInfo();' /></p><p></p>"
%>
</form>
<script type="text/javascript">
	
		var str17="<%=ZC_MSG118%>";
	
		function checkCateInfo(){
			document.getElementById("edit").action="../cmd.asp?act=CategoryPst";
	
			if(!$("#edtName").val()){
				alert(str17);
				return false
			}
	
		}
</script>
<script type="text/javascript">ActiveLeftMenu("aCategoryMng");</script>
</div>
</div>
<!--#include file="admin_footer.asp"-->
<%
Call System_Terminate()

If Err.Number<>0 then
	'Call ShowError(0)
End If
%>
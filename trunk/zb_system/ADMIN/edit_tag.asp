<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog 彩虹网志个人版
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    edit_tag.asp
'// 开始时间:    2005.04.07
'// 最后修改:    
'// 备    注:    
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
For Each sAction_Plugin_Edit_Tag_Begin in Action_Plugin_Edit_Tag_Begin
	If Not IsEmpty(sAction_Plugin_Edit_Tag_Begin) Then Call Execute(sAction_Plugin_Edit_Tag_Begin)
Next


'检查非法链接
Call CheckReference("")

'检查权限
If Not CheckRights("TagEdt") Then Call ShowError(6)

GetCategory()
GetUser()

Dim EditTag
Set EditTag=New TTag

If Not IsEmpty(Request.QueryString("id")) Then

	If EditTag.LoadInfoByID(Request.QueryString("id"))=False Then Call ShowError(35)

End If


BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG066

%>
<!--#include file="admin_header.asp"-->
<!--#include file="admin_top.asp"-->   
<div id="divMain">
<%	Call GetBlogHint()	%>
<div class="divHeader2"><%=ZC_MSG241%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_TagEdt_SubMenu & "</div>"
%>
<div id="divMain2">
<form id="edit" name="edit" method="post" action="">
<%
	Response.Write "<input id=""edtID"" name=""edtID""  type=""hidden"" value="""& EditTag.ID &""" />"
	Response.Write "<p>"& ZC_MSG001 &":<br/><input id=""edtName"" style='width:300px;' name=""edtName""  type=""text"" value="""& TransferHTML(EditTag.Name,"[html-format]") &""" />(*)</p>"
	Response.Write "<p>"& ZC_MSG016 &":<br/><input id=""edtIntro"" style='width:300px;' name=""edtIntro""  type=""text"" value="""& TransferHTML(EditTag.Intro,"[html-format]") &""" /></p>"
	Response.Write "<p>"& ZC_MSG147 &":<br/><input id=""edtAlias"" style='width:300px;' name=""edtAlias""  type=""text"" value="""& TransferHTML(EditTag.Alias,"[html-format]") &""" /></p>"

	Response.Write "<p>"&ZC_MSG324&":<br/><select style='width:310px;' class='edit' size='1' id='cmbTemplate' onchange='edtTemplate.value=this.options[this.selectedIndex].value'>"

	'Response.Write "<option value="""">"&ZC_MSG325&"</option>"

	Dim aryFileList

	aryFileList=LoadIncludeFiles("zb_users\theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY)

	If IsArray(aryFileList) Then
		Dim i,j,t
		j=UBound(aryFileList)
		For i=1 to j
			t=UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
			If EditTag.TemplateName=t Then
				Response.Write "<option value="""&t&""" selected=""selected"">"&t&"</option>"
			Else 
				Response.Write "<option value="""&t&""">"&t&"</option>"
			End If
		Next
	End If

	If EditTag.TemplateName="" Then
	Response.Write "<option value='' selected='selected'>"&ZC_MSG325&"(CATALOG)</option>"
	Else
	Response.Write "<option value=''>"&ZC_MSG325&"(CATALOG)</option>"
	End If 

	Response.Write "</select><input type='hidden' name='edtTemplate' id='edtTemplate' value='"&EditTag.TemplateName&"' /></p>"

	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='return checkTagInfo();' /></p><p></p>"
%>
</form>
</div>
</div>
<script type="text/javascript">
	
		var str17="<%=ZC_MSG118%>";
	
		function checkTagInfo(){
			document.getElementById("edit").action="../cmd.asp?act=TagPst";
	
			if(!document.getElementById("edtName").value){
				alert(str17);
				return false
			}
	
		}
		ActiveLeftMenu("aTagMng");
		</script>
<!--#include file="admin_footer.asp"-->
<% 
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
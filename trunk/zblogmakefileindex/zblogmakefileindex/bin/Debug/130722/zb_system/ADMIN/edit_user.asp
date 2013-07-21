<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog 彩虹网志个人版
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    edit_user.asp
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
For Each sAction_Plugin_Edit_User_Begin in Action_Plugin_Edit_User_Begin
	If Not IsEmpty(sAction_Plugin_Edit_User_Begin) Then Call Execute(sAction_Plugin_Edit_User_Begin)
Next


'检查非法链接
Call CheckReference("")

'检查权限
If Not CheckRights("UserEdt") Then Call ShowError(6)

Dim EditUser
Set EditUser=New TUser

If Not IsEmpty(Request.QueryString("id")) Then

	If EditUser.LoadInfoByID(Request.QueryString("id")) Then
		If (EditUser.ID<>BlogUser.ID) And (CheckRights("Root")=False) Then Call ShowError(6)
	Else
		Call ShowError(16)
	End If

Else

	EditUser.Level=3
	EditUser.Name=""
	EditUser.Email="null@null.com"
	EditUser.HomePage=bloghost

End If

BlogTitle=ZC_MSG070

'为1号输出输口准备的Action接口
'plugin node
For Each sAction_Plugin_EditUser_Form in Action_Plugin_EditUser_Form
	If Not IsEmpty(sAction_Plugin_EditUser_Form) Then Call Execute(sAction_Plugin_EditUser_Form)
Next

%>
<!--#include file="admin_header.asp"-->
<script type="text/javascript" src="../script/md5.js"></script>
<!--#include file="admin_top.asp"-->
			<div id="divMain">
<%	Call GetBlogHint()	%>
<div class="divHeader2"><%=ZC_MSG242%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_UserEdt_SubMenu & "</div>"
%>
<div id="divMain2">
<form id="edit" name="edit" method="post" action="">
<%
	Dim i
	Response.Write "<p><span class='title'>"& ZC_MSG249 &":</span><input id=""edtID"" name=""edtID""  type=""hidden"" value="""& EditUser.ID &""" />"
	Response.Write "<input id=""edtLevel"" name=""edtLevel"" type=""hidden"" value="""& EditUser.Level &""" /> <select "&IIF(CheckRights("root"),"","disabled=""disabled""")&" size=""1"" id=""cmbUserLevel"" onchange=""edtLevel.value=this.options[this.selectedIndex].value"">"
	Dim UserLevel
	i=1
	If EditUser.Level<>1 Then i=BlogUser.Level
	For Each UserLevel in ZVA_User_Level_Name
		Response.Write "<option value="""& i &""" "
		If i=EditUser.Level Then Response.Write "selected=""selected"""
		Response.Write ">"& TransferHTML(ZVA_User_Level_Name(i),"[html-format]") &"</option>"
		i=i+1
		If i=5 Then Exit For
	Next
	Response.Write "</select></p>"
	Response.Write "<p><span class='title'>"& ZC_MSG003 &":</span><span class='star'>(*)</span><br/><input id=""edtName"" size=""50"" name=""edtName""  type=""text"" value="""& TransferHTML(EditUser.Name,"[html-format]") &""" /></p>"
	Response.Write "<p><span class='title'>"& ZC_MSG002 &":</span>"&IIF(EditUser.ID<>0,"","<span class='star'>(*)</span>")&"<br/><input id=""edtPassWord"" size=""50"" name=""edtPassWord""  type=""password"" value="""" />"

	Response.Write "</p>"
	Response.Write "<p><span class='title'>"& ZC_MSG237 &":</span>"&IIF(EditUser.ID<>0,"","<span class='star'>(*)</span>")&"<br/><input id=""edtPassWordRe"" size=""50"" name=""edtPassWordRe""  type=""password"" value="""" />"

	Response.Write "</p>"
	Response.Write "<p><span class='title'>"& ZC_MSG147 &":</span><br/><input id=""edtAlias"" size=""50"" name=""edtAlias""  type=""text"" value="""& TransferHTML(EditUser.Alias,"[html-format]") &""" /></p>"
	Response.Write "<p><span class='title'>"& ZC_MSG053 &":</span><span class='star'>(*)</span><br/><input id=""edtEmail"" size=""50"" name=""edtEmail""  type=""text"" value="""& TransferHTML(EditUser.Email,"[html-format]") &""" /></p>"
	Response.Write "<p><span class='title'>"& ZC_MSG054 &":</span><br/><input id=""edtHomePage"" size=""50"" name=""edtHomePage""  type=""text"" value="""& TransferHTML(EditUser.HomePage,"[html-format]") &""" /></p>"
	Response.Write "<p><span class='title'>"& ZC_MSG198 &":</span><br/><textarea "& IIF(BlogUser.Level>3,"readonly=""readonly""","") &" cols=""3"" rows=""6"" id=""edtIntro"" name=""edtIntro"" style=""width:600px;"">"& TransferHTML(EditUser.Intro,"[html-format]") &"</textarea></p>"

	Response.Write "<p><span class='title'>"&ZC_MSG188&":</span><br/><select style='width:310px;' class='edit' size='1' id='cmbTemplate' onchange='edtTemplate.value=this.options[this.selectedIndex].value'>"

	Dim aryFileList
	aryFileList=LoadIncludeFilesOnlyType("zb_users\theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY)

	If IsArray(aryFileList) Then
		Dim j,t
		j=UBound(aryFileList)
		For i=1 to j
			t=UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
			If Left(t,2)<>"B_" AND t<>"FOOTER" And t<>"HEADER" And t<>"SINGLE" And t<>"PAGE" Then
				If EditUser.GetDefaultTemplateName=t Then
					Response.Write "<option value="""&t&""" selected=""selected"">"&t&IIF(EditUser.TemplateName=""," ("&ZC_MSG187&")","")&"</option>"
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

	Response.Write "</select><input type='hidden' name='edtTemplate' id='edtTemplate' value='"&EditUser.TemplateName&"' />"
	Response.Write "</p>"

	'<!-- 1号输出接口 -->
	If Response_Plugin_EditUser_Form<>"" Then Response.Write "<div id=""divEditForm1"">"&Response_Plugin_EditUser_Form&"</div>"

	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG078 &""" id=""btnPost"" onclick='return checkUserInfo();' /></p>"
%>
</form>
</div>

			</div>
<script type="text/javascript">
<!-- 

	var str13="<%=ZC_MSG118%>";
	var str14="<%=ZC_MSG119%>";
	var str15="<%=ZC_MSG120%>";
	var str16="<%=ZC_MSG038%>";
	var str17="<%=ZC_MSG237%>";

	function checkUserInfo(){

		if(<%=EditUser.ID%>==0){

			document.getElementById("edit").action="../cmd.asp?act=UserCrt";

			if(document.getElementById("edtName").value==""){
				alert(str13);
				return false;
			}
			if(document.getElementById("edtPassWord").value==""){
				alert(str14);
				return false;
			}
			else{
				if((document.getElementById("edtPassWord").value).length<=7){
					alert(str16);
					return false;
				}
				if((document.getElementById("edtPassWord").value)!==(document.getElementById("edtPassWordRe").value)){
					alert(str17);
					return false;
				}
			}
			if(document.getElementById("edtEmail").value==""){
				alert(str15);
				return false;
			}

		}
		else{

			document.getElementById("edit").action="../cmd.asp?act=UserMod"

			if(document.getElementById("edtName").value==""){
				alert(str13);
				return false;
			}

			if(document.getElementById("edtEmail").value==""){
				alert(str15);
				return false;
			}

			if(document.getElementById("edtPassWord").value){
				if((document.getElementById("edtPassWord").value).length<=7){
					alert(str16);
					return false;
				}
				if((document.getElementById("edtPassWord").value)!==(document.getElementById("edtPassWordRe").value)){
					alert(str17);
					return false;
				}
			}
		};

		//if(document.getElementById("edtPassWord").value){
		//	document.getElementById("edtPassWord").value=MD5(document.getElementById("edtPassWord").value);
		//}

		return true;
	}

 -->
</script>
<script type="text/javascript">ActiveLeftMenu("aUserMng");</script>

<!--#include file="admin_footer.asp"-->
<% 
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
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

GetCategory()
GetUser()

Dim EditUser
Set EditUser=New TUser

If Not IsEmpty(Request.QueryString("id")) Then

	If EditUser.LoadInfoByID(Request.QueryString("id")) Then
		If (EditUser.ID<>BlogUser.ID) And (CheckRights("Root")=False) Then Call ShowError(6)
	Else
		Call ShowError(16)
	End If

Else

	EditUser.Level=4
	EditUser.Name=""
	EditUser.Email="null@null.com"

End If



BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG070

%>
<!--#include file="admin_header.asp"-->
<!--#include file="admin_top.asp"-->
			<div id="divMain">
<%	Call GetBlogHint()	%>
<div class="divHeader2"><%=ZC_MSG242%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_UserEdt_SubMenu & "</div>"
%>
<div id="divMain2">
<form id="edit" name="edit" method="post">
<%
	Dim i
	Response.Write "<p>"& ZC_MSG249 &":<input id=""edtID"" name=""edtID""  type=""hidden"" value="""& EditUser.ID &""" />"
	Response.Write "<input id=""edtLevel"" name=""edtLevel"" type=""hidden"" value="""& EditUser.Level &""" /><select size=""1"" id=""cmbUserLevel"" onchange=""edtLevel.value=this.options[this.selectedIndex].value"">"
	Dim UserLevel
	i=0
	If EditUser.Level<>1 Then i=2
	For Each UserLevel in ZVA_User_Level_Name
		Response.Write "<option value="""& i &""" "
		If i=EditUser.Level Then Response.Write "selected=""selected"""
		Response.Write ">"& TransferHTML(ZVA_User_Level_Name(i),"[html-format]") &"</option>"
		i=i+1
		If i=5 Then Exit For
	Next
	Response.Write "</select></p>"
	Response.Write "<p>"& ZC_MSG001 &":<br/><input id=""edtName"" size=""50"" name=""edtName""  type=""text"" value="""& TransferHTML(EditUser.Name,"[html-format]") &""" />(*)</p>"
	Response.Write "<p>"& ZC_MSG002 &":<br/><input id=""edtPassWord"" size=""50"" name=""edtPassWord""  type=""password"" value="""" />"
	If EditUser.ID=0 Then
	Response.Write "(*)"
	End If
	Response.Write "</p>"
	Response.Write "<p>"& ZC_MSG282 &":<br/><input id=""edtPassWordRe"" size=""50"" name=""edtPassWordRe""  type=""password"" value="""" />"
	If EditUser.ID=0 Then
	Response.Write "(*)"
	End If
	Response.Write "</p>"
	Response.Write "<p>"& ZC_MSG147 &":<br/><input id=""edtAlias"" size=""50"" name=""edtAlias""  type=""text"" value="""& TransferHTML(EditUser.Alias,"[html-format]") &""" /></p>"
	Response.Write "<p>"& ZC_MSG053 &":<br/><input id=""edtEmail"" size=""50"" name=""edtEmail""  type=""text"" value="""& TransferHTML(EditUser.Email,"[html-format]") &""" />(*)</p>"
	Response.Write "<p>"& ZC_MSG054 &":<br/><input id=""edtHomePage"" size=""100"" name=""edtHomePage""  type=""text"" value="""& TransferHTML(EditUser.HomePage,"[html-format]") &""" /></p>"
	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG078 &""" id=""btnPost"" onclick='return checkUserInfo();' /></p>"
%>
</form>
</div>

			</div>
<script>


	var str13="<%=ZC_MSG118%>";
	var str14="<%=ZC_MSG119%>";
	var str15="<%=ZC_MSG120%>";
	var str16="<%=ZC_MSG038%>";
	var str17="<%=ZC_MSG282%>";

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
				if((document.getElementById("edtPassWord").value).length<=5){
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

			document.getElementById("edit").action="../cmd.asp?act=UserEdt"

			if(document.getElementById("edtName").value==""){
				alert(str13);
				return false;
			}

			if(document.getElementById("edtEmail").value==""){
				alert(str15);
				return false;
			}

			if(document.getElementById("edtPassWord").value){
				if((document.getElementById("edtPassWord").value).length<=5){
					alert(str16);
					return false;
				}
				if((document.getElementById("edtPassWord").value)!==(document.getElementById("edtPassWordRe").value)){
					alert(str17);
					return false;
				}
			}
		};

		if(document.getElementById("edtPassWord").value){
			document.getElementById("edtPassWord").value=MD5(document.getElementById("edtPassWord").value);
		}
	}
</script>
<script type="text/javascript">ActiveLeftMenu("aUserMng");</script>

<!--#include file="admin_footer.asp"-->
<% 
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
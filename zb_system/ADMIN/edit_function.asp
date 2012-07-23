<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    
'// 开始时间:    
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
For Each sAction_Plugin_Edit_Link_Begin in Action_Plugin_Edit_Link_Begin
	If Not IsEmpty(sAction_Plugin_Edit_Link_Begin) Then Call Execute(sAction_Plugin_Edit_Link_Begin)
Next

'检查非法链接
Call CheckReference("")

'检查权限
If Not CheckRights("FunctionEdt") Then Call ShowError(6)

GetCategory()
GetUser()
GetFunction()

Dim EditFunction

If Not  (IsEmpty(Request.QueryString("id")) Or Request.QueryString("id")="") Then
	Set EditFunction=Functions(Request.QueryString("id"))
Else
	Set EditFunction=New TFunction
	EditFunction.FileName="function"&EditFunction.GetNewID
	EditFunction.HtmlID="divFunction"&EditFunction.GetNewID
End If



BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG347

%><!--#include file="admin_header.asp"-->
<!--#include file="admin_top.asp"-->
			<div id="divMain">
<div class="divHeader2"><%=ZC_MSG347%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_Function_SubMenu & "</div>"
%>

<div id="divMain2">
<% Call GetBlogHint() %>
<form id="edit" name="edit" method="post" action="../cmd.asp?act=FunctionSav">
<%
	Dim s
	s=EditFunction.Content
	s=Replace(s,"<#ZC_BLOG_HOST#>",ZC_BLOG_HOST)
	s=Replace(s,"</li>","</li>"&vbCrlf)
	s=TransferHTML(s,"[html-format]")

	Response.Write "<input id=""edtID"" name=""edtID""  type=""hidden"" value="""& EditFunction.ID &""" />"
	Response.Write "<p><span class='title'>"& ZC_MSG001 &"</span>:<br/><input type=""text"" id=""inpName"" name=""inpName"" value="""& EditFunction.Name &""" size=""40"" />(*)</p>"
	Response.Write "<p><span class='title'>"& ZC_MSG170 &"</span>:<br/><input type=""text"" id=""inpFileName"" name=""inpFileName"" value="""& EditFunction.FileName &""" size=""40"" />(*)</p>"
	Response.Write "<p><span class='title'>"& "HTML ID" &"</span>:<br/><input type=""text"" name=""intHtmlID"" value="""&  EditFunction.HtmlId &""" size=""40""  />(*)<br/>("&ZC_MSG351&")</p>"

	Response.Write "<p><span class='title'>"& ZC_MSG061 &"</span>:<br/>"
	Response.Write "<input name=""intFtype"" type=""radio"" value=""div"" "&IIF(EditFunction.Ftype="div","checked=""checked""","")&" onclick=""$('#pMaxLi').css('display','none');"" />DIV &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""radio"" name=""intFtype"" value=""ul"" "&IIF(EditFunction.Ftype<>"div","checked=""checked""","")&" onclick=""$('#pMaxLi').css('display','block');"" />UL"
	Response.Write "</p>"
	Response.Write "<p id=""pMaxLi"" "&IIF(EditFunction.Ftype="div","style='display:none;'","")&"><span class='title'>"& ZC_MSG348 &"</span>:<br/><input type=""text"" name=""inpMaxLi"" value="""& EditFunction.MaxLi &""" size=""40""  />("&ZC_MSG350&")</p>"
	
	Response.Write "<p><span class='title'>"& ZC_MSG090 &":</span><br/><textarea name=""inpContent"" id=""inpContent"" onchange=""GetActiveText(this.id);"" onclick=""GetActiveText(this.id);"" onfocus=""GetActiveText(this.id);"" cols=""80"" rows=""12"">"&s&"</textarea></p>"

	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='return checkCateInfo();' /></p>"
%>
</form>
</div>

</div>
<script type="text/javascript">ActiveLeftMenu("aFunctionMng");</script>
<!--#include file="admin_footer.asp"-->
<% 
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
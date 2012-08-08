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
'// 开始时间:    ‎2012‎年‎7‎月‎23‎日
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

GetFunction()

Dim EditFunction

If Not (IsEmpty(Request.QueryString("id")) Or Request.QueryString("id")="") Then
	Set EditFunction=Functions(Request.QueryString("id"))
Else
	Set EditFunction=New TFunction
	EditFunction.FileName="function"&EditFunction.GetNewID
	EditFunction.HtmlID="divFunction"&EditFunction.GetNewID
	EditFunction.Order=EditFunction.GetNewOrder
End If



BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG144

%><!--#include file="admin_header.asp"-->
<!--#include file="admin_top.asp"-->
			<div id="divMain">
<div class="divHeader2"><%=ZC_MSG144%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_Function_SubMenu & "</div>"
%>

<div id="divMain2">
<% Call GetBlogHint() %>
<form id="edit" name="edit" method="post" action="../cmd.asp?act=FunctionSav">
<%
	Dim s,t,u
	s=EditFunction.Content
	s=Replace(s,"<#ZC_BLOG_HOST#>",ZC_BLOG_HOST)
	s=Replace(s,"</li>","</li>"&vbCrlf)
	s=TransferHTML(s,"[html-format]")
	If EditFunction.IsSystem=True Then t="readonly=""readonly"""
	If EditFunction.IsSystem=True Then u="disabled=""disabled"""

	Response.Write "<input id=""inpID"" name=""inpID""  type=""hidden"" value="""& EditFunction.ID &""" />"
	Response.Write "<input id=""inpOrder"" name=""inpOrder""  type=""hidden"" value="""& EditFunction.Order &""" />"
	Response.Write "<input id=""inpSidebarID"" name=""inpSidebarID""  type=""hidden"" value="""& EditFunction.SidebarID &""" />"
	Response.Write "<p><span class='title'>"& ZC_MSG001 &":</span><span class='star'>(*)</span><br/><input type=""text"" id=""inpName"" name=""inpName"" value="""& EditFunction.Name &""" size=""40"" /></p>"
	Response.Write "<p><span class='title'>"& ZC_MSG170 &":</span><span class='star'>(*)</span><br/><input "&t&" type=""text"" id=""inpFileName"" name=""inpFileName"" value="""& EditFunction.FileName &""" size=""40"" /></p>"
	Response.Write "<p><span class='title'>"& "HTML ID" &":</span><span class='star'>(*)</span><br/><input "&t&" type=""text"" name=""inpHtmlID"" value="""&  EditFunction.HtmlId &""" size=""40""  /><br/>("&ZC_MSG137&")</p>"

	Response.Write "<p><span class='title'>"& ZC_MSG061 &":</span><br/>"
	Response.Write "<label><input "&u&" name=""inpFtype"" type=""radio"" value=""div"" "&IIF(EditFunction.Ftype="div","checked=""checked""","")&" onclick=""$('#pMaxLi').css('display','none');"" />&nbsp;DIV </label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<label><input "&u&"  type=""radio"" name=""inpFtype"" value=""ul"" "&IIF(EditFunction.Ftype<>"div","checked=""checked""","")&" onclick=""$('#pMaxLi').css('display','block');"" />&nbsp;UL</label>"
	Response.Write "</p>"
	Response.Write "<p id=""pMaxLi"" "&IIF(EditFunction.Ftype="div","style='display:none;'","")&"><span class='title'>"& ZC_MSG143 &":</span><br/><input type=""text"" name=""inpMaxLi"" value="""& EditFunction.MaxLi &""" size=""40""  />("&ZC_MSG140&")</p>"

	Response.Write "<p><span class='title'>"& ZC_MSG017 &":</span><br/><label><input id=""inpNoSidebar"" type=""checkbox"" "&IIf(EditFunction.SidebarID=0, "checked=""checked""","") & " />&nbsp;&nbsp;"&ZC_MSG074&"</label><br/>"
	
	Response.Write "<label><input id=""inpSidebar""  type=""checkbox"" "&IIf(EditFunction.InSidebar=True, "checked=""checked""","") & " />&nbsp;&nbsp;"  & ZC_MSG008 & "&nbsp;&nbsp;&nbsp;&nbsp;</label>"
	Response.Write "<label><input id=""inpSidebar2"" type=""checkbox"" "&IIf(EditFunction.InSidebar2=True,"checked=""checked""","") & " />&nbsp;&nbsp;"  & ZC_MSG008 & "2&nbsp;&nbsp;&nbsp;&nbsp;</label>"
	Response.Write "<label><input id=""inpSidebar3"" type=""checkbox"" "&IIf(EditFunction.InSidebar3=True,"checked=""checked""","") & " />&nbsp;&nbsp;"  & ZC_MSG008 & "3&nbsp;&nbsp;&nbsp;&nbsp;</label>"
	Response.Write "<label><input id=""inpSidebar4"" type=""checkbox"" "&IIf(EditFunction.InSidebar4=True,"checked=""checked""","") & " />&nbsp;&nbsp;"  & ZC_MSG008 & "4&nbsp;&nbsp;&nbsp;&nbsp;</label>"
	Response.Write "<label><input id=""inpSidebar5"" type=""checkbox"" "&IIf(EditFunction.InSidebar5=True,"checked=""checked""","") & " />&nbsp;&nbsp;"  & ZC_MSG008 & "5&nbsp;&nbsp;&nbsp;&nbsp;</label>"

	Response.Write "<br/>"&ZC_MSG232&"</p>"

	
	Response.Write "<p><span class='title'>"& ZC_MSG090 &":</span><br/><textarea name=""inpContent"" id=""inpContent"" onchange=""GetActiveText(this.id);"" onclick=""GetActiveText(this.id);"" onfocus=""GetActiveText(this.id);"" cols=""80"" rows=""12"">"&s&"</textarea></p>"

	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" /></p>"
%>
</form>
</div>

</div>
<script type="text/javascript">

function CountSideBarID(){

	if($("#inpNoSidebar").attr("checked")){
		$("#inpSidebarID").val(0);
	}else{
		var s=""
		if($("#inpSidebar5").attr("checked")){s=s+"1"}else{s=s+"0"}
		if($("#inpSidebar4").attr("checked")){s=s+"1"}else{s=s+"0"}
		if($("#inpSidebar3").attr("checked")){s=s+"1"}else{s=s+"0"}
		if($("#inpSidebar2").attr("checked")){s=s+"1"}else{s=s+"0"}
		if( $("#inpSidebar").attr("checked")){s=s+"1"}else{s=s+"0"}
		$("#inpSidebarID").val(new Number(s.valueOf()));
	}
};

$("#inpNoSidebar").click(function(){
	if($(this).attr("checked")){
		$("#inpSidebar").removeAttr("checked")
		$("#inpSidebar2").removeAttr("checked")
		$("#inpSidebar3").removeAttr("checked")
		$("#inpSidebar4").removeAttr("checked")
		$("#inpSidebar5").removeAttr("checked")
	}else{
		$("#inpSidebar").attr("checked","checked")
		//$("#inpSidebar2").attr("checked","checked")
		//$("#inpSidebar3").attr("checked","checked")
		//$("#inpSidebar4").attr("checked","checked")
		//$("#inpSidebar5").attr("checked","checked")
	}
	CountSideBarID();
});

$("#inpSidebar,#inpSidebar2,#inpSidebar3,#inpSidebar4,#inpSidebar5").click(function(){
	if($(this).attr("checked")){
		$("#inpNoSidebar").removeAttr("checked");
	}
	CountSideBarID();
});


</script>


<script type="text/javascript">ActiveLeftMenu("aFunctionMng");</script>
<!--#include file="admin_footer.asp"-->
<% 
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
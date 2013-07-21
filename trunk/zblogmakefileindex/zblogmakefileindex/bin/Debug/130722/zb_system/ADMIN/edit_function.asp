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
	EditFunction.Source="users"
	If Request.QueryString("source")<>"" Then
		EditFunction.Source=Request.QueryString("source")
	End If
End If



BlogTitle=ZC_MSG144

%>
<!--#include file="admin_header.asp"-->
<!--#include file="admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <% Call GetBlogHint() %>
          </div>
          <div class="divHeader2"><%=ZC_MSG144%></div>
          <%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_Function_SubMenu & "</div>"
%>
          <div id="divMain2">
            <form id="edit" name="edit" method="post" action="../cmd.asp?act=FunctionSav">
              <%
	Dim s,t,u
	s=EditFunction.Content
	s=Replace(s,"<#ZC_BLOG_HOST#>",BlogHost)
	s=Replace(s,"</li>","</li>"&vbCrlf)
	s=TransferHTML(s,"[html-format]")
	If EditFunction.IsSystem=True Or EditFunction.IsPlugin=True Or EditFunction.IsTheme=True Then t="readonly=""readonly"""
	If EditFunction.IsSystem=True Or EditFunction.IsPlugin=True Or EditFunction.IsTheme=True Then u="disabled=""disabled"""
	If EditFunction.ID=0 Then  u="":t=""

	Response.Write "<input id=""inpID"" name=""inpID""  type=""hidden"" value="""& EditFunction.ID &""" />"
	Response.Write "<input id=""inpOrder"" name=""inpOrder""  type=""hidden"" value="""& EditFunction.Order &""" />"
	Response.Write "<input id=""inpSidebarID"" name=""inpSidebarID""  type=""hidden"" value="""& EditFunction.SidebarID &""" />"
	Response.Write "<input id=""inpSource"" name=""inpSource""  type=""hidden"" value="""& EditFunction.Source &""" />"
	Response.Write "<p><span class='title'>"& ZC_MSG001 &":</span><span class='star'>(*)</span><br/><input type=""text"" id=""inpName"" name=""inpName"" value="""& TransferHTML(EditFunction.Name,"[html-format]") &""" size=""40"" />&nbsp;&nbsp;"&ZC_MSG298&":<input id=""inpIsHideTitle"" name=""inpIsHideTitle"" style="""" type=""text"" value="""& EditFunction.IsHideTitle &""" class=""checkbox""/></p>"
	Response.Write "<p><span class='title'>"& ZC_MSG170 &":</span><span class='star'>(*)</span><br/><input "&t&" type=""text"" id=""inpFileName"" name=""inpFileName"" value="""& EditFunction.FileName &""" size=""40"" /></p>"
	Response.Write "<p><span class='title'>"& "HTML ID" &":</span><span class='star'>(*)</span><br/><input type=""text"" name=""inpHtmlID"" value="""&  EditFunction.HtmlId &""" size=""40""  /><br/>("&ZC_MSG137&")</p>"

	Response.Write "<p><span class='title'>"& ZC_MSG061 &":</span><br/>"
	Response.Write "<label><input "&u&" name=""inpFtype"" type=""radio"" value=""div"" "&IIF(EditFunction.Ftype="div","checked=""checked""","")&" onclick=""$('#pMaxLi').css('display','none');"" />&nbsp;DIV </label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<label><input "&u&"  type=""radio"" name=""inpFtype"" value=""ul"" "&IIF(EditFunction.Ftype<>"div","checked=""checked""","")&" onclick=""$('#pMaxLi').css('display','block');"" />&nbsp;UL</label>"
	Response.Write "</p>"
	Response.Write "<p id=""pMaxLi"" "&IIF(EditFunction.Ftype="div","style='display:none;'","")&"><span class='title'>"& ZC_MSG143 &":</span><br/><input type=""text"" name=""inpMaxLi"" value="""& EditFunction.MaxLi &""" size=""40""  />("&ZC_MSG140&")</p>"

	Response.Write "<p><span class='title'>"& ZC_MSG279 &":</span></p>"
	
	Response.Write "<label><input id=""viewtype1"" name=""inpViewType"" value="""" type=""radio"" "&IIf(EditFunction.ViewType="" Or EditFunction.ViewType="auto", "checked=""checked""","") & " />&nbsp;&nbsp;"& ZC_MSG280 &"&nbsp;&nbsp;&nbsp;&nbsp;</label>"
	Response.Write "<label><input id=""viewtype2"" name=""inpViewType"" value=""js"" type=""radio"" "&IIf(EditFunction.ViewType="js", "checked=""checked""","") & " />&nbsp;&nbsp;JavaScript&nbsp;&nbsp;&nbsp;&nbsp;</label>"
	Response.Write "<label><input id=""viewtype3"" name=""inpViewType"" value=""html"" type=""radio"" "&IIf(EditFunction.ViewType="html", "checked=""checked""","") & " />&nbsp;&nbsp;HTML&nbsp;&nbsp;&nbsp;&nbsp;</label>"

	Response.Write "<p><span class='title'>"& ZC_MSG017 &":</span></p><input id=""inpIsHidden"" name=""inpIsHidden"" style="""" type=""text"" value="""& EditFunction.IsDisplay &""" class=""checkbox""/><hr/>"
	
	Response.Write "<p><span class='title'>"& ZC_MSG090 &":</span><br/><textarea name=""inpContent"" id=""inpContent"" onchange=""GetActiveText(this.id);"" onclick=""GetActiveText(this.id);"" onfocus=""GetActiveText(this.id);"" cols=""80"" rows=""12"" " & IIF(InStr("|calendar|catalog|comments|previous|archives|authors|tags|statistics|","|"&EditFunction.FileName&"|")>0,"disabled=""disabled""","") & " >"&s&"</textarea></p>"

If Not (IsEmpty(Request.QueryString("id")) Or Request.QueryString("id")="") Then
	Response.Write ZC_MSG299 & "&lt;#CACHE_INCLUDE_"&UCase(EditFunction.FileName)&"#&gt;"
End If

	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" /></p>"

%>
            </form>
          </div>
        </div>
        <script type="text/javascript">
        ActiveLeftMenu("aFunctionMng");
        $("#inpContent").change(function(){
        	if(new RegExp("<"+"script","ig").test($(this).val())){
        		$("#viewtype3").click();
        	}
        })


        </script> 
        <!--#include file="admin_footer.asp"-->
<% 
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
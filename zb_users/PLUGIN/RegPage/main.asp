<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize
'检查非法链接
Call CheckReference("")

If CheckPluginState("RegPage")=False Then Call ShowError(48)
BlogTitle="注册管理"
Dim c
Set c=New TConfig
c.Load "RegPage"
If Request.QueryString("act")="save" Then c.Write "Level",Request.Form("defaultlevel"):c.Save:call SetBlogHint(True,True,False):Response.Redirect "main.asp"
Dim d
d=c.Read("Level")
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<script type="text/javascript">
function js(){
	var a=$("#defaultlevel").attr("value");
	if(a>4||a<1){return false}
	if(a!=4){return confirm("您选择的注册用户默认的权限为\n\n         "+$("#defaultlevel").children("option:eq("+(a-2)+")").html()+"\n\n这样给用户的权限太高，可能有未知的风险！\n点击确定继续，取消返回")}
}
$(document).ready(function(){ActiveLeftMenu("aPlugInMng")})

</script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader"><%=BlogTitle%></div>
<div class="SubMenu"><a href="main.asp"><span class="m-left m-now">注册管理</span></a></div>
<div id="divMain2">
<form action="main.asp?act=save" method="post" onsubmit="return js()">
<p>
<label for="defaultlevel">默认创建用户等级</label>
<select name="defaultlevel" id="defaultlevel">
<option value="2"<%=checked(2)%>><%=ZVA_User_Level_Name(2)%></option>
<option value="3"<%=checked(3)%>><%=ZVA_User_Level_Name(3)%></option>
<option value="4"<%=checked(4)%>><%=ZVA_User_Level_Name(4)%></option>
</select>
</p>
<p>
<input type="submit" name="button" id="button" value="提交" class="button" />
</p>
</form>

</div></div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%
function checked(a)
if cstr(a)=d then
response.write  "selected=""selected"""
elseif cstr(d)="" and cstr(a)="4" then 
response.write  "selected=""selected"""
end if
end function
%>
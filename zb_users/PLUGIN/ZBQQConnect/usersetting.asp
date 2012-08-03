<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
'
Call System_Initialize()
Call ZBQQConnect_Initialize()

Call CheckReference("")
If CheckPluginState("ZBQQConnect")=False Then Call ShowError(48)
If BlogUser.Level=5 Then Response.End

BlogTitle="ZBQQConnect-插件配置"
If Request.QueryString("act")="save" Then
	Dim c
	Set c=New TUser
	c.LoadInfoById BlogUser.Id
	c.Meta.SetValue "ZBQQConnect_a",IIf(Request.Form("a")="on",True,False)
	c.Edit BlogUser
	Response.Redirect "main.asp"
End If
Call SetBlogHint_Custom("修改配置需要重新登录")
%>
<%=ZBQQConnect_Config.Load("ZBQQConnect")%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<script type="text/javascript">
function showqk(){
	$("#how").toggleClass('hidden')
}
</script>

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader">ZBQQConnect</div>
          <div class="SubMenu"><%=ZBQQConnect_SBar(4)%></div>
          <form id="form1" name="form1" method="post" action="usersetting.asp?act=save">
            <div id="divMain2">


            <input name="a" id="a" type="checkbox" <%=d(BlogUser.Meta.GetValue("ZBQQConnect_a"))%> />
            <label for="a">评论同步到QQ空间</label>
			<p><input type="submit" class="button" value="提交"/></p>
            </div>
          </form>
        </div>
        <%
function d(v)
	d=iif(v="true"," checked=""checked"" ","")
end function
function e(s,b)
	e=iif(cint(s)=b," checked=""checked"" ","")
end function
%>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
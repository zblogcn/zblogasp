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
If BlogUser.Level<>1 Then
	Set ZBQQConnect_DB.objUser=BlogUser
	ZBQQConnect_DB.LoadInfo 2
Else
	ZBQQConnect_DB.ID=Request.QueryString("id")
	ZBQQConnect_DB.LoadInfo 1
End If
If Request.QueryString("act")="save" Then
	ZBQQConnect_DB.Email=TransferHTML(FilterSQL(Request.Form("a")),"[nohtml][""]")
	Call CheckParameter(ZBQQConnect_DB.ID,"int",0)
	If CheckRegExp(ZBQQConnect_DB.Email,"[email]") Then
		
		objConn.Execute "UPDATE [blog_Plugin_ZBQQConnect] SET [QQ_Eml]='"&ZBQQConnect_DB.Email&"' WHERE [QQ_ID]="&ZBQQConnect_DB.ID
		SetBlogHint_Custom  "ID为"&ZBQQConnect_DB.ID&"的绑定邮箱已经被修改为"&ZBQQConnect_DB.Email
	Else
		SetBlogHint_Custom "不符合E-Mail格式"
	End If
End If
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
          <div class="SubMenu"><%=ZBQQConnect_SBar(2)%></div>
          <form id="form1" name="form1" method="post" action="edit.asp?act=save&id=<%=ZBQQConnect_DB.ID%>">
            <div id="divMain2">
            <p>用户ID：<%=BlogUser.ID%></p>
            <p>QQ连接ID：<%=ZBQQConnect_DB.ID%></p>
            <p>用户名：<%=BlogUser.Name%></p>
            <p>
            电子邮箱(<a href="javascript:$('#a').attr('value','<%=BlogUser.email%>')">插入当前账户邮箱</a>)</p><p>
            <input type="text" id="a" value="<%=TransferHTML(ZBQQConnect_DB.Email,"[nohtml]")%>" name="a" style="width:100%"/>
            </p>
            <input type="submit" class="button" value="提交"/>
            </div>
          </form>
        </div>
        <%
function d(v)
	d=iif(v="True"," checked=""checked"" ","")
end function
function e(s,b)
	e=iif(cint(s)=b," checked=""checked"" ","")
end function
%>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
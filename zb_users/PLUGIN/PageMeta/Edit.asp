<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="function.asp" -->

<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_manage.asp" -->

<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("PageMeta")=False Then Call ShowError(48)
BlogTitle="PageMeta"


%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
      
    
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
<span class="m-left m-now"><a href="javascript:void(0)">编辑页</a> </span>
  </div>
  <div id="divMain2">
 <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
    <!--<table border="1" width="100%" cellspacing="1" cellpadding="1">
	<tr><td>Name</td><td>Value</td></tr>-->
    <%
	Call GetUser
	Call GetCategory
	Dim oA,j,k,a,f
	a=Request.QueryString("act")
	Call CheckParameter(a,"int",1)
	f=Array("","Article","Category","User","Tag","ArticleList","Comment","UploadFile")

	Execute "Set Oa=New T" & f(a)
	Oa.LoadInfoById request.QueryString("id")
	j=vbsunescape2(Oa.Meta.GetValue("pagemeta"))
	if j="null" then j=""
   %>        <p>修改格式：一个meta标签占一行，格式为name---value。如test---abcde</p>

    <form id="edit" name="edit" method="post" action="savedata.asp">
      <p>
         标题\名字
        <INPUT TYPE="text" Value="<%=g(Oa)%>" style="width:70%" name="path" id="path" >
      </p>
      <p>
        <textarea name="txaContent" id="txaContent" cols="45" rows="5" class="resizable" style="height:300px;width:100%"><%=TransferHTML(j,"[html]")%></textarea>
      </p>
      <input type="hidden" name="id" value="<%=oa.id%>"/>
      <input type="hidden" name="type" value="<%=a%>"/>
      <input class="button" type="submit" value="提交" id="btnPost"/>
      <input class="button" type="button" value="返回"  onclick="history.go(-1)"/>
      </p>
    </form>

    <!--    </table>-->
    <%Set oA=Nothing%>
  </div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->


<%
Call System_Terminate()

function g(a)
	on error resume next
	g=a.htmltitle
	if err.number>0 then
		err.clear
		g=a.HtmlName
		if err.number>0 then
			err.clear
			g=a.name
		end if
	end if
end function
%>

<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("duoshuo")=False Then Call ShowError(48)
BlogTitle="多说社会化评论"
Call DuoShuo_Initialize
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=duoshuo_SubMenu(duoshuo.get("act"))%></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("aPlugInMng");</script>
            <%
			If duoshuo.config.Read("short_name")="" Then
			%>
            <iframe id="duoshuo-remote-window" src="http://duoshuo.com/connect-site/?name=<%=Server.URLEncode(ZC_BLOG_TITLE)%>&description=<%=Server.URLEncode(ZC_BLOG_SUBTITLE)%>&url=<%=Server.URLEncode(ZC_BLOG_HOST)%>&siteurl=<%=Server.URLEncode(ZC_BLOG_HOST)%>&system_version=<%=BlogVersion%>&plugin_version=<%=Server.URLEncode(duoshuo.config.Read("ver"))%>&system=zblog&callback=<%=Server.URLEncode(BlogHost &"zb_users/plugin/duoshuo/noresponse.asp?act=callback")%>&user_key=<%=BlogUser.ID%>&user_name=<%=Server.URLEncode(BlogUser.Name)%>" style="border:0; width:100%; height:580px;"></iframe>
            <%
			Else
			Select Case duoshuo.get("act")
			Case "setting"
			%>
            <input name="" type="button" class="button" onclick="if(confirm('这是一个很占资源的过程，你确定要继续吗？')){location.href='noresponse.asp?act=export'}" value="导出评论至多说" />
            <%
			Case Else
			%>
			<iframe id="duoshuo-remote-window" src="http://<%=duoshuo.config.Read("short_name")%>.duoshuo.com/admin" style="width:100%; border:0;"></iframe>
			<%
			End Select
			End If
			%>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<script type="text/javascript">
ActiveLeftMenu("aCommentMng");
$(document).ready(function(){
	var iframe = $('#duoshuo-remote-window');
	resetIframeHeight = function(){
		iframe.height($(window).height() - iframe.offset().top - 70);
	};
	resetIframeHeight();
	$(window).resize(resetIframeHeight);
});
$('#duoshuo_manage').addClass('sidebarsubmenu1');
</script>

<%Call System_Terminate()%>

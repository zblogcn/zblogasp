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
If CheckPluginState("Wap")=False Then Call ShowError(48)
BlogTitle="WAP插件配置"

If Request.QueryString("act")="save" Then
	Call SaveWAPConfig2DB
End If

Set ConfigMetas=New TMeta
IsRunConfigs=False
Call GetConfigs()

Dim c
Set c = New TConfig
c.Load "Wap"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"></div>
          <div id="divMain2"> 
			<form  method="post" action="?act=save">
            <table width="100%">
			<tbody>
				<tr height="40" class="color1"><th width="30%">配置项</th><th width="69%">配置</th></tr>
				<tr><td><p align="left"><b> · 禁用WAP浏览模式(只保留PAD模式)</b></p></td><td><input type="text" name="WAP_DISABLE" id="WAP_DISABLE" value="<%=c.Read("WAP_DISABLE")%>" class="checkbox"></td></tr>
				<tr><td><p align="left"><b> · WAP文章列表单页显示文章数量</b></p></td><td><input type="text" name="WAP_DISPLAY_COUNT" id="WAP_DISPLAY_COUNT" value="<%=c.Read("WAP_DISPLAY_COUNT")%>"></td></tr>
				<tr><td><p align="left"><b> · WAP单页显示评论数量</b></p></td><td><input type="text" name="WAP_COMMENT_COUNT" id="WAP_COMMENT_COUNT" value="<%=c.Read("WAP_COMMENT_COUNT")%>"></td></tr>
				<tr><td><p align="left"><b> · WAP评论分页条显示条数</b></p></td><td><input type="text" name="WAP_PAGEBAR_COUNT" id="WAP_PAGEBAR_COUNT" value="<%=c.Read("WAP_PAGEBAR_COUNT")%>"></td></tr>
				<tr><td><p align="left"><b> · WAP相关文章数量</b></p></td><td><input type="text" name="WAP_MUTUALITY_LIMIT" id="WAP_MUTUALITY_LIMIT" value="<%=c.Read("WAP_MUTUALITY_LIMIT")%>"></td></tr>
				 <tr><td><p align="left"><b> · 打开WAP评论</b></p></td><td><input type="text" name="WAP_COMMENT_ENABLE" id="WAP_COMMENT_ENABLE" value="<%=c.Read("WAP_COMMENT_ENABLE")%>" class="checkbox"></td></tr>
				<tr><td><p align="left"><b> · WAP文章页全文显示模式</b></p></td><td><input type="text" name="WAP_DISPLAY_MODE_ALL" id="WAP_DISPLAY_MODE_ALL" value="<%=c.Read("WAP_DISPLAY_MODE_ALL")%>" class="checkbox"></td></tr>
				<tr><td><p align="left"><b> · 分页查看文章时单页字数</b></p></td><td><input type="text" name="WAP_SINGLE_SIZE" id="WAP_SINGLE_SIZE" value="<%=c.Read("WAP_SINGLE_SIZE")%>"></td></tr>
				<tr><td><p align="left"><b> · 显示分类导航</b></p></td><td><input type="text" name="WAP_DISPLAY_CATE_ALL" id="WAP_DISPLAY_CATE_ALL" value="<%=c.Read("WAP_DISPLAY_CATE_ALL")%>" class="checkbox"></td></tr>
				<tr><td><p align="left"><b> · 数字分页条模式</b></p></td><td><input type="text" name="WAP_DISPLAY_PAGEBAR_ALL" id="WAP_DISPLAY_PAGEBAR_ALL" value="<%=c.Read("WAP_DISPLAY_PAGEBAR_ALL")%>" class="checkbox"></td></tr>

			</tbody></table>
			<hr/>
			<p>
                <input type="submit" value="提交">
              </p>
			  </form>
          </div>
        </div>
		
		<script type="text/javascript">
			ActiveLeftMenu("aPlugInMng");
			$(".SubMenu a[href*='zb_users/plugin/wap/main.asp']").find("span").addClass("m-now")
		</script> 
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>

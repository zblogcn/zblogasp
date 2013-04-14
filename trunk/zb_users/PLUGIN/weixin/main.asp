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
If CheckPluginState("weixin")=False Then Call ShowError(48)
BlogTitle="微信搜索"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=weixin_SubMenu(0)%></div>
          <div id="divMain2"> 
			<form id="form1" name="form1" method="post">
			
			<table width="100%" style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
		  <tr>
			<th width='10%'><p align="center">设置项目</p></th>
			<th width='5%'><p align="center">设置内容</p></th>
			<th width='50%'><p align="left">设置说明</p></th>
			
		  </tr>
		  <tr>
			<td><b><label for="WelcomeStr"><p align="center">欢迎语</p></label></b></td>
			<td><p align="center"><textarea name="WelcomeStr" type="text" id="WelcomeStr"  style="width:218px;font-size:8px;" rows="3" cols="20" />欢迎关注《未寒博客》！！！<br/>您可发送“最新文章”来查看博客最新的10篇文章，或者直接发送关键词来搜索博客中已发表的文章。更多使用帮助请输入英文“help”或者数字“0”来查看。</textarea></p></td>
			<td><b><label for="WelcomeStr"><p align="left">&nbsp;&nbsp;设置被用户微信关注后发送的默认欢迎语</p></label></b></td>
		  </tr>
		  <tr>
			<td><b><label for="SearchNum"><p align="center">搜索结果数量</p></label></b></td>
			<td><p align="center"><input name="SearchNum" type="text" id="SearchNum"  style="width:80px;" value="10" /></p></td>
			<td><b><label for="SearchNum"><p align="left">&nbsp;&nbsp;设置微信关键词搜索返回的文章数量</p></label></b></td>
		  </tr>
		  <tr>
			<td><b><label for="LastPostNum"><p align="center">最新文章数量</p></label></b></td>
			<td><p align="center"><input name="LastPostNum" type="text" id="LastPostNum"  style="width:80px;" value="5" /></p></td>
			<td><b><label for="LastPostNum"><p align="left">&nbsp;&nbsp;设置微信中查看最新文章数量（最多为10篇）。</p></label></b></td>
		  </tr>
		  <tr>
			<td><b><label for="ShowMeta"><p align="center">文章查看方式</p></label></b></td>
			<td><p align="center">
				<select name="ShowMeta" id="ShowMeta" style="width:100px;">
					<option value="30" >文字版文章</option>
					<option value="90" >微信版页面章</option>
					<option value="180">博客主题模式</option>
				</select>
			</p></td>
			<td><b><label for="ShowMeta"><p align="left">&nbsp;&nbsp;设置微信中查看为重的方式</p></label></b></td>
		  </tr>
		</table>
		 <br />
		   <input name="" type="submit" class="button" value="保存" onclick='document.getElementById("form1").action="?act=Save";'/>
		  
			</form>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>

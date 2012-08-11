<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%
'On Error Resume Next
 %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_manage.asp" -->

<!-- #include file="../p_config.asp" -->

<% 
Call System_Initialize() 

'检查非法链接
Call CheckReference("") 

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 

If CheckPluginState("dztaotao")=False Then Call ShowError(48)

BlogTitle="dztaotao - 查看/操作淘淘" 
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

	<div id="divMain">
		<div class="divHeader"><%=BlogTitle%></div>
        <div id="ShowBlogHint"><%Call GetBlogHint()%></div>
			<div class="SubMenu">
				<a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/admin.asp?a=list&page=1"><span class="m-left">淘淘管理</span></a>
                <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/admin_cmt.asp?a=list&page=1"><span class="m-left">评论管理</span></a>
                <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/setting.asp"><span class="m-left">配置管理</span></a>
				<a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/help.asp"><span class="m-left m-now">帮助说明</span></a>
			</div>
	<div id="divMain2">
<form name="update_form1" id="update_form1" action="admin_cmt.asp?a=updatelist" method="post">
<table border="1" width="100%" cellpadding="2" cellspacing="0" bordercolordark="#f7f7f7" bordercolorlight="#cccccc">
<tr>
  <td height="30" bgcolor="#f7f7f7">这里是帮助中心</td>
  </tr>
<tr>
  <td><p>其实也没啥需要帮助的，只是感觉插件应该有个帮助说明，于是就象征性的整一个吧，</p>
    <p>从哪里开始呢？还是说说文件结构吧。</p>
    <p>滔滔用了独立的数据表，但未用独立的库，目的是为了数据独立，在数据量大的时候也能减轻一些服务在器压力，在安装滔滔的时候会自动建立新表到ZB的主数据库中，两个数据表分别是[dz_comment]和[dz_taotao]，感兴趣的同鞋可以研究一下。</p>
    <p>滔滔默认使用当前风格的文章内容页模板single.html，感觉不妥可以自行修改，文件就是index.asp，</p>
    <p>滔滔默认是没有调用功能的，但为了扩展性，这里有一个小小的调用，调用的内容是最新的10条滔滔信息，会在发表滔滔后自动生成在“zb_users/include/dztaotao.asp”，可以使用&lt;#CACHE_INCLUDE_DZTAOTAO#&gt;进行调用，调用出来的内容列表是&lt;UL&gt;....&lt;/UL&gt;</p></td>
</tr>
<tr>
  <td height="40">PS：估计会有朋友看到我博客上的滔滔有幻灯广告和BANNER广告，其实那就是传说中的收费版，既然挂广告能赚到钱，为啥不分点给大猪呢^_^<br />
保证价格童叟无欺，感性趣的可以联系QQ：38053383</td>
</tr>
</table>
</form>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>


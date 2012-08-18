<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8及以上的版本
'// 插件制作:  zblog管理员之家(www.zbadmin.com)
'// 备    注:   Mini缩略图插件代码
'// 最后修改：   2012/2/20
'// 最后版本:    0.1
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%

Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
'检查插件是否安装
If CheckPluginState("MiniTu")=False Then Call ShowError(48)

BlogTitle="Mini缩略图 for z-blog 2.0 后台设置"
MiniTu_Initialize

	Dim s_MiniTu_MiniImgWidth,s_MiniTu_MiniImgHeight
	s_MiniTu_MiniImgWidth=MiniTu_MiniImgWidth
	s_MiniTu_MiniImgHeight=MiniTu_MiniImgHeight

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

<style type="text/css">
<!--
.STYLE1 {
	color: #FF0000;
	font-weight: bold;
}
.STYLE2 {color: #009900}
.STYLE3 {
	color: #000000;
	font-weight: bold;
}
.STYLE4 {color: #FF0000}
-->
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
</div>
<div id="divMain2">
  <form id="edit" name="edit" method="post" action="save.asp">
    <p><b>关于[Mini缩略图 for Z-BLOG 2.0] </b></p>
    <p>提取文章中的第一张图片并生成缩略图, 使用 &lt;#article/intro/minitu#&gt; 标签.</p>
    <p>在YT标签里使用 &lt;#eval/MiniTu_Build(row(2),row(4),row(10))#&gt; 调用</p>
    <p><a href="http://www.zsxsoft.com" target="_blank">ZSXSOFT</a>将其升级到了for 2.0版本</p>
    <p>AspJPEG安装检查：<%ON error resume next:Dim a:Set a=Server.CreateObject("Persits.Jpeg"):Response.write IIf(Err.Number=0,"<font color='green'>可以使用本插件</font>","<font color='red'>不可使用本插件</font>"):Set a=Nothing%></p>
    <table width="90%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="144">缩略图宽度</td>
        <td><input name="MiniTu_MiniImgWidth" style="width:40px" type="text" value="<%=s_MiniTu_MiniImgWidth%>"/></td>
      </tr>
      <tr>
        <td>缩略图高度</td>
        <td><input name="MiniTu_MiniImgHeight" style="width:40px" type="text" value="<%=s_MiniTu_MiniImgHeight%>"/></td>
      </tr>
    </table>
    <p>
      <input type="submit" class="button" value=" 保存 " id="btnPost" />
    </p>
    <br/>
    <p>这是修改小飞龙的摘要产品图片插件ArticleIntroIllustration</p>
	<p>原插件地址：<a href="http://www.ecworker.com/resources/articleintroillustration.html">Zblog图文插件免费发布</a></p>
    <p>流年的修改记录：</p>
    <p>1、修改插件函数名，由原来的ArticleIntroIllustration（太长了 - -），改为MiniTu</p>
    <p>2、在老大（瑜廷）的帮助下，添加了后台自定义缩略图大小管理界面，正是你看到的这个，原插件需要修改Config.asp，比较麻烦</p>
    <p>3、通过与YT的结合，可以在侧栏使用缩略图功能，只是只能定义一种尺寸！</p>
    <p>更多zblog插件、教程、模板，请访问<a href="http://bbs.zbadmin.com">zblog管理员之家交流社区</a></p>
  </form>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->


<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("STACentre")=False Then Call ShowError(48)
BlogTitle="静态管理中心"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"> <a href="main.asp"><span class="m-left">配置页面</span></a><a href="list.asp"><span class="m-left">ReWrite规则</span></a> <a href="help.asp"><span class="m-right m-now">帮助</span></a> </div>
          <div id="divMain2"> 
          <table width="100%" border="1">
            <tr height="32">
              <td>1. 本插件只能实现Z-Blog的伪静态配置，而无法实现完全静态。</td>
            </tr>
            <tr height="32">
              <td>2. 伪静态规则全部会影响到整站，根目录与子目录想要同时使用伪静态必须手动修改Rewrite规则。</td>
            </tr>
            <tr height="32">
              <td>3. ISAPI_Rewrite 2.x和3.x是由Helicon出品的伪静态组件，均分Lite和Full两个版本。Lite版本免费，Full版本收费。URL Rewrite为微软官方出品的免费的应用于IIS7+的伪静态组件。</td>
            </tr>
            <tr height="32">
              <td>4. Windows Server 2003(r2)一般使用IIS6+ISAPI Rewrite 2.x或3.x ， Windows Server 2008(r2)及2012一般使用IIS7、7.5、8+URL Rewrite组件。虚拟主机用户请咨询你的空间商。<a href="http://www.dbshost.cn/" target="_blank">DBS主机</a>目前使用ISAPI Rewrite 2.x。</a></td>
            </tr>
            <tr height="32">
              <td>5. VPS、独立服务器若未安装组件，可点击右边链接下载。 <a href="http://www.helicontech.com/download-isapi_rewrite.htm" target="_blank">ISAPI_Rewrite 2.x </a> &nbsp; &nbsp; <a href="http://www.helicontech.com/download-isapi_rewrite3.htm" target="_blank">ISAPI_Rewrite 3.x </a> &nbsp; &nbsp; <a href="http://www.iis.net/downloads/microsoft/url-rewrite" target="_blank">URL Rewrite（需安装Microsoft Web Platform）</td>
            </tr>
            <tr height="32">
              <td>6. 其他伪静态组件（如<a href="http://iirf.codeplex.com/" target="_blank">Ionics Isapi Rewrite Filter</a>）未经测试，不保证生成的规则可以正常使用。</td>
            </tr>
          </table>
          <script type="text/javascript">ActiveLeftMenu("aPlugInMng");bmx2table();</script> 
          </div>
        </div>
        
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
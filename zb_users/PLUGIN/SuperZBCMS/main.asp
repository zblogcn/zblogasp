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
If CheckPluginState("SuperZBCMS")=False Then Call ShowError(48)
BlogTitle="超级Z-Blog"
SetBlogHint_Custom "本插件所做任何修改均可通过<a href='javascript:alert(""不先玩会儿吗？放心这只对你的后台有影响，没修改任何数据\n\n如您真要停用，请尽您的最大努力打开插件管理停用本插件，嗯！"")'>停用插件</a>恢复，使用前请注意关闭其它的浏览器标签页，否则可能会出现意外情况！"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
tr {
	height: 32px
}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=SuperZBCMS_SubMenu(0)%></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("aPlugInMng");</script>
            <form>
              <table width="100%" border="1">
                <tr>
                  <th>配置</th>
                  <th>论坛</th>
                  <th>电影CMS</th>
                  <th>Wiki</th>
                  <th>内容CMS</th>
                  <th>网址导航</th>
                </tr>
                <tr>
                  <td>使用模式</td>
                  <td><input type="radio" id="_a" name="_a"/>
                    <label for="_a"> 使用此模式</label></td>
                  <td><input type="radio" id="_b" name="_a"/>
                    <label for="_b"> 使用此模式</label></td>
                  <td><input type="radio" id="_c" name="_a"/>
                    <label for="_c"> 使用此模式</label></td>
                  <td><input type="radio" id="_d" name="_a"/>
                    <label for="_d"> 使用此模式</label></td>
                  <td><input type="radio" id="_e" name="_a"/>
                    <label for="_e"> 使用此模式</label></td>
                </tr>
                <tr>
                  <td>使用风格</td>
                  <td><label for="select"></label>
                    <select name="select" id="select">
                      <option>vBulletin风格</option>
                      <option>Discuz风格</option>
                      <option>PHPWind风格</option>
                      <option>PHPBB风格</option>
                      <option>DvBBS风格</option>
                      <option>百度贴吧风格</option>
                      <option>天涯社区风格</option>
                      <option>猫扑社区风格</option>
                  </select></td>
                  <td><label for="select2"></label>
                    <select name="select2" id="select2">
                      <option>MaxCMS风格</option>
                      <option>光线CMS风格</option>
                      <option>飞飞CMS风格</option>
                      <option>晴天CMS风格</option>
                      <option>搜狐影音风格</option>
                      <option>爱奇艺风格</option>
                      <option>优酷土豆风格</option>
                      <option>YouTube风格</option>
                  </select></td>
                  <td><select name="select3" id="select3">
                    <option>MediaWiki风格</option>
                    <option>HDWiki风格</option>
                    <option>DokuWiki风格</option>
                    <option>Confluence风格</option>
                    <option>百度百科风格</option>
                  </select></td>
                  <td><label for="select4"></label>
                    <select name="select4" id="select4">
                      <option>DeDeCMS风格</option>
                      <option>ASPCMS风格</option>
                      <option>PHPCMS风格</option>
                      <option>帝国CMS风格</option>
                      <option>KESION风格</option>
                      <option>SiteWeaver风格</option>
                      <option>腾讯风格</option>
                      <option>新浪风格</option>
                      <option>网易风格</option>
                      <option>搜狐风格</option>
                  </select></td>
                  <td><label for="select5"></label>
                    <select name="select5" id="select5">
                      <option>hao123</option>
                      <option>2345</option>
                      <option>搜狗</option>
                      <option>114啦</option>
                      <option>金山</option>
                      <option>QQ</option>
                  </select></td>
                </tr>
                <tr>
                  <td rowspan="2">其他选项</td>
                  
                  <td>数据库引擎</td>
                  <td colspan="4"><label for="select6"></label>
                    <select name="select6" id="select6">
                      <option>Access</option>
                      <option>Microsoft SQL Server</option>
                      <option>MYSQL</option>
                      <option>Oracle</option>
                      <option>NoSQL</option>
                      <option>PostgreSQL</option>
                      <option>OceanBase</option>
                      <option>SQLite</option>
                      <option>MemSQL</option>
                  </select></td>
                </tr>
                <tr>
                  <td>辅语言</td>
                  <td colspan="4"><label for="select7"></label>
                    <select name="select7" id="select7">
                      <option>.net 2.0</option>
                      <option>.net 3.5</option>
                      <option>.net 4.0</option>
                      <option>PHP 5.3</option>
                      <option>NodeJS</option>
                  </select></td>
                </tr>
              </table>
              <p>
                <input type="button" class="button" name="button" id="button" value="立即下载需求组件" />
              </p>
              
            </form>
          
          </div>
        </div>
        <script type="text/javascript">
		$("#button").click(function(){$.get("result.asp",{data:$("[name=_a]:checked").attr("id")},function(){alert("愚人节快乐！");location.reload();})})
        </script>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>

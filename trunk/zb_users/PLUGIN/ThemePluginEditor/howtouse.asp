<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
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
If CheckPluginState("ThemePluginEditor")=False Then Call ShowError(48)
BlogTitle="主题插件生成器"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<script type="text/javascript">
function newtr(This){
	var m=window.prompt("请输入新文件名","xxxx.html");
	if(m==null) return false;
	$(This).parent().parent().before("<tr><td>"+m+"</td><td>"+"<select name=\"type_"+m+"\"><option value=\"1\" selected=\"selected\">HTML</option><option value=\"2\">文件</option></select></td><td><input type=\"text\" id=\""+m+"\" name=\"include_"+m+"\" value=\"\" style=\"width:98%\"/><input type=\"hidden\" id=\""+m+"_2\" name=\"new_"+m+"\"/></td></tr>");bmx2table();
}
$("head").append("<sc"+"ript src='http://test.zsxsoft.com/zb/zb_system/admin/ueditor/third-party/SyntaxHighlighter/shCore.js' type='text/javascript'></sc"+"ript>");
$("head").append("<link rel='stylesheet' type='text/css' href='http://test.zsxsoft.com/zb/zb_system/admin/ueditor/third-party/SyntaxHighlighter/shCoreDefault.css'/>");
$(document).ready(function(e) {
    SyntaxHighlighter.all();
});
</script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <h1>恭喜你，插件生成成功！</h1>
            <p>不过，插件生成后不是马上就可以用的。你还需要对主题进行修改。</p>
            <p>编辑plugin.xml，在最后面的&lt;/plugin&gt;之前加入下列代码：</p>
            <p><pre class="brush:xml;toolbar:false;">
       &lt;plugin&gt;
           &lt;path&gt;editor.asp&lt;/path&gt;
               &lt;include&gt;include.asp&lt;/include&gt;
           &lt;level&gt;1&lt;/level&gt;
       &lt;/plugin&gt;       
            </pre>
            <p>示例的XML代码：
  </p>
            </p>
            <p><pre class="brush:xml;toolbar:false;">
&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot; standalone=&quot;yes&quot;?&gt;
&lt;plugin version=&quot;2.0&quot;&gt;
    &lt;id&gt;default&lt;/id&gt;
    &lt;name&gt;Default主题&lt;/name&gt;
    &lt;url&gt;http://www.rainbowsoft.org/&lt;/url&gt;
    &lt;note&gt;Z-Blog的默认主题&lt;/note&gt;
    &lt;author&gt;
        &lt;name&gt;zx.asd&lt;/name&gt;
        &lt;email&gt;zxasd@rainbowsoft.org&lt;/email&gt;
        &lt;url&gt;http://www.zdevo.com/&lt;/url&gt;
    &lt;/author&gt;
    &lt;source&gt;
    &lt;name&gt;jiaojiao&lt;/name&gt;
    &lt;email&gt;luheou@126.com&lt;/email&gt;
    &lt;url&gt;http://imjiao.com/&lt;/url&gt;
    &lt;/source&gt;
    &lt;adapted&gt;Z-Blog 2.0&lt;/adapted&gt;
    &lt;version&gt;1.0&lt;/version&gt;
    &lt;pubdate&gt;2005-2-18&lt;/pubdate&gt;
    &lt;modified&gt;2012-8-12&lt;/modified&gt;
    &lt;description&gt;
    &lt;![CDATA[Z-Blog的默认主题.模板由zx制作,娇娇设计.]]&gt;
    &lt;/description&gt;
    &lt;price&gt;&lt;/price&gt;
    &lt;plugin&gt;
        &lt;path&gt;editor.asp&lt;/path&gt;
        &lt;include&gt;include.asp&lt;/include&gt;
        &lt;level&gt;1&lt;/level&gt;
    &lt;/plugin&gt;
&lt;/plugin&gt;
</pre>
           接着，您可以视实际情况对生成的插件进行修改，这里不再赘述。 </p>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>
<%

%>
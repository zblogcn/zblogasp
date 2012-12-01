<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Devo Or Newer
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    批量管理文章插件 - 跳转页
'// 最后修改：   2008-10-24
'// 最后版本:    1.4
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!-- #include file="config.asp" -->
<%


'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>4 Then Call ShowError(6) 

If CheckPluginState("BatchArticles")=False Then Call ShowError(48)

BlogTitle="Batch Articles"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

	<style>
		ol {line-height:220%;}
		ol li {margin:0 0 0 -18px;text-decoration: none;}
		b {color:Navy;font-weight:Normal;text-decoration: underline;}
		p {line-height:160%;}
	</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->


<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader">批量管理文章插件 - 免责声明条款</div>
<div class="SubMenu"><a href="help.asp"><span class="m-right">[帮助]</span></a></div>
<div id="divMain2">

<hr />
<p>通过此插件可以实现文章在不同分类下的批量移动及批量更改文章的类型. 在使用前请仔细阅读以下免责声明条款:</p>

<ol>
<li>
此"批理管理文章插件"(以下简称插件)由 <a href="http://haphic.com/" target="_blank">haphic</a> 编写, 使用时请保留相关信息. 原作者(haphic)对更改插件内容所产生的问题不承担责任.
</li>
<li>
此插件仅适用于 2.0 以上版本 , 而对于其它版本的Z-blog未做调试, 无法对兼容性作出担保.
</li>
<li>
此插件的安全性基于Z-blog相关内容, 对于发生在本插件中但由Z-blog导致的安全问题, 插件原作者(haphic)不承担责任, 且没有解决问题的义务.
</li>
<li>
正常情况下, 此插件对文章的管理都是可逆的, 如果您不小心进行了错误的操作, 只要有足够的耐心, 都可通过此插件还原. 恢复备份的数据库也是好办法. 如果您还是因此而弄乱了您的博客数据, 相关责任只能由您自己承担.
</li>
<li>
如果您启用了自定义静态目录或月光静态目录插件, 则认为您了解这些功能的工作方式, 对因此而产生的问题(如移动分类后原文章仍存在于原路径下等), 只能由您自行解决, 而您也要同时考虑到主机空间(占用,整洁), 及搜索引擎收录等多方面的因素. 对于使用批量移动而产生的主机空间占用及SEO等相关问题, 插件原作者(haphic)不承担责任, 且没有解决问题的义务.
</li>
<li>
对于您因批量删除文章所造成的数据不可逆问题, 插件原作者(haphic)不承担责任, 且没有解决问题的义务.
</li>
<li>
非本插件所导致的技术问题, 如网络, 服务器等, 插件原作者(haphic)不承担责任, 且没有解决问题的义务.
</li>
<li>
此插件虽经多种测试未发现异常, 但无法保证这些功能在所有情况下(如数据库巨大或文章巨多等)都能稳定执行. 因为批量修改影响范围较大, 请在使用此插件前备份您的数据库, 以防止意外出现. 当您翻过此页时, 则认为您已经有了数据库的备份. 对于未及时备份而由使用此插件产生的不可恢复之问题, 将为您解决插件本身的问题并向您表示歉意, 但对于丢失的数据, 插件原作者(haphic)不承担责任, 且没有解决问题的义务.
</li>
<li>
免责声明的条款可能随着版本变动而变动, 任何时候均以最新版本的条款为准.
</li>
<li>
此页面将会在您启动此插件时出现, 直到您选择今后不再提示您. 您只有同意以上声明才可以进入此插件, 因此当您开始执行此插件的功能时, 则认定您已经同意了以上声明并做好了准备.
</li>
<li>
您可以随时在右上角的"[帮助]"中查看更多信息.
</li>
</ol>

<hr />

<p style="text-align:center;">
<a href="articlelist.asp">[同意,并开始使用]</a> <a href="articleMng.asp?act=SkipWarning">[同意,但今后别再烦我!]</a> <a href="../../cmd.asp?act=PlugInDisable&name=BatchArticles">[我不同意,宁愿退出并停用此插件!]</a>
</p>

</div>
</div>
<script>

	//斑马线
	var tables=document.getElementsByTagName("ol");
	var b=false;
	for (var j = 0; j < tables.length; j++){

		var cells = tables[j].getElementsByTagName("li");

		for (var i = 0; i < cells.length; i++){
			if(b){
				cells[i].style.color="#333366";
				cells[i].style.background="#F1F4F7";
				b=false;
			}
			else{
				cells[i].style.color="#666699";
				cells[i].style.background="#FFFFFF";
				b=true;
			};
		};
	}

document.close();

</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>


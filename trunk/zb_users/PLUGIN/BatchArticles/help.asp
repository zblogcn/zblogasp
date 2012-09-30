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

%><!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

	<style>
		/*ol {line-height:220%;}
		ol li {margin:0 0 0 -18px;text-decoration: none;}
		*/b {color:Navy;font-weight:Normal;text-decoration: underline;}
		p {line-height:160%;}
	</style><!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader">批量管理文章插件 - 使用说明</div>
<div class="SubMenu"><a href="javascript:history.back(-1);"><span class="m-left">[返回]</span></a></div>
<div id="divMain2">

<hr />
<p>此页介绍了批量管理文章插件(以下简称插件)在使用中的注意事项, 页面下部还包含了一些对插件的设置. 如果有时间, 建议简单浏览一下, 并按照您的需要进行设置.</p>

<ol>

<li>
推荐在分辨率大于 1024*768 的情况下使用此插件, 否则显示效果很糟.
</li>

<li>
<b>插件用途</b>: 
使用此插件适合于将较多的文章移动分类, 更改类型, 划分用户, 修改Tag, 以满足特殊的要求. 但更改极少量文章时此插件可能并不适用.
</li>

<li>
<b>插件稳定性</b>: 
这个插件应该是稳定的, 经过理论上的推敲和近乎疯狂的测试并未发现问题. 尽管如此, 也仍然无法保证这些推敲和测试涵盖了所有情况. 因此强列建议在使用此插件做大量更改前备份数据库.
</li>


<li>
<b>同时使用自定义静态目录?</b><br />
您应该知道, 改变了自定义目录后原目录中的文件不会被删除. 实际上考虑到SEO的需要您也不应该立即删除它们. 批量移动文章及删除文章也是一样. 如果您的自定义目录中有"{%category%}", 那么将文章"A"从分类"1"移动到分类"2"后, "文件重建"只是在分类"2"中新建了文章"A", "索引重建"只是将站内指向文章"A"的链接由分类"1"改为分类"2", 但分类"1"中存在的文章"A"并不会被删除.
如果您有洁癖的需要, 可以在"文件重建"前删除"日志存放目录"中的全部文件. 但您要确保您考虑过了SEO因素. 相信您不会经常的更改分类, 就如同您不会经常的更改自定义目录一样.
</li>

<li>
<b>关于管理Tags?</b><br />
此功能用来精确的为指定文章添加指定的 Tag, 或删除指定文章中的指定 Tag. 因为各人使用 Tags 的习惯不同, 此功能作用有限, Tags 过多或使用 "朦胧流" Tags 者不易使用. 因此, 下面额外提供了对插件的一些设置:
<br />
<%If UseTagMng Then%>
<a href="articleMng.asp?act=DisableTagMng">[停用Tags管理]</a> 如果您的 Tags 过多, 会造成管理页加载过慢, 建议在不需要时停用 Tags 管理.
<%Else%>
<a href="articleMng.asp?act=EnableTagMng">[启用Tags管理]</a> 启用 Tags 管理后, 您将可以对选择的文章的 Tags 进行操作.
<%End If%>
<br />
<%If UseTagCloud Then%>
<a href="articleMng.asp?act=DisableTagCloud">[停用TagClout显示]</a> TagCloud 显示方式会造成管理页加载进一步变慢, 您可以在不必要时停用.
<%Else%>
<a href="articleMng.asp?act=EnableTagCloud">[启用TagCloud显示]</a> TagCloud 显示方式提供了一个直观的面板以抽取含有某 Tag 的文章.
<%End If%>
<br />
<%If UseTagHint Then%>
<a href="articleMng.asp?act=DisableTagHint">[停用TagHint方式]</a> 如果您认为 TagHint 方式示能给你带来方便且干扰了您的操作, 您可以关闭之.
<%Else%>
<a href="articleMng.asp?act=EnableTagHint">[启用TagHint方式]</a> TagHint 方式的作用是: 当您抽取了含有某 Tag 的文章后, 自动将 "删除该Tag" 这种操作方式选中.
<%End If%>
</li>

<li>
<b>关于批量删除?</b><br />
应广大使用者的要求加入了批量删除功能. 批量删除会删除Tags关连, 分类关联,文章静态文件及摘要缓存文件等, 但不包括: 1, 对启用自定义静态目录后根据文章创建的目录; 2, 对数据库进行压缩修复以减少空间占用. 这两点并不影响博客的使用, 如果您有完美主义的需要, 请自行用 FTP 和 Access 解决.
</li>

<li>
<b>关于免责声明:</b><br />
《免责声明条款》在使用插件前已经要求您查看过了, 您仍可以<a href="warning.asp">再次查看《免责声明》</a> 还可以无聊的选择在启动插件时<a href="articleMng.asp?act=ShowWarning">恢复提示《免责声明》</a>.
</li>

</ol>

<p>此页说明可能随需要或插件版本变动而变动, 变动之后将随新版本发布, 对于此前的用户, 不会(也无法)另行通知.</p>

<hr />

<p style="text-align:center;"><a href="javascript:history.back(1);">[返回]</a></p>

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


<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Pre Terminator 及以上版本, 其它版本的Z-blog未知
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    插件管理插件
'// 最后修改：   2008-6-28
'// 最后版本:    1.2
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="c_sapper.asp" -->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)

If CheckPluginState("PluginSapper")=False Then Call ShowError(48)

BlogTitle="Plugin Sapper"

%><!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
	<style>
		ul {list-style:Upper-Alpha;line-height:200%;}
		ol {line-height:220%;}
		ol li {margin:0 0 0 -18px;text-decoration: none;}
		b {color:Navy;font-weight:Normal;text-decoration: underline;}
		sup {color:Red;}
	</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
	<div class="divHeader">Plugin Sapper - 帮助说明页</div>
	<%Call SapperMenu("8")%>
<div id="divMain2">

<form id="edit" name="edit">


<p><strong>说明文档目录:</strong></p>
<ul>
<li>
<a href="#pluginintro">插件简介.</a>
</li>
<li>
<a href="#pluginlist">插件管理扩展面板说明.</a>
</li>
<li>
<a href="#editinfo">如何编辑插件信息.</a>
</li>
<li>
<a href="#pluginxml">关于插件信息文档 (Plugin.xml).</a>
</li>
<li>
<a href="#exportzpi">导出插件为 ZPI 插件安装包文件 (以下简称 ZPI 文件).</a>
</li>
<li>
<a href="#importzpi">从本地上传 ZPI 文件并导入插件.</a>
</li>
<li>
<a href="#restorzpi">管理保存在主机上的 ZPI 文件.</a>
</li>
<li>
<a href="#aboutzpi">关于 ZPI 文件 ( <u><b>Z</b></u>-Blog <u><b>P</b></u>lugin <u><b>I</b></u>nstallation Pack ).</a>
</li>
<li>
<a href="#installonline">"获取更多插件" (在线安装插件) 使用指南.</a>
</li>
<li>
<a href="#checkupdate">为插件查找可用的更新版本.</a>
</li>
</ul>

<ul>

<a name="pluginintro"></a><br />
<li><strong>插件简介:</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>
Plugin Sapper, (以下简称 PS), 提供了一些有关插件的辅助功能, 属于此插件功能的页面, 会在页面标题中看到 "Plugin Sapper" 的字样.
</li>
<li>
激活此插件后, 会在"插件管理"中多出此插件的菜单. 停用此插件后, 这些菜单会消失.
</li>
<li>
插件提供有三大类功能: 一, 管理插件(编辑查看插件信息, 导出插件为 ZPI 文件, 删除插件); 二, 从本地上传并导入插件(从本地上传 ZPI 文件并导入该文件中的插件); 三, 管理主机上的 ZPI 文件(从主机上的 ZPI 文件恢复插件到 Blog, 下载主机上的 ZPI 文件, 删除主机上的 ZPI 文件); 四, 在线安装插件.
</li>
</ol>


<a name="pluginlist"></a><br />
<li><strong>插件管理扩展面板说明:</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>
<b>插件列表</b>: 这里列出了所有装在 Plugin 目录下的插件, 为每个插件提供了简要的信息, 并在每个插件缩略图右上方提供了四个功能按扭.
</li>
<li>
<b><img src="images/update.gif" alt="↓"> 升级修复插件</b>: 用来重新安装覆盖该插件以实现升级和修复.
</li>
<li>
<b><img src="images/info.gif" alt="i"> 查看插件信息</b>: 点击可以查看该插件的详细信息.
</li>
<li>
<b><img src="images/edit.gif" alt="√"> 编辑插件信息</b>: 用来生成或编辑该插件的信息文档 (Plugin.xml).
</li>
<li>
<b><img src="images/export.gif" alt="↑"> 导出插件</b>: 将该插件导出成 ZPI 文件 (关于 ZPI 文件).
</li>
<li>
<b><img src="images/delete.gif" alt="×"> 删除插件</b>: 删除该插件 (位于 Plugin 目录下的该插件文件夹), 正在使用的插件无法删除.
</li>
<!--
<li>
<b>导入插件</b>: 列表中最后一个插件, 被用作导入本地 ZPI 文件的表单.
</li>
-->
</ol>


<a name="editinfo"></a><br />
<li><strong>如何编辑插件信息:</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>
<b>进入编辑页面</b>: 在插件管理扩展面板中点击 <img src="images/edit.gif" alt="√"> 即可进入插件信息编辑页面. 在插件详细信息页面的下方也可找到 [编辑信息] 的菜单.
</li>
<li>
<b>编辑修改与全新生成</b>: 当该插件包含插件信息时, PS 会在进入编辑页时将其载入. 这时您看到的文本框内的文字为原有的插件信息. 当您更改并保存后, 原有的插件信息将被新信息覆盖; 而当该插件不包含插件信息时, 大部分文本框内的文字为空, 当您填写并保存后, PS 根据您填写的内容为您全新生成插件信息.
</li>
<li>
<b>插件信息和作者信息</b>: 按照提示填写即可, 作者信息如不想填写可以留空. <u>注意 <sup>notice</sup>: "适用版本"与"发布日期"的写法要标准, 不然系统可能无法识别. "插件版本", "发布日期", "最后修改日期" 三项关系到在线查到更新时的版本识别, 一定要正确填写.</u>
</li>
<li>
<b>插件的系统信息</b>: 此选项只适用于含有自带插件的插件, 如果插件不包含有自带插件, 请留空.
</li>
<li>
<b>插件信息的保存</b>: 当您点击按扭"完成编辑并保存信息"后, PS 会保存您当前填写的插件信息, 并在该插件目录下生成插件信息文档. 原有的插件信息将被覆盖.
</li>
</ol>


<a name="pluginxml"></a><br />
<li><strong>关于插件信息文档 (Plugin.xml):</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>
<b>什么是插件信息文档?</b> 每个插件都需要有插件信息以供后台的 "插件管理器" 使用. 这些信息以 XML 文档的形式保存于该插件目录下. 名称为 Plugin.xml.
</li>
<li>
<b>插件信息文档规范</b>: <a href="#" target="_blank">查看 Z-Blog 插件制作规范</a>
</li>
<li>
<b>如何得到标准的插件信息文档</b>: 在当前 PS 中使用 "编辑插件信息" 功能, 可以得到 (规范版本为 0.1 的) 标准插件信息文档.
</li>
</ol>


<a name="exportZPI"></a><br />
<li><strong>导出插件为 ZPI 文件<a href="#aboutzpi"> (什么是 ZPI 插件安装包文件?)</a>:</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>
<b>进入导出插件页面</b>: 在插件管理扩展面板中点击 <img src="images/export.gif" alt="↑"> 即可进入导出插件页面. 在插件详细信息页面的下方也可找到 [导出插件] 的菜单.
</li>
<li>
<b>编写 ZPI 文件的信息</b>: 进入导出页面后, 先要编辑 ZPI 文件的信息, 这些信息默认由插件信息中取得, 所以一般只要点击按扭 "确认信息并打包插件" 即可进入打包过程.
</li>
<li>
<b>发布与备份</b>: 用于发布插件, 指的是导出的文件将被放到资源中心下载, 这时要求主文件名必须为插件的 ID. 如仅用作备份插件, 则文件名随意, PS 会自生成不同的文件名.
</li>
<li>
<b>备份技巧</b> <sup>tip</sup>: 在选择导出类型为备份的同时, 可以修改一些信息, 如最后更新时间, 简介等, 这些信息将会在 <a href="XML_Restor.asp">"管理主机上的 ZPI 文件"</a> 中显示出来. 这相当于为这个备份做了备注.
</li>
<li>
<b>打包过程的执行</b>: 点击按扭 "确认信息并打包插件" 后, 打包程序将会启动. 将所选插件的所有文件打包进 ZPI 文件. 然后将 ZPI 文件保存在 PS 插件的 Export 目录下. 所以, <u>请确认此 Export 目录的存在, 不然打包无法完成</u>.
</li>
<li>
<b>下载 ZPI 文件</b>: 打包过程执行成功后, 会弹出下载页面, 这时您可以下载 ZPI 文件到本地. 另外, 所有导出在 Export 目录下的 ZPI 文件均可在 <a href="XML_Restor.asp">"管理主机上的 ZPI 文件"</a> 中下载.
</li>
<li>
<b>注意</b> <sup>notice</sup>: 如果你的浏览器无法直接下载, <u>请按照提示操作</u>. Opera 下载的 ZPI 文件扩展名可能为 XML, 并不影响使用. 但发布时请更改扩展名为 ZPI.
</li>
</ol>


<a name="importzpi"></a><br />
<li><strong>从本地上传 ZPI 文件并导入插件<a href="#aboutzpi"> (什么是 ZPI 插件安装包文件?)</a>:</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>
<b>从本地上传并导入插件</b>: 点击浏览, 从本地选择 ZPI 文件, 然后提交. PS 将会导入此 ZPI 文件中的插件, 并为您安装到博客上 (Plugin 目录下).
</li>
<li>
<b>是否覆盖提示</b>: 如果导入插件时发现该插件已存在于 Plugin 目录下. 会有 "是否覆盖掉已安装插件" 的提示.
</li>
</ol>


<a name="restorzpi"></a><br />
<li><strong>管理保存在主机上的 ZPI 文件<a href="#aboutzpi"> (什么是 ZPI 插件安装包文件?)</a>:</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>
<b>何为 "保存在主机上的 ZPI 文件" ?</b> 当您导出插件时, 无论是用作发布还是用作备份, 在 PS 插件中的 Exprot 目录下都会有相应名称的 ZPI 文件生成. 对于这些 ZPI 文件, PS 提供了后台管理功能, 如下:
</li>
<li>
<b><strong style="color:green;">←</strong> - 恢复</b>: 从主机上的 ZPI 文件恢复插件到 Blog, 即将该 ZPI 文件中的插件覆盖安装到 Plugin 目录下.
</li>
<li>
<b><strong style="color:blue;">↓</strong> - 下载</b>: 下载保存在主机上的该 ZPI 文件.
</li>
<li>
<b><strong style="color:red;">×</strong> - 删除</b>: 删除保存在主机上的该 ZPI 文件.
</li>
<li>
<b>是否覆盖提示</b>: 如果恢复插件时发现该插件已存在于 Plugin 目录下. 会有 "是否覆盖掉已安装插件" 的提示.
</li>
<li>
<b>注意</b> <sup>notice</sup>: 如果你的浏览器无法直接下载, <u>请按照提示操作</u>. 将鼠标悬停在链接上可看到提示. Opera 下载的 ZPI 文件扩展名可能为 XML, 并不影响使用. 但发布时请更改扩展名为 ZPI.
</li>
</ol>


<a name="aboutzpi"></a><br />
<li><strong>关于 ZPI 文件 ( <u><b>Z</b></u>-Blog <u><b>P</b></u>lugin <u><b>I</b></u>nstallation Pack ):</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>
<b>什么是 ZPI 文件?</b> ZPI 是 <u><b>Z</b></u>-Blog <u><b>P</b></u>lugin <u><b>I</b></u>nstallation Pack Document 的缩写. 意为 Z-Blog 插件安装包文件. 是由 Plugin Sapper 插件导出的一种 XML 格式的数据文件, 扩展名为 ZPI. Plugin Sapper 的导出导入插件功能, 在线安装功能等, 使用的都是这种文件.
</li>
<li>
<b>ZPI 文件的好处</b>: 使用 PS 的导入功能可以直接从本地的 ZPI 文件导入插件, 而不必使用 FTP 上传整个插件目录和文件. PS 还通过 ZPI 文件, 以及服务端的配合实现了直接从资源中心在线安装插件. 总之, ZPI 文件的出现方便了插件的备份和交流.
</li>
<li>
<b>如何得到 ZPI 文件</b>: 方法一, 可以使用 PS 的导出插件功能, 生成并下载 ZPI 文件; 方法二, 从资源中心的下载的插件安装包, 均为 ZPI 文件.
</li>
<li>
<b>Z-Wiki 上关于 ZPI 文件的解释</b>: <a target="_blank" href="http://wiki.rainbowsoft.org/doku.php?id=Plugins:pack">什么是 ZPI 插件安装包文件?</a>
</li>
</ol>


<a name="installonline"></a><br />
<li><strong>"获取更多插件" (在线安装插件) 使用指南:</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>
<b>浏览资源中心的插件</b>: 点击菜单 <a href="XML_List.asp">"获取更多插件"</a>, 等待服务器数据载入完成, 即可浏览资源中心中的插件. 其中, 您已经安装在博客内的插件, 会被打上 "已安装" 之类的标记, 以示区别.
</li>
<li>
<b>安装插件</b>: 点击每个插件缩略图下方的 "安装插件", 将会进入安装页面. 等待安装页面执行完成 - 这一过程所需要的时间要视网络状况和插件大小而定 - 即可在 "插件管理" 中找到该插件.
</li>
<li>
<b>覆盖提示</b>: 如果您的博客中已装有您正在安装的插件, 在安装时会有 "是否覆盖" 的提示. 如果选择 "继续安装", 则原有插件会被完全覆盖.
</li>
</ol>


<a name="checkupdate"></a><br />
<li><strong>为插件查找可用的更新版本.</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>
<b>查看插件的可用更新</b>: 点击菜单 <a href="XML_ChkVer.asp">"查看插件的可用更新"</a>, 即可看到已找到可用更新的插件.
</li>
<li>
<b>查找插件的可用更新 - 手动</b>: 在 "插件管理扩展面板" 页面, "查看插件的可用更新" 页面的下方, 均有 "查找更新" 的按扭. 点击即开始为您安装的插件(无论是否激活)查找可用更新版本.
</li>
<li>
<b>查找插件的可用更新 - 自动</b>: 当您或其它博客成员在后台活动的时候, PS 也会为您查找更新, 这种查找是自动的但是极为缓慢.
</li>
<li>
<b>插件更新提示</b>: 当插件有可用更新时, "插件管理" 页面, PS 中的 "插件管理扩展面板" 页面中均会有提示.
</li>
<li>
<b>不支持在线更新的插件</b>: "菠萝的海" 中没有收录的插件不具有在线更新的功能, 在查找更新后这些插件会被标示出来. "查看插件的可用更新" 页面中也提供了列出这些插件的功能.
</li>
<li>
<b>清除更新提示</b>: 点击 "查看插件的可用更新" 页面下方的 "清除更新提示" 按扭, "插件更新提示" 和 "不支持在线更新" 的提示均会被清除.
</li>
</ol>

</ul>
<p>
如果 PS 在使用过程中出错, 一般会有比较详细的错误提示. 有其它相关问题可 <a href="http://bbs.rainbowsoft.org/thread-19711-1-1.html" target="_blank">到论坛上提出</a> <a href="mailto:haphic@gmail.com">发我邮件</a> 或 <a href="http://haphic.com/blog/guestbook.asp" target="_blank">给我留言</a>.
</p>

</form>

<%Dim i : For i=0 To 26 : Response.Write "<br />" : Next%>
<a href="javascript:window.scrollTo(0,0);">[↑]</a>

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


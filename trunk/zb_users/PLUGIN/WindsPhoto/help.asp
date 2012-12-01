<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 spirit 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备    注:    WindsPhoto
'// 最后修改：   2011.8.22
'// 最后版本:    2.7.3
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
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

BlogTitle = "WindsPhoto"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
	<style>
		ul {list-style:Upper-Alpha;line-height:200%;}
		ol {line-height:220%;}
		ol li {margin:0 0 0 -18px;text-decoration: none;}
		b {color:Navy;font-weight:Normal;text-decoration: underline;}
		sup {color:Red;}
    </style>
</head>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<body>
<div id="divMain">
	<div class="divHeader">WindsPhoto 相册帮助</div>
		<div class="SubMenu">
<script type="text/javascript">ActiveLeftMenu("aWindsPhoto")</script>
        <a class="m-a-left m-now" href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_main.asp"><span>相册管理</span></a>
        <a class="m-a-left" href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_addtype.asp"><span>新建相册</span></a>
        <a class="m-a-left" href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_setting.asp"><span>系统设置</span></a>
        <a class="m-a-right" href="<%=ZC_BLOG_HOST%>zb_system/admin/admin.asp?act=PlugInMng"><span>退出</span></a>
        <a class="m-a-right" href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp"><span class="m-now">帮助说明</span></a>
        <a class="m-a-right" href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp#more"><span>更多功能</span></a>
		</div>
<div id="divMain2">
<p><strong>说明文档目录:</strong></p>
<div style="float:right;padding:5px;"><img src="images/logo.gif"></div>
<ul>
<li><a href="#pluginintro">插件简介.</a></li>
<li><a href="#source">功能介绍.</a></li>
<li><a href="#attention">注意事项.</a></li>
<li><a href="#qanda">常见问题.</a></li>
<li><a href="#more">更多功能.</a></li>
</ul>
<br />
<a name="pluginintro"></a>
<ul><li><strong>插件简介:</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
<li>WindsPhoto是本人同awinds基于朱朱相册的合作开发的Z-Blog相册插件。</li>
<li>新的相册依然沿用WindsPhoto的名字，版本号为2.x。</li>
<li>在保留原程序功能的基础上，最大程度的精简了代码，并引入了模板、特效、封面、RSS、静态化等概念。</li>
<li>使得WindsPhoto成为Z-Blog目前最佳的相册解决方案。</li>
<li>For 2.0版本为ZSXSOFT升级。</li>
</ol>

<a name="source"></a><br />
<li><strong>功能介绍:</strong> <a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
  <li>支持本地上传和远程图片，自动盗链国内门户提供的博客/空间/相册中的图片；</li>
  <li>缩略图和列表两种方式显示，上传时自动生成缩略图，可以设置缩略图大小、文字水印和LOGO水印；</li>
  <li>一键设置分类相册封面，自定义相册排序方式，正序或倒序；</li>
  <li>相册集成HighSlide、LightBox、GreyBox、ThickBox特效，并可以使用外部的Z-Blog插件；</li>  
  <li>停用时自动删除生成的文件，替换添加的导航，一键安装，一键卸载；</li>
  <li>可以添加不同的相册，方便的设置和修改相册简介，并且可以设置加密相册；</li>
  <li>DIY模板中，相册首页模版wp_index.html，分类页模版wp_album.html，如果存在则优先使用，其次读取tags.html模板；</li>
  <li>完整的Media RSS输出；</li>
  <li>贴图相册与插入相册图片功能，让WindsPhoto相册与Z-Blog紧密结合，成为Z-Blog不可或缺一部分；</li>
  <li>可设置单张图片的介绍，删除照片时，同时删除服务器中的图片和缩略图文件，不会产生垃圾文件。</li>
  </ol>
 
<a name="attention"></a><br />
<li><strong>注意事项:</strong><a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a></li>
<ol>
  <li>缩略图模式每行显示三个，列表显示时图片默认最大宽度为470px。</li>
  <li>相册内整合了四种特效，如果默认特效无效，请尝试使用其他的。</li>  
  <li>请不要尝试设置多个图片为封面，按正序排列，插件只认顺序第一个为封面。</li>
  <li>默认模板与本相册插件是完全配套的，你下载的其它模板或修改过的模板可能会让相册不正常。如果你不能确定原因，请换上默认模板试试。</li>
</ol>
<a name="qanda"></a><br />
<li><strong>常见问题:<a href="http://www.wilf.cn/post/WindsPhoto_Q_A.html" target="_blank">http://www.wilf.cn/post/WindsPhoto_Q_A.html</a></strong></li>
<a name="more"></a><br />
<li><strong>更多功能:</strong><a href="javascript:window.scrollTo(0,0);">[↑返回目录]</a><br />
您现在使用的是WindsPhoto普通版本，如果需要更新功能，可以<a href="http://www.wilf.cn/windsphoto_pro/" target="_blank">订购WindsPhoto Pro</a>。<br />
更多功能如下：
</li>
<ol>
  <li>更加简洁方便的后台设置，如直接设置水印位置等；</li>
  <li>单张图片采用网页展示，可视化编辑器填写详细的图片说明文字；</li>  
  <li>完善的图片评论功能，调用Z-Blog系统自带的反spam插件屏蔽黑词；</li>
  <li>幻灯片演示图片的slideshow效果，非常漂亮；</li>
  <li>域名绑定，作为插件，你可以绑定域名到插件目录；</li>
  <li>数码照片的EXIF信息显示；</li>
  <li>全局Media RSS及各分类Media RSS输出；</li>
  <li>Google Sitemap输出，更利于搜索引擎收录；</li>
  <li>免费的技术支持服务。</li>
</ol>
<p>更多精彩功能等你一一发掘，<a href="http://www.wilf.cn/windsphoto_pro/" target="_blank">现在就来订购吧！</a></p>
<p>如果在使用该插件中有任何问题 ，可以<a href="mailto:yangcheng@i0554.com">E-mail我</a> 或 <a href="http://www.wilf.cn/guestbook.asp" target="_blank">留言到我的Blog</a>。</p>

<%Dim i : For i=0 To 2 : Response.Write "<br />" : Next%>
<a href="javascript:window.scrollTo(0,0);">[↑]</a>

</div>
</div>


<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<id>			gbook_gravatar
<插件名称>		首页留言带gravatar头像调用插件
<摘要说明>		些插件只能在 z-blog 1.8 上使用, 启用这个插件以后,将会在用户每次留言后更新首页留言调用列表，而你只要在模板里像使用侧边栏日历等侧边栏目一样来使用
<version>		V1.0 beta for z-blog 1.8 
<作者信息>
<名称>		大猪
<网站>		http://www.dazhuer.cn/
<email>		myllop@gmail.com
有任何总是请到论坛或者我的博客上留言说明

使用方法:
	1:下载,解压,上传到plugin目录下
	2:进入后台,插件管理中,激活这个插件
	3:打开要显示文件排行的模板,在你想要显示的地方加上代码,详细方法见下面说明
	4:重建索引
	5:什么？看不到效果，那你随便留个言再刷新一下首页看看是不是就有了？
	
	第3步详细说明:
		就是和调整使用侧边栏完全一样,请看示例:
			<div class="function" id="divGuestComments"> <!-- 最近留言-->
			<h3><#ZC_MSG274#></h3>
			<ul>
			<#CACHE_INCLUDE_DZGUESTCOMMENTS#>
			</ul>
			</div>
			上面这一段是模板里系统自带的侧边栏:"最近留言"的代码
			
			下面这一段是需要你在风格css文件里添加的一段样式



			/*设置首页调用最新留言样式*/
			#divGuestComments ul li{height:46px; clear:both;} /*设置每条留言列表的高度*/
			.rc_avatar { padding:0 5px 0 0;FLOAT: left; }
			.rc_avatar img{border:1px #D4D4D4 solid; padding:1px;} /*gravatar头像边框样式*/
			.rc_comment{ line-height:13px !important;} /*留言内容行高*/
			.rc_comment A { COLOR: #2173af; height:13px !important;} /*链接样式*/
			.rc_comment A:hover {TEXT-DECORATION: underline}
			/*留言样式结束*



到这里就结束鸟，啥也没了，有问题就去我博客问吧！







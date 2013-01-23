var template_tags = {
	"filename": [{
		"filename": "default",
		"data": "首页主模板文件"
	},
	{
		"filename": "catalog",
		"data": "列表页模板文件((指由catalog.asp展现的页面。))"
	},
	{
		"filename": "b_article-multi",
		"data": "摘要文章模板"
	},
	{
		"filename": "b_article-istop",
		"data": "置顶文章模板((置顶文章会自动显示于首页及列表页中，无须标签调用))"
	},
	{
		"filename": "b_pagebar",
		"data": "页面底部分页条模板，可以改分页条样式"
	},
	{
		"filename": "page",
		"data": "独立页面模板，默认tags、search均用该模板"
	},
	{
		"filename": "b_function",
		"data": "侧栏模块模板  （[[themes:sidebar2.0|更多...]]）"
	},
	{
		"filename": "b_article-search-content",
		"data": "搜索结果内文模板"
	},
	{
		"filename": "b_article_comment_pagebar",
		"data": "评论分页条模板"
	},
	{
		"filename": "single",
		"data": "日志页主模板文件"
	},
	{
		"filename": "b_article-single",
		"data": "日志页文章模板"
	},
	{
		"filename": "b_article_nvabar_l",
		"data": "日志页面文章导航，显示“上一篇”日志链接((文章导航功能启用时才可以调用))  // （:!:Z-Blog 2.0默认删除，改为可选模板）//"
	},
	{
		"filename": "b_article_nvabar_r",
		"data": "日志页面文章导航，显示“下一篇”日志链接 // （:!:Z-Blog 2.0默认删除，改为可选模板）//"
	},
	{
		"filename": "b_article_trackback",
		"data": "每条引用通告显示模板"
	},
	{
		"filename": "b_article_mutuality",
		"data": "每条相关文章显示模板"
	},
	{
		"filename": "b_article_comment",
		"data": "每条评论内容显示模板((“私人文章”不显示评论及评论发送表单))"
	},
	{
		"filename": "b_article_commentpost",
		"data": "评论发送表单模板((“锁定文章”不显示评论发送表单))"
	},
	{
		"filename": "b_article_commentpost-verify",
		"data": "评论验证码显示模板((启用“验证码”功能选项才显示))"
	},
	{
		"filename": "b_article_tag",
		"data": "每个tag 的显示样式"
	},
	{
		"filename": "search",
		"data": "搜索页模板，显示搜索页面 // （ :!:Z-Blog 2.0默认删除，改为可选模板）//"
	},
	{
		"filename": "tags",
		"data": "标签页模板，显示TagCloud // （:!:Z-Blog 2.0默认删除，改为可选模板）//"
	},
	],
	"tags": [{
		"message": "文章基本数据",
		"file": ["b_article-istop.html", "b_article-multi.html", "b_article-single.html", "b_article-page.html"],
		"tags": [{
			"tag": "article/id",
			"note": "编号",
			"msg": "   "
		},
		{
			"tag": "article/url",
			"note": "链接",
			"msg": "  "
		},
		{
			"tag": "article/level",
			"note": "权限",
			"msg": "  "
		},
		{
			"tag": "article/title",
			"note": "标题",
			"msg": "  "
		},
		{
			"tag": "article/staticname",
			"note": "别名",
			"msg": "  "
		},
		{
			"tag": "article/intro",
			"note": "摘要",
			"msg": "  "
		},
		{
			"tag": "article/content",
			"note": "全文",
			"msg": "  "
		},
		{
			"tag": "article/posttime",
			"note": "时间",
			"msg": "更多时间格式请看[[#文章发布时间]]"
		},
		{
			"tag": "article/commnums",
			"note": "评论数",
			"msg": "  "
		},
		{
			"tag": "article/viewnums",
			"note": "浏览数",
			"msg": "推荐使用JS显示浏览数"
		},
		{
			"tag": "article/commentposturl",
			"note": "",
			"msg": ""
		},
		{
			"tag": "article/commentkey",
			"note": "",
			"msg": ""
		},
		{
			"tag": "article/commentrss",
			"note": "评论RSS",
			"msg": "  "
		},
		{
			"tag": "article/tagtoname",
			"note": "tags",
			"msg": "文本形式，可用于<head>区内作为关键词"
		},
		{
			"tag": "article/firsttagintro",
			"note": "第一个tag摘要",
			"msg": "将摘要设置图片链接，可做自动图文混排"
		}]
	},
	{
		"message": "文章分类数据",
		"file": ["b_article-istop.html", "b_article-multi.html", "b_article-single.html", "b_article-page.html"],
		"tags": [{
			"tag": "article/category/id",
			"note": "编号",
			"msg": ""
		},
		{
			"tag": "article/category/order",
			"note": "序号",
			"msg": ""
		},
		{
			"tag": "article/category/name",
			"note": "名称",
			"msg": ""
		},
		{
			"tag": "article/category/staticname",
			"note": "别名",
			"msg": ""
		},
		{
			"tag": "article/category/url",
			"note": "链接",
			"msg": ""
		},
		{
			"tag": "article/category/count",
			"note": "该分类下文章数",
			"msg": ""
		}]
	},
	{
		"message": "文章作者数据：",
		"file": ["b_article-istop.html", "b_article-multi.html", "b_article-single.html", "b_article-page.html"],
		"tags": [{
			"tag": "article/author/id",
			"note": "编号",
			"msg": ""
		},
		{
			"tag": "article/author/name",
			"note": "名称",
			"msg": ""
		},
		{
			"tag": "article/author/staticname",
			"note": "别名",
			"msg": ""
		},
		{
			"tag": "article/author/level",
			"note": "等级",
			"msg": ""
		},
		{
			"tag": "article/author/email",
			"note": "Email",
			"msg": ""
		},
		{
			"tag": "article/author/homepage",
			"note": "网站链接",
			"msg": ""
		},
		{
			"tag": "article/author/count",
			"note": "文章数",
			"msg": ""
		},
		{
			"tag": "article/author/url",
			"note": "链接",
			"msg": ""
		},
		]
	},
	{
		"message": "文章块级标签：",
		"file": ["b_article-istop.html", "b_article-multi.html", "b_article-single.html", "b_article-page.html"],
		"tags": [{
			"tag": "template:article_tag",
			"note": "tags",
			"msg": "链接形式，其定制模板为 b_article_tag.html"
		},
		{
			"tag": "template:article_mutuality",
			"note": "相关文章",
			"msg": "    b_article_mutuality.html"
		},
		{
			"tag": "template:article_trackback",
			"note": "引用列表",
			"msg": "    article_trackback.html"
		},
		{
			"tag": "template:article_comment",
			"note": "文章评论",
			"msg": "    article_comment.html"
		},
		{
			"tag": "template:article_commentpost",
			"note": "评论发送表单",
			"msg": "    article_commentpost.html"
		},
		{
			"tag": "template:article_navbar_l",
			"note": "上一篇文章",
			"msg": "  //（2.0中已弃用）// article_navbar_l.html"
		},
		{
			"tag": "template:article_navbar_r",
			"note": "下一篇文章",
			"msg": "  //（2.0中已弃用）// article_navbar_r.html"
		}]
	},
	{
		"message": "文章Tag数据：",
		"file": ["b_article_tag.html"],
		"tags": [{
			"tag": "article/tag/id",
			"note": "编号",
			"msg": ""
		},
		{
			"tag": "article/tag/name",
			"note": "名称",
			"msg": ""
		},
		{
			"tag": "article/tag/intro",
			"note": "摘要",
			"msg": ""
		},
		{
			"tag": "article/tag/count",
			"note": "文章数",
			"msg": ""
		},
		{
			"tag": "article/tag/url",
			"note": "链接",
			"msg": ""
		},
		{
			"tag": "article/tag/encodename",
			"note": "  ",
			"msg": ""
		}]
	},
	{
		"message": "相关文章数据：",
		"file": ["b_article_mutuality.html"],
		"tags": [{
			"tag": "article/mutuality/id",
			"note": "编号",
			"msg": ""
		},
		{
			"tag": "article/mutuality/url",
			"note": "链接",
			"msg": ""
		},
		{
			"tag": "article/mutuality/posttime",
			"note": "时间",
			"msg": ""
		},
		{
			"tag": "article/mutuality/name",
			"note": "文章名",
			"msg": ""
		}]
	},
	{
		"message": "文章评论数据：",
		"file": ["b_article_comment.html"],
		"tags": [{
			"tag": "article/comment/id",
			"note": "编号",
			"msg": ""
		},
		{
			"tag": "article/comment/count",
			"note": "序号",
			"msg": "  "
		},
		{
			"tag": "article/comment/name",
			"note": "名称",
			"msg": "评论者名称，下同"
		},
		{
			"tag": "article/comment/url",
			"note": "网址",
			"msg": "  "
		},
		{
			"tag": "article/comment/email",
			"note": "Email",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime",
			"note": "时间",
			"msg": "更多时间格式请看[[#评论发布时间]]"
		},
		{
			"tag": "article/comment/content",
			"note": "内容",
			"msg": ""
		},
		{
			"tag": "article/comment/authorid",
			"note": "作者编号",
			"msg": ""
		},
		{
			"tag": "article/comment/firstcontact",
			"note": "首要联系",
			"msg": "优先显示网址，若无网址则显示邮箱地址"
		},
		{
			"tag": "article/comment/emailmd5",
			"note": "Email的MD5码",
			"msg": ""
		},
		{
			"tag": "article/comment/urlencoder",
			"note": "经过加扰的URL链接",
			"msg": "防SPAM优先使用"
		},
		{
			"tag": "article/comment/parentid",
			"note": "父评论ID",
			"msg": ""
		},
		{
			"tag": "article/comment/avatar",
			"note": "头像",
			"msg": "格式为blogsite/zb_users/AVATAR/UserID.png，非注册用户的userID为0  "
		},
		{
			"tag": "article/loadviewcount",
			"note": "加载阅读数",
			"msg": "用于列表页展示  "
		},
		{
			"tag": "article/addviewcount",
			"note": "添加并显示阅读数",
			"msg": "用于文章页  "
		}]
	},
	{
		"message": "文章页“前后篇文章导航条”数据：",
		"file": ["b_article_nvabar_l.html", "b_article_nvabar_r.html"],
		"tags": [{
			"tag": "article/nav_l/url",
			"note": "上一篇文章链接",
			"msg": ""
		},
		{
			"tag": "article/nav_l/name",
			"note": "上一篇文章标题",
			"msg": ""
		},
		{
			"tag": "article/nav_r/url",
			"note": "下一篇文章链接",
			"msg": ""
		},
		{
			"tag": "article/nav_r/name",
			"note": "下一篇文章标题",
			"msg": ""
		}]
	},
	{
		"message": "文章发布时间：",
		"file": ["文章发布时间"],
		"tags": [{
			"tag": "article/posttime/longdate",
			"note": "2007年1月13日",
			"msg": "  "
		},
		{
			"tag": "article/posttime/shortdate",
			"note": "2007-1-13",
			"msg": "  "
		},
		{
			"tag": "article/posttime/longtime",
			"note": "15:31:13",
			"msg": "  "
		},
		{
			"tag": "article/posttime/shorttime",
			"note": "15:31",
			"msg": "  "
		},
		{
			"tag": "article/posttime/year",
			"note": "2007",
			"msg": "  "
		},
		{
			"tag": "article/posttime/month",
			"note": "1",
			"msg": "  "
		},
		{
			"tag": "article/posttime/monthname",
			"note": "January",
			"msg": "LANGUAGE文件中对应 ZVA_Month（1~12）全称"
		},
		{
			"tag": "article/posttime/monthnameabbr",
			"note": "Jan",
			"msg": "LANGUAGE文件中对应 ZVA_Month_Abbr（1~12）"
		},
		{
			"tag": "article/posttime/day",
			"note": "13",
			"msg": "  "
		},
		{
			"tag": "article/posttime/weekday",
			"note": "7",
			"msg": "  "
		},
		{
			"tag": "article/posttime/weekdayname",
			"note": "Saturday",
			"msg": "LANGUAGE文件中对应 ZVA_Week（1~7）全称"
		},
		{
			"tag": "article/posttime/weekdaynameabbr",
			"note": "sat",
			"msg": "LANGUAGE文件中对应 ZVA_Week_Abbr（1~7）"
		},
		{
			"tag": "article/posttime/hour",
			"note": "15",
			"msg": "  "
		},
		{
			"tag": "article/posttime/minute",
			"note": "31",
			"msg": "  "
		},
		{
			"tag": "article/posttime/second",
			"note": "13",
			"msg": "  "
		}]
	},
	{
		"message": "评论发布时间：",
		"file": ["评论发布时间"],
		"tags": [{
			"tag": "article/comment/posttime/longdate",
			"note": "2007年1月13日",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime/shortdate",
			"note": "2007-1-13",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime/longtime",
			"note": "15:31:13",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime/shorttime",
			"note": "15:31",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime/year",
			"note": "2007",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime/month",
			"note": "1",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime/monthname",
			"note": "January",
			"msg": "LANGUAGE文件中对应 ZVA_Month（1~12）全称"
		},
		{
			"tag": "article/comment/posttime/monthnameabbr",
			"note": "Jan",
			"msg": "LANGUAGE文件中对应 ZVA_Month_Abbr（1~12）"
		},
		{
			"tag": "article/comment/posttime/day",
			"note": "13",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime/weekday",
			"note": "7",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime/weekdayname",
			"note": "Saturday",
			"msg": "LANGUAGE文件中对应 ZVA_Week（1~7）全称"
		},
		{
			"tag": "article/comment/posttime/weekdaynameabbr",
			"note": "sat",
			"msg": "LANGUAGE文件中对应 ZVA_Week_Abbr（1~7）"
		},
		{
			"tag": "article/comment/posttime/hour",
			"note": "15",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime/minute",
			"note": "31",
			"msg": "  "
		},
		{
			"tag": "article/comment/posttime/second",
			"note": "13",
			"msg": "  "
		}]
	},
	{
		"message": "分类数据",
		"file": ["catalog.html", "tags.html"],
		"tags": [{
			"tag": "articlelist/category/id",
			"note": "分类ID",
			"msg": ""
		},
		{
			"tag": "articlelist/category/name",
			"note": "分类名",
			"msg": ""
		},
		{
			"tag": "articlelist/category/order",
			"note": "分类序号",
			"msg": ""
		},
		{
			"tag": "articlelist/category/count",
			"note": "分类下文章数",
			"msg": ""
		},
		{
			"tag": "articlelist/category/url",
			"note": "分类链接",
			"msg": ""
		},
		{
			"tag": "articlelist/category/staticname",
			"note": "分类静态别名，默认为分类别名，别名为空时则为分类名",
			"msg": ""
		}]
	},
	{
		"message": "作者数据",
		"file": ["catalog.html", "tags.html"],
		"tags": [{
			"tag": "articlelist/author/id",
			"note": "作者ID",
			"msg": ""
		},
		{
			"tag": "articlelist/author/name",
			"note": "作者名",
			"msg": ""
		},
		{
			"tag": "articlelist/author/level",
			"note": "作者等级",
			"msg": ""
		},
		{
			"tag": "articlelist/author/email",
			"note": "作者邮箱",
			"msg": ""
		},
		{
			"tag": "articlelist/author/homepage",
			"note": "作者网址",
			"msg": ""
		},
		{
			"tag": "articlelist/author/count",
			"note": "作者下文章数",
			"msg": ""
		},
		{
			"tag": "articlelist/author/url",
			"note": "作者页面地址",
			"msg": ""
		},
		{
			"tag": "articlelist/author/staticname",
			"note": "作者静态别名，默认为作者别名，别名为空时则为作者名",
			"msg": ""
		},
		{
			"tag": "articlelist/author/intro",
			"note": "作者简介",
			"msg": ""
		}]
	},
	{
		"message": "标签数据",
		"file": ["catalog.html", "tags.html"],
		"tags": [{
			"tag": "articlelist/tags/id",
			"note": "标签编号",
			"msg": ""
		},
		{
			"tag": "articlelist/tags/name",
			"note": "标签名",
			"msg": ""
		},
		{
			"tag": "articlelist/tags/intro",
			"note": "标签摘要",
			"msg": ""
		},
		{
			"tag": "articlelist/tags/count",
			"note": "标签下文章数",
			"msg": ""
		},
		{
			"tag": "articlelist/tags/url",
			"note": "标签页面地址",
			"msg": ""
		},
		{
			"tag": "articlelist/tags/encodename",
			"note": "URL编码后的标签名",
			"msg": ""
		}]
	},
	{
		"message": "日期数据",
		"file": ["catalog.html", "tags.html"],
		"tags": [{
			"tag": "articlelist/date/year",
			"note": "（年）2007",
			"msg": ""
		},
		{
			"tag": "articlelist/date/month",
			"note": "（月）1",
			"msg": ""
		},
		{
			"tag": "articlelist/date/day",
			"note": "（日）4",
			"msg": ""
		},
		{
			"tag": "articlelist/date/shortdate",
			"note": "2007-1-4",
			"msg": ""
		}]
	},
	{
		"message": "列表页分页条：",
		"file": ["catalog.html", "default.html"],
		"tags": [{
			"tag": "template:pagebar",
			"note": "分页条",
			"msg": "完整的分页条,后台设置显示条数"
		},
		{
			"tag": "template:pagebar_next",
			"note": "下一页",
			"msg": "<span class=\"pagebar-next\"><a href=NextUrl><span>« 更早的文章 \ </span></a></span>"
		},
		{
			"tag": "template:pagebar_previous",
			"note": "上一页",
			"msg": "<span class=\"pagebar-previous\"><a href=PrevUrl><span>之后的文章 » \ </span></a></span>"
		},
		{
			"tag": "articlelist/page/now",
			"note": "当前页码",
			"msg": "  "
		},
		{
			"tag": "articlelist/page/all",
			"note": "总页数",
			"msg": "  "
		},
		{
			"tag": "articlelist/page/count",
			"note": "每页显示条数",
			"msg": "*2.0  "
		}]
	},
	{
		"message": "分页条样式",
		"file": ["b_pagebar.html"],
		"tags": [{
			"tag": "pagebar/page/url",
			"note": "分页链接",
			"msg": ""
		},
		{
			"tag": "pagebar/page/number",
			"note": "分页码",
			"msg": ""
		}]
	},
	{
		"message": "侧栏模块",
		"file": ["b_function.html"],
		"tags": [{
			"tag": "function/id",
			"note": "模块ID",
			"msg": "  "
		},
		{
			"tag": "function/htmlid",
			"note": "模块HTML ID",
			"msg": "  "
		},
		{
			"tag": "function/name",
			"note": "模块名",
			"msg": "  "
		},
		{
			"tag": "function/content",
			"note": "模块内容",
			"msg": "  "
		},
		{
			"tag": "function/filename",
			"note": "模块引用文件名",
			"msg": "指[[#文件缓存]]目录下保存模块内容的文件名  "
		},
		]
	},
	{
		"message": "文件缓存",
		"file": ["all"],
		"tags": [{
			"tag": "CACHE_INCLUDE_文件名（全大写且不要后缀）",
			"note": "",
			"msg": "可以引用**系统目录**((“系统目录”在1.8版中指根目录，在2.0版中指zb_system目录))INCLUDE文件夹下的文本文件内容"
		},
		{
			"tag": "CACHE_INCLUDE_文件名_HTML",
			"note": "",
			"msg": "表示强制引用**系统目录**INCLUDE文件夹下的文件内容"
		},
		{
			"tag": "CACHE_INCLUDE_文件名_JS",
			"note": "",
			"msg": "表示强制JS方式动态引用**系统目录**INCLUDE文件夹下的文件"
		},
		{
			"tag": "TEMPLATE_INCLUDE_文件名（全大写且不要后缀）",
			"note": "",
			"msg": "可以引用zb_users/THEME/INCLUDE文件夹下的文本文件内容"
		},
		{
			"tag": "CACHE_INCLUDE_CATALOG",
			"note": "分类目录",
			"msg": "   "
		},
		{
			"tag": "CACHE_INCLUDE_AUTHORS",
			"note": "用户列表",
			"msg": "   "
		},
		{
			"tag": "CACHE_INCLUDE_TAGS",
			"note": "Tags",
			"msg": "从使用数多到少排列，最多显示50个 "
		},
		{
			"tag": "CACHE_INCLUDE_STATISTICS",
			"note": "站点统计",
			"msg": "    "
		},
		{
			"tag": "CACHE_INCLUDE_PREVIOUS",
			"note": "最近发表",
			"msg": "    "
		},
		{
			"tag": "CACHE_INCLUDE_COMMENTS",
			"note": "最新评论",
			"msg": "    "
		},
		{
			"tag": "CACHE_INCLUDE_GUESTCOMMENTS",
			"note": "最近留言",
			"msg": "指留言本中的最新留言  "
		},
		{
			"tag": "CACHE_INCLUDE_TRACKBACKS",
			"note": "最近引用",
			"msg": "    "
		},
		{
			"tag": "CACHE_INCLUDE_CALENDAR",
			"note": "日    历",
			"msg": "    "
		},
		{
			"tag": "CACHE_INCLUDE_CALENDAR_NOW",
			"note": "当前日历",
			"msg": "catalog.html中引用这个日历标签才能显示当前日期 "
		},
		{
			"tag": "CACHE_INCLUDE_ARCHIVES",
			"note": "文章归档",
			"msg": " "
		},
		{
			"tag": "CACHE_INCLUDE_NAVBAR",
			"note": "导 航 条",
			"msg": "   "
		},
		{
			"tag": "CACHE_INCLUDE_LINK",
			"note": "友情链接",
			"msg": "   "
		},
		{
			"tag": "CACHE_INCLUDE_FAVORITE",
			"note": "网站收藏",
			"msg": "   "
		},
		{
			"tag": "CACHE_INCLUDE_MISC",
			"note": "图标汇集",
			"msg": "   "
		},
		]
	}]
}
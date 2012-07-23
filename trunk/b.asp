<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

Dim t

Set t=new Tfunction
t.Name="导航栏"
t.FileName="navbar"
t.IsSystem=True
t.SidebarID=0
t.Order=1
t.Content="<li><a href=""<#ZC_BLOG_HOST#>"">首页</a></li><li><a href=""<#ZC_BLOG_HOST#>tags.asp"">标签</a></li><li><a href=""<#ZC_BLOG_HOST#>guestbook.asp"">留言本</a></li><li><a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=login"">管理</a></li>"
t.HtmlID="divNavBar"
t.Ftype="ul"
t.post


Set t=new Tfunction
t.Name="日历"
t.FileName="calendar"
t.IsSystem=True
t.SidebarID=1
t.Order=2
t.Content=""
t.HtmlID="divCalendar"
t.Ftype="div"
t.post




Set t=new Tfunction
t.Name="控制面板"
t.FileName="controlpanel"
t.IsSystem=True
t.SidebarID=1
t.Order=3
t.Content="<a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=login"">[<#ZC_MSG009#>]</a>&nbsp;&nbsp;<a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=vrs"">[<#ZC_MSG021#>]</a>"
t.HtmlID="divContorPanel"
t.Ftype="div"
t.post




Set t=new Tfunction
t.Name="网站分类"
t.FileName="catalog"
t.IsSystem=True
t.SidebarID=1
t.Order=4
t.Content=""
t.HtmlID="divCatalog"
t.Ftype="ul"
t.post


Set t=new Tfunction
t.Name="搜索"
t.FileName="searchpanel"
t.IsSystem=True
t.SidebarID=1
t.Order=5
t.Content="<form method=""post"" action=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=Search""><input type=""text"" name=""edtSearch"" id=""edtSearch"" size=""12"" /><input type=""submit"" value=""<#ZC_MSG087#>"" name=""btnPost"" id=""btnPost"" /></form>"
t.HtmlID="divSearchPanel"
t.Ftype="div"
t.post


Set t=new Tfunction
t.Name="最新评论及回复"
t.FileName="comments"
t.IsSystem=True
t.SidebarID=1
t.Order=6
t.Content=""
t.HtmlID="divComments"
t.Ftype="ul"
t.post




Set t=new Tfunction
t.Name="文章归档"
t.FileName="archives"
t.IsSystem=True
t.SidebarID=1
t.Order=7
t.Content=""
t.HtmlID="divArchives"
t.Ftype="ul"
t.post



Set t=new Tfunction
t.Name="站点统计"
t.FileName="statistics"
t.IsSystem=True
t.SidebarID=0
t.Order=8
t.Content=""
t.HtmlID="divStatistics"
t.Ftype="ul"
t.post




Set t=new Tfunction
t.Name="网站收藏"
t.FileName="favorite"
t.IsSystem=True
t.SidebarID=1
t.Order=9
t.Content="<li><a href=""http://bbs.rainbowsoft.org/"" target=""_blank"">ZBlogger社区</a></li><li><a href=""http://download.rainbowsoft.org/"" target=""_blank"">菠萝的海</a></li><li><a href=""http://t.qq.com/zblogcn"" target=""_blank"">Z-Blog微博</a></li>"
t.HtmlID="divFavorites"
t.Ftype="ul"
t.post




Set t=new Tfunction
t.Name="友情链接"
t.FileName="link"
t.IsSystem=True
t.SidebarID=1
t.Order=10
t.Content="<li><a href=""http://www.dbshost.cn/"" target=""_blank"" title=""独立博客服务 Z-Blog官方主机"">DBS主机</a></li><li><a href=""http://www.dutory.com/blog/"" target=""_blank"">Dutory官方博客</a></li>"
t.HtmlID="divLinkage"
t.Ftype="ul"
t.post



Set t=new Tfunction
t.Name="图标汇集"
t.FileName="misc"
t.IsSystem=True
t.SidebarID=1
t.Order=11
t.Content="<li><a href=""http://www.rainbowsoft.org/"" target=""_blank""><img src=""<#ZC_BLOG_HOST#>zb_system/image/logo/zblog.gif"" height=""31"" width=""88"" border=""0"" alt=""RainbowSoft Studio Z-Blog"" /></a></li><li><a href=""<#ZC_BLOG_HOST#>feed.asp"" target=""_blank""><img src=""<#ZC_BLOG_HOST#>zb_system/image/logo/rss-big-sq.png"" height=""48"" width=""48"" border=""0"" alt=""订阅本站的 RSS 2.0 新闻聚合"" /></a></li>"
t.HtmlID="divMisc"
t.Ftype="ul"
t.post




Set t=new Tfunction
t.Name="作者列表"
t.FileName="authors"
t.IsSystem=True
t.SidebarID=0
t.Order=12
t.Content=""
t.HtmlID="divAuthors"
t.Ftype="ul"
t.post




Set t=new Tfunction
t.Name="最近发表"
t.FileName="previous"
t.IsSystem=True
t.SidebarID=0
t.Order=13
t.Content=""
t.HtmlID="divPrevious"
t.Ftype="ul"
t.post



Set t=new Tfunction
t.Name="标签列表"
t.FileName="tags"
t.IsSystem=True
t.SidebarID=0
t.Order=14
t.Content=""
t.HtmlID="divTags"
t.Ftype="ul"
t.post



Call System_Terminate()

%>
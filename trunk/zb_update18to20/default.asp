<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../zb_users/c_option.asp" -->
<!-- #include file="../zb_system/function/c_function.asp" -->
<!-- #include file="../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../zb_system/function/c_system_base.asp" -->
<!-- #include file="../zb_system/function/c_system_plugin.asp" -->
<%

Dim username,password,userguid
Dim dbtype,dbpath,dbserver,dbname,dbusername,dbpassword


Dim zblogstep
zblogstep=Request.QueryString("step")

If ZC_DATABASE_PATH<>"" Or ZC_MSSQL_DATABASE<>"" Then
	zblogstep=0
End If

If zblogstep="" Then zblogstep=1

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
	<meta name="generator" content="Z-Blog <%=ZC_BLOG_VERSION%>" />
	<meta name="robots" content="nofollow" />
	<script language="JavaScript" src="../zb_system/script/common.js" type="text/javascript"></script>
	<script language="JavaScript" src="../zb_system/script/md5.js" type="text/javascript"></script>
	<link rel="stylesheet" rev="stylesheet" href="../zb_system/css/admin3.css" type="text/css" media="screen" />
	<title>Z-Blog 1.8 to 2.0 升级程序</title>
</head>
<body>
  <div class="setup"><form method="post" action="?step=<%=zblogstep+1%>">
<%

Select Case zblogstep
Case 0 Call Setup0
Case 1 Call Setup1
Case 2 Call Setup2
Case 3 Call Setup3
End  Select
%>
  </form></div>

<script language="JavaScript" type="text/javascript">
</script>
</body>
</html>
<%




Function Setup0()
%>
<dl>
<dd id="ddleft">
<img src='../zb_system/image/admin/update.png' alt='' />
<p>安装进度:<span><font color='#3d69aa'></font><font color='white'>██████████████████</font></span></p>
<p>升级选项及说明&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;»&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;升级结果</p>
</dd>
<dd id="ddright">
<div id='title'>安装提示</div>
<div id='content'>
通过配置文件的检验,您已经安装并配置好Z-Blog了,不能再重复使用升级程序.
</div>
<div id='bottom'>
<input type="button" name="next" onClick="window.location.href='<%=BlogHost%>'" id="netx" value="退出" />
</div>
</dd>
</dl>
<%
End Function












Function Setup1()
%>
<dl>
<dd id="ddleft">
<img src='../zb_system/image/admin/update.png' alt='' />
<p>安装进度:<span><font color='#3d69aa'>█████████</font><font color='white'>█████████</font></span></p>
<p><b>升级选项及说明</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;»&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;升级结果</p>
</dd>
<dd id="ddright">
<div id='title'>选择需要升级的数据库</div>
<div id='content'>
<input type="hidden" name="dbpath" id="dbpath" value="" />
<%

	Dim fso, f, f1, fc
	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FolderExists(BlogPath & "zb_users\data")=True Then

		Set f = fso.GetFolder(BlogPath & "zb_users\data")
		Set fc = f.Files
		For Each f1 in fc
			If Right(f1.name,4)=".mdb" Then
				Response.Write "<p><label>&nbsp;&nbsp;<input type='radio' name='files'/>&nbsp;&nbsp;<span>" & f1.name & "</span></label></p>"
			End If
		Next

	End If

%>
<p class='title'>升级说明</p>
<p>1.将Z-Blog 1.8目录中DATA目录下的文件转移至Z-Blog 2.0的zb_users\DATA目录下.</p>
<p>2.将Z-Blog 1.8目录中UPLOAD目录下的文件转移至Z-Blog 2.0的zb_users\UPLOAD目录下.</p>
<p>3.将Z-Blog 1.8目录中INCLUDE目录下的文件转移至Z-Blog 2.0的zb_users\INCLUDE目录下.</p>
<p>4.将Z-Blog 1.8目录中c_custom.asp和c_option.asp文件转移至Z-Blog 2.0的根目录下.</p>
<p>5.运行http://你的网站/zb_update18to20/进入升级程序.</p>
</div>
<div id='bottom'>
<input type="submit" name="next" id="netx" value="下一步" onClick="if($('#dbpath').val()==''){alert('选择一个需要升级的数据库.');return false;}" />
</div>
</dd>
</dl>
<script language="JavaScript" type="text/javascript">
$('label').live('click', function() {
  $('#dbpath').val($(this).find('span').html());
});
</script>
<%
End Function











Function Setup2()
On Error Resume Next
%>
<dl>
<dd id="ddleft">
<img src='../zb_system/image/admin/update.png' alt='' />
<p>安装进度:<span><font color='#3d69aa'>█████████████████</font><font color='white'></font></span></p>
<p><b>升级选项及说明</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;»&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>升级结果</b></p>
</dd>
<dd id="ddright">

<div id='title'>安装结果</div>
<div id='content'>
<%

dbpath=Request.Form("dbpath")


ZC_DATABASE_PATH="zb_users\data\" & dbpath

ZC_MSSQL_ENABLE=False




If OpenConnect()=False Then

	Response.Write("<script language=javascript>alert('数据库连接错误!');</script>")
	Response.Write("<script language=javascript>history.go(-1);</script>")
	Response.End

End If

Response.Cookies("password")=""
Response.Cookies("username")=""

Application.Contents.RemoveAll

%>
<%Call UpdateAccessTable():Response.Flush%><p>数据库表和数据升级成功!</p>
<%Call InsertOptions():Response.Flush%><p>默认配置数据导入成功!</p>
<%Call InsertFunctions():Response.Flush%><p>默认侧栏数据导入成功!</p>
<%Call LoadOldOptions():Response.Flush%><p>读取旧配置数据成功!</p>
<%Call LoadOldFunctions():Response.Flush%><p>读取旧侧栏模块数据成功!</p>
<%Call FixOldTotoro():Response.Flush%><p>调整旧Totoro评论成功!</p>
<%Call RevToComment():Response.Flush%><p>将1.8版的回复转换2.0版评论成功!</p>
<%Call SaveConfigs():Response.Flush%><p>配置文件c_option.asp保存成功!</p>
<%

%>

<p>Z-Blog 2.0升级成功了,现在您可以点击"完成"进入网站首页.</p>

</div>
<div id='bottom'>
<input type="button" name="next" onClick="window.location.href='<%=BlogHost%>'" id="netx" value="完成" />
</div>


</dd>
</dl>
<%
End Function






Function Setup3()
	Response.Redirect BlogHost
End Function







Function CheckUpdateDB(a,b)
	Err.Clear
	On Error Resume Next
	Dim Rs
	Set Rs=objConn.execute("SELECT "&a&" FROM "&b)
	Set Rs=Nothing
	If Err.Number=0 Then
	CheckUpdateDB=True
	Else
	Err.Clear
	CheckUpdateDB=False
	End If	
End Function


Function UpdateAccessTable()

	Dim s,t

	If Not CheckUpdateDB("[log_IsTop]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_IsTop] YESNO DEFAULT FALSE")
		objConn.execute("UPDATE [blog_Article] SET [log_IsTop]=0")
	End If

	If Not CheckUpdateDB("[log_Tag]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Tag] VARCHAR(255) default """"")
	End If

	If Not CheckUpdateDB("[tag_ID]","[blog_Tag]") Then
		objConn.execute("CREATE TABLE [blog_Tag] (tag_ID AutoIncrement primary key,tag_Name VARCHAR(255) default """",tag_Intro text default """",tag_ParentID int DEFAULT 0,tag_URL VARCHAR(255) default """",tag_Order int DEFAULT 0,tag_Count int DEFAULT 0)")
	End If

	If Not CheckUpdateDB("[coun_ID]","[blog_Counter]") Then
		objConn.execute("CREATE TABLE [blog_Counter] (coun_ID AutoIncrement primary key,coun_IP VARCHAR(20) default """",coun_Agent text default """",coun_Refer VARCHAR(255) default """",coun_PostTime TIME DEFAULT Now())")
	End If

	If Not CheckUpdateDB("[key_ID]","[blog_Keyword]") Then
		objConn.execute("CREATE TABLE [blog_Keyword] (key_ID AutoIncrement primary key,key_Name VARCHAR(255) default """",key_Intro text default """",key_URL VARCHAR(255) default """")")
	End If

	If Not CheckUpdateDB("[ul_Quote]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_Quote] VARCHAR(255) default """"")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_Quote]=''")
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_DownNum] int DEFAULT 0")
	End If

	If Not CheckUpdateDB("[ul_FileIntro]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_FileIntro] VARCHAR(255) default """"")
	End If

	If Not CheckUpdateDB("[ul_DirByTime]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_DirByTime] YESNO DEFAULT FALSE")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_DirByTime]=[ul_Quote]")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_Quote]=''")
	End If

	If Not CheckUpdateDB("[log_Meta]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Yea] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Nay] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Ratting] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Template] VARCHAR(50) default """"")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_FullUrl] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Type] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Meta] text default """"")

		objConn.execute("UPDATE [blog_Article] SET [log_FullUrl]=''")
		objConn.execute("UPDATE [blog_Article] SET [log_Type]=0")
	End If

	If Not CheckUpdateDB("[cate_Meta]","[blog_Category]") Then
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Url] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_ParentID] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Template] VARCHAR(50) default """"")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_LogTemplate] VARCHAR(50) default """"")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_FullUrl] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Meta] text default """"")
		objConn.execute("ALTER TABLE [blog_Category] ALTER COLUMN [cate_Intro] text default """"")

		objConn.execute("UPDATE [blog_Category] SET [cate_ParentID]=0")
	End If

	If Not CheckUpdateDB("[comm_Meta]","[blog_Comment]") Then
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Reply] text default """"")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_LastReplyIP] VARCHAR(15) default """"")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_LastReplyTime] datetime default now()")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Yea] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Nay] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Ratting] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_ParentID] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_IsCheck] YESNO DEFAULT FALSE")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Meta] text default """"")

		objConn.execute("UPDATE [blog_Comment] SET [comm_ParentID]=0")
	End If
	
	If Not CheckUpdateDB("[mem_Meta]","[blog_Member]") Then
		objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_Guid] VARCHAR(36) default """"")
		objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_Template] VARCHAR(50) default """"")
		objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_FullUrl] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_Url] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_Meta] text default """"")

		Dim objRS
		Set objRS=objConn.Execute("SELECT * FROM [blog_Member]")
		If (Not objRS.bof) And (Not objRS.eof) Then

			Do While Not objRS.eof
				s=RndGuid
				t=md5(objRS("mem_Password") & s)
				objConn.execute("UPDATE [blog_Member] SET [mem_Guid]='"&s&"' WHERE [mem_ID]="& objRS("mem_ID"))
				objConn.execute("UPDATE [blog_Member] SET [mem_Password]='"&t&"' WHERE [mem_ID]="& objRS("mem_ID"))
				objRS.MoveNext
			Loop

		End If

	End If

	If Not CheckUpdateDB("[ul_Meta]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_Meta] text default """"")
		Call objConn.Execute("ALTER TABLE [blog_UpLoad] ALTER COLUMN [ul_FileName] NVARCHAR(255) ")
	End If

	If Not CheckUpdateDB("[tb_Meta]","[blog_TrackBack]") Then
		objConn.execute("ALTER TABLE [blog_TrackBack] ADD COLUMN [tb_Meta] text default """"")
	End If

	If Not CheckUpdateDB("[tag_Meta]","[blog_Tag]") Then
		objConn.execute("ALTER TABLE [blog_Tag] ADD COLUMN [tag_Template] VARCHAR(50) default """"")
		objConn.execute("ALTER TABLE [blog_Tag] ADD COLUMN [tag_FullUrl] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Tag] ADD COLUMN [tag_Meta] text default """"")
	End If

	If Not CheckUpdateDB("[conf_Name]","[blog_Config]") Then
		objConn.execute("CREATE TABLE [blog_Config] (conf_Name VARCHAR(255) default """" not null,conf_Value text default """")")
		objConn.execute("CREATE TABLE [blog_Function] (fn_ID AutoIncrement primary key,fn_Name VARCHAR(50) default """",fn_FileName VARCHAR(50) default """",fn_Order int default 0,fn_Content text default """",fn_IsHidden YESNO DEFAULT 0,fn_SidebarID int default 0,fn_HtmlID VARCHAR(50) default """",fn_Ftype VARCHAR(5) default """",fn_MaxLi int default 0,fn_Source VARCHAR(50) default """",fn_ViewType VARCHAR(50) default """",fn_Meta text default """")")
	End If

	If Not CheckUpdateDB("[coun_Content]","[blog_Counter]") Then
		objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_Content text default """"")
		objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_UserID int default 0")
		objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_PostData  text default """"")
		objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_URL  text default """"")
		objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_AllRequestHeader  text default """"")
		objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_LogName text default """"")
	ENd If

End Function


Function InsertFunctions()

Dim t

Set t=new Tfunction
t.Name="导航栏"
t.FileName="navbar"
t.IsHidden=False
t.Source="system"
t.SidebarID=0
t.Order=1
t.Content="<li><a href=""<#ZC_BLOG_HOST#>"">首页</a></li><li><a href=""<#ZC_BLOG_HOST#>tags.asp"">标签</a></li><li><a href=""<#ZC_BLOG_HOST#>guestbook.html"">留言本</a></li>"
t.HtmlID="divNavBar"
t.Ftype="ul"
t.post


Set t=new Tfunction
t.Name="日历"
t.FileName="calendar"
t.IsHidden=False
t.Source="system"
t.SidebarID=1
t.Order=2
t.Content=""
t.HtmlID="divCalendar"
t.Ftype="div"
t.post




Set t=new Tfunction
t.Name="控制面板"
t.FileName="controlpanel"
t.IsHidden=False
t.Source="system"
t.SidebarID=1
t.Order=3
t.Content="<span class=""cp-login""><a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=login"">[<#ZC_MSG009#>]</a></span>&nbsp;&nbsp;<span class=""cp-vrs""><a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=vrs"">[<#ZC_MSG021#>]</a></span>"
t.HtmlID="divContorPanel"
t.Ftype="div"
t.post




Set t=new Tfunction
t.Name="网站分类"
t.FileName="catalog"
t.IsHidden=False
t.Source="system"
t.SidebarID=1
t.Order=4
t.Content=""
t.HtmlID="divCatalog"
t.Ftype="ul"
t.post


Set t=new Tfunction
t.Name="搜索"
t.FileName="searchpanel"
t.IsHidden=False
t.Source="system"
t.SidebarID=1
t.Order=5
t.Content="<form method=""post"" action=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=Search""><input type=""text"" name=""edtSearch"" id=""edtSearch"" size=""12"" /> <input type=""submit"" value=""<#ZC_MSG087#>"" name=""btnPost"" id=""btnPost"" /></form>"
t.HtmlID="divSearchPanel"
t.Ftype="div"
t.post


Set t=new Tfunction
t.Name="最新留言"
t.FileName="comments"
t.IsHidden=False
t.Source="system"
t.SidebarID=1
t.Order=6
t.Content=""
t.HtmlID="divComments"
t.Ftype="ul"
t.post




Set t=new Tfunction
t.Name="文章归档"
t.FileName="archives"
t.IsHidden=False
t.Source="system"
t.SidebarID=1
t.Order=7
t.Content=""
t.HtmlID="divArchives"
t.Ftype="ul"
t.post



Set t=new Tfunction
t.Name="站点统计"
t.FileName="statistics"
t.IsHidden=False
t.Source="system"
t.SidebarID=0
t.Order=8
t.Content=""
t.HtmlID="divStatistics"
t.Ftype="ul"
t.post




Set t=new Tfunction
t.Name="网站收藏"
t.FileName="favorite"
t.IsHidden=False
t.Source="system"
t.SidebarID=1
t.Order=9
t.Content="<li><a href=""http://bbs.rainbowsoft.org/"" target=""_blank"">ZBlogger社区</a></li><li><a href=""http://download.rainbowsoft.org/"" target=""_blank"">菠萝的海</a></li><li><a href=""http://t.qq.com/zblogcn"" target=""_blank"">Z-Blog微博</a></li>"
t.HtmlID="divFavorites"
t.Ftype="ul"
t.post




Set t=new Tfunction
t.Name="友情链接"
t.FileName="link"
t.IsHidden=False
t.Source="system"
t.SidebarID=1
t.Order=10
t.Content="<li><a href=""http://www.dbshost.cn/"" target=""_blank"" title=""独立博客服务 Z-Blog官方主机"">DBS主机</a></li><li><a href=""http://www.dutory.com/blog/"" target=""_blank"">Dutory官方博客</a></li>"
t.HtmlID="divLinkage"
t.Ftype="ul"
t.post



Set t=new Tfunction
t.Name="图标汇集"
t.FileName="misc"
t.IsHidden=False
t.Source="system"
t.SidebarID=1
t.Order=11
t.Content="<li><a href=""http://www.rainbowsoft.org/"" target=""_blank""><img src=""<#ZC_BLOG_HOST#>zb_system/image/logo/zblog.gif"" height=""31"" width=""88"" border=""0"" alt=""RainbowSoft Studio Z-Blog"" /></a></li><li><a href=""<#ZC_BLOG_HOST#>feed.asp"" target=""_blank""><img src=""<#ZC_BLOG_HOST#>zb_system/image/logo/rss.png"" height=""31"" width=""88"" border=""0"" alt=""订阅本站的 RSS 2.0 新闻聚合"" /></a></li>"
t.HtmlID="divMisc"
t.Ftype="ul"
t.post




Set t=new Tfunction
t.Name="作者列表"
t.FileName="authors"
t.IsHidden=False
t.Source="system"
t.SidebarID=0
t.Order=12
t.Content=""
t.HtmlID="divAuthors"
t.Ftype="ul"
t.post




Set t=new Tfunction
t.Name="最近发表"
t.FileName="previous"
t.IsHidden=False
t.Source="system"
t.SidebarID=0
t.Order=13
t.Content=""
t.HtmlID="divPrevious"
t.Ftype="ul"
t.post



Set t=new Tfunction
t.Name="Tags列表"
t.FileName="tags"
t.IsHidden=False
t.Source="system"
t.SidebarID=0
t.Order=14
t.Content=""
t.HtmlID="divTags"
t.Ftype="ul"
t.post


End Function





Function InsertOptions()

BlogConfig.Load("Blog")

'---------------------------------网站基本设置-----------------------------------
Call BlogConfig.Write("ZC_BLOG_HOST","http://localhost/")
Call BlogConfig.Write("ZC_BLOG_TITLE","My Blog")
Call BlogConfig.Write("ZC_BLOG_SUBTITLE","Hello, world!")
Call BlogConfig.Write("ZC_BLOG_NAME","My Blog")
Call BlogConfig.Write("ZC_BLOG_SUB_NAME","Hello, world!")
Call BlogConfig.Write("ZC_BLOG_THEME","default")
Call BlogConfig.Write("ZC_BLOG_CSS","default")
Call BlogConfig.Write("ZC_BLOG_COPYRIGHT","Copyright Your WebSite. Some Rights Reserved.")
Call BlogConfig.Write("ZC_BLOG_MASTER","zblogger")
Call BlogConfig.Write("ZC_BLOG_LANGUAGE","zh-CN")





'----------------------------数据库配置---------------------------------------
Call BlogConfig.Write("ZC_DATABASE_PATH","")
Call BlogConfig.Write("ZC_MSSQL_ENABLE",False)
Call BlogConfig.Write("ZC_MSSQL_DATABASE","")
Call BlogConfig.Write("ZC_MSSQL_USERNAME","")
Call BlogConfig.Write("ZC_MSSQL_PASSWORD","")
Call BlogConfig.Write("ZC_MSSQL_SERVER","(local)")





'---------------------------------插件----------------------------------------
Call BlogConfig.Write("ZC_USING_PLUGIN_LIST","AppCentre|FileManage|GuestBook|Totoro")








'-------------------------------全局配置-----------------------------------
Call BlogConfig.Write("ZC_BLOG_CLSID","BB1C5669-6E37-460C-F415-D287D7BBB59E")
Call BlogConfig.Write("ZC_TIME_ZONE","+0800")
Call BlogConfig.Write("ZC_HOST_TIME_ZONE","+0800")
Call BlogConfig.Write("ZC_UPDATE_INFO_URL","http://update.rainbowsoft.org/info/")
Call BlogConfig.Write("ZC_MULTI_DOMAIN_SUPPORT",False)
Call BlogConfig.Write("ZC_PERMANENT_DOMAIN_ENABLE",False)



'留言评论
Call BlogConfig.Write("ZC_COMMENT_TURNOFF",False)
Call BlogConfig.Write("ZC_COMMENT_VERIFY_ENABLE",False)
Call BlogConfig.Write("ZC_COMMENT_REVERSE_ORDER_EXPORT",False)
Call BlogConfig.Write("ZC_COMMNET_MAXFLOOR",4)


'验证码
Call BlogConfig.Write("ZC_VERIFYCODE_STRING","0123456789")
Call BlogConfig.Write("ZC_VERIFYCODE_WIDTH",60)
Call BlogConfig.Write("ZC_VERIFYCODE_HEIGHT",20)


Call BlogConfig.Write("ZC_DISPLAY_COUNT",10)
Call BlogConfig.Write("ZC_RSS2_COUNT",10)
Call BlogConfig.Write("ZC_SEARCH_COUNT",25)
Call BlogConfig.Write("ZC_PAGEBAR_COUNT",15)
Call BlogConfig.Write("ZC_MUTUALITY_COUNT",10)
Call BlogConfig.Write("ZC_COMMENTS_DISPLAY_COUNT",10)





Call BlogConfig.Write("ZC_USE_NAVIGATE_ARTICLE",True)

Call BlogConfig.Write("ZC_RSS_EXPORT_WHOLE",False)




'后台管理
Call BlogConfig.Write("ZC_MANAGE_COUNT",50)
Call BlogConfig.Write("ZC_REBUILD_FILE_COUNT",50)
Call BlogConfig.Write("ZC_REBUILD_FILE_INTERVAL",1)










'UBB转换
Call BlogConfig.Write("ZC_UBB_ENABLE",False)
Call BlogConfig.Write("ZC_UBB_LINK_ENABLE",False)
Call BlogConfig.Write("ZC_UBB_FONT_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_CODE_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_FACE_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_IMAGE_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_MEDIA_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_FLASH_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_TYPESET_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_AUTOLINK_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_AUTOKEY_ENABLE",False)




'表情相关
Call BlogConfig.Write("ZC_EMOTICONS_FILENAME","neutral|grin|happy|slim|smile|tongue|wink|surprised|confuse|cool|cry|evilgrin|fat|mad|red|roll|unhappy|waii|yell")
Call BlogConfig.Write("ZC_EMOTICONS_FILETYPE","png")
Call BlogConfig.Write("ZC_EMOTICONS_FILESIZE",16)




'上传相关
Call BlogConfig.Write("ZC_UPLOAD_FILETYPE","jpg|gif|png|jpeg|bmp|psd|wmf|ico|rpm|deb|tar|gz|sit|7z|bz2|zip|rar|xml|xsl|svg|svgz|doc|xls|wps|chm|txt|pdf|mp3|avi|mpg|rm|ra|rmvb|mov|wmv|wma|swf|fla|torrent|zpi|zti|zba")
Call BlogConfig.Write("ZC_UPLOAD_FILESIZE",10485760)
Call BlogConfig.Write("ZC_UPLOAD_DIRBYMONTH",True)
Call BlogConfig.Write("ZC_UPLOAD_DIRECTORY","zb_users\upload")



'当前 Z-Blog 版本
Call BlogConfig.Write("ZC_BLOG_VERSION","2.0 Doomsday Build 121221")



'用户名,密码,评论长度等限制
Call BlogConfig.Write("ZC_USERNAME_MIN",4)
Call BlogConfig.Write("ZC_USERNAME_MAX",20)
Call BlogConfig.Write("ZC_PASSWORD_MIN",8)
Call BlogConfig.Write("ZC_PASSWORD_MAX",14)
Call BlogConfig.Write("ZC_EMAIL_MAX",30)
Call BlogConfig.Write("ZC_HOMEPAGE_MAX",100)
Call BlogConfig.Write("ZC_CONTENT_MAX",1000)




Call BlogConfig.Write("ZC_UNCATEGORIZED_NAME","未分类")
Call BlogConfig.Write("ZC_UNCATEGORIZED_ALIAS","")
Call BlogConfig.Write("ZC_UNCATEGORIZED_COUNT",0)
Call BlogConfig.Write("ZC_SYNTAXHIGHLIGHTER_ENABLE",True)
Call BlogConfig.Write("ZC_CODEMIRROR_ENABLE",True)
Call BlogConfig.Write("ZC_ARTICLE_EXCERPT_MAX",250)
Call BlogConfig.Write("ZC_HTTP_LASTMODIFIED",False)





'---------------------------------静态化配置-----------------------------------


'{asp html shtml}
Call BlogConfig.Write("ZC_STATIC_TYPE","html")

Call BlogConfig.Write("ZC_STATIC_DIRECTORY","post")

Call BlogConfig.Write("ZC_TEMPLATE_DIRECTORY","template")



'ACTIVE MIX REWRITE
Call BlogConfig.Write("ZC_STATIC_MODE","ACTIVE")
Call BlogConfig.Write("ZC_POST_STATIC_MODE","STATIC")
Call BlogConfig.Write("ZC_ARTICLE_REGEX","{%host%}/{%post%}/{%alias%}.html")
Call BlogConfig.Write("ZC_PAGE_REGEX","{%host%}/{%alias%}.html")
Call BlogConfig.Write("ZC_CATEGORY_REGEX","{%host%}/catalog.asp?cate={%id%}")
Call BlogConfig.Write("ZC_USER_REGEX","{%host%}/catalog.asp?auth={%id%}")
Call BlogConfig.Write("ZC_TAGS_REGEX","{%host%}/catalog.asp?tags={%alias%}")
Call BlogConfig.Write("ZC_DATE_REGEX","{%host%}/catalog.asp?date={%date%}")
Call BlogConfig.Write("ZC_DEFAULT_REGEX","{%host%}/catalog.asp")





'--------------------------WAP----------------------------------------
Call BlogConfig.Write("ZC_DISPLAY_COUNT_WAP",5)
Call BlogConfig.Write("ZC_COMMENT_COUNT_WAP",5)
Call BlogConfig.Write("ZC_PAGEBAR_COUNT_WAP",5)
Call BlogConfig.Write("ZC_SINGLE_SIZE_WAP",1000)
Call BlogConfig.Write("ZC_SINGLE_PAGEBAR_COUNT_WAP",5)

Call BlogConfig.Write("ZC_FILENAME_WAP","wap.asp")
Call BlogConfig.Write("ZC_WAPCOMMENT_ENABLE",True)
'全文
Call BlogConfig.Write("ZC_DISPLAY_MODE_ALL_WAP",True)
'显示分类导航
Call BlogConfig.Write("ZC_DISPLAY_CATE_ALL_WAP",True)
'分页条
Call BlogConfig.Write("ZC_DISPLAY_PAGEBAR_ALL_WAP",True)
'数量
Call BlogConfig.Write("ZC_WAP_MUTUALITY_LIMIT",5)

BlogConfig.Save


End Function



Function LoadOldOptions()

	ZC_BLOG_HOST=BlogHost
	Dim tmp

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If (fso.FileExists(BlogPath & "c_custom.asp")) Then
		tmp=LoadFromFile(BlogPath & "c_custom.asp","utf-8")
		Call LoadValueForSetting(tmp,True,"String","ZC_BLOG_TITLE",ZC_BLOG_TITLE)
		Call LoadValueForSetting(tmp,True,"String","ZC_BLOG_SUBTITLE",ZC_BLOG_SUBTITLE)
		Call LoadValueForSetting(tmp,True,"String","ZC_BLOG_NAME",ZC_BLOG_NAME)
		Call LoadValueForSetting(tmp,True,"String","ZC_BLOG_SUB_NAME",ZC_BLOG_SUB_NAME)
		Call LoadValueForSetting(tmp,True,"String","ZC_BLOG_COPYRIGHT",ZC_BLOG_COPYRIGHT)

		Call fso.DeleteFile(BlogPath & "c_custom.asp")

	End If

	If (fso.FileExists(BlogPath & "c_option.asp")) Then
		tmp=LoadFromFile(BlogPath & "c_option.asp","utf-8")
		Call LoadValueForSetting(tmp,True,"String","ZC_TIME_ZONE",ZC_TIME_ZONE)
		Call LoadValueForSetting(tmp,True,"String","ZC_HOST_TIME_ZONE",ZC_HOST_TIME_ZONE)

		Call fso.DeleteFile(BlogPath & "c_option.asp")

	End If

End Function


Function LoadOldFunctions()

	Call GetFunction()

	Functions(FunctionMetas.GetValue("navbar")).Content=LoadFromFile(BlogPath & "zb_users\INCLUDE\navbar.asp","utf-8")
	Functions(FunctionMetas.GetValue("navbar")).Post()

	Functions(FunctionMetas.GetValue("favorite")).Content=LoadFromFile(BlogPath & "zb_users\INCLUDE\favorite.asp","utf-8")
	Functions(FunctionMetas.GetValue("favorite")).Post()

	Functions(FunctionMetas.GetValue("link")).Content=LoadFromFile(BlogPath & "zb_users\INCLUDE\link.asp","utf-8")
	Functions(FunctionMetas.GetValue("link")).Post()

	Functions(FunctionMetas.GetValue("misc")).Content=LoadFromFile(BlogPath & "zb_users\INCLUDE\misc.asp","utf-8")
	Functions(FunctionMetas.GetValue("misc")).Post()

End Function



Function SaveConfigs()

	On Error Resume Next
	Dim a,b
	b=LoadFromFile(BlogPath &"zb_users\c_option.asp","utf-8")
	For Each a In BlogConfig.Meta.Names
		If InStr(b,"Dim "& a)>0 Then
			Call Execute("Call BlogConfig.Write("""&a&""","&a&")")
		End If
	Next

	Call BlogConfig.Save()
	Err.Clear


	Call SaveConfig2Option()

End Function


Function FixOldTotoro()

objConn.Execute "UPDATE [blog_Comment] SET [comm_isCheck]=1 WHERE [log_ID]<0"
objConn.Execute "UPDATE [blog_Comment] SET [log_ID]=-1-[log_ID] WHERE [log_ID]<0"

End Function


Function RevToComment()

	ShowError_Custom="StarTime = Timer()"

	Call GetUser()

	Dim objRS2,objComment,s,t,Match,Matches,t1,t2,t3,t4,c,u

	Dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True


	Set objRS2=objConn.Execute("SELECT * FROM [blog_Comment] WHERE ([comm_isCheck]=0) AND (InStr(1,LCase([comm_Content]),LCase('[/REVERT]'),0)<>0)")
	If (Not objRS2.bof) And (Not objRS2.eof) Then
		Do While Not objRS2.eof

			Set objComment=New TComment
			objComment.LoadInfoByArray(Array(objRS2("comm_ID"),objRS2("log_ID"),objRS2("comm_AuthorID"),objRS2("comm_Author"),objRS2("comm_Content"),objRS2("comm_Email"),objRS2("comm_HomePage"),objRS2("comm_PostTime"),objRS2("comm_IP"),objRS2("comm_Agent"),objRS2("comm_Reply"),objRS2("comm_LastReplyIP"),objRS2("comm_LastReplyTime"),objRS2("comm_ParentID"),objRS2("comm_IsCheck"),objRS2("comm_Meta")))



			s=objComment.Content
			objRegExp.Pattern="(\[REVERT=)(.+?)(\])([\u0000-\uffff]+?)(\[\/REVERT\])"
			Set Matches = objRegExp.Execute(s)

			For Each Match in Matches

				t=Match
				s=Replace(s,t,"")


				t1=Match.SubMatches(1)
				t2=Match.SubMatches(3)

				t4= Left(t1,InStr(1,t1," ",0)-1)
				t3=Mid(t1,InStr(InStr(1,t1," ",0)+1,t1," ",0),InstrRev(t1," ",Len(t1),0)-InStr(InStr(1,t1," ",0)+1,t1," ",0) )

				t3=CDate(t3)

				u=0
				Dim User
				For Each User in Users
					If IsObject(User) Then
						If User.Name=t4 Then
							u=User.ID
						End If
					End If
				Next

				Set c=new TComment

				c.log_ID=objComment.log_ID
				c.AuthorID=u
				c.Author=t4
				c.Content=t2
				c.Email=""
				c.HomePage=""
				c.ParentID=objComment.ID
				c.PostTime=t3

				c.post
				Set c=Nothing


			Next
			Set Matches = Nothing



			objComment.Content=s

			objConn.Execute("UPDATE [blog_Comment] SET [comm_Content]='"&FilterSQL(s)&"' WHERE [comm_ID] =" & objComment.ID)

			Set objComment=Nothing
			objRS2.MoveNext
		Loop
	End If
	objRS2.Close
	Set objRS2=Nothing

End Function
%>
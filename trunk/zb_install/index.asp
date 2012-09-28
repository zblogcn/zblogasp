<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
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

'If (ZC_DATABASE_PATH<>"" And ZC_MSSQL_ENABLE=False) Or (ZC_MSSQL_SERVER<>"" And ZC_MSSQL_ENABLE=True) Then
'	zblogstep=0
'End If

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
	<title>Z-Blog 2.0 安装程序</title>
</head>
<body>
  <div class="setup"><form method="post" action="?step=<%=zblogstep+1%>">
<%

Select Case zblogstep
Case 0 Call Setup0
Case 1 Call Setup1
Case 2 Call Setup2
Case 3 Call Setup3
Case 4 Call Setup4
End  Select
%>
  </form></div>

<script language="JavaScript" type="text/javascript">
function setup2(){
	if($("#dbtype").val()=="mssql"){
		if($("#dbserver").val()==""){alert('数据库服务器需要填写.');return false;};
		if($("#dbname").val()==""){alert('数据库名称需要填写.');return false;};
		if($("#dbusername").val()==""){alert('数据库用户名需要填写.');return false;};
	}



if($("#blogtitle").val()==""){alert('网站标题需要填写.');return false;};
if($("#username").val()==""){alert('管理员名称需要填写.');return false;};
if($("#password").val()==""){alert('管理员密码需要填写.');return false;};
if($("#password").val().toString().search("^[A-Za-z0-9`~!@#\$%\^&\*\-_]{8,}$")==-1){alert('管理员密码必须是8位或更长的数字和字母,字符组合.');return false;};
if($("#password").val()!==$("#repassword").val()){alert('必须确认密码.');return false;};

}

</script>
</body>
</html>
<%




Function Setup0()
%>
<dl>
<dd id="ddleft">
<img src='../zb_system/image/admin/install.png' alt='' />
<p>安装进度:<span><font color='#3d69aa'></font><font color='white'>█████████████████</font></span></p>
<p>安装协议&nbsp;&nbsp;»&nbsp;&nbsp;数据库建立与设置&nbsp;&nbsp;»&nbsp;&nbsp;安装结果</p>
</dd>
<dd id="ddright">
<div id='title'>安装提示</div>
<div id='content'>
通过配置文件的检验,您已经安装并配置好Z-Blog了,不能再重复使用安装程序.
</div>
<div id='bottom'>
<input type="button" name="next" onclick="window.location.href='<%=BlogHost%>'" id="netx" value="退出" />
</div>
</dd>
</dl>
<%
End Function












Function Setup1()
%>
<dl>
<dd id="ddleft">
<img src='../zb_system/image/admin/install.png' alt='' />
<p>安装进度:<span><font color='#3d69aa'>██████</font><font color='white'>███████████</font></span></p>
<p><b>安装协议</b>&nbsp;&nbsp;»&nbsp;&nbsp;数据库建立与设置&nbsp;&nbsp;»&nbsp;&nbsp;安装结果</p>
</dd>
<dd id="ddright">
<div id='title'>安装协议</div>
<div id='content'>
  <textarea readonly="readonly">
本《Z-Blog软件最终用户许可协议》（以下简称《协议》）是您与RainbowSoft Studio之间关于下载、安装、使用、复制Z-Blog软件的法律协议。本《协议》描述RainbowSoft Studio与您之间关于Z-Blog许可使用及相关方面的权利义务。

请您仔细阅读本《协议》，用户可选择不使用Z-Blog，用户使用Z-Blog的行为将被视为对本《协议》全部内容的认可，并同意接受本《协议》各项条款的约束。

  </textarea>
</div>
<div id='bottom'>
 <label><input type="checkbox" onclick="$('input').prop('disabled',false);$(this).prop('disabled',true);" />我已阅读并同意此协议.</label>&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="next" id="netx" value="下一步" disabled="disabled" />
</div>
</dd>
</dl>
<%
End Function











Function Setup2()
%>
<dl>
<dd id="ddleft">
<img src='../zb_system/image/admin/install.png' alt='' />
<p>安装进度:<span><font color='#3d69aa'>███████████</font><font color='white'>██████</font></span></p>
<p><b>安装协议</b>&nbsp;&nbsp;»&nbsp;&nbsp;<b>数据库建立与设置</b>&nbsp;&nbsp;»&nbsp;&nbsp;安装结果</p>
</dd>
<dd id="ddright">
<div id='title'>数据库建立与设置</div>
<div id='content'>
<input type="hidden" name="dbtype" id="dbtype" value="access" />
<p><b>类型选择</b>:&nbsp;&nbsp;<label onclick="$('#mssql').hide();$('#access').show();$('#dbtype').val('access');"><input type="radio" name="db" checked="checked" />Access</label>&nbsp;&nbsp;&nbsp;&nbsp;<label onclick="$('#access').hide();$('#mssql').show();$('#dbtype').val('mssql');"><input type="radio" name="db" />MSSQL</label></p>
<div id='access'>
<p><b>数&nbsp;据&nbsp;库:</b>&nbsp;&nbsp;<input type="text" name="dbpath" id="dbpath" value="#%20<%=LCase(Replace(RndGuid(),"-",""))%>.mdb" readonly="readonly" style='width:350px;' /></p>
</div>
<div id='mssql' style='display:none;'>
<p><b>数据库主机:</b><input type="text" name="dbserver" id="dbserver" value="localhost" style='width:350px;' /></p>
<p><b>数据库名称:</b><input type="text" name="dbname" id="dbname" value="" style='width:350px;' /></p>
<p><b>用户名称:</b>&nbsp;&nbsp;<input type="text" name="dbusername" id="dbusername" value="" style='width:350px;' /></p>
<p><b>用户密码:</b>&nbsp;&nbsp;<input type="text" name="dbpassword" id="dbpassword" value="" style='width:350px;' /></p>
</div>
<p class='title'>网站设置</p>
<p><b>网站名称:</b>&nbsp;&nbsp;<input type="text" name="blogtitle" id="blogtitle" value="" style='width:350px;' /></p>
<p><b>用&nbsp;户&nbsp;名:</b>&nbsp;&nbsp;<input type="text" name="username" id="username" value="" style='width:250px;' />&nbsp;(英文,数字,汉字和._的组合)</p>
<p><b>密&nbsp;&nbsp;&nbsp;&nbsp;码:</b>&nbsp;&nbsp;<input type="password" name="password" id="password" value="" style='width:250px;' />&nbsp;(8位或更长的数字和字母,字符组合)</p>
<p><b>确认密码:</b>&nbsp;&nbsp;<input type="password" name="repassword" id="repassword" value="" style='width:250px;' /></p>
</div>
<div id='bottom'>
<input type="submit" name="next" id="netx" onclick="return setup2()" value="下一步" />
</div>
</dd>
</dl>
<%
End Function













Function Setup3()
On Error Resume Next
%>
<dl>
<dd id="ddleft">
<img src='../zb_system/image/admin/install.png' alt='' />
<p>安装进度:<span><font color='#3d69aa'>█████████████████</font><font color='white'></font></span></p>
<p><b>安装协议</b>&nbsp;&nbsp;»&nbsp;&nbsp;<b>数据库建立与设置</b>&nbsp;&nbsp;»&nbsp;&nbsp;<b>安装结果</b></p>
</dd>
<dd id="ddright">

<div id='title'>安装结果</div>
<div id='content'>
<%
ZC_BLOG_TITLE=Request.Form("blogtitle")
ZC_BLOG_NAME=ZC_BLOG_TITLE


userguid=RndGuid()
password=MD5(MD5(Request.Form("password")) & userguid)
username=Request.Form("username")

dbtype=Request.Form("dbtype")
dbpath=Request.Form("dbpath")
dbserver=Request.Form("dbserver")
dbname=Request.Form("dbname")
dbusername=Request.Form("dbusername")
dbpassword=Request.Form("dbpassword")

If dbtype="access" Then

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

	fso.CopyFile BlogPath & "\zb_install\zblog.mdb", BlogPath & "\zb_users\data\" & dbpath

	ZC_DATABASE_PATH="zb_users\data\" & dbpath

	ZC_MSSQL_ENABLE=False

ElseIf dbtype="mssql" Then

	ZC_MSSQL_DATABASE=dbname

	ZC_MSSQL_USERNAME=dbusername

	ZC_MSSQL_PASSWORD=dbpassword

	ZC_MSSQL_SERVER=dbserver

	ZC_MSSQL_ENABLE=True

End If
Response.Write dbpath


If OpenConnect()=False Then

	Response.Write("<script language=javascript>alert('数据库连接错误!');</script>")
	Response.Write("<script language=javascript>history.go(-1);</script>")
	Response.End

End If


If dbtype="access" Then
	Call CreateAccessTable()
ElseIf dbtype="mssql" Then
	Call CreateMssqlTable()
End If

Call InsertFunctions()

Call InsertOptions()

Call InsertArticleAndPage()

Call SaveConfigs()

%>
<p>数据库表创建成功!</p>
<p>默认配置数据导入成功!</p>
<p>默认侧栏数据导入成功!</p>
<p>用户信息导入成功!</p>
<p>Hell World文章导入成功!</p>
<p>留言本页面导入成功!</p>
<p>配置文件c_option.asp保存成功!</p>
<p>Z-Blog 2.0安装成功了,现在您可以点击"完成"进入网站首页.</p>

</div>
<div id='bottom'>
<input type="submit" name="next" id="netx" value="完成" />
</div>


</dd>
</dl>
<%
End Function

Function Setup4()
	Response.Redirect BlogHost
End Function


Function CreateAccessTable()

	objConn.BeginTrans

	objConn.execute("CREATE TABLE [blog_Tag] (tag_ID AutoIncrement primary key,tag_Name VARCHAR(255) default """",tag_Intro text default """",tag_ParentID int default 0,tag_URL VARCHAR(255) default """",tag_Order int default 0,tag_Count int default 0,tag_Template VARCHAR(50) default """",tag_FullUrl VARCHAR(255) default """",tag_Meta text default """")")

	objConn.execute("CREATE TABLE [blog_Article] (log_ID AutoIncrement primary key,log_CateID int default 0,log_AuthorID int default 0,log_Level int default 0,log_Url VARCHAR(255) default """",log_Title VARCHAR(255) default """",log_Intro text default """",log_Content text default """",log_IP VARCHAR(15) default """",log_PostTime datetime default now(),log_CommNums int default 0,log_ViewNums int default 0,log_TrackBackNums int default 0,log_Tag VARCHAR(255) default """",log_IsTop YESNO DEFAULT 0,log_Yea int default 0,log_Nay int default 0,log_Ratting int default 0,log_Template VARCHAR(50) default """",log_FullUrl VARCHAR(255) default """",log_Type int DEFAULT 0,log_Meta text default """")")

	objConn.execute("CREATE TABLE [blog_Category] (cate_ID AutoIncrement primary key,cate_Name VARCHAR(50) default """",cate_Order int default 0,cate_Intro VARCHAR(255) default """",cate_Count int default 0,cate_URL VARCHAR(255) default """",cate_ParentID int default 0,cate_Template VARCHAR(50) default """",cate_LogTemplate VARCHAR(50) default """",cate_FullUrl VARCHAR(255) default """",cate_Meta text default """")")

	objConn.execute("CREATE TABLE [blog_Comment] (comm_ID AutoIncrement primary key,log_ID int default 0,comm_AuthorID int default 0,comm_Author VARCHAR(20) default """",comm_Content text default """",comm_Email VARCHAR(50) default """",comm_HomePage VARCHAR(255) default """",comm_PostTime datetime default now(),comm_IP VARCHAR(15) default """",comm_Agent text default """",comm_Reply text default """",comm_LastReplyIP VARCHAR(15) default """",comm_LastReplyTime datetime default now(),comm_Yea int default 0,comm_Nay int default 0,comm_Ratting int default 0,comm_ParentID int default 0,comm_IsCheck YESNO DEFAULT FALSE,comm_Meta text default """")")

	objConn.execute("CREATE TABLE [blog_TrackBack] (tb_ID AutoIncrement primary key,log_ID int default 0,tb_URL VARCHAR(255) default """",tb_Title VARCHAR(100) default """",tb_Blog VARCHAR(50) default """",tb_Excerpt text default """",tb_PostTime datetime default now(),tb_IP VARCHAR(15) default """",tb_Agent text default """",tb_Meta text default """")")

	objConn.execute("CREATE TABLE [blog_UpLoad] (ul_ID AutoIncrement primary key,ul_AuthorID int default 0,ul_FileSize int default 0,ul_FileName VARCHAR(255) default """",ul_PostTime datetime default now(),ul_Quote VARCHAR(255) default """",ul_DownNum int default 0,ul_FileIntro VARCHAR(255) default """",ul_DirByTime YESNO DEFAULT 0,ul_Meta text default """")")

	objConn.execute("CREATE TABLE [blog_Counter] (coun_ID AutoIncrement primary key,coun_IP VARCHAR(15) default """",coun_Agent text default """",coun_Refer VARCHAR(255) default """",coun_PostTime datetime default now(),coun_Content text default """",coun_UserID int default 0,coun_PostData  text default """",coun_URL  text default """",coun_AllRequestHeader  text default """",coun_LogName text default """")")

	objConn.execute("CREATE TABLE [blog_Keyword] (key_ID AutoIncrement primary key,key_Name VARCHAR(255) default """",key_Intro text default """",key_URL VARCHAR(255) default """")")

	objConn.execute("CREATE TABLE [blog_Member] (mem_ID AutoIncrement primary key,mem_Level int default 0,mem_Name VARCHAR(20) default """",mem_Password VARCHAR(32) default """",mem_Sex int default 0,mem_Email VARCHAR(50) default """",mem_MSN VARCHAR(50) default """",mem_QQ VARCHAR(50) default """",mem_HomePage VARCHAR(255) default """",mem_LastVisit datetime default now(),mem_Status int default 0,mem_PostLogs int default 0,mem_PostComms int default 0,mem_Intro text default """",mem_IP VARCHAR(15) default """",mem_Count int default 0,mem_Template VARCHAR(50) default """",mem_FullUrl VARCHAR(255) default """",mem_Guid VARCHAR(36) default """",mem_Meta text default """")")

	objConn.execute("CREATE TABLE [blog_Config] (conf_Name VARCHAR(255) default """" not null,conf_Value text default """")")
	'objConn.execute("CREATE UNIQUE INDEX index_conf_Name ON [blog_Config](conf_Name)")

	objConn.execute("CREATE TABLE [blog_Function] (fn_ID AutoIncrement primary key,fn_Name VARCHAR(50) default """",fn_FileName VARCHAR(50) default """",fn_Order int default 0,fn_Content text default """",fn_IsSystem YESNO DEFAULT 0,fn_SidebarID int default 0,fn_HtmlID VARCHAR(50) default """",fn_Ftype VARCHAR(5) default """",fn_MaxLi int default 0,fn_Meta text default """")")

	objConn.Execute("INSERT INTO [blog_Member]([mem_Level],[mem_Name],[mem_PassWord],[mem_Email],[mem_HomePage],[mem_Intro],[mem_Guid]) VALUES (1,'"&username&"','"&password&"','null@null.com','','','"&userguid&"')")

	objConn.CommitTrans

End Function


Function CreateMssqlTable()

	objConn.BeginTrans

	objConn.execute("CREATE TABLE [blog_Tag] (tag_ID int identity(1,1) not null primary key,tag_Name nvarchar(255) default '',tag_Intro ntext default '',tag_ParentID int default 0,tag_URL nvarchar(255) default '',tag_Order int default 0,tag_Count int default 0,tag_Template nvarchar(50) default '',tag_FullUrl nvarchar(255) default '',tag_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_Article] (log_ID int identity(1,1) not null primary key,log_CateID int default 0,log_AuthorID int default 0,log_Level int default 0,log_Url nvarchar(255) default '',log_Title nvarchar(255) default '',log_Intro ntext default '',log_Content ntext default '',log_IP nvarchar(15) default '',log_PostTime datetime default getdate(),log_CommNums int default 0,log_ViewNums int default 0,log_TrackBackNums int default 0,log_Tag nvarchar(255) default '',log_IsTop bit DEFAULT 0,log_Yea int default 0,log_Nay int default 0,log_Ratting int default 0,log_Template nvarchar(50) default '',log_FullUrl nvarchar(255) default '',log_Type int default 0,log_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_Category] (cate_ID int identity(1,1) not null primary key,cate_Name nvarchar(50) default '',cate_Order int default 0,cate_Intro nvarchar(255) default '',cate_Count int default 0,cate_URL nvarchar(255) default '',cate_ParentID int default 0,cate_Template nvarchar(50) default '',cate_LogTemplate nvarchar(50) default '',cate_FullUrl nvarchar(255) default '',cate_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_Comment] (comm_ID int identity(1,1) not null primary key,log_ID int default 0,comm_AuthorID int default 0,comm_Author nvarchar(20) default '',comm_Content ntext default '',comm_Email nvarchar(50) default '',comm_HomePage nvarchar(255) default '',comm_PostTime datetime default getdate(),comm_IP nvarchar(15) default '',comm_Agent ntext default '',comm_Reply ntext default '',comm_LastReplyIP nvarchar(15) default '',comm_LastReplyTime datetime default getdate(),comm_Yea int default 0,comm_Nay int default 0,comm_Ratting int default 0,comm_ParentID int default 0,comm_IsCheck bit default 0,comm_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_TrackBack] (tb_ID int identity(1,1) not null primary key,log_ID int default 0,tb_URL nvarchar(255) default '',tb_Title nvarchar(100) default '',tb_Blog nvarchar(50) default '',tb_Excerpt ntext default '',tb_PostTime datetime default getdate(),tb_IP nvarchar(15) default '',tb_Agent ntext default '',tb_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_UpLoad] (ul_ID int identity(1,1) not null primary key,ul_AuthorID int default 0,ul_FileSize int default 0,ul_FileName nvarchar(255) default '',ul_PostTime datetime default getdate(),ul_Quote nvarchar(255) default '',ul_DownNum int default 0,ul_FileIntro nvarchar(255) default '',ul_DirByTime bit DEFAULT 0,ul_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_Counter] (coun_ID int identity(1,1) not null primary key,coun_IP nvarchar(15) default '',coun_Agent ntext default '',coun_Refer nvarchar(255) default '',coun_PostTime datetime default getdate(),coun_Content ntext default '',coun_UserID int default 0,coun_PostData ntext default '',coun_URL ntext default '',coun_AllRequestHeader ntext default '',coun_LogName ntext default '')")


	objConn.execute("CREATE TABLE [blog_Keyword] (key_ID int identity(1,1) not null primary key,key_Name nvarchar(255) default '',key_Intro ntext default '',key_URL nvarchar(255) default '')")

	objConn.execute("CREATE TABLE [blog_Member] (mem_ID int identity(1,1) not null primary key,mem_Level int default 0,mem_Name nvarchar(20) default '',mem_Password nvarchar(32) default '',mem_Sex int default 0,mem_Email nvarchar(50) default '',mem_MSN nvarchar(50) default '',mem_QQ nvarchar(50) default '',mem_HomePage nvarchar(255) default '',mem_LastVisit datetime default getdate(),mem_Status int default 0,mem_PostLogs int default 0,mem_PostComms int default 0,mem_Intro ntext default '',mem_IP nvarchar(15) default '',mem_Count int default 0,mem_Template nvarchar(50) default '',mem_FullUrl nvarchar(255) default '',mem_Guid  nvarchar(36) default '',mem_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_Config] (conf_Name nvarchar(255) not null default '',conf_Value text default '')")

	objConn.execute("CREATE TABLE [blog_Function] (fn_ID int identity(1,1) not null primary key,fn_Name nvarchar(50) default '',fn_FileName nvarchar(50) default '',fn_Order int default 0,fn_Content ntext default '',fn_IsSystem bit DEFAULT 0,fn_SidebarID int default 0,fn_HtmlID nvarchar(50) default '',fn_Ftype nvarchar(5) default '',fn_MaxLi int default 0,fn_Meta ntext default '')")

	objConn.Execute("INSERT INTO [blog_Member]([mem_Level],[mem_Name],[mem_PassWord],[mem_Email],[mem_HomePage],[mem_Intro],[mem_Guid]) VALUES (1,'"&username&"','"&password&"','null@null.com','','','"&userguid&"')")

	objConn.CommitTrans

End Function


Function InsertFunctions()

Dim t

Set t=new Tfunction
t.Name="导航栏"
t.FileName="navbar"
t.IsSystem=True
t.SidebarID=0
t.Order=1
t.Content="<li><a href=""<#ZC_BLOG_HOST#>"">首页</a></li><li><a href=""<#ZC_BLOG_HOST#>tags.asp"">标签</a></li><li><a href=""<#ZC_BLOG_HOST#>guestbook.html"">留言本</a></li>"
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
t.Content="<span class=""cp-login""><a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=login"">[<#ZC_MSG009#>]</a></span>&nbsp;&nbsp;<span class=""cp-vrs""><a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=vrs"">[<#ZC_MSG021#>]</a></span>"
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
t.Content="<form method=""post"" action=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=Search""><input type=""text"" name=""edtSearch"" id=""edtSearch"" size=""12"" /> <input type=""submit"" value=""<#ZC_MSG087#>"" name=""btnPost"" id=""btnPost"" /></form>"
t.HtmlID="divSearchPanel"
t.Ftype="div"
t.post


Set t=new Tfunction
t.Name="最新留言"
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
t.Content="<li><a href=""http://www.rainbowsoft.org/"" target=""_blank""><img src=""<#ZC_BLOG_HOST#>zb_system/image/logo/zblog.gif"" height=""31"" width=""88"" border=""0"" alt=""RainbowSoft Studio Z-Blog"" /></a></li><li><a href=""<#ZC_BLOG_HOST#>feed.asp"" target=""_blank""><img src=""<#ZC_BLOG_HOST#>zb_system/image/logo/rss.png"" height=""31"" width=""88"" border=""0"" alt=""订阅本站的 RSS 2.0 新闻聚合"" /></a></li>"
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
t.Name="Tags列表"
t.FileName="tags"
t.IsSystem=True
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
Call BlogConfig.Write("ZC_DATABASE_PATH","zb_users\data\#%20768d53283c63b13403f0.mdb")
Call BlogConfig.Write("ZC_MSSQL_ENABLE",False)
Call BlogConfig.Write("ZC_MSSQL_DATABASE","zb")
Call BlogConfig.Write("ZC_MSSQL_USERNAME","sa")
Call BlogConfig.Write("ZC_MSSQL_PASSWORD","")
Call BlogConfig.Write("ZC_MSSQL_SERVER","(local)\SQLEXPRESS")





'---------------------------------插件----------------------------------------
Call BlogConfig.Write("ZC_USING_PLUGIN_LIST","")








'-------------------------------全局配置-----------------------------------
Call BlogConfig.Write("ZC_BLOG_CLSID","BB1C5669-6E37-460C-F415-D287D7BBB59E")
Call BlogConfig.Write("ZC_TIME_ZONE","+0800")
Call BlogConfig.Write("ZC_HOST_TIME_ZONE","+0800")
Call BlogConfig.Write("ZC_UPDATE_INFO_URL","http://update.rainbowsoft.org/info/")
Call BlogConfig.Write("ZC_MULTI_DOMAIN_SUPPORT",False)




'留言评论
Call BlogConfig.Write("ZC_COMMENT_TURNOFF",False)
Call BlogConfig.Write("ZC_COMMENT_VERIFY_ENABLE",False)
Call BlogConfig.Write("ZC_COMMENT_NOFOLLOW_ENABLE",True)
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





Call BlogConfig.Write("ZC_IMAGE_WIDTH",520)

Call BlogConfig.Write("ZC_USE_NAVIGATE_ARTICLE",True)

Call BlogConfig.Write("ZC_RSS_EXPORT_WHOLE",False)




'后台管理
Call BlogConfig.Write("ZC_MANAGE_COUNT",50)
Call BlogConfig.Write("ZC_REBUILD_FILE_COUNT",50)
Call BlogConfig.Write("ZC_REBUILD_FILE_INTERVAL",1)










'UBB转换
Call BlogConfig.Write("ZC_UBB_ENABLE",False)
Call BlogConfig.Write("ZC_UBB_LINK_ENABLE",True)
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
Call BlogConfig.Write("ZC_BLOG_VERSION","2.0 Beta Build 120819")



'用户名,密码,评论长度等限制
Call BlogConfig.Write("ZC_USERNAME_MIN",4)
Call BlogConfig.Write("ZC_USERNAME_MAX",14)
Call BlogConfig.Write("ZC_PASSWORD_MIN",8)
Call BlogConfig.Write("ZC_PASSWORD_MAX",14)
Call BlogConfig.Write("ZC_EMAIL_MAX",30)
Call BlogConfig.Write("ZC_HOMEPAGE_MAX",100)
Call BlogConfig.Write("ZC_CONTENT_MAX",1000)










'---------------------------------静态化配置-----------------------------------


'{asp html shtml}
Call BlogConfig.Write("ZC_STATIC_TYPE","html")

Call BlogConfig.Write("ZC_STATIC_DIRECTORY","post")

Call BlogConfig.Write("ZC_TEMPLATE_DIRECTORY","template")



'ACTIVE MIX REWRITE
Call BlogConfig.Write("ZC_STATIC_MODE","ACTIVE")

Call BlogConfig.Write("ZC_ARTICLE_REGEX","{%host%}/{%post%}/{%alias%}.html")
Call BlogConfig.Write("ZC_PAGE_REGEX","{%host%}/{%alias%}.html")
Call BlogConfig.Write("ZC_CATEGORY_REGEX","{%host%}/catalog.asp?cate={%id%}")
Call BlogConfig.Write("ZC_USER_REGEX","{%host%}/catalog.asp?user={%id%}")
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
'相关文章
Call BlogConfig.Write("ZC_WAP_MUTUALITY",True)
'数量
Call BlogConfig.Write("ZC_WAP_MUTUALITY_LIMIT",5)

'Response.Write BlogConfig.Count
'Response.Write BlogConfig.Count
BlogConfig.Save


End Function










Function InsertArticleAndPage()

Set Categorys(0)=New TCategory

Dim a

Set a = New TArticle
a.AuthorID=1
a.CateID=0
a.id=0
a.Title="Hello, world!"
a.FType=0
a.Content="<p>欢迎使用Z-Blog,这是程序自动生成的文章.您可以删除或是编辑它,在没有进行&quot;文件重建&quot;前,无法打开该文章页面的,这不是故障:)</p><p>系统总共生成了一个&quot;留言本&quot;页面,和一个&quot;Hello, world!&quot;文章,祝您使用愉快!</p>"'<p>默认管理员账号和密码为:zblogger.</p>"
a.Intro=a.Content
a.Level=4
a.post
Set a=Nothing

Set a = New TArticle
a.AuthorID=1
a.CateID=0
a.id=0
a.title="留言本"
a.FType=1
a.Content="<p>这是我的留言本,欢迎给我留言.</p>"
a.Level=4
a.Alias="guestbook"
a.post
Set a=Nothing



End Function




Function SaveConfigs()

	On Error Resume Next
	Dim a
	For Each a In BlogConfig.Meta.Names
		Call Execute("Call BlogConfig.Write("""&a&""","&a&")")
	Next

	Call BlogConfig.Save()
	Err.Clear


	Call SaveConfig2Option()

End Function

%>
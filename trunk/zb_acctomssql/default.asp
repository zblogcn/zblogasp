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
<!-- #include file="function.asp" -->

<%Const num="num"
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
    <style type="text/css">
	#loading, #loading2, #loading3 {position:absolute;top:1px;right:1px;width:380px;margin:0;padding:5px 10px;background:rgb(0,102,204);color:#FFFFFF;font-size:12px;}
#loading a, #loading2 a, #loading3 a {color:white;}</style>
    <script>
		var dbtype="<%=dbtype%>";
		var dbpath="<%=dbpath%>";
		var dbserver="<%=dbserver%>";
		var dbname="<%=dbname%>";
		var dbusername="<%=dbusername%>";
		var dbpassword="<%=dbpassword%>";
    </script>
    
	<script language="JavaScript" type="text/javascript">
    function setup2(){
        if($("#dbtype").val()=="mssql"){
            if($("#dbserver").val()==""){alert('数据库服务器需要填写.');return false;};
            if($("#dbname").val()==""){alert('数据库名称需要填写.');return false;};
            if($("#dbusername").val()==""){alert('数据库用户名需要填写.');return false;};
            return true;
        }
    }
    function createtable(){
        if(setup2()){
            $("#db").html("wait...");
            $.post("?step=3&t=c",{"dbtype":"mssql","dbserver":$("#dbserver").val(),"dbname":$("#dbname").val(),"dbusername":$("#dbusername").val(),"dbpassword":$("#dbpassword").val()},function(data){$("#db").html(data)})
        }
    }
    function redirect(step,r){
        if(r){
            location.href="?step="+step+"&dbtype="+dbtype+"&dbserver="+dbserver+"&dbname="+dbname+"&dbusername="+dbusername+"&dbpassword="+dbpassword
        }
        else{
            return "?step="+step+"&dbtype="+dbtype+"&dbserver="+dbserver+"&dbname="+dbname+"&dbusername="+dbusername+"&dbpassword="+dbpassword
        }
    }
    </script>
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
Case 5 Call Setup5
Case 10 Call Setup10
End  Select
%>
  </form></div>

</body>
</html>
<%
Function Setup1()
Response.Redirect "?step=2"
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
<input type="hidden" name="dbtype" id="dbtype" value="mssql" />

<div id='mssql'>
<p><b>MSSQL地址:</b><input type="text" name="dbserver" id="dbserver" value="(local)" style='width:350px;' /></p>
<p><b>MSSQL库名:</b><input type="text" name="dbname" id="dbname" value="zblog" style='width:350px;' /></p>
<p><b>MSSQL帐号:</b><input type="text" name="dbusername" id="dbusername" value="sa" style='width:350px;' /></p>
<p><b>MSSQL密码:</b><input type="text" name="dbpassword" id="dbpassword" value="" style='width:350px;' /></p>
<p><a href='javascript:' onclick='createtable()'>如果您对MSSQL有管理权限，您可以直接点击这里新建数据库</a></p>
</div>
<div id='db'></div>
<div id='bottom'>
<input type="submit" name="next" id="netx" onClick="return setup2()" value="下一步" />
</div>
</dd>
</dl>
<%
End Function













Function Setup3()
'On Error Resume Next
%>
<dl>
<dd id="ddleft">
<img src='../zb_system/image/admin/install.png' alt='' />
<p>安装进度:<span><font color='#3d69aa'>██████████████</font><font color='white'>███</font></span></p>
<p><b>安装协议</b>&nbsp;&nbsp;»&nbsp;&nbsp;<b>数据库建立与设置</b>&nbsp;&nbsp;»&nbsp;&nbsp;<b>安装结果</b></p>
</dd>
<dd id="ddright">

<div id='title'>安装结果</div>
<div id='content'>
<%




Dim isCreate
isCreate=IIf(Request.QueryString("t")="c",True,False)

If dbtype="mssql" Then
	ZC_MSSQL_DATABASE=dbname
	ZC_MSSQL_USERNAME=dbusername
	ZC_MSSQL_PASSWORD=dbpassword
	ZC_MSSQL_SERVER=dbserver
	ZC_MSSQL_ENABLE=True
End If

If isCreate Then 
	Response.Clear
	
	If OpenConnect2(1)=False Then
		Response.Write("<p>抱歉，连接数据库失败！</p>"&IIf(ZC_MSSQL_ENABLE,"<p>您提供的数据库用户名和密码可能不正确，或者无法连接到 "&ZC_MSSQL_SERVER&" 上的数据库服务器，这意味着您的主机数据库服务器已停止工作。</p><p><ul><li>您确认您提供的用户名和密码正确么？</li><li>您确认您提供的主机名正确么？</li><li>您确认数据库服务器运行正常么？</li><li>您确认您购买的数据库是MSSQL而不是MYSQL么？</li></ul></p>","")&"<p>请您联系您的空间商。</div>")
		Response.End
	Else
		If isCreate Then
			On Error Resume Next
			objConn.Execute "CREATE DATABASE ["&dbname&"]"
			If Err.Number=0 Then Response.Write "创建数据表"&dbname&"成功" Else Response.Write dbname&"已存在或没有建表权限！"
			Response.End
		End If
	End If
	Response.End
End If
CloseConnect
OpenConnect2 0
server.scripttimeout=100000
Response.Write "<p>正在为您创建数据表..</p>"
If CreateMssqlTable() Then
	Response.Write "<p>数据表创建成功！五秒钟后将开始导入数据！</p>" 
	Response.Write "<script>setTimeout(""redirect(4,true)"",5000)</sc"&"ript>"
Else
	Response.Write "<p>数据表创建失败！可能您已经创建过了！</p>"
	Response.Write "<a href='javascript:redirect(4,true)'>继续</a>"
End If

Response.End

%>






</div>
<div id='bottom'>
<input type="button" name="next" onClick="redirect(4,true)'" id="netx" value="下一步" />
</div>


</dd>
</dl>
<%
End Function

Function Setup4()
'On Error Resume Next
%>
<dl>
<dd id="ddleft">
<img src='../zb_system/image/admin/install.png' alt='' />
<p>安装进度:<span><font color='#3d69aa'>███████████████</font><font color='white'>██</font></span></p>
<p><b>安装协议</b>&nbsp;&nbsp;»&nbsp;&nbsp;<b>数据库建立与设置</b>&nbsp;&nbsp;»&nbsp;&nbsp;<b>安装结果</b></p>
</dd>
<dd id="ddright">

<div id='title'>安装结果</div>
<div id='content'>
<%
If dbtype="mssql" Then
	ZC_MSSQL_DATABASE=dbname
	ZC_MSSQL_USERNAME=dbusername
	ZC_MSSQL_PASSWORD=dbpassword
	ZC_MSSQL_SERVER=dbserver
	ZC_MSSQL_ENABLE=True
End If
CloseConnect
OpenConnect2 0
server.scripttimeout=100000
Response.Write "<p>正在尝试使用OpenDataSource导入数据..</p>"
On Error Resume Next
'Response.Write 
objConn.Execute Replace(Replace(tempString,"<#MDBPath#>",BlogPath & ZC_DATABASE_PATH),"\n",vbCrlf)
If Err.Number=&H80040e14 Then
	Response.Write "<p>服务器可能不允许使用OpenDataSource，这个函数不安全，但是可以让您快速完成导入过程。</p>"
	Response.Write "<p>您可以不打开此函数，但是速度可能会变慢。</p>"
	Response.Write "<p>是否打开此函数？&nbsp;&nbsp;<span><a href='javascript:$(""#showhs"").show()'>是</a>&nbsp;&nbsp;<a href='javascript:redirect(5,true)'>不打开或无法打开</a></span></p>"
	Response.Write "<div id='showhs' style='display:none'><p>如果您是系统管理员，您可以打开SQL Server Management Studio连接到数据库后，点击新建查询，输入以下代码并执行后刷新本页面。</p>"
%><textarea style="border:1px solid;height:130px">exec sp_configure 'show advanced options',1  
 
reconfigure  
 
exec sp_configure 'Ad Hoc Distributed Queries',1  
 
reconfigure </textarea>
<p>执行完毕以后，再输入以下代码关闭该函数：</p>
<textarea style="border:1px solid;height:130px">exec sp_configure 'Ad Hoc Distributed Queries',0  
 
reconfigure  
 
exec sp_configure 'show advanced options',0  
 
reconfigure </textarea>
</div></dd></dl><%

ElseIf Err.Number=0 Then 
	Response.Write "<p>导入成功！5秒后进行最后配置！</p>"
	Response.Write "<script>setTimeout(5,""redirect(10,true)"")</sc"&"ript>"
	Response.Write "</div><div id='bottom'><input type=""button"" name=""next"" onClick=""redirect(10,true)'"" id=""netx"" value=""下一步"" /></div></dd></dl>"
End If

Response.End
End Function



Function Setup5()
'On Error Resume Next
%>
<dl>
<dd id="ddleft">
<img src='../zb_system/image/admin/install.png' alt='' />
<p>安装进度:<span><font color='#3d69aa'>███████████████</font><font color='white'>██</font></span></p>
<p><b>安装协议</b>&nbsp;&nbsp;»&nbsp;&nbsp;<b>数据库建立与设置</b>&nbsp;&nbsp;»&nbsp;&nbsp;<b>安装结果</b></p>
</dd>
<dd id="ddright">

<div id='title'>安装结果</div>
<div id='content'><p id="loading" style="display:none"></p>
<%
If dbtype="mssql" Then
	ZC_MSSQL_DATABASE=dbname
	ZC_MSSQL_USERNAME=dbusername
	ZC_MSSQL_PASSWORD=dbpassword
	ZC_MSSQL_SERVER=dbserver
	ZC_MSSQL_ENABLE=True
End If
CloseConnect
OpenConnect2 0
server.scripttimeout=100000
Response.Write "<p>正在尝试使用循环导入数据..</p><script>$(""#loading"").show()</script>"

Dim intCount,objConn2,objRs2
Set objConn2=Server.CreateObject("adodb.connection")
objConn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BlogPath & ZC_DATABASE_PATH
Set objRs2=Server.CreateObject("adodb.recordset")
ExportLog "清空目标MSSQL数据表.."
objConn.execute replace("DELETE FROM [blog_Article]  \nDELETE FROM [blog_Category] \nDELETE FROM [blog_Comment] \nDELETE FROM [blog_Counter]  \nDELETE FROM [blog_Keyword]  \nDELETE FROM [blog_Member]\nDELETE FROM [blog_Tag]\nDELETE FROM [blog_TrackBack]\nDELETE FROM [blog_Upload]\nDELETE FROM [blog_Function]\nDELETE FROM [blog_Config]","\n",vbcrlf)

ExportLog "开始复制blog_Article"
makeloading "正在复制blog_Article..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_Article]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_Article"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_Article] On "&vbcrlf)&"INSERT INTO [blog_Article]([log_ID],[log_CateID],[log_AuthorID],[log_Level],[log_Url],[log_Title],[log_Intro],[log_Content],[log_IP],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Tag],[log_IsTop],[log_Yea],[log_Nay],[log_Ratting],[log_Template],[log_FullUrl],[log_Type],[log_Meta]) VALUES ("&FilterSQL2(objRs2("log_ID"),num)&","&FilterSQL2(objRs2("log_CateID"),num)&","&FilterSQL2(objRs2("log_AuthorID"),num)&","&FilterSQL2(objRs2("log_Level"),"")&","&FilterSQL2(objRs2("log_Url"),"")&","&FilterSQL2(objRs2("log_Title"),"")&","&FilterSQL2(objRs2("log_Intro"),"")&","&FilterSQL2(objRs2("log_Content"),"")&","&FilterSQL2(objRs2("log_IP"),"")&","&FilterSQL2(objRs2("log_PostTime"),"")&","&FilterSQL2(objRs2("log_CommNums"),"")&","&FilterSQL2(objRs2("log_ViewNums"),"")&","&FilterSQL2(objRs2("log_TrackBackNums"),"")&","&FilterSQL2(objRs2("log_Tag"),"")&","&FilterSQL2(objRs2("log_IsTop"),"")&","&FilterSQL2(objRs2("log_Yea"),"")&","&FilterSQL2(objRs2("log_Nay"),"")&","&FilterSQL2(objRs2("log_Ratting"),"")&","&FilterSQL2(objRs2("log_Template"),"")&","&FilterSQL2(objRs2("log_FullUrl"),"")&","&FilterSQL2(objRs2("log_Type"),"")&","&FilterSQL2(objRs2("log_Meta"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_Article..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_Article"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_Article] Off"
loop
ExportLog "blog_Article复制成功"
objRs2.close
intCount=1
CleanLog objConn

ExportLog "开始复制blog_Category"
makeloading "正在复制blog_Category..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_Category]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_Category"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_Category] On "&vbcrlf)&"INSERT INTO [blog_Category]([cate_ID],[cate_Name],[cate_Order],[cate_Intro],[cate_Count],[cate_URL],[cate_ParentID],[cate_Template],[cate_LogTemplate],[cate_FullUrl],[cate_Meta]) VALUES ("&FilterSQL2(objRs2("cate_ID"),num)&","&FilterSQL2(objRs2("cate_Name"),"")&","&FilterSQL2(objRs2("cate_Order"),"")&","&FilterSQL2(objRs2("cate_Intro"),"")&","&FilterSQL2(objRs2("cate_Count"),"")&","&FilterSQL2(objRs2("cate_URL"),"")&","&FilterSQL2(objRs2("cate_ParentID"),num)&","&FilterSQL2(objRs2("cate_Template"),"")&","&FilterSQL2(objRs2("cate_LogTemplate"),"")&","&FilterSQL2(objRs2("cate_FullUrl"),"")&","&FilterSQL2(objRs2("cate_Meta"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_Category..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_Category"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_Category] Off"
loop
ExportLog "blog_Category复制成功"
objRs2.close
intCount=1
CleanLog objConn

ExportLog "开始复制blog_Comment"
makeloading "正在复制blog_Comment..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_Comment]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_Comment"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_Comment] On "&vbcrlf)&"INSERT INTO [blog_Comment]([comm_ID],[log_ID],[comm_AuthorID],[comm_Author],[comm_Content],[comm_Email],[comm_HomePage],[comm_PostTime],[comm_IP],[comm_Agent],[comm_Reply],[comm_LastReplyIP],[comm_LastReplyTime],[comm_Yea],[comm_Nay],[comm_Ratting],[comm_ParentID],[comm_IsCheck],[comm_Meta]) VALUES ("&FilterSQL2(objRs2("comm_ID"),num)&","&FilterSQL2(objRs2("log_ID"),num)&","&FilterSQL2(objRs2("comm_AuthorID"),num)&","&FilterSQL2(objRs2("comm_Author"),"")&","&FilterSQL2(objRs2("comm_Content"),"")&","&FilterSQL2(objRs2("comm_Email"),"")&","&FilterSQL2(objRs2("comm_HomePage"),"")&","&FilterSQL2(objRs2("comm_PostTime"),"")&","&FilterSQL2(objRs2("comm_IP"),"")&","&FilterSQL2(objRs2("comm_Agent"),"")&","&FilterSQL2(objRs2("comm_Reply"),"")&","&FilterSQL2(objRs2("comm_LastReplyIP"),"")&","&FilterSQL2(objRs2("comm_LastReplyTime"),"")&","&FilterSQL2(objRs2("comm_Yea"),"")&","&FilterSQL2(objRs2("comm_Nay"),"")&","&FilterSQL2(objRs2("comm_Ratting"),"")&","&FilterSQL2(objRs2("comm_ParentID"),num)&","&FilterSQL2(objRs2("comm_IsCheck"),"")&","&FilterSQL2(objRs2("comm_Meta"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_Comment..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_Comment"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_Comment] Off"
loop
ExportLog "blog_Comment复制成功"
objRs2.close
intCount=1
CleanLog objConn

ExportLog "开始复制blog_Config"
makeloading "正在复制blog_Config..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_Config]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_Config"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_Config] On "&vbcrlf)&"INSERT INTO [blog_Config]([conf_Name],[conf_Value]) VALUES ("&FilterSQL2(objRs2("conf_Name"),"")&","&FilterSQL2(objRs2("conf_Value"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_Config..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_Config"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_Config] Off"
loop
ExportLog "blog_Config复制成功"
objRs2.close
intCount=1
CleanLog objConn

ExportLog "开始复制blog_Counter"
makeloading "正在复制blog_Counter..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_Counter]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_Counter"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_Counter] On "&vbcrlf)&"INSERT INTO [blog_Counter]([coun_ID],[coun_IP],[coun_Agent],[coun_Refer],[coun_PostTime],[coun_Content],[coun_UserID],[coun_PostData],[coun_URL],[coun_AllRequestHeader],[coun_LogName]) VALUES ("&FilterSQL2(objRs2("coun_ID"),num)&","&FilterSQL2(objRs2("coun_IP"),"")&","&FilterSQL2(objRs2("coun_Agent"),"")&","&FilterSQL2(objRs2("coun_Refer"),"")&","&FilterSQL2(objRs2("coun_PostTime"),"")&","&FilterSQL2(objRs2("coun_Content"),"")&","&FilterSQL2(objRs2("coun_UserID"),num)&","&FilterSQL2(objRs2("coun_PostData"),"")&","&FilterSQL2(objRs2("coun_URL"),"")&","&FilterSQL2(objRs2("coun_AllRequestHeader"),"")&","&FilterSQL2(objRs2("coun_LogName"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_Counter..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_Counter"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_Counter] Off"
loop
ExportLog "blog_Counter复制成功"
objRs2.close
intCount=1
CleanLog objConn

ExportLog "开始复制blog_Function"
makeloading "正在复制blog_Function..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_Function]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_Function"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_Function] On "&vbcrlf)&"INSERT INTO [blog_Function]([fn_ID],[fn_Name],[fn_FileName],[fn_Order],[fn_Content],[fn_IsSystem],[fn_SidebarID],[fn_HtmlID],[fn_Ftype],[fn_MaxLi],[fn_Meta]) VALUES ("&FilterSQL2(objRs2("fn_ID"),num)&","&FilterSQL2(objRs2("fn_Name"),"")&","&FilterSQL2(objRs2("fn_FileName"),"")&","&FilterSQL2(objRs2("fn_Order"),"")&","&FilterSQL2(objRs2("fn_Content"),"")&","&FilterSQL2(objRs2("fn_IsSystem"),"")&","&FilterSQL2(objRs2("fn_SidebarID"),num)&","&FilterSQL2(objRs2("fn_HtmlID"),num)&","&FilterSQL2(objRs2("fn_Ftype"),"")&","&FilterSQL2(objRs2("fn_MaxLi"),"")&","&FilterSQL2(objRs2("fn_Meta"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_Function..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_Function"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_Function] Off"
loop
ExportLog "blog_Function复制成功"
objRs2.close
intCount=1
CleanLog objConn

ExportLog "开始复制blog_Keyword"
makeloading "正在复制blog_Keyword..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_Keyword]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_Keyword"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_Keyword] On "&vbcrlf)&"INSERT INTO [blog_Keyword]([key_ID],[key_Name],[key_Intro],[key_URL]) VALUES ("&FilterSQL2(objRs2("key_ID"),num)&","&FilterSQL2(objRs2("key_Name"),"")&","&FilterSQL2(objRs2("key_Intro"),"")&","&FilterSQL2(objRs2("key_URL"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_Keyword..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_Keyword"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_Keyword] Off"
loop
ExportLog "blog_Keyword复制成功"
objRs2.close
intCount=1
CleanLog objConn

ExportLog "开始复制blog_Member"
makeloading "正在复制blog_Member..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_Member]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_Member"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_Member] On "&vbcrlf)&"INSERT INTO [blog_Member]([mem_ID],[mem_Level],[mem_Name],[mem_Password],[mem_Sex],[mem_Email],[mem_MSN],[mem_QQ],[mem_HomePage],[mem_LastVisit],[mem_Status],[mem_PostLogs],[mem_PostComms],[mem_Intro],[mem_IP],[mem_Count],[mem_Template],[mem_FullUrl],[mem_Guid],[mem_Meta]) VALUES ("&FilterSQL2(objRs2("mem_ID"),num)&","&FilterSQL2(objRs2("mem_Level"),"")&","&FilterSQL2(objRs2("mem_Name"),"")&","&FilterSQL2(objRs2("mem_Password"),"")&","&FilterSQL2(objRs2("mem_Sex"),"")&","&FilterSQL2(objRs2("mem_Email"),"")&","&FilterSQL2(objRs2("mem_MSN"),"")&","&FilterSQL2(objRs2("mem_QQ"),"")&","&FilterSQL2(objRs2("mem_HomePage"),"")&","&FilterSQL2(objRs2("mem_LastVisit"),"")&","&FilterSQL2(objRs2("mem_Status"),"")&","&FilterSQL2(objRs2("mem_PostLogs"),"")&","&FilterSQL2(objRs2("mem_PostComms"),"")&","&FilterSQL2(objRs2("mem_Intro"),"")&","&FilterSQL2(objRs2("mem_IP"),"")&","&FilterSQL2(objRs2("mem_Count"),"")&","&FilterSQL2(objRs2("mem_Template"),"")&","&FilterSQL2(objRs2("mem_FullUrl"),"")&","&FilterSQL2(objRs2("mem_Guid"),"")&","&FilterSQL2(objRs2("mem_Meta"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_Member..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_Member"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_Member] Off"
loop
ExportLog "blog_Member复制成功"
objRs2.close
intCount=1
CleanLog objConn

ExportLog "开始复制blog_Tag"
makeloading "正在复制blog_Tag..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_Tag]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_Tag"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_Tag] On "&vbcrlf)&"INSERT INTO [blog_Tag]([tag_ID],[tag_Name],[tag_Intro],[tag_ParentID],[tag_URL],[tag_Order],[tag_Count],[tag_Template],[tag_FullUrl],[tag_Meta]) VALUES ("&FilterSQL2(objRs2("tag_ID"),num)&","&FilterSQL2(objRs2("tag_Name"),"")&","&FilterSQL2(objRs2("tag_Intro"),"")&","&FilterSQL2(objRs2("tag_ParentID"),num)&","&FilterSQL2(objRs2("tag_URL"),"")&","&FilterSQL2(objRs2("tag_Order"),"")&","&FilterSQL2(objRs2("tag_Count"),"")&","&FilterSQL2(objRs2("tag_Template"),"")&","&FilterSQL2(objRs2("tag_FullUrl"),"")&","&FilterSQL2(objRs2("tag_Meta"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_Tag..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_Tag"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_Tag] Off"
loop
ExportLog "blog_Tag复制成功"
objRs2.close
intCount=1
CleanLog objConn

ExportLog "开始复制blog_TrackBack"
makeloading "正在复制blog_TrackBack..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_TrackBack]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_TrackBack"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_TrackBack] On "&vbcrlf)&"INSERT INTO [blog_TrackBack]([tb_ID],[log_ID],[tb_URL],[tb_Title],[tb_Blog],[tb_Excerpt],[tb_PostTime],[tb_IP],[tb_Agent],[tb_Meta]) VALUES ("&FilterSQL2(objRs2("tb_ID"),num)&","&FilterSQL2(objRs2("log_ID"),num)&","&FilterSQL2(objRs2("tb_URL"),"")&","&FilterSQL2(objRs2("tb_Title"),"")&","&FilterSQL2(objRs2("tb_Blog"),"")&","&FilterSQL2(objRs2("tb_Excerpt"),"")&","&FilterSQL2(objRs2("tb_PostTime"),"")&","&FilterSQL2(objRs2("tb_IP"),"")&","&FilterSQL2(objRs2("tb_Agent"),"")&","&FilterSQL2(objRs2("tb_Meta"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_TrackBack..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_TrackBack"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_TrackBack] Off"
loop
ExportLog "blog_TrackBack复制成功"
objRs2.close
intCount=1
CleanLog objConn

ExportLog "开始复制blog_UpLoad"
makeloading "正在复制blog_UpLoad..<br/>当前进度：0%<br/>"
Server.ScriptTimeout=100000
objRs2.CursorType = 1
objRs2.LockType = 1
objRs2.ActiveConnection=objConn2
objRs2.Source="SELECT * FROM [blog_UpLoad]"
objRs2.open
if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
Do until objRs2.eof
	intCount=intCount+1
	objConn.execute IIf("blog_UpLoad"="blog_Config","","SET IDENTITY_INSERT [dbo].[blog_UpLoad] On "&vbcrlf)&"INSERT INTO [blog_UpLoad]([ul_ID],[ul_AuthorID],[ul_FileSize],[ul_FileName],[ul_PostTime],[ul_Quote],[ul_DownNum],[ul_FileIntro],[ul_DirByTime],[ul_Meta]) VALUES ("&FilterSQL2(objRs2("ul_ID"),num)&","&FilterSQL2(objRs2("ul_AuthorID"),num)&","&FilterSQL2(objRs2("ul_FileSize"),"")&","&FilterSQL2(objRs2("ul_FileName"),"")&","&FilterSQL2(objRs2("ul_PostTime"),"")&","&FilterSQL2(objRs2("ul_Quote"),"")&","&FilterSQL2(objRs2("ul_DownNum"),"")&","&FilterSQL2(objRs2("ul_FileIntro"),"")&","&FilterSQL2(objRs2("ul_DirByTime"),"")&","&FilterSQL2(objRs2("ul_Meta"),"")&")"
	objRs2.movenext
	if intCount=200 then 
		CleanLog(objConn)
		MakeLoading "正在复制blog_UpLoad..<br/>当前进度："&Formatnumber(objRs2.AbsolutePosition/objRs2.recordcount,2)*100&"%<br/>"
		intCount=1
	end if
	if err.number<>0 then ExportErr "出现错误：" & Err.Number & Err.Description:Response.End
	If "blog_UpLoad"<>"blog_Config" Then objConn.Execute "SET IDENTITY_INSERT [dbo].[blog_UpLoad] Off"
loop
ExportLog "blog_UpLoad复制成功"
objRs2.close
intCount=1
CleanLog objConn



ExportLog "复制成功！5s后配置程序！<script>$('#loading').hide();setTimeout(""redirect(10,true)"",5000)</script>"

Response.Write "</div><div id='bottom'><input type=""button"" name=""next"" onClick=""redirect(10,true)"" id=""netx"" value=""下一步"" /></div></dd></dl>"



Response.End
End Function



Function Setup10
	Call System_Initialize
	BlogConfig.Load("Blog")
	BlogConfig.Write "ZC_MSSQL_DATABASE",dbname
	BlogConfig.Write "ZC_MSSQL_USERNAME",dbusername
	BlogConfig.Write "ZC_MSSQL_PASSWORD",dbpassword
	BlogConfig.Write "ZC_MSSQL_SERVER",dbserver
	BlogConfig.Write "ZC_MSSQL_ENABLE",True
	BlogConfig.Save
	
	SaveConfig2Option
	Call CloseConnect
	SetBlogHint_Custom "转换成功！"
	Response.Redirect BlogHost
End Function







Function CreateMssqlTable()

	On Error Resume Next
	
	CreateMssqlTable=False
	
	objConn.BeginTrans

	objConn.execute("CREATE TABLE [blog_Tag] (tag_ID int identity(1,1) not null primary key,tag_Name nvarchar(255) default '',tag_Intro ntext default '',tag_ParentID int default 0,tag_URL nvarchar(255) default '',tag_Order int default 0,tag_Count int default 0,tag_Template nvarchar(50) default '',tag_FullUrl nvarchar(255) default '',tag_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_Article] (log_ID int identity(1,1) not null primary key,log_CateID int default 0,log_AuthorID int default 0,log_Level int default 0,log_Url nvarchar(255) default '',log_Title nvarchar(255) default '',log_Intro ntext default '',log_Content ntext default '',log_IP nvarchar(15) default '',log_PostTime datetime default getdate(),log_CommNums int default 0,log_ViewNums int default 0,log_TrackBackNums int default 0,log_Tag nvarchar(255) default '',log_IsTop bit DEFAULT 0,log_Yea int default 0,log_Nay int default 0,log_Ratting int default 0,log_Template nvarchar(50) default '',log_FullUrl nvarchar(255) default '',log_Type int default 0,log_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_Category] (cate_ID int identity(1,1) not null primary key,cate_Name nvarchar(50) default '',cate_Order int default 0,cate_Intro ntext default '',cate_Count int default 0,cate_URL nvarchar(255) default '',cate_ParentID int default 0,cate_Template nvarchar(50) default '',cate_LogTemplate nvarchar(50) default '',cate_FullUrl nvarchar(255) default '',cate_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_Comment] (comm_ID int identity(1,1) not null primary key,log_ID int default 0,comm_AuthorID int default 0,comm_Author nvarchar(20) default '',comm_Content ntext default '',comm_Email nvarchar(50) default '',comm_HomePage nvarchar(255) default '',comm_PostTime datetime default getdate(),comm_IP nvarchar(15) default '',comm_Agent ntext default '',comm_Reply ntext default '',comm_LastReplyIP nvarchar(15) default '',comm_LastReplyTime datetime default getdate(),comm_Yea int default 0,comm_Nay int default 0,comm_Ratting int default 0,comm_ParentID int default 0,comm_IsCheck bit default 0,comm_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_TrackBack] (tb_ID int identity(1,1) not null primary key,log_ID int default 0,tb_URL nvarchar(255) default '',tb_Title nvarchar(100) default '',tb_Blog nvarchar(50) default '',tb_Excerpt ntext default '',tb_PostTime datetime default getdate(),tb_IP nvarchar(15) default '',tb_Agent ntext default '',tb_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_UpLoad] (ul_ID int identity(1,1) not null primary key,ul_AuthorID int default 0,ul_FileSize int default 0,ul_FileName nvarchar(255) default '',ul_PostTime datetime default getdate(),ul_Quote nvarchar(255) default '',ul_DownNum int default 0,ul_FileIntro nvarchar(255) default '',ul_DirByTime bit DEFAULT 0,ul_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_Counter] (coun_ID int identity(1,1) not null primary key,coun_IP nvarchar(15) default '',coun_Agent ntext default '',coun_Refer nvarchar(255) default '',coun_PostTime datetime default getdate(),coun_Content ntext default '',coun_UserID int default 0,coun_PostData ntext default '',coun_URL ntext default '',coun_AllRequestHeader ntext default '',coun_LogName ntext default '')")


	objConn.execute("CREATE TABLE [blog_Keyword] (key_ID int identity(1,1) not null primary key,key_Name nvarchar(255) default '',key_Intro ntext default '',key_URL nvarchar(255) default '')")

	objConn.execute("CREATE TABLE [blog_Member] (mem_ID int identity(1,1) not null primary key,mem_Level int default 0,mem_Name nvarchar(20) default '',mem_Password nvarchar(32) default '',mem_Sex int default 0,mem_Email nvarchar(50) default '',mem_MSN nvarchar(50) default '',mem_QQ nvarchar(50) default '',mem_HomePage nvarchar(255) default '',mem_LastVisit datetime default getdate(),mem_Status int default 0,mem_PostLogs int default 0,mem_PostComms int default 0,mem_Intro ntext default '',mem_IP nvarchar(15) default '',mem_Count int default 0,mem_Template nvarchar(50) default '',mem_FullUrl nvarchar(255) default '',mem_Url nvarchar(255) default '',mem_Guid  nvarchar(36) default '',mem_Meta ntext default '')")

	objConn.execute("CREATE TABLE [blog_Config] (conf_Name nvarchar(255) not null default '',conf_Value text default '')")

	objConn.execute("CREATE TABLE [blog_Function] (fn_ID int identity(1,1) not null primary key,fn_Name nvarchar(50) default '',fn_FileName nvarchar(50) default '',fn_Order int default 0,fn_Content ntext default '',fn_IsSystem bit DEFAULT 0,fn_SidebarID int default 0,fn_HtmlID nvarchar(50) default '',fn_Ftype nvarchar(5) default '',fn_MaxLi int default 0,fn_Meta ntext default '')")

	objConn.CommitTrans

	If Err.Number=0 Then CreateMssqlTable=True
End Function






%>
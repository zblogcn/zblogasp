<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:   
'// 开始时间:   
'// 最后修改:    
'// 备    注:   
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn" lang="zh-cn">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="zh-cn" />
	<meta http-equiv="pragma" content="no-cache">
	<meta http-equiv="cache-control" content="no-cache,must-revalidate">
	<meta http-equiv="expires" content="0">
	<title>Z-Blog密码重置工具</title>
<style type="text/css">
<!--
*{
	font-size:14px;
}
body{
	margin:0;
	padding:0;
	color: #FFFFFF;
	font-size:12px;
	background:#FFFFFF;
	font-family:"微软雅黑","黑体","宋体";
}
h1,h2,h3,h4,h5,h6{
	font-size:18px;
	padding:0;
	margin:0;
}
a{
	text-decoration: none;
}
a:link {
	color:#FFFFFF;
	text-decoration: none;
}
a:visited {
	color:#FFFFFF;
	text-decoration: none;
}
a:hover {
	color:yellow;
	text-decoration: underline;
}
a:active {
	color:yellow;
	text-decoration: underline;
}
p{
	margin:0;
	padding:5px;
}
table {
	border-collapse: collapse;
	border:0px solid #333333;
	background:#ffffff;
	margin-top:10px;
}
td{
	border:0px solid #333333;
	margin:0;
	padding:3px;
}
img{
	border:0;
}
hr{
	border:0px;
	border-top:1px solid #666666;
	background:#666666;
	margin:2px 0 4px 0;
	padding:0;
	height:0px;
}
img{
	margin:0;
	padding:0;
}
form{
	margin:0;
	padding:0;
}


#frmLogin{
	position:absolute;
	left: 50%;
	top: 40%;
	margin: -150px 0px 0px -300px;
	padding:0;
	overflow:hidden;
	width:600px;
	height:400px;
	background-color:#3a6ea5;
	border:0px solid #B3C3CD;
	box-shadow: 0px 0px 15px black;
}

#frmLogin h3{
	padding:15px 0 5px 0;
	margin:0;
	text-align:center;
	color:white;
	font-size:24px;
	height:30px;
}

#divHeader{
	margin:0 0;
	padding:8px;
}
#divMain{
	height:280px;
}
#divFooter{
	margin:5px 0px 0 0px;
	text-align:center;
	padding:2px;
}

#divMain_Top{
	padding:8px;
	padding-bottom:0;
}
#divMain_Center{
	padding:5px;
}
#divMain_Bottom{
	text-align:right;
	padding:5px;
}
#txaContent{
	border:1px solid #A1B0B9;
	background:#FFFFFF;
}
-->
</style>
</head>
<body>


<form id="frmLogin" method="post">
<h3>Z-Blog免输入密码登陆工具</h3>
<div id="divHeader">&nbsp;&nbsp;<a href="http://www.zblogcn.com/" target="_blank">Z-Blog主页</a>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://bbs.zblogcn.com" class="here" target="_blank">Zblogger社区</a>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://wiki.zblogcn.com/" target="_blank">Z-Wiki</a>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://blog.zblogcn.com/" target="_blank">菠萝阁</a>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://www.dbshost.cn/" target="_blank">DBS主机</a></div>
<div id="divMain">
<input type="hidden" name="userid" id="userid" value="0" />
<%
	Call System_Initialize()

	Call GetUser()

Dim objRS,i,j
Set objRS=Server.CreateObject("ADODB.Recordset")
objRS.CursorType = adOpenKeyset
objRS.LockType = adLockReadOnly
objRS.ActiveConnection=objConn


If CInt(Request.Form("userid"))>0 Then

	Response.Cookies("username") = vbsescape(Users(CInt(Request.Form("userid"))).Name)
	Response.Cookies("username").Expires = DateAdd("y", 1, now)
	Response.Cookies("username").Path = "/"


	Response.Cookies("password") = Users(CInt(Request.Form("userid"))).PassWord
	Response.Cookies("password").Expires = DateAdd("y", 1, now)
	Response.Cookies("password").Path = "/"

	Response.Redirect "zb_system/cmd.asp?act=admin&nbsp;"

End If

		objRS.Source="SELECT * FROM [blog_Member] WHERE [mem_Level]<5"
		objRS.Open()
		If (Not objRS.bof) And (Not objRS.eof) Then
				For i=1 To objRS.RecordCount

					Response.Write "<p style='clear:both;padding:10px;'><span style='font-size:20px;float:left;'>"+ZVA_User_Level_Name(objRS("mem_level"))+":"+objRS("mem_Name")+"</span><input type='submit' value='登录' style='float:right;' onclick='document.getElementById(""userid"").value="&objRS("mem_ID")&"'></input></p>"

					objRS.MoveNext
					If objRS.eof Then Exit For
				Next
		End If
		objRS.Close()


	Call System_Terminate()
%>
</div>
<div id="divFooter"><b><font color="red">[使用必看]</font>&nbsp;<font color="yellow">注意:此工具非常危险,使用后请立刻通过<u>FTP</u>删除或改名.</font></b></div>
</form>
</body>
</html>
<%
If Err.Number<>0 then
	Call ShowError(0)
End If
%>
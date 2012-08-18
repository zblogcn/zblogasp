<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../zb_users/c_option.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../ZB_SYSTEM/function/c_system_event.asp" -->
<%

'Call System_Initialize
Dim Step,fso,path
path=BlogPath & "zb_users\data\"&Request.QueryString("mdb")
Step=Request.QueryString("Step")
If Not IsNumeric(Step) Or IsEmpty(Step) Then Step=1
Set fso=New a
Dim objRS,rndPwd,Guid
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-CN" lang="zh-CN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="zh-CN" />
<title>Z-Blog 数据库升级</title>
<link rel="stylesheet" rev="stylesheet" href="<%=GetCurrentHost%>ZB_SYSTEM/css/admin.css" type="text/css" media="screen" />
<script language="JavaScript" src="<%=GetCurrentHost%>ZB_SYSTEM/SCRIPT/common.js" type="text/javascript"></script>
<script language="JavaScript" src="<%=GetCurrentHost%>ZB_SYSTEM/SCRIPT/md5.js" type="text/javascript"></script>
</head>
<body class="short">
<div class="bg"></div>
<div id="wrapper">
  <div class="logo"><img src="<%=GetCurrentHost%>ZB_SYSTEM/image/admin/none.gif" title="Z-Blog<%=ZC_MSG009%>" alt="Z-Blog<%=ZC_MSG009%>"/></div>
  <div class="login">
    <div class="divHeader">升级[Step<%=Step%>]</div>
    <p>
<%select case Step
	case 1%>
  1.请将数据库直接拷贝到zb_users下的data文件夹<br/>2.将UPLOAD拷贝到zb_users文件夹下。</p>
<p>&nbsp;</p>
<p>注意：升级数据库可能会有未知的风险，在升级数据库之前请做好备份！<br/>
<br/>
<input type="button" style="width:100%" value="如果准备好了，请点击这里升级数据库" class="button" onClick="location.href='?step=2'"/>
<%case 2%>
<%
	Dim objFile,objFolder,i
	i=0
	Set objFolder=fso.b.GetFolder(path) 
	for each objFile in objFolder.files
		If LCase(Right(objFile.name,4))=".mdb" Then
			Response.write "<a href=""update.asp?step=3&mdb="&Server.URLEncode(objFile.name)&""">"&objFile.Name&"</a><br/>"
			i=i+1
		End If
    Next
	Response.Write "<br/>在"&path&"发现"&i&"个数据库，请选择需要升级的数据库。"
	%>
<%case 3%>
<%
	
	If fso.Exists(path) Then
		ExportOK "找到数据库"
		ZC_MSSQL_ENABLE=False
		ZC_DATABASE_PATH="zb_users\data\"&Request.QueryString("mdb")
		If OpenConnect Then
			ExportOK "初始化数据库成功！"
			If Not CheckUpdateDB("[log_IsTop]","[blog_Article]") Then
				Call UpdateDB("ALTER TABLE [blog_Article] ADD COLUMN [log_IsTop] YESNO DEFAULT FALSE","[blog_Article].[log_IsTop]")
				objConn.execute("UPDATE [blog_Article] SET [log_IsTop]=0")
			End If
			If Not CheckUpdateDB("[log_Tag]","[blog_Article]") Then
				Call UpdateDB("ALTER TABLE [blog_Article] ADD COLUMN [log_Tag] VARCHAR(255) default """"","[blog_Article].[log_Tag]")
			End If
			If Not CheckUpdateDB("[tag_ID]","[blog_Tag]") Then
				Call UpdateDB("CREATE TABLE [blog_Tag] (tag_ID AutoIncrement primary key,tag_Name VARCHAR(255) default """",tag_Intro text default """",tag_ParentID int DEFAULT 0,tag_URL VARCHAR(255) default """",tag_Order int DEFAULT 0,tag_Count int DEFAULT 0)","[blog_Tag]")
			End If
		
			If Not CheckUpdateDB("[coun_ID]","[blog_Counter]") Then
				Call UpdateDB("CREATE TABLE [blog_Counter] (coun_ID AutoIncrement primary key,coun_IP VARCHAR(20) default """",coun_Agent text default """",coun_Refer VARCHAR(255) default """",coun_PostTime TIME DEFAULT Now())","[blog_Counter]")
			End If
		
			If Not CheckUpdateDB("[key_ID]","[blog_Keyword]") Then
				Call UpdateDB("CREATE TABLE [blog_Keyword] (key_ID AutoIncrement primary key,key_Name VARCHAR(255) default """",key_Intro text default """",key_URL VARCHAR(255) default """")","[blog_KeyWord]")
			End If
		
			If Not CheckUpdateDB("[ul_Quote]","[blog_UpLoad]") Then
				Call UpdateDB("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_Quote] VARCHAR(255) default """"","[blog_Upload].[ul.Quote]")
				objConn.execute("UPDATE [blog_UpLoad] SET [ul_Quote]=''")
				Call UpdateDB("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_DownNum] int DEFAULT 0","[blog_Upload].[ul_DownNum]")
			End If
		
			If Not CheckUpdateDB("[ul_FileIntro]","[blog_UpLoad]") Then
				Call UpdateDB("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_FileIntro] VARCHAR(255) default """"","[blog_Upload].[ul_FileIntro]")
			End If
		
			If Not CheckUpdateDB("[ul_DirByTime]","[blog_UpLoad]") Then
				Call UpdateDB("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_DirByTime] YESNO DEFAULT FALSE","[blog_Upload].[ul_DirByTime]")
				objConn.execute("UPDATE [blog_UpLoad] SET [ul_DirByTime]=[ul_Quote]")
				objConn.execute("UPDATE [blog_UpLoad] SET [ul_Quote]=''")
			End If
		
			If Not CheckUpdateDB("[log_Meta]","[blog_Article]") Then
				Call UpdateDB("ALTER TABLE [blog_Article] ADD COLUMN [log_Yea] int DEFAULT 0","[blog_Article].[log_Yea]")
				Call UpdateDB("ALTER TABLE [blog_Article] ADD COLUMN [log_Nay] int DEFAULT 0","[blog_Article].[log_Nay]")
				Call UpdateDB("ALTER TABLE [blog_Article] ADD COLUMN [log_Ratting] int DEFAULT 0","[blog_Article].[log_Ratting]")
				Call UpdateDB("ALTER TABLE [blog_Article] ADD COLUMN [log_Template] VARCHAR(50) default """"","[blog_Article].[log_Template]")
				Call UpdateDB("ALTER TABLE [blog_Article] ADD COLUMN [log_FullUrl] VARCHAR(255) default """"","[blog_Article].[log_FullUrl]")
				Call UpdateDB("ALTER TABLE [blog_Article] ADD COLUMN [log_Type] int DEFAULT 0","[blog_Article].[log_Type]")
				Call UpdateDB("ALTER TABLE [blog_Article] ADD COLUMN [log_Meta] text default """"","[blog_Article].[log_Meta]")
			End If
		
			If Not CheckUpdateDB("[log_Type]","[blog_Article]") Then
				Call UpdateDB("ALTER TABLE [blog_Article] ADD [log_Type] int default 0","[blog_Article].[log_Type]")
				Call objConn.Execute("UPDATE [blog_Article] SET [log_Type]=0")
				Call objConn.Execute("UPDATE [blog_Article] SET [log_Type]=1 WHERE [log_CateID]=0")
			End If
			
			If Not CheckUpdateDB("[cate_Meta]","[blog_Category]") Then
				objconn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Temp] VARCHAR(255) default """"")
				Call UpdateDB("ALTER TABLE [blog_Category] ADD COLUMN [cate_URL] VARCHAR(255) default """"","[blog_Category].[cate_URL]")
				objConn.execute("UPDATE [blog_Category] SET [cate_Temp]=[cate_Url]")
				objConn.execute("UPDATE [blog_Category] SET [cate_URL]=[cate_Intro]")
				objConn.execute("UPDATE [blog_Category] SET [cate_Intro]=[cate_Temp]")
				objConn.execute("ALTER TABLE [blog_Category] DROP COLUMN [cate_Temp]")
				Call UpdateDB("ALTER TABLE [blog_Category] ADD COLUMN [cate_ParentID] int DEFAULT 0","[blog_Category].[cate_ParentID]")
				Call UpdateDB("ALTER TABLE [blog_Category] ADD COLUMN [cate_Template] VARCHAR(50) default """"","[blog_Category].[cate_Template]")
				Call UpdateDB("ALTER TABLE [blog_Category] ADD COLUMN [cate_FullUrl] VARCHAR(255) default """"","[blog_Category].[cate_FullUrl]")
				Call UpdateDB("ALTER TABLE [blog_Category] ADD COLUMN [cate_Meta] text default """"","[blog_Category].[cate_Meta]")
			End If
		
			If Not CheckUpdateDB("[comm_Meta]","[blog_Comment]") Then
				Call UpdateDB("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Reply] text default """"","[blog_Comment].[comm_Reply")
				Call UpdateDB("ALTER TABLE [blog_Comment] ADD COLUMN [comm_LastReplyIP] VARCHAR(15) default """"","[blog_Comment].[comm_LastReplyIP]")
				Call UpdateDB("ALTER TABLE [blog_Comment] ADD COLUMN [comm_LastReplyTime] datetime default now()","[blog_Comment].[comm_LastReplyTime]")
				Call UpdateDB("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Yea] int DEFAULT 0","[blog_Comment].[comm_Yea]")
				Call UpdateDB("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Nay] int DEFAULT 0","[blog_Comment].[comm_Nay]")
				Call UpdateDB("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Ratting] int DEFAULT 0","[blog_Comment].[comm_Ratting]")
				Call UpdateDB("ALTER TABLE [blog_Comment] ADD COLUMN [comm_ParentID] int DEFAULT 0","[blog_Comment].[comm_ParentID]")
				Call UpdateDB("ALTER TABLE [blog_Comment] ADD COLUMN [comm_IsCheck] YESNO DEFAULT FALSE","[blog_Comment].[comm_IsCheck]")
				Call UpdateDB("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Meta] text default """"","[blog_Comment].[comm_Meta]")
		
				objConn.execute("UPDATE [blog_Comment] SET [comm_ParentID]=0")
			End If
			
			If Not CheckUpdateDB("[mem_Meta]","[blog_Member]") Then
				Call UpdateDB("ALTER TABLE [blog_Member] ADD COLUMN [mem_Guid] VARCHAR(36) default """"","[blog_Member].[mem_Guid]")
				Call UpdateDB("ALTER TABLE [blog_Member] ADD COLUMN [mem_Template] VARCHAR(50) default """"","[blog_Member].[mem_Template]")
				Call UpdateDB("ALTER TABLE [blog_Member] ADD COLUMN [mem_FullUrl] VARCHAR(255) default """"","[blog_Member].[mem_FullUrl]")
				Call UpdateDB("ALTER TABLE [blog_Member] ADD COLUMN [mem_Meta] text default """"","[blog_Member].[mem_Meta]")
			End If
		
			If Not CheckUpdateDB("[ul_Meta]","[blog_UpLoad]") Then
				Call UpdateDB("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_Meta] text default """"","[blog_UpLoad].[ul.Meta]")
				Call UpdateDB("ALTER TABLE [blog_UpLoad] ALTER COLUMN [ul_FileName] NVARCHAR(255) ","[blog_UpLoad].[ul_FileName]")
			End If
		
			If Not CheckUpdateDB("[tb_Meta]","[blog_TrackBack]") Then
				Call UpdateDB("ALTER TABLE [blog_TrackBack] ADD COLUMN [tb_Meta] text default """"","[blog_TrackBack].[tb_Meta]")
			End If
		
			If Not CheckUpdateDB("[tag_Meta]","[blog_Tag]") Then
				Call UpdateDB("ALTER TABLE [blog_Tag] ADD COLUMN [tag_Template] VARCHAR(50) default """"","[blog_Tag].[tag_Template]")
				Call UpdateDB("ALTER TABLE [blog_Tag] ADD COLUMN [tag_FullUrl] VARCHAR(255) default """"","[blog_Tag].[tag_FullUrl]")
				Call UpdateDB("ALTER TABLE [blog_Tag] ADD COLUMN [tag_Meta] text default """"","[blog_Tag].[tag_Meta]")
			End If
		
			If Not CheckUpdateDB("[conf_Name]","[blog_Config]") Then
				Call UpdateDB("CREATE TABLE [blog_Config] (conf_Name VARCHAR(255) default """" not null,conf_Value text default """")","[blog_Config]")
				Call UpdateDB("CREATE TABLE [blog_Function] (fn_ID AutoIncrement primary key,fn_Name VARCHAR(50) default """",fn_FileName VARCHAR(50) default """",fn_Order int default 0,fn_Content text default """",fn_IsSystem YESNO DEFAULT 0,fn_SidebarID int default 0,fn_HtmlID VARCHAR(50) default """",fn_Ftype VARCHAR(5) default """",fn_MaxLi int default 0,fn_Meta text default """")","[blog_Function]")
			End If
			
			If Not CheckUpdateDB("[cate_LogTemplate]","[blog_Category]") Then				
				Call UpdateDB("ALTER TABLE [blog_Category] ADD [cate_LogTemplate] nvarchar(50) default ''","[blog_Category].[cate_LogTemplate]")
				objConn.execute("UPDATE [blog_Category] SET [cate_LogTemplate]=''")
			End If
			ExportOK "升级数据库结构成功！"
			ExportOK "正在自动跳转到升级数据库内容"
			%><input type="button" style="width:100%" value="请点击这里升级内容" class="button" onClick="location.href='?step=4&mdb=<%=Server.URLEncode(Request.QueryString("mdb"))%>'"/>
<script>location.href='?step=4&mdb=<%=Server.URLEncode(Request.QueryString("mdb"))%>'</script>
<%
		Else
			ExportError "初始化数据库失败！"
		End If
	Else
		ExportError "找不到"&Path&"！"
	End If
	%>
<%case 4%>
<%
	Dim objArticle
	If fso.Exists(path) Then
		ExportOK "找到数据库"
		ZC_MSSQL_ENABLE=False
		ZC_DATABASE_PATH="zb_users\data\"&Request.QueryString("mdb")
		If OpenConnect Then
			ExportOK "初始化数据库成功！"
			Set objRs=objConn.Execute("SELECT MAX([log_ID]) FROM [blog_Article]")
			Call SetAllFullUrl(objRs(0))
			ExportOK "设置Template成功"
			
			
			Set objRs=Server.CreateObject("adodb.recordset")
			objRs.Open "SELECT [mem_Password],[mem_Guid] FROM [blog_Member]",obJConn,1,3
			Do Until objRs.Eof
				Guid=RndGuid()
				objRs("mem_Password")=MD5(objRs("mem_Password")&Guid)
				objRs("mem_Guid")=Guid
				objRs.MoveNext
			Loop
			ExportOK "升级密码成功！"
			Set objRs=Nothing
			Set objRs=objConn.Execute("SELECT [mem_Name],[mem_Password] FROM [blog_Member] WHERE mem_Level=1")

			Response.Cookies("password")=objRs("mem_Password")
			Response.Cookies("password").Expires = DateAdd("d", 1, now)
			Response.Cookies("password").Path = "/"
			Response.Cookies("username")=escape(objRs("mem_Name"))
			Response.Cookies("username").Expires = DateAdd("d", 1, now)
			Response.Cookies("username").Path = "/"
			Call SetBlogHint(True,True,True)
			Response.Write "<br/><br/><a href='../zb_system/cmd.asp?act=login'>点击这里进入后台</a>"

		Else
			ExportError "初始化数据库失败！"
		End If
	Else
		ExportError "找不到"&Path&"！"
	End If%>
<%end select%>
</p>
  </div>
</div>


</body>
</html>

<%
Class a
Public b
Sub Class_Initialize
	Set b=CreateObject("Scripting.FileSystemObject")
End Sub

Public Function Exists(str)
	Exists=b.FileExists(str)
End Function


End Class
Set fso=Nothing

Sub ExportOK(str)
	Response.Write "<span style='color:green'>√ 【"&Now&"】"&str&"</span><br/>"
End Sub

Sub ExportError(str)
	Response.Write "<span style='color:red'>× 【"&Now&"】"&str&"</span>"
	Response.Write "<br/><br/><a href='javascript:history.go(-1)'>点击这里返回</a>"
	Response.End

End Sub

Sub UpdateDB(sql,str)
	On Error Resume Next
	objConn.Execute sql
	If Err.Number=0 Then
		ExportOK str
		Response.Flush()
	Else
		ExportError "升级"&str&"时出现错误"&Err.Number&"<br/>"&Err.description&"<br/>"&sql
	End If
End Sub

Function CheckUpdateDB(a,b)
	Err.Clear
	On Error Resume Next
	objConn.execute "SELECT "&a&" FROM "&b	
	
	If Err.Number=0 Then
		CheckUpdateDB=True
	Else
	Err.Clear
		CheckUpdateDB=False
	End If	
End Function

Function GetRndPassword()
	Dim i,j,k
	Randomize
	j="!@#$%^&*()_+|qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM,./:;""{}[]7894563210+-*`"
	For i=0 to 12
		k=k&Mid(j,Int(89 * Rnd + 1),1)
		
	Next
	GetRndPassword=k
End Function

Sub SetAllFullUrl(j)
	Dim i,a
	For i=0 to j	
		Set a=New TArticle
		If (a.LoadInfoById(i)) Then a.TemplateName="":a.Post
		Set a=Nothing
	Next
End Sub
%>
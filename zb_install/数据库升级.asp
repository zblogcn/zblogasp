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
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<%Response.Buffer=True %>
<!-- #include file="c_custom.asp" -->
<%

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


Function UpdateDateBase()
	
	If Not CheckUpdateDB("[log_IsTop]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_IsTop] YESNO DEFAULT FALSE")
		objConn.execute("UPDATE [blog_Article] SET [log_IsTop]=0")
	End If

	If Not CheckUpdateDB("[log_Tag]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Tag] VARCHAR(255)")
	End If

	If Not CheckUpdateDB("[tag_ID]","[blog_Tag]") Then
		objConn.execute("CREATE TABLE [blog_Tag] (tag_ID AutoIncrement primary key,tag_Name VARCHAR(255),tag_Intro text,tag_ParentID int,tag_URL VARCHAR(255),tag_Order int,tag_Count int)")
	End If

	If Not CheckUpdateDB("[coun_ID]","[blog_Counter]") Then
		objConn.execute("CREATE TABLE [blog_Counter] (coun_ID AutoIncrement primary key,coun_IP VARCHAR(20),coun_Agent text,coun_Refer VARCHAR(255),coun_PostTime TIME DEFAULT Now())")
	End If

	If Not CheckUpdateDB("[key_ID]","[blog_Keyword]") Then
		objConn.execute("CREATE TABLE [blog_Keyword] (key_ID AutoIncrement primary key,key_Name VARCHAR(255),key_Intro text,key_URL VARCHAR(255))")
	End If

	If Not CheckUpdateDB("[ul_Quote]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_Quote] VARCHAR(255)")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_Quote]=''")
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_DownNum] int DEFAULT 0")
	End If

	If Not CheckUpdateDB("[ul_FileIntro]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_FileIntro] VARCHAR(255)")
	End If

	If Not CheckUpdateDB("[ul_DirByTime]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_DirByTime] YESNO DEFAULT FALSE")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_DirByTime]=[ul_Quote]")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_Quote]=''")
	End If

	If Not CheckUpdateDB("[tag_Template]","[blog_Tag]") Then
		objConn.execute("ALTER TABLE [blog_Tag] ADD COLUMN [tag_Template] VARCHAR(255)")
	End If

	If Not CheckUpdateDB("[cate_Template]","[blog_Category]") Then
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Template] VARCHAR(255)")
	End If

	If Not CheckUpdateDB("[log_Template]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Template] VARCHAR(255)")
	End If

	If Not CheckUpdateDB("[log_Meta]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Yea] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Nay] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Delate] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Meta] text")
	End If

	If Not CheckUpdateDB("[tag_Meta]","[blog_Tag]") Then
		objConn.execute("ALTER TABLE [blog_Tag] ADD COLUMN [tag_Meta] text")
	End If

	If Not CheckUpdateDB("[ul_Meta]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_Meta] text")
	End If

	If Not CheckUpdateDB("[cate_Meta]","[blog_Category]") Then
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Temp] VARCHAR(255)")
		objConn.execute("UPDATE [blog_Category] SET [cate_Temp]=[cate_Url]")
		objConn.execute("UPDATE [blog_Category] SET [cate_URL]=[cate_Intro]")
		objConn.execute("UPDATE [blog_Category] SET [cate_Intro]=[cate_Temp]")
		objConn.execute("ALTER TABLE [blog_Category] DROP COLUMN [cate_Temp]")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_ParentID] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Meta] text")
	End If

	If Not CheckUpdateDB("[comm_Meta]","[blog_Comment]") Then
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Reply] text")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_LastReplyIP] VARCHAR(15)")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_LastReplyTime] datetime default now()")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Yea] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Nay] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Delate] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Meta] text")
	End If
	
	If Not CheckUpdateDB("[mem_Meta]","[blog_Member]") Then
		objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_Meta] text")
	End If

	If Not CheckUpdateDB("[tb_Meta]","[blog_TrackBack]") Then
		objConn.execute("ALTER TABLE [blog_TrackBack] ADD COLUMN [tb_Meta] text")
	End If

	If CheckUpdateDB("[comm_Reply]","[blog_Comment]") Then

	End If

End Function




Dim objConn	
Dim objCat


'这段是升级数据库的



Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(ZC_DATABASE_PATH)

Call UpdateDateBase()





%>
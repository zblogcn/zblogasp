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
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_FullUrl] VARCHAR(255)")
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


End Function




Dim objConn	
Dim objCat


'这段是升级数据库的



'Set objConn = Server.CreateObject("ADODB.Connection")
'objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("data/#%20d7d52c946f1f84b79884.mdb")

'Call UpdateDateBase()













'这段是创建一个全新的空的ACCESS数据库及默认数据


Set objCat=Server.CreateObject("ADOX.Catalog")   
objCat.Create  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("zblog.mdb")


Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("zblog.mdb")

objConn.execute("CREATE TABLE [blog_Tag] (tag_ID AutoIncrement primary key,tag_Name VARCHAR(255),tag_Intro text,tag_ParentID int default 0,tag_URL VARCHAR(255),tag_Order int default 0,tag_Count int default 0,tag_Template VARCHAR(255),tag_Meta text)")

objConn.execute("CREATE TABLE [blog_Article] (log_ID AutoIncrement primary key,log_CateID int default 0,log_AuthorID int default 0,log_Level int default 0,log_Url VARCHAR(255),log_Title VARCHAR(255),log_Intro text,log_Content text,log_IP VARCHAR(15),log_PostTime datetime default now(),log_CommNums int default 0,log_ViewNums int default 0,log_TrackBackNums int default 0,log_Tag VARCHAR(255),log_IsTop YESNO DEFAULT 0,log_Yea int default 0,log_Nay int default 0,log_Delate int default 0,log_Template VARCHAR(255),log_FullUrl VARCHAR(255),log_Meta text)")

objConn.execute("CREATE TABLE [blog_Category] (cate_ID AutoIncrement primary key,cate_Name VARCHAR(50),cate_Order int default 0,cate_Intro VARCHAR(255),cate_Count int default 0,cate_URL VARCHAR(255),cate_ParentID int default 0,cate_Template VARCHAR(255),cate_Meta text)")

objConn.execute("CREATE TABLE [blog_Comment] (comm_ID AutoIncrement primary key,log_ID int default 0,comm_AuthorID int default 0,comm_Author VARCHAR(20),comm_Content text,comm_Email VARCHAR(50),comm_HomePage VARCHAR(255),comm_PostTime datetime default now(),comm_IP VARCHAR(15),comm_Agent text,comm_Reply text,comm_LastReplyIP VARCHAR(15),comm_LastReplyTime datetime default now(),comm_Yea int default 0,comm_Nay int default 0,comm_Delate int default 0,comm_Meta text)")

objConn.execute("CREATE TABLE [blog_TrackBack] (tb_ID AutoIncrement primary key,log_ID int default 0,tb_URL VARCHAR(255),tb_Title VARCHAR(100),tb_Blog VARCHAR(50),tb_Excerpt text,tb_PostTime datetime,tb_IP VARCHAR(15),tb_Agent text,tb_Meta text)")

objConn.execute("CREATE TABLE [blog_UpLoad] (ul_ID AutoIncrement primary key,ul_AuthorID int default 0,ul_FileSize int default 0,ul_FileName VARCHAR(50),ul_PostTime datetime default now(),ul_Quote VARCHAR(255),ul_DownNum int default 0,ul_FileIntro VARCHAR(255),ul_DirByTime YESNO DEFAULT 0,ul_Meta text)")

objConn.execute("CREATE TABLE [blog_Counter] (coun_ID AutoIncrement primary key,coun_IP VARCHAR(15),coun_Agent text,coun_Refer VARCHAR(255),coun_PostTime datetime default now() )")

objConn.execute("CREATE TABLE [blog_Keyword] (key_ID AutoIncrement primary key,key_Name VARCHAR(255),key_Intro text,key_URL VARCHAR(255) )")

objConn.execute("CREATE TABLE [blog_Member] (mem_ID AutoIncrement primary key,mem_Level int default 0,mem_Name VARCHAR(20),mem_Password VARCHAR(32),mem_Sex int default 0,mem_Email VARCHAR(50),mem_MSN VARCHAR(50),mem_QQ VARCHAR(50),mem_HomePage VARCHAR(255),mem_LastVisit datetime,mem_Status int default 0,mem_PostLogs int default 0,mem_PostComms int default 0,mem_Intro text,mem_IP VARCHAR(15),mem_Count int default 0,mem_Meta text)")

objConn.Execute("INSERT INTO [blog_Member]([mem_Level],[mem_Name],[mem_PassWord],[mem_Email],[mem_HomePage],[mem_Intro]) VALUES (1,'zblogger','aa055c6d7875a18fa49058c2c48f2140','null@null.com','','')")





'这段是在指定的MSSQL数据库里创建新表及默认数据





Set objConn = Server.CreateObject("ADODB.Connection")

objConn.Open "Provider=SqlOLEDB;Data Source=localhost;Initial Catalog=zb;Persist Security Info=True;User ID=sa;Password=123456;"

objConn.execute("CREATE TABLE [blog_Tag] (tag_ID int identity(1,1) not null primary key,tag_Name nvarchar(255),tag_Intro ntext,tag_ParentID int default 0,tag_URL nvarchar(255),tag_Order int default 0,tag_Count int default 0,tag_Template nvarchar(255),tag_Meta ntext)")

objConn.execute("CREATE TABLE [blog_Article] (log_ID int identity(1,1) not null primary key,log_CateID int default 0,log_AuthorID int default 0,log_Level int default 0,log_Url nvarchar(255),log_Title nvarchar(255),log_Intro ntext,log_Content ntext,log_IP nvarchar(15),log_PostTime datetime default getdate(),log_CommNums int default 0,log_ViewNums int default 0,log_TrackBackNums int default 0,log_Tag nvarchar(255),log_IsTop bit DEFAULT 0,log_Yea int default 0,log_Nay int default 0,log_Delate int default 0,log_Template nvarchar(255),log_Meta ntext)")

objConn.execute("CREATE TABLE [blog_Category] (cate_ID int identity(1,1) not null primary key,cate_Name nvarchar(50),cate_Order int default 0,cate_Intro nvarchar(255),cate_Count int default 0,cate_URL nvarchar(255),cate_ParentID int default 0,cate_Template nvarchar(255),cate_Meta ntext)")

objConn.execute("CREATE TABLE [blog_Comment] (comm_ID int identity(1,1) not null primary key,log_ID int default 0,comm_AuthorID int default 0,comm_Author nvarchar(20),comm_Content ntext,comm_Email nvarchar(50),comm_HomePage nvarchar(255),comm_PostTime datetime default getdate(),comm_IP nvarchar(15),comm_Agent ntext,comm_Reply ntext,comm_LastReplyIP nvarchar(15),comm_LastReplyTime datetime default getdate(),comm_Yea int default 0,comm_Nay int default 0,comm_Delate int default 0,comm_Meta ntext)")

objConn.execute("CREATE TABLE [blog_TrackBack] (tb_ID int identity(1,1) not null primary key,log_ID int default 0,tb_URL nvarchar(255),tb_Title nvarchar(100),tb_Blog nvarchar(50),tb_Excerpt ntext,tb_PostTime datetime,tb_IP nvarchar(15),tb_Agent ntext,tb_Meta ntext)")

objConn.execute("CREATE TABLE [blog_UpLoad] (ul_ID int identity(1,1) not null primary key,ul_AuthorID int default 0,ul_FileSize int default 0,ul_FileName nvarchar(50),ul_PostTime datetime default getdate(),ul_Quote nvarchar(255),ul_DownNum int default 0,ul_FileIntro nvarchar(255),ul_DirByTime bit DEFAULT 0,ul_Meta ntext)")

objConn.execute("CREATE TABLE [blog_Counter] (coun_ID int identity(1,1) not null primary key,coun_IP nvarchar(15),coun_Agent ntext,coun_Refer nvarchar(255),coun_PostTime datetime default getdate() )")

objConn.execute("CREATE TABLE [blog_Keyword] (key_ID int identity(1,1) not null primary key,key_Name nvarchar(255),key_Intro ntext,key_URL nvarchar(255) )")

objConn.execute("CREATE TABLE [blog_Member] (mem_ID int identity(1,1) not null primary key,mem_Level int default 0,mem_Name nvarchar(20),mem_Password nvarchar(32),mem_Sex int default 0,mem_Email nvarchar(50),mem_MSN nvarchar(50),mem_QQ nvarchar(50),mem_HomePage nvarchar(255),mem_LastVisit datetime,mem_Status int default 0,mem_PostLogs int default 0,mem_PostComms int default 0,mem_Intro ntext,mem_IP nvarchar(15),mem_Count int default 0,mem_Meta ntext)")

objConn.Execute("INSERT INTO [blog_Member]([mem_Level],[mem_Name],[mem_PassWord],[mem_Email],[mem_HomePage],[mem_Intro]) VALUES (1,'zblogger','aa055c6d7875a18fa49058c2c48f2140','null@null.com','','')")

%>
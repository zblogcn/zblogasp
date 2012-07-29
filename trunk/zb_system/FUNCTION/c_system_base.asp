<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:
'// 程序版本:
'// 单元名称:    c_system_base.asp
'// 开始时间:    2005.02.11
'// 最后修改:
'// 备    注:
'///////////////////////////////////////////////////////////////////////////////

'定义全局变量
Dim objConn

Dim BlogTitle
Dim BlogUser
Set BlogUser =New TUser

Dim BlogPath
BlogPath=GetReallyDirectory()

Dim StarTime
Dim EndTime
StarTime = Timer()

Dim Categorys()
Dim Users()
Dim Tags()
Dim Functions()

ReDim Categorys(0)
ReDim Users(0)
ReDim Tags(0)
ReDim Functions(0)

Dim FunctionMetas
Set FunctionMetas=New TMeta

Dim ConfigMetas
Set ConfigMetas=New TMeta

Dim BlogConfig
Set BlogConfig = New TConfig

Dim PluginName()
Dim PluginActiveFunction()
ReDim PluginName(0)
ReDim PluginActiveFunction(0)


Dim TemplateDic
Dim TemplateTagsDic

Set TemplateDic=CreateObject("Scripting.Dictionary")
Set TemplateTagsDic=CreateObject("Scripting.Dictionary")


Const ZC_BLOG_WEBEDIT="ueditor"
Const ZC_TB_EXCERPT_MAX=250
Const ZC_TRACKBACK_TURNOFF=True

Const ZC_AUTO_NEWLINE=False
Const ZC_JAPAN_TO_HTML=False

Const ZC_DISPLAY_MODE_ALL=1
Const ZC_DISPLAY_MODE_INTRO=2
Const ZC_DISPLAY_MODE_HIDE=3
Const ZC_DISPLAY_MODE_LIST=4
Const ZC_DISPLAY_MODE_ONTOP=5
Const ZC_DISPLAY_MODE_SEARCH=6

'如果连接数据库为MSSQL，则应为'，默认连接Access数据库则为#
Dim ZC_SQL_POUND_KEY
ZC_SQL_POUND_KEY="#"


'*********************************************************
' 目的：    System 初始化
'*********************************************************
Sub System_Initialize()

	'On Error Resume Next

	'plugin node
	bAction_Plugin_System_Initialize=False
	For Each sAction_Plugin_System_Initialize in Action_Plugin_System_Initialize
		If Not IsEmpty(sAction_Plugin_System_Initialize) Then Call Execute(sAction_Plugin_System_Initialize)
		If bAction_Plugin_System_Initialize=True Then Exit Sub
	Next

	If OpenConnect()=False Then
		If Err.Number<>0 Then Call ShowError(4)
	End If

	BlogUser.Verify()

	'Call GetCategory()
	'Call GetUser()
	'Call GetTags()
	'Call GetKeyWords()
	Call GetConfigs()
	'Call GetFunction()


	BlogConfig.Load("Blog")

	Call LoadGlobeCache()

	Dim bolRebuildIndex
	Application.Lock
	bolRebuildIndex=Application(ZC_BLOG_CLSID & "SIGNAL_REBUILDINDEX")
	Application.UnLock
	If IsEmpty(bolRebuildIndex)=False Then
		If bolRebuildIndex=True Then
			Call MakeBlogReBuild_Core()
		End If
	End If

	Call CreateAdminLeftMenu()
	Call CreateAdminTopMenu()

	'将激活插件后移
	Call ActivePlugin()

	'plugin node
	bAction_Plugin_System_Initialize_Succeed=False
	For Each sAction_Plugin_System_Initialize_Succeed in Action_Plugin_System_Initialize_Succeed
		If Not IsEmpty(sAction_Plugin_System_Initialize_Succeed) Then Call Execute(sAction_Plugin_System_Initialize_Succeed)
		If bAction_Plugin_System_Initialize_Succeed=True Then Exit Sub
	Next

	If Err.Number<>0 Then Call ShowError(10)

End Sub
'*********************************************************




'*********************************************************
' 目的：    System 释放
'*********************************************************
Sub System_Terminate()

	'plugin node
	bAction_Plugin_System_Terminate=False
	For Each sAction_Plugin_System_Terminate in Action_Plugin_System_Terminate
		If Not IsEmpty(sAction_Plugin_System_Terminate) Then Call Execute(sAction_Plugin_System_Terminate)
		If bAction_Plugin_System_Terminate=True Then Exit Sub
	Next

	Call CloseConnect()

End Sub
'*********************************************************




'*********************************************************
' 目的：    System 初始化 WithOutDB
'*********************************************************
Sub System_Initialize_WithOutDB()

	'plugin node
	bAction_Plugin_System_Initialize_WithOutDB=False
	For Each sAction_Plugin_System_Initialize_WithOutDB in Action_Plugin_System_Initialize_WithOutDB
		If Not IsEmpty(sAction_Plugin_System_Initialize_WithOutDB) Then Call Execute(sAction_Plugin_System_Initialize_WithOutDB)
		If bAction_Plugin_System_Initialize_WithOutDB=True Then Exit Sub
	Next

	Call LoadGlobeCache()

	'将激活插件后移
	Call ActivePlugin()

	'plugin node
	bAction_Plugin_System_Initialize_WithOutDB_Succeed=False
	For Each sAction_Plugin_System_Initialize_WithOutDB_Succeed in Action_Plugin_System_Initialize_WithOutDB_Succeed
		If Not IsEmpty(sAction_Plugin_System_Initialize_WithOutDB_Succeed) Then Call Execute(sAction_Plugin_System_Initialize_WithOutDB_Succeed)
		If bAction_Plugin_System_Initialize_WithOutDB_Succeed=True Then Exit Sub
	Next

End Sub
'*********************************************************




'*********************************************************
' 目的：    System 释放 WithOutDB
'*********************************************************
Sub System_Terminate_WithOutDB()

	'plugin node
	bAction_Plugin_System_Terminate_WithOutDB=False
	For Each sAction_Plugin_System_Terminate_WithOutDB in Action_Plugin_System_Terminate_WithOutDB
		If Not IsEmpty(sAction_Plugin_System_Terminate_WithOutDB) Then Call Execute(sAction_Plugin_System_Terminate_WithOutDB)
		If bAction_Plugin_System_Terminate_WithOutDB=True Then Exit Sub
	Next

End Sub
'*********************************************************




'*********************************************************
' 目的：    数据库连接
'*********************************************************
Dim IsDBConnect '数据库是否已连接
IsDBConnect=False
Function OpenConnect()

	If IsDBConnect=True Then
		Exit Function
	End If

	'plugin node
	bAction_Plugin_OpenConnect=False
	For Each sAction_Plugin_OpenConnect in Action_Plugin_OpenConnect
		If Not IsEmpty(sAction_Plugin_OpenConnect) Then Call Execute(sAction_Plugin_OpenConnect)
		If bAction_Plugin_OpenConnect=True Then Exit Function
	Next

	'判定是否为子目录调用
	Dim strDbPath

	strDbPath=BlogPath & ZC_DATABASE_PATH

	Set objConn = Server.CreateObject("ADODB.Connection")
	If ZC_MSSQL_ENABLE=True Then
		objConn.Open "Provider=SqlOLEDB;Data Source="&ZC_MSSQL_SERVER&";Initial Catalog="&ZC_MSSQL_DATABASE&";Persist Security Info=True;User ID="&ZC_MSSQL_USERNAME&";Password="&ZC_MSSQL_PASSWORD&";"
		ZC_SQL_POUND_KEY="'"
	Else
		objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDbPath
	End If

	IsDBConnect=True

	OpenConnect=True

End Function
'*********************************************************




'*********************************************************
' 目的：    DB Disable Connect
'*********************************************************
Function CloseConnect()

	If IsDBConnect=False Then
		Exit Function
	End If

	objConn.Close
	Set objConn=Nothing

	IsDBConnect=False

	CloseConnect=True

End Function
'*********************************************************




'*********************************************************
' 目的：    时间计长
'*********************************************************
Function RunTime()

	EndTime=Timer()
	RunTime = CLng(FormatNumber((EndTime-StarTime)*1000,3))

End Function
'*********************************************************




'*********************************************************
' 目的：    分类读取
'*********************************************************
'IsRunGetCategory默认为false,如果运行过一次GetCategory则为True,之后再GetCategory则不执行
Dim IsRunGetCategory
IsRunGetCategory=False
Function GetCategory()

	If IsRunGetCategory=True Then Exit Function

	Dim i,j,k,l

	Dim aryAllData
	Dim arySingleData()

	Erase Categorys
	ReDim Categorys(0)

	Dim objRS

	Set objRS=objConn.Execute("SELECT TOP 1 [cate_ID] FROM [blog_Category] ORDER BY [cate_ID] DESC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		i=objRS("cate_ID")
		ReDim Categorys(i)
	End If
	objRS.Close
	Set objRS=Nothing

	Set objRS=objConn.Execute("SELECT [cate_ID],[cate_Name],[cate_Intro],[cate_Order],[cate_Count],[cate_ParentID],[cate_URL],[cate_Template],[cate_FullUrl],[cate_Meta] FROM [blog_Category] ORDER BY [cate_ID] ASC")
	If (Not objRS.bof) And (Not objRS.eof) Then

		aryAllData=objRS.GetRows(objRS.RecordCount)
		objRS.Close
		Set objRS=Nothing

		k=UBound(aryAllData,1)
		l=UBound(aryAllData,2)
		For i=0 To l
			Set Categorys(aryAllData(0,i))=New TCategory
			Categorys(aryAllData(0,i)).LoadInfoByArray(Array(aryAllData(0,i),aryAllData(1,i),aryAllData(2,i),aryAllData(3,i),aryAllData(4,i),aryAllData(5,i),aryAllData(6,i),aryAllData(7,i),aryAllData(8,i),aryAllData(9,i)))
		Next
	End If

	Set Categorys(0)=New TCategory

	IsRunGetCategory=True

	GetCategory=True

End Function
'*********************************************************




'*********************************************************
' 目的：    用户读取
'*********************************************************
Dim IsRunGetUser
IsRunGetUser=False
Function GetUser()

	If IsRunGetUser=True Then Exit Function

	Dim i,j,k,l

	Dim aryAllData
	Dim arySingleData()

	Erase Users
	ReDim Users(0)

	Dim objRS

	Set objRS=objConn.Execute("SELECT TOP 1 [mem_ID] FROM [blog_Member] ORDER BY [mem_ID] DESC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		i=objRS("mem_ID")
		ReDim Users(i)
	End If
	objRS.Close
	Set objRS=Nothing


	Set objRS=objConn.Execute("SELECT [mem_ID],[mem_Name],[mem_Level],[mem_Password],[mem_Email],[mem_HomePage],[mem_PostLogs],[mem_Intro],[mem_Meta] FROM [blog_Member] ORDER BY [mem_ID] ASC")
	If (Not objRS.bof) And (Not objRS.eof) Then

		aryAllData=objRS.GetRows(objRS.RecordCount)
		objRS.Close
		Set objRS=Nothing

		k=UBound(aryAllData,1)
		l=UBound(aryAllData,2)
		For i=0 To l
			Set Users(aryAllData(0,i))=New TUser
			Users(aryAllData(0,i)).LoadInfoByArray(Array(aryAllData(0,i),aryAllData(1,i),aryAllData(2,i),aryAllData(3,i),aryAllData(4,i),aryAllData(5,i),aryAllData(6,i),aryAllData(7,i),aryAllData(8,i)))
		Next

	End If

	Set Users(0)=New TUser

	IsRunGetUser=True

	Getuser=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Tags读取
'*********************************************************
Dim IsRunGetTags
IsRunGetTags=False
Function GetTags()

	If IsRunGetTags=True Then Exit Function

	Dim i,j,k,l

	Dim aryAllData
	Dim arySingleData()

	Erase Tags
	ReDim Tags(0)

	Dim objRS

	Set objRS=objConn.Execute("SELECT TOP 1 [tag_ID] FROM [blog_Tag] ORDER BY [tag_ID] DESC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		i=objRS("tag_ID")
		ReDim Tags(i)
	End If

	Set objRS=objConn.Execute("SELECT [tag_ID],[tag_Name],[tag_Intro],[tag_Order],[tag_Count],[tag_ParentID],[tag_URL],[tag_Template],[tag_FullUrl],[tag_Meta] FROM [blog_Tag] ORDER BY [tag_ID] ASC")
	If (Not objRS.bof) And (Not objRS.eof) Then

		aryAllData=objRS.GetRows(objRS.RecordCount)
		objRS.Close
		Set objRS=Nothing

		k=UBound(aryAllData,1)
		l=UBound(aryAllData,2)
		For i=0 To l
			Set Tags(aryAllData(0,i))=New TTag
			Tags(aryAllData(0,i)).LoadInfoByArray(Array(aryAllData(0,i),aryAllData(1,i),aryAllData(2,i),aryAllData(3,i),aryAllData(4,i),aryAllData(5,i),aryAllData(6,i),aryAllData(7,i),aryAllData(8,i),aryAllData(9,i)))
		Next

	End If

	Set Tags(0)=New TTag

	IsRunGetTags=True

	GetTags=True

End Function
'*********************************************************



'*********************************************************
' 目的：    Tags读取
'*********************************************************
Dim IsRunConfigs
IsRunConfigs=False
Function GetConfigs()

	If IsRunConfigs=True Then Exit Function

	Dim objRS
	Set objRS=objConn.Execute("SELECT [conf_Name],[conf_Value] FROM [blog_Config]")
	If (Not objRS.bof) And (Not objRS.eof) Then

		Do While Not objRS.eof
			Call ConfigMetas.SetValue(objRS(0),objRS(1))
			objRS.MoveNext
		Loop

	End If

	IsRunConfigs=True

	GetConfigs=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Functions读取
'*********************************************************
Dim IsRunFunctions
IsRunFunctions=False
Function GetFunction()

	If IsRunFunctions=True Then Exit Function

	Dim i,j,k,l

	Dim aryAllData
	Dim arySingleData()

	Erase Tags
	ReDim Tags(0)

	Dim objRS

	Set objRS=objConn.Execute("SELECT TOP 1 [fn_ID] FROM [blog_Function] ORDER BY [fn_ID] DESC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		i=objRS("fn_ID")
		ReDim Functions(i)
	End If

	Set objRS=objConn.Execute("SELECT [fn_ID],[fn_Name],[fn_FileName],[fn_Order],[fn_Content],[fn_IsSystem],[fn_SidebarID],[fn_HtmlID],[fn_Ftype],[fn_MaxLi],[fn_Meta] FROM [blog_Function] ORDER BY [fn_ID] ASC")
	If (Not objRS.bof) And (Not objRS.eof) Then

		aryAllData=objRS.GetRows(objRS.RecordCount)
		objRS.Close
		Set objRS=Nothing

		k=UBound(aryAllData,1)
		l=UBound(aryAllData,2)
		For i=0 To l
			Set Functions(aryAllData(0,i))=New TFunction
			Functions(aryAllData(0,i)).LoadInfoByArray(Array(aryAllData(0,i),aryAllData(1,i),aryAllData(2,i),aryAllData(3,i),aryAllData(4,i),aryAllData(5,i),aryAllData(6,i),aryAllData(7,i),aryAllData(8,i),aryAllData(9,i),aryAllData(10,i)))
			Call FunctionMetas.SetValue(aryAllData(2,i),aryAllData(0,i))
		Next

	End If

	Set Functions(0)=New TFunction

	IsRunFunctions=True

	GetFunction=True

End Function
'*********************************************************




'*********************************************************
' 目的：    读取权限
' 备注:     权限最高为1 最低为5 不是则非法
'           "Root"一定只能为1
'           权限配置方式可以变通
'*********************************************************
Function GetRights(strAction)

	'plugin node
	bAction_Plugin_GetRights_Begin=False
	For Each sAction_Plugin_GetRights_Begin in Action_Plugin_GetRights_Begin
		If Not IsEmpty(sAction_Plugin_GetRights_Begin) Then Call Execute(sAction_Plugin_GetRights_Begin)
		If bAction_Plugin_GetRights_Begin=True Then Exit Function
	Next

	Select Case strAction

		Case "Root"
			GetRights=1
		Case "login"
			GetRights=5
		Case "verify"
			GetRights=5
		Case "logout"
			GetRights=5
		Case "admin"
			GetRights=4
		Case "cmt"
			GetRights=5
		'Case "tb"
		'	GetRights=5
		Case "vrs"
			GetRights=5
		Case "rss"
			GetRights=5
		Case "batch"
			GetRights=4
		'Case "gettburl"
		'	GetRights=5
		Case "ArticleMng"
			GetRights=3
		Case "ArticleEdt"
			GetRights=3
		Case "ArticlePst"
			GetRights=3
		Case "ArticleDel"
			GetRights=3
		Case "ArticleBud"
			GetRights=3
		Case "CategoryMng"
			GetRights=2
		Case "CategoryEdt"
			GetRights=2
		Case "CategoryPst"
			GetRights=2
		Case "CategoryDel"
			GetRights=2
		Case "TagMng"
			GetRights=1
		Case "TagEdt"
			GetRights=1
		Case "TagPst"
			GetRights=1
		Case "TagDel"
			GetRights=1
		'Case "KeyWordMng"
		'	GetRights=1
		'Case "KeyWordEdt"
		'	GetRights=1
		'Case "KeyWordPst"
		'	GetRights=1
		'Case "KeyWordDel"
		'	GetRights=1
		'Case "GuestBookMng"
		'	GetRights=2
		Case "CommentMng"
			GetRights=4
		Case "CommentDel"
			GetRights=4
		Case "CommentEdt"
			GetRights=4
		Case "CommentSav"
			GetRights=4
		Case "CommentDelBatch"
			GetRights=4
		'Case "TrackBackMng"
		'	GetRights=3
		'Case "TrackBackDel"
		'	GetRights=3
		'Case "TrackBackDelBatch"
		'	GetRights=3
		'Case "TrackBackSnd"
		'	GetRights=0
		Case "UserMng"
			GetRights=4
		Case "UserEdt"
			GetRights=4
		Case "UserDel"
			GetRights=1
		Case "UserCrt"
			GetRights=1
		Case "BlogReBuild"
			GetRights=3
		Case "DirectoryReBuild"
			GetRights=3
		Case "FileReBuild"
			GetRights=1
		Case "AskFileReBuild"
			GetRights=1
		Case "FileMng"
			GetRights=2
		Case "FileSnd"
			GetRights=2
		Case "FileUpload"
			GetRights=2
		Case "FileDel"
			GetRights=2
		Case "FileDelBatch"
			GetRights=2
		Case "Search"
			GetRights=5
		'Case "BlogMng"
		'	GetRights=4
		Case "SettingMng"
			GetRights=1
		Case "SettingSav"
			GetRights=1
		Case "PlugInMng"
			GetRights=2
		Case "SiteInfo"
			GetRights=4
		'Case "Update"
		'	GetRights=1
		Case "ThemeMng"
			GetRights=1
		Case "ThemeSav"
			GetRights=1
		'Case "LinkMng"
		'	GetRights=1
		'Case "LinkSav"
		'	GetRights=1
		Case "PlugInActive"
			GetRights=1
		Case "PlugInDisable"
			GetRights=1
		Case "FunctionMng"
			GetRights=1
		Case "FunctionEdt"
			GetRights=1
		Case "FunctionSav"
			GetRights=1
		Case "FunctionDel"
			GetRights=1
		Case Else
			GetRights=Null
	End Select

End Function
'*********************************************************




'*********************************************************
' 目的：    检查权限
'*********************************************************
Function CheckRights(strAction)

	'plugin node
	bAction_Plugin_CheckRights_Begin=False
	For Each sAction_Plugin_CheckRights_Begin in Action_Plugin_CheckRights_Begin
		If Not IsEmpty(sAction_Plugin_CheckRights_Begin) Then Call Execute(sAction_Plugin_CheckRights_Begin)
		If bAction_Plugin_CheckRights_Begin=True Then Exit Function
	Next

	If BlogUser.Level>GetRights(strAction) Then
		CheckRights=False
	Else
		CheckRights=True
	End If

End Function
'*********************************************************




'*********************************************************
' 目的：    检查作者是否存在
'*********************************************************
Function CheckAuthorByID(intAuthorId)

	CheckAuthorByID=False
	If intAuthorId<=UBound(Users) Then
		If IsObject(Users(intAuthorId)) Then CheckAuthorByID=True
	End If

End Function
'*********************************************************




'*********************************************************
' 目的：    检查分类是否存在
'*********************************************************
Function CheckCateByID(intCateId)

	CheckCateByID=False
	If intCateId<=UBound(Categorys) Then
		If IsObject(Categorys(intCateId)) Then CheckCateByID=True
	End If

End Function
'*********************************************************




'*********************************************************
' 目的：    Get Category Order 输出数组.
'*********************************************************
Function GetCategoryOrder()

	Dim i
	Dim objRS

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	Dim aryCateInOrder()
	i=0

	objRS.Open("SELECT * FROM [blog_Category] ORDER BY [cate_Order] ASC,[cate_ID] ASC")
	Do While Not objRS.eof
		i=i+1
		ReDim Preserve aryCateInOrder(i)
		aryCateInOrder(i)=objRS("cate_ID")
		objRS.MoveNext
	Loop
	objRS.Close
	Set objRS=Nothing

	If i>0 Then GetCategoryOrder=aryCateInOrder
	If i=0 Then GetCategoryOrder=Array()

	Erase aryCateInOrder

End Function
'*********************************************************




'*********************************************************
' 目的：    取得所有子分类的ID 输出数组.
'*********************************************************
Function GetSubCateID(intCateId,bolIncludePare)

	Dim i
	Dim Category
	Dim arySubCateID()
	i=0

	If bolIncludePare=True Then
		ReDim Preserve arySubCateID(i)
		arySubCateID(i)=intCateId
		i=i+1
	End If

	For Each Category In Categorys
		If IsObject(Category) Then
			If Category.ParentID=intCateId Then
				ReDim Preserve arySubCateID(i)
				arySubCateID(i)=Category.ID
				i=i+1
			End If
		End If
	Next

	If i>0 Then GetSubCateID=arySubCateID

	Erase arySubCateID

End Function
'*********************************************************




'*********************************************************
' 目的：    Make Calendar
'*********************************************************
Function MakeCalendar(dtmYearMonth)

	'plugin node
	bAction_Plugin_MakeCalendar_Begin=False
	For Each sAction_Plugin_MakeCalendar_Begin in Action_Plugin_MakeCalendar_Begin
		If Not IsEmpty(sAction_Plugin_MakeCalendar_Begin) Then Call Execute(sAction_Plugin_MakeCalendar_Begin)
		If bAction_Plugin_MakeCalendar_Begin=True Then Exit Function
	Next

	Dim strCalendar

	Dim y
	Dim m
	Dim d
	Dim firw
	Dim lasw
	Dim ny
	Dim nm

	Dim i
	Dim j
	Dim k
	Dim b
	Dim s
	Dim t

	Call CheckParameter(dtmYearMonth,"dtm",Date())

	y=year(dtmYearMonth)
	m=month(dtmYearMonth)
	ny=y
	nm=m+1
	If m=12 Then ny=ny+1:nm=1

	firw=Weekday(Cdate(y&"-"&m&"-1"))

	For i=28 to 32
		If IsDate(y&"-"&m&"-"&i) Then
			lasw=Weekday(Cdate(y&"-"&m&"-"&i))
		Else
			Exit For
		End If
	Next

	d=i-1
	k=1

	If firw>5 Then b=42 Else b=35
	If (d=28) And (firw=1) Then b=28
	If (firw>5) And (d<31) And (d-firw<>23) Then b=35


'//////////////////////////////////////////////////////////
'	逻辑处理
	Dim aryDateLink(32)
	Dim aryDateID(32)
	Dim aryDateArticle(32)
	Dim objRS

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""
	objRS.Open("select [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_IsAnonymous],[log_Meta] from [blog_Article] where ([log_CateID]>0) And ([log_Level]>2) And ([log_PostTime] BETWEEN "& ZC_SQL_POUND_KEY &y&"-"&m&"-1"& ZC_SQL_POUND_KEY &" AND "& ZC_SQL_POUND_KEY &ny&"-"&nm&"-1"& ZC_SQL_POUND_KEY &")")

	If (Not objRS.bof) And (Not objRS.eof) Then
		For i=1 To objRS.RecordCount
			j=CInt(Day(CDate(objRS("log_PostTime"))))
			aryDateLink(j)=True
			aryDateID(j)=objRS("log_ID")
			Set aryDateArticle(j)=New TArticle
			aryDateArticle(j).LoadInfobyArray Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))
			objRS.MoveNext
			If objRS.eof Then Exit For
		Next
	End If
	objRS.Close
	Set objRS=Nothing
'//////////////////////////////////////////////////////////

	s="catalog.asp?date="&y&"-"&(m-1)
	t="catalog.asp?date="&y&"-"&(m+1)
	If m=1 Then s="catalog.asp?date="&(y-1)&"-12"
	If m=12 Then t="catalog.asp?date="&(y+1)&"-1"

	strCalendar=strCalendar & "<div class=""year"&y&" month"&m&""">"
	strCalendar=strCalendar & "<p class=""y""><a href=""<#ZC_BLOG_HOST#>"&s&""">&lt;&lt;</a>  <a href=""<#ZC_BLOG_HOST#>"&"catalog.asp?date="&y&"-"&m&""">"&y&"-"&m&"</a>  <a href=""<#ZC_BLOG_HOST#>"&t&""">&gt;&gt;</a></p>"
	strCalendar=strCalendar & "<p class=""w"">"&ZVA_Week_Abbr(1)&"</p><p class=""w"">"&ZVA_Week_Abbr(2)&"</p><p class=""w"">"&ZVA_Week_Abbr(3)&"</p><p class=""w"">"&ZVA_Week_Abbr(4)&"</p><p class=""w"">"&ZVA_Week_Abbr(5)&"</p><p class=""w"">"&ZVA_Week_Abbr(6)&"</p><p class=""w"">"&ZVA_Week_Abbr(7)&"</p>"
	j=0
	For i=1 to b
		If (j=>firw-1) and (k=<d) Then
			If aryDateLink(k) Then
				strCalendar=strCalendar & "<p id=""pCalendar_"&y&"_"&m&"_"&k&""" class=""yd""><a class=""l"" href=""<#ZC_BLOG_HOST#>"&"catalog.asp?date="&Year(aryDateArticle(k).PostTime)&"-"&Month(aryDateArticle(k).PostTime)&"-"&Day(aryDateArticle(k).PostTime)& """>"&(k)&"</a></p>"
			Else
				strCalendar=strCalendar & "<p id=""pCalendar_"&y&"_"&m&"_"&k&""" class=""d"">"&(k)&"</p>"
			End If

			k=k+1
		Else
			strCalendar=strCalendar & "<p class=""nd""></p>"
		End If
		j=j+1
	Next

	strCalendar=strCalendar & "</div>"
	MakeCalendar=strCalendar

End Function
'*********************************************************




'*********************************************************
' 目的：    加载默认的主题模板
'*********************************************************
Function LoadDefaultTemplates()

If TemplateDic.Exists("TEMPLATE_WAP_ARTICLE_COMMENT")=False Then Call TemplateDic.add("TEMPLATE_WAP_ARTICLE_COMMENT",LoadFromFile(BlogPath &"zb_system\defend\wap\wap_article_comment","utf-8"))
If TemplateDic.Exists("TEMPLATE_WAP_ARTICLE-MULTI")=False Then Call TemplateDic.add("TEMPLATE_WAP_ARTICLE-MULTI",LoadFromFile(BlogPath &"zb_system\defend\wap\wap_article-multi.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_WAP_SINGLE")=False Then Call TemplateDic.add("TEMPLATE_WAP_SINGLE",LoadFromFile(BlogPath &"zb_system\defend\wap\wap_single.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE-ISTOP")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE-ISTOP",LoadFromFile(BlogPath &"zb_system\defend\default\b_article-istop.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE-MULTI")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE-MULTI",LoadFromFile(BlogPath &"zb_system\defend\default\b_article-multi.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE-SINGLE")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE-SINGLE",LoadFromFile(BlogPath &"zb_system\defend\default\b_article-single.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_COMMENT")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_COMMENT",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_comment.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_COMMENTPOST-VERIFY")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_COMMENTPOST-VERIFY",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_commentpost-verify.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_COMMENTPOST")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_COMMENTPOST",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_commentpost.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_comment_pagebar.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_MUTUALITY")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_MUTUALITY",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_mutuality.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_NVABAR_L")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_NVABAR_L",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_nvabar_l.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_NVABAR_R")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_NVABAR_R",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_nvabar_r.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_TAG")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_TAG",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_tag.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_FUNCTION")=False Then Call TemplateDic.add("TEMPLATE_B_FUNCTION",LoadFromFile(BlogPath &"zb_system\defend\default\b_function.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_PAGEBAR")=False Then Call TemplateDic.add("TEMPLATE_B_PAGEBAR",LoadFromFile(BlogPath &"zb_system\defend\default\b_pagebar.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_CATALOG")=False Then Call TemplateDic.add("TEMPLATE_CATALOG",LoadFromFile(BlogPath &"zb_system\defend\default\catalog.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_DEFAULT")=False Then Call TemplateDic.add("TEMPLATE_DEFAULT",LoadFromFile(BlogPath &"zb_system\defend\default\default.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_FOOTER")=False Then Call TemplateDic.add("TEMPLATE_FOOTER",LoadFromFile(BlogPath &"zb_system\defend\default\footer.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_HEADER")=False Then Call TemplateDic.add("TEMPLATE_HEADER",LoadFromFile(BlogPath &"zb_system\defend\default\header.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_SEARCH")=False Then Call TemplateDic.add("TEMPLATE_SEARCH",LoadFromFile(BlogPath &"zb_system\defend\default\search.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_SIDEBAR")=False Then Call TemplateDic.add("TEMPLATE_SIDEBAR",LoadFromFile(BlogPath &"zb_system\defend\default\.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_SINGLE")=False Then Call TemplateDic.add("TEMPLATE_SINGLE",LoadFromFile(BlogPath &"zb_system\defend\default\single.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_TAGS")=False Then Call TemplateDic.add("TEMPLATE_TAGS",LoadFromFile(BlogPath &"zb_system\defend\default\tags.html","utf-8"))

	Dim i,j
	'在模板文件中先替换一次模板INCLUDE里的文件标签
	For Each i In TemplateDic.Keys
		For Each j In TemplateDic.Keys
			TemplateDic.Item(j)=Replace(TemplateDic.Item(j),"<#"+i+"#>",TemplateDic.Item(i))
		Next
	Next


End Function
'*********************************************************




'*********************************************************
' 目的：    加载指定目录的文件列表
'*********************************************************
Function LoadIncludeFiles(strDir)

	Dim aryFileList()
	ReDim aryFileList(-1)

	Dim fso, f, f1, fc, s, i
	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FolderExists(BlogPath & strDir)=False Then
		LoadIncludeFiles=Array()
		Exit Function
	End If

	Set f = fso.GetFolder(BlogPath & strDir)
	Set fc = f.Files

	i=0

	For Each f1 in fc
		i=i+1
		ReDim Preserve aryFileList(i)
		aryFileList(i)=f1.name
	Next

	LoadIncludeFiles=aryFileList

	Set fso=nothing

End Function
'*********************************************************



'*********************************************************
' 目的：    加载指定目录的文件列表
'*********************************************************
Function LoadIncludeFilesOnlyType(strDir)

	Dim aryFileList()
	ReDim aryFileList(-1)

	Dim fso, f, f1, fc, s, i
	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FolderExists(BlogPath & strDir)=False Then
		LoadIncludeFilesOnlyType=Array()
		Exit Function
	End If

	Set f = fso.GetFolder(BlogPath & strDir)
	Set fc = f.Files

	i=0

	For Each f1 in fc
		If Right(f1.name,5)=".html" Or Right(f1.name,4)=".asp" Or Right(f1.name,4)=".htm"  Then
			i=i+1
			ReDim Preserve aryFileList(i)
			aryFileList(i)=f1.name
		End If
	Next

	LoadIncludeFilesOnlyType=aryFileList

	Set fso=nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Get Template by Name
'*********************************************************
Function GetTemplate(Name)

	GetTemplate=TemplateDic.Item(Name)

End Function
'*********************************************************




'*********************************************************
' 目的：    Set Template by Name
'*********************************************************
Function SetTemplate(Name,Value)

	TemplateDic.Item(Name)=Value

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function GetTemplateTags(Name)
	GetTemplateTags=TemplateTagsDic.Item(Name)
End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function SetTemplateTags(Name,Value)
	TemplateTagsDic.Item(Name)=Value
End Function
'*********************************************************




'*********************************************************
' 目的：    Check Template Modified Date
'*********************************************************
Function CheckTemplateModified()

	Dim fso, f, f1, fc, s
	Dim d,nd

	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(BlogPath & "zb_users\" & "theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY)=False Then Exit Function
	Set f = fso.GetFolder(BlogPath & "zb_users\" & "theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY)
	Set fc = f.Files

	For Each f1 in fc
	  d=f1.DateLastModified
	  If nd="" Then nd=d
	  If DateDiff("s",nd,d)>0 Then nd=d
	Next

	CheckTemplateModified=nd

End Function
'*********************************************************




'*********************************************************
' 目的：    Load 全局 Cache
'*********************************************************
Function LoadGlobeCache()

	'On Error Resume Next


	Dim ii,jj

	Dim bolReLoadCache
	Application.Lock
	bolReLoadCache=Application(ZC_BLOG_CLSID & "SIGNAL_RELOADCACHE")
	Application.UnLock

	If IsEmpty(bolReLoadCache)=True Then
		bolReLoadCache="ok"
	Else
		Dim TemplatesName
		Dim TemplatesContent

		Dim TemplateTagsName
		Dim TemplateTagsValue

		TemplateTagsDic.RemoveAll
		TemplateDic.RemoveAll

		Application.Lock
		TemplateTagsName=Application(ZC_BLOG_CLSID & "TemplateTagsName")
		TemplateTagsValue=Application(ZC_BLOG_CLSID & "TemplateTagsValue")

		TemplatesName=Application(ZC_BLOG_CLSID & "TemplatesName")
		TemplatesContent=Application(ZC_BLOG_CLSID & "TemplatesContent")
		Application.UnLock

		jj=UBound(TemplatesName)
		For ii=0 to jj-1
			If TemplateDic.Exists(TemplatesName(ii))=False Then TemplateDic.Add TemplatesName(ii), TemplatesContent(ii)
		Next

		jj=UBound(TemplateTagsName)
		For ii=0 to jj-1
			If TemplateTagsDic.Exists(TemplateTagsName(ii))=False Then TemplateTagsDic.Add TemplateTagsName(ii), TemplateTagsValue(ii)
		Next

		'Call LoadDefaultTemplates()

		If IsEmpty(TemplateTagsValue)=False And IsEmpty(TemplateTagsValue)=False And IsEmpty(TemplatesName)=False And IsEmpty(TemplatesContent)=False Then
			Exit Function
		End If
	End If

	Dim i,j

	'加载模板
	Dim objStream
	Dim strContent


	Dim aryTemplatesName()
	Dim aryTemplatesContent()

	ReDim Preserve aryTemplatesName(3)
	ReDim Preserve aryTemplatesContent(3)

	'加载WAP
	'Application.Lock
	aryTemplatesName(1)="TEMPLATE_WAP_ARTICLE_COMMENT"
	aryTemplatesName(2)="TEMPLATE_WAP_ARTICLE-MULTI"
	aryTemplatesName(3)="TEMPLATE_WAP_SINGLE"
	aryTemplatesContent(1)=LoadFromFile(BlogPath & "zb_system\defend\wap\wap_article_comment.html","utf-8")
	aryTemplatesContent(2)=LoadFromFile(BlogPath & "zb_system\defend\wap\wap_article-multi.html","utf-8")
	aryTemplatesContent(3)=LoadFromFile(BlogPath & "zb_system\defend\wap\wap_single.html","utf-8")


	'读取Template目录下的所有文件并写入Cache
	Dim aryFileList
	Dim aryFileNameTemplate()
	Dim aryFileNameTemplate_Variable()

	aryFileList=LoadIncludeFilesOnlyType("zb_users\theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY)

	If IsArray(aryFileList) Then

		j=UBound(aryFileList)

		If j>0 Then

			ReDim aryFileNameTemplate(j)
			ReDim aryFileNameTemplate_Variable(j)

			ReDim Preserve aryTemplatesName(3+j)
			ReDim Preserve aryTemplatesContent(3+j)

			For i=1 to j

				aryFileNameTemplate(i)="theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY & "/" & aryFileList(i)
				aryFileNameTemplate_Variable(i)="TEMPLATE_" & UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
				If InStr(aryFileList(i),".")=0 Then
					aryFileNameTemplate_Variable(i)="TEMPLATE_" & UCase(aryFileList(i))
				End If
				aryTemplatesName(3+i)=aryFileNameTemplate_Variable(i)

				strContent=""
				strContent=LoadFromFile(BlogPath & "zb_users\" & aryFileNameTemplate(i),"utf-8")

				'Application.Lock
				'Application(ZC_BLOG_CLSID & aryFileNameTemplate_Variable(i))=strContent
				'Application.UnLock

				aryTemplatesContent(3+i)=strContent
			Next

			'在模板文件中先替换当前模版内的文件标签
			For i=1 To UBound(aryTemplatesName)
				For j=1 to UBound(aryTemplatesName)
					aryTemplatesContent(i)=Replace(aryTemplatesContent(i),"<#"+aryTemplatesName(j)+"#>",aryTemplatesContent(j))
				Next
				aryTemplatesContent(i)=Replace(aryTemplatesContent(i),"<#ZC_BLOG_HOST#>themes/","<#ZC_BLOG_HOST#>zb_users/theme/")
			Next
			j=UBound(aryFileList)

		Else
			j=0
		End If

	End If


	'读取Cache目录下的所有侧栏文件并写入Cache

	ReDim Preserve aryTemplatesName(3+j+5)
	ReDim Preserve aryTemplatesContent(3+j+5)


	aryTemplatesName(3+j+1)="CACHE_SIDEBAR"
	aryTemplatesName(3+j+2)="CACHE_SIDEBAR2"
	aryTemplatesName(3+j+3)="CACHE_SIDEBAR3"
	aryTemplatesName(3+j+4)="CACHE_SIDEBAR4"
	aryTemplatesName(3+j+5)="CACHE_SIDEBAR5"

	aryTemplatesContent(3+j+1)=LoadFromFile(BlogPath & "zb_users\cache" & "\sidebar.html","utf-8" )
	aryTemplatesContent(3+j+2)=LoadFromFile(BlogPath & "zb_users\cache" & "\sidebar2.html","utf-8")
	aryTemplatesContent(3+j+3)=LoadFromFile(BlogPath & "zb_users\cache" & "\sidebar3.html","utf-8")
	aryTemplatesContent(3+j+4)=LoadFromFile(BlogPath & "zb_users\cache" & "\sidebar4.html","utf-8")
	aryTemplatesContent(3+j+5)=LoadFromFile(BlogPath & "zb_users\cache" & "\sidebar5.html","utf-8")


	'加载标签
	Dim a,b,c,d,e,a2,a3,f
	Dim t()
	Dim s()

	a=0
	b=21
	c=1
	d=300
	e=0
	a2=0
	a3=0
	f=1

	'读取TEMPLATE下的Include目录下的所有文件并写入Cache
	'Dim aryFileList
	Dim aryFileNameTemplateInclude()
	Dim aryFileNameTemplateInclude_Variable()

	aryFileList=LoadIncludeFilesOnlyType("zb_users\theme" & "/" & ZC_BLOG_THEME & "/" & "INCLUDE")

	If IsArray(aryFileList) Then

		e=UBound(aryFileList)

		If e>0 Then

			ReDim aryFileNameTemplateInclude(e)
			ReDim aryFileNameTemplateInclude_Variable(e)
			ReDim aryFileNameTemplateInclude_Content(e)
			ReDim s(e)
			ReDim Preserve aryTemplateTagsName(e)
			ReDim Preserve aryTemplateTagsValue(e)

			For i=1 to e

				aryFileNameTemplateInclude(i)="theme" & "/" & ZC_BLOG_THEME & "/" & "INCLUDE" & "/" & aryFileList(i)
				aryFileNameTemplateInclude_Variable(i)="TEMPLATE_INCLUDE_" & UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
				If InStr(aryFileList(i),".")=0 Then
					aryFileNameTemplateInclude_Variable(i)="TEMPLATE_INCLUDE_" & UCase(aryFileList(i))
				End If

				s(i)=aryFileNameTemplateInclude_Variable(i)

				strContent=""
				strContent=LoadFromFile(BlogPath & "zb_users\" & aryFileNameTemplateInclude(i),"utf-8")
				strContent=Replace(strContent,"<"&"%=ZC_BLOG_HOST%"&">","<#ZC_BLOG_HOST#>")

				aryFileNameTemplateInclude_Content(i)=strContent

				aryTemplateTagsName(i)=s(i)
				aryTemplateTagsValue(i)=aryFileNameTemplateInclude_Content(i)
			Next

		Else
			e=0
		End If

	End If


	'读取Include目录下的所有文件并写入Cache
	'Dim aryFileList
	Dim aryFileNameInclude()
	Dim aryFileNameInclude_Variable()
	Dim aryFileNameInclude_Content()

	aryFileList=LoadIncludeFilesOnlyType("zb_users\INCLUDE")

	If IsArray(aryFileList) Then

		a=UBound(aryFileList)
		
		If a>0 Then

			ReDim aryFileNameInclude(a)
			ReDim aryFileNameInclude_Variable(a)
			ReDim aryFileNameInclude_Content(a)
			ReDim s(a)
			ReDim Preserve aryTemplateTagsName(e+a)
			ReDim Preserve aryTemplateTagsValue(e+a)

			For i=1 to a

				aryFileNameInclude(i)="INCLUDE/" & aryFileList(i)
				aryFileNameInclude_Variable(i)="CACHE_INCLUDE_" & UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
				If InStr(aryFileList(i),".")=0 Then
					aryFileNameInclude_Variable(i)="CACHE_INCLUDE_" & UCase(aryFileList(i))
				End If

				s(i)=aryFileNameInclude_Variable(i)

				strContent=""
				strContent=LoadFromFile(BlogPath & "zb_users\" & aryFileNameInclude(i),"utf-8")
				strContent=Replace(strContent,"<"&"%=ZC_BLOG_HOST%"&">","<#ZC_BLOG_HOST#>")
				aryFileNameInclude_Content(i)=strContent

				aryTemplateTagsName(e+i)=s(i)
				aryTemplateTagsValue(e+i)=aryFileNameInclude_Content(i)

			Next


			a2=a
			ReDim Preserve aryTemplateTagsName(e+a+a2)
			ReDim Preserve aryTemplateTagsValue(e+a+a2)

			For i=1 to a

				aryTemplateTagsName(e+i+a)=aryFileNameInclude_Variable(i) & "_JS"

				Dim modname
				modname=LCase(Replace(aryFileNameInclude_Variable(i),"CACHE_INCLUDE_",""))

				Dim functionstype

				Set functionstype=New TMeta
				functionstype.LoadString=LoadFromFile(BlogPath & "zb_users\cache\functionstype.html","utf-8")
				
				If functionstype.GetValue(modname)="div" Then
					aryTemplateTagsValue(e+i+a)="<div id=""mod_"+modname+""" style=""display:none;""><script type=""text/javascript"">LoadFunction('"&modname&"');</script></div>"
				Else
					aryTemplateTagsValue(e+i+a)="<li id=""mod_"+modname+""" style=""display:none;""><script type=""text/javascript"">LoadFunction('"&modname&"');</script></li>"
				End If

			Next

			a3=a
			ReDim Preserve aryTemplateTagsName(e+a+a2+a3)
			ReDim Preserve aryTemplateTagsValue(e+a+a2+a3)

			For i=1 to a
				aryTemplateTagsName(e+i+a+a2)=aryFileNameInclude_Variable(i) & "_HTML"
				aryTemplateTagsValue(e+i+a+a2)=aryFileNameInclude_Content(i)
			Next

		Else
			a=0
		End If

	End If


	ReDim Preserve aryTemplateTagsName(a+a2+a3+e+d)
	ReDim Preserve aryTemplateTagsValue(a+a2+a3+e+d)
	For j=1 to d
		i=Right("000" & CStr(j),3)
		aryTemplateTagsName(a+a2+a3+e+j)="ZC_MSG" & i
		Call Execute("aryTemplateTagsValue(a+a2+a3+e+j)=ZC_MSG" & i)
	Next


	ReDim t(b)
	t(1)="ZC_BLOG_VERSION"
	t(2)="ZC_BLOG_LANGUAGE"
	t(3)="ZC_BLOG_HOST"
	t(4)="ZC_BLOG_TITLE"
	t(5)="ZC_BLOG_SUBTITLE"
	t(6)="ZC_BLOG_NAME"
	t(7)="ZC_BLOG_SUB_NAME"
	t(8)="ZC_BLOG_CSS"
	t(9)="ZC_BLOG_COPYRIGHT"
	t(10)="ZC_BLOG_MASTER"
	t(11)="ZC_CONTENT_MAX"
	t(12)="ZC_EMOTICONS_FILENAME"
	t(13)="ZC_EMOTICONS_FILESIZE"
	t(14)="ZC_EMOTICONS_FILETYPE"
	t(15)="ZC_GUESTBOOK_CONTENT"
	t(16)="ZC_BLOG_CLSID"
	t(17)="ZC_TIME_ZONE"
	t(18)="ZC_IMAGE_WIDTH"
	t(19)="ZC_BLOG_THEME"
	t(20)="ZC_VERIFYCODE_WIDTH"
	t(21)="ZC_VERIFYCODE_HEIGHT"



	ReDim Preserve aryTemplateTagsName(a+a2+a3+e+d+b)
	ReDim Preserve aryTemplateTagsValue(a+a2+a3+e+d+b)
	For j=1 to b
		aryTemplateTagsName(a+a2+a3+e+d+j)=t(j)
		Call Execute("aryTemplateTagsValue(a+a2+a3+e+d+j)="& t(j))
	Next

	ReDim Preserve aryTemplateTagsName(a+a2+a3+e+d+b+c)
	ReDim Preserve aryTemplateTagsValue(a+a2+a3+e+d+b+c)
	aryTemplateTagsName(a+a2+a3+e+d+b+c)="BLOG_CREATE_TIME"
	aryTemplateTagsValue(a+a2+a3+e+d+b+c)=GetTime(Now())

	ReDim Preserve aryTemplateTagsName(a+a2+a3+e+d+b+c+f)
	ReDim Preserve aryTemplateTagsValue(a+a2+a3+e+d+b+c+f)

	aryTemplateTagsName(a+a2+a3+e+d+b+c+f)="CACHE_INCLUDE_CALENDAR_NOW"
	aryTemplateTagsValue(a+a2+a3+e+d+b+c+f)=""



	Application.Lock
	Application(ZC_BLOG_CLSID & "TEMPLATEMODIFIED")=CheckTemplateModified()
	Application.UnLock

	Application.Lock
	Application(ZC_BLOG_CLSID & "SIGNAL_RELOADCACHE")=bolReLoadCache
	Application.UnLock


	TemplateTagsDic.RemoveAll
	TemplateDic.RemoveAll
	TemplateTagsDic.add "BlogTitle",""

	jj=UBound(aryTemplatesName)
	For ii=1 to jj
		If TemplateDic.Exists(aryTemplatesName(ii))=False Then TemplateDic.Add aryTemplatesName(ii), aryTemplatesContent(ii)
	Next

	jj=UBound(aryTemplateTagsName)
	For ii=1 to jj
		If TemplateTagsDic.Exists(aryTemplateTagsName(ii))=False Then TemplateTagsDic.Add aryTemplateTagsName(ii), aryTemplateTagsValue(ii)
	Next

	Call LoadDefaultTemplates()


	Application.Lock
	Application(ZC_BLOG_CLSID & "TemplateTagsName")=TemplateTagsDic.Keys
	Application(ZC_BLOG_CLSID & "TemplateTagsValue")=TemplateTagsDic.Items


	Application(ZC_BLOG_CLSID & "TemplatesName")=TemplateDic.Keys
	Application(ZC_BLOG_CLSID & "TemplatesContent")=TemplateDic.Items

	Application.UnLock


	Set TemplatesName=Nothing
	Set TemplatesContent=Nothing

	Set TemplateTagsName=Nothing
	Set TemplateTagsValue=Nothing

	LoadGlobeCache=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Clear Cache
'*********************************************************
Function ClearGlobeCache()

	Application.Lock


	Application(ZC_BLOG_CLSID & "TemplateTagsName")=Empty
	Application(ZC_BLOG_CLSID & "TemplateTagsValue")=Empty

	Application(ZC_BLOG_CLSID & "TemplatesName")=Empty
	Application(ZC_BLOG_CLSID & "TemplatesContent")=Empty


	Application(ZC_BLOG_CLSID & "SIGNAL_RELOADCACHE")=Empty

	Application(ZC_BLOG_CLSID & "TEMPLATEMODIFIED")=Empty

	Application(ZC_BLOG_CLSID & "CACHE_ARTICLE_VIEWCOUNT")=Empty


	Application.UnLock

	ClearGlobeCache=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Parse Tag 并格式化
'*********************************************************
Function ParseTag(strTag)

	Dim s
	Dim t
	Dim i
	Dim Tag
	Dim b
	Dim objTag

	strTag=Trim(strTag)
	strTag=Replace(strTag,",",vbCrlf)
	strTag=Replace(strTag,"，",vbCrlf)
	strTag=TransferHTML(strTag,"[normalname]")
	strTag=Replace(strTag,vbCrlf,",")

	If strTag="" Then ParseTag="":Exit Function

	t=Split(strTag,",")

	Call GetTagsbyTagNameList(strTag)

	For i=LBound(t) To UBound(t)
		t(i)=Trim(t(i))
	Next

	For i=LBound(t) To UBound(t)

		b=False

		For Each Tag in Tags
			If IsObject(Tag) Then
				If UCase(Tag.Name)=UCase(t(i)) Then
					b=True
				End If
			End If
		Next

		If b=False Then
			Set objTag=New TTag
			objTag.ID=0
			objTag.Name=t(i)
			objTag.Order=0
			objTag.Intro=""
			If objTag.Post Then
				ReDim Preserve Tags(objTag.ID)
				Set Tags(objTag.ID)=objTag
			End If
			Set objTag=Nothing
		End If

	Next

	For i=LBound(t) To UBound(t)
		For Each Tag in Tags
			If IsObject(Tag) Then
				If UCase(Tag.Name)=UCase(t(i)) Then
					t(i)="{"&Tag.ID&"}"
				End If
			End If
		Next
	Next

	s=Join(t)
	s=Replace(s," ","")

	ParseTag=s

End Function
'*********************************************************




'*********************************************************
' 目的：    得到实际上的BLOG的真实物理目录
'*********************************************************
Dim IsRunGetReallyDirectory
IsRunGetReallyDirectory=False
Function GetReallyDirectory()

	If IsRunGetReallyDirectory=True Then Exit Function

	'On Error Resume Next

	BlogPath=Server.MapPath(".")

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(BlogPath & "ZB_SYSTEM\") Then
		BlogPath=BlogPath
	ElseIf fso.FolderExists(BlogPath & "..\ZB_SYSTEM\") Then
		BlogPath=BlogPath & "..\"
	ElseIf fso.FolderExists(BlogPath & "..\..\ZB_SYSTEM\") Then
		BlogPath=BlogPath & "..\..\"
	ElseIf fso.FolderExists(BlogPath & "..\..\..\ZB_SYSTEM\") Then
		BlogPath=BlogPath & "..\..\..\"
	ElseIf fso.FolderExists(BlogPath & "..\..\..\..\ZB_SYSTEM\") Then
		BlogPath=BlogPath & "..\..\..\..\"
	ElseIf fso.FolderExists(BlogPath & "..\..\..\..\..\ZB_SYSTEM\") Then
		BlogPath=BlogPath & "..\..\..\..\..\"
	ElseIf fso.FolderExists(BlogPath & "..\..\..\..\..\..\ZB_SYSTEM\") Then
		BlogPath=BlogPath & "..\..\..\..\..\..\"
	ElseIf fso.FolderExists(BlogPath & "..\..\..\..\..\..\..\ZB_SYSTEM\") Then
		BlogPath=BlogPath & "..\..\..\..\..\..\..\"
	End If
	Set fso=Nothing

	GetReallyDirectory=CreateObject("Scripting.FileSystemObject").GetFolder(BlogPath).Path & "\"

	IsRunGetReallyDirectory=True

	'Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    设置提示标志
'*********************************************************
Function SetBlogHint(bolOperateSuccess,bolRebuildIndex,bolRebuildFiles)

	Call SetBlogHintWithCLSID(bolOperateSuccess,bolRebuildIndex,bolRebuildFiles,ZC_BLOG_CLSID)

End Function
'*********************************************************


'*********************************************************
' 目的：    设置提示自定义标志
'*********************************************************
Function SetBlogHint_Custom(strInfo)

	Application.Lock

	Application(ZC_BLOG_CLSID & "SIGNAL_OPERATEINFO")=Application(ZC_BLOG_CLSID & "SIGNAL_OPERATEINFO") & vbCrlf &  strInfo

	Application.UnLock

End Function
'*********************************************************



'*********************************************************
' 目的：    设置提示标志withCLSID
'*********************************************************
Function SetBlogHintWithCLSID(bolOperateSuccess,bolRebuildIndex,bolRebuildFiles,newCLSID)

	Application.Lock

	Application(newCLSID & "SIGNAL_OPERATESUCCESS")=bolOperateSuccess

	If IsEmpty(bolRebuildIndex)=False Then
		Application(newCLSID & "SIGNAL_REBUILDINDEX")=bolRebuildIndex
	End If

	If IsEmpty(bolRebuildFiles)=False Then
		Application(newCLSID & "SIGNAL_REBUILDFILES")=bolRebuildFiles
	End If

	Application.UnLock

End Function
'*********************************************************




'*********************************************************
' 目的：    输出提示
'*********************************************************
Function GetBlogHint()

	Dim bolOperateSuccess,bolRebuildIndex,bolRebuildFiles,strOperateInfo

	Application.Lock
	bolOperateSuccess=Application(ZC_BLOG_CLSID & "SIGNAL_OPERATESUCCESS")
	bolRebuildIndex=Application(ZC_BLOG_CLSID & "SIGNAL_REBUILDINDEX")
	bolRebuildFiles=Application(ZC_BLOG_CLSID & "SIGNAL_REBUILDFILES")
	strOperateInfo=Application(ZC_BLOG_CLSID & "SIGNAL_OPERATEINFO")
	Application(ZC_BLOG_CLSID & "SIGNAL_OPERATEINFO")=Empty
	Application.UnLock


	If IsEmpty(bolOperateSuccess)=False Then

		If bolOperateSuccess=True Then
			Response.Write "<div class='hint'><p class='hint hint_green'><font color='green'>" & ZC_MSG266 & "</font></p></div>"
		End If

		If bolOperateSuccess=False Then
			Response.Write "<div class='hint'><p class='hint hint_red'><font color='red'>" & ZC_MSG267 & "</font></p></div>"
		End If

		Application.Lock
		Application(ZC_BLOG_CLSID & "SIGNAL_OPERATESUCCESS")=Empty
		Application.UnLock

	End If

	If IsEmpty(strOperateInfo)=False Then

		Dim s,t
		t=Split(strOperateInfo,vbCrlf)

		For Each s In t
			If s<>"" Then
			Response.Write "<div class='hint'><p class='hint hint_teal'><font color='orangered'>" & s & "</font></p></div>"
			End If
		Next

	End If


	If IsEmpty(bolRebuildIndex)=False Then

		If bolRebuildIndex=True Then
			Response.Write "<div class='hint'><p class='hint hint_blue'><font color='blue'>" & ZC_MSG268 & "</font></p></div>"
		End If

	End If

	If IsEmpty(bolRebuildFiles)=False Then

		If bolRebuildFiles=True Then
			Response.Write "<div class='hint'><p class='hint hint_blue'><font color='blue'>" & Replace(ZC_MSG269,"%u",GetCurrentHost()&"zb_system/admin/admin.asp?act=AskFileReBuild") & "</font></p></div>"
		End If

	End If


End Function
'*********************************************************




'*********************************************************
' 目的：    解析ZC_CUSTOM_DIRECTORY_REGEX
'*********************************************************
Function ParseCustomDirectory(strRegex,strPost,strCategory,strUser,strYear,strMonth,strDay,strID,strAlias)

	On Error Resume Next

	Dim s
	s=strRegex

	s=Replace(s,"{%post%}",strPost)
	s=Replace(s,"{%category%}",strCategory)
	s=Replace(s,"{%user%}",strUser)
	s=Replace(s,"{%year%}",strYear)
	s=Replace(s,"{%month%}",Right("0" & strMonth,2))
	s=Replace(s,"{%day%}",Right("0" & strDay,2))
	s=Replace(s,"{%id%}",strID)
	s=Replace(s,"{%alias%}",strAlias)

	ParseCustomDirectory=s

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    按照CustomDirectory指示创建相应的目录
'*********************************************************
Sub CreatDirectoryByCustomDirectory(strCustomDirectory)
	On Error Resume Next

	Dim s
	Dim t
	Dim i
	Dim j

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

	s=BlogPath

	strCustomDirectory=Replace(strCustomDirectory,"/","\")

	t=Split(strCustomDirectory,"\")

	j=0
	For i=LBound(t) To UBound(t)
		If (IsEmpty(t(i))=False) And (t(i)<>"") Then
			'If j=0 And LCase(Left(t(i),3))="zb_" Then Exit For
			s=s & t(i) & "\"
			If (fso.FolderExists(fldr)=False) Then
				Call fso.CreateFolder(s)
			End If
			j=j+1
		End If
	Next

	Set fso = Nothing

	Err.Clear

End Sub
'*********************************************************




'*********************************************************
' 目的：  生成左侧导航栏
'*********************************************************
Dim AdminLeftMenuCount
AdminLeftMenuCount=0
Function MakeLeftMenu(requireLevel,strName,strUrl,strLiId,strAName,strImgUrl)

	If BlogUser.Level>requireLevel Then Exit Function

	AdminLeftMenuCount=AdminLeftMenuCount+1
	dim tmp
	If Trim(strImgUrl)<>"" Then
		tmp="<li id="""&strLiId&"""><a id="""&strAName&""" href="""&strUrl&"""><span style=""background-image:url('"&strImgUrl&"')"">"&strName&"</span></a></li>"
	Else
		tmp="<li id="""&strLiId&"""><a id="""&strAName&""" href="""&strUrl&"""><span>"&strName&"</span></a></li>"
	End If
	MakeLeftMenu=tmp
	
End Function
'*********************************************************




'*********************************************************
' 目的：  生成头部菜单
'*********************************************************
Dim AdminTopMenuCount
AdminTopMenuCount=0
Function MakeTopMenu(strName,strUrl,strTarget)
	Dim tmp
	If strTarget="" Then strTarget="_self"
	AdminTopMenuCount=AdminTopMenuCount+1
	tmp="<li id=""topmenu"&AdminTopMenuCount&"""><a href="""&strUrl&""" target="""&strTarget&""">"&strName&"</a></li>"
	MakeTopMenu=tmp
End Function
'*********************************************************




'*********************************************************
' 目的： 加入二级菜单项
'*********************************************************
Function MakeSubMenu(strName,strUrl,strType,isNewWindows)

	Dim strSource

	strSource=strSource & "<span class=""" & strType & """>"

	strSource=strSource & "<a " & "href=""" & strUrl  & """"

	If isNewWindows=True Then strSource=strSource & " target=""_blank"""

	strSource=strSource & ">" & strName

	strSource=strSource & "</a></span>"

	MakeSubMenu=strSource

End Function
'*********************************************************




'*********************************************************
' 目的： 注册插件函数
'*********************************************************
Function RegisterPlugin(strPluginName,strPluginActiveFunction)

	Dim i
	i=UBound(PluginName)

	ReDim Preserve PluginName(i+1)
	ReDim Preserve PluginActiveFunction(i+1)

	PluginName(i)=strPluginName
	PluginActiveFunction(i)=strPluginActiveFunction

End Function
'*********************************************************




'*********************************************************
' 目的： 激活插件函数
'*********************************************************
Function ActivePlugin()

	On Error Resume Next

	Dim i
	For i=0 To UBound(PluginActiveFunction)-1

		Call Execute(PluginActiveFunction(i))

	Next

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的： 安装插件函数，只运行一次
'*********************************************************
Function InstallPlugin(strPluginName)
	On Error Resume Next
	Call Execute("Call InstallPlugin_" & strPluginName & "()")
	Err.Clear
End Function
'*********************************************************




'*********************************************************
' 目的： 删除插件函数，只运行一次
'*********************************************************
Function UninstallPlugin(strPluginName)
	On Error Resume Next
	Call Execute("Call UninstallPlugin_" & strPluginName & "()")
	Err.Clear
End Function
'*********************************************************




'*********************************************************
' 目的： 检测插件是否已激活
'*********************************************************
Function CheckPluginState(strPluginName)

	CheckPluginState=CheckPluginStateByNewValue(strPluginName,ZC_BLOG_THEME & "|" & ZC_USING_PLUGIN_LIST)

End Function
'*********************************************************



'*********************************************************
' 目的： 检测插件是否已激活 by new value
'*********************************************************
Function CheckPluginStateByNewValue(strPluginName,newZC_USING_PLUGIN_LIST)

	Dim s,i
	s=Split(newZC_USING_PLUGIN_LIST,"|")

	For i=LBound(s) To UBound(s)
		If UCase(s(i))=UCase(strPluginName) Then
			CheckPluginStateByNewValue=True
			Exit Function
		End If
	Next

	CheckPluginStateByNewValue=False

End Function
'*********************************************************




'*********************************************************
' 目的：挂上Action接口
' 参数：'plugname:接口名称
		'actioncode:要执行的语句，要转义为Execute可执行语句
'*********************************************************
Function Add_Action_Plugin(plugname,actioncode)
	On Error Resume Next
	actioncode=Replace(actioncode,"Exit Function","b" & plugname & "=True")
	actioncode=Replace(actioncode,"Exit Sub","b" & plugname & "=True")
	Call Execute("ReDim Preserve " & plugname & "(UBound("& plugname &")+1)")
	Call Execute(plugname & "(UBound("& plugname &"))=" & plugname & "(UBound("& plugname &"))&""" & Replace(actioncode,"""","""""") & """" & ":")
	Err.Clear
End Function
'*********************************************************




'*********************************************************
' 目的：挂上Filter接口
' 参数：'plugname:接口名称
		'functionname:要挂接的函数名
'*********************************************************
Function Add_Filter_Plugin(plugname,functionname)
	On Error Resume Next
	Call Execute("s" & plugname & "=" & "s" & plugname & "&""" & functionname & """" & "& ""|""")
	Err.Clear
End Function
'*********************************************************




'*********************************************************
' 目的：挂上Response接口
' 参数：'plugname:接口名称
		'parameter:要写入的内容
'*********************************************************
Function Add_Response_Plugin(plugname,parameter)
	On Error Resume Next
	Call Execute(plugname & "=" & plugname & "&""" & Replace(parameter,"""","""""") & """")
	Err.Clear
End Function
'*********************************************************


'*********************************************************
' 目的：GetSettingFormName
'*********************************************************
Function GetSettingFormName(s)
	On Error Resume Next
	Dim x
	Call Execute("x=" & s)
	GetSettingFormName=x
	Err.Clear
End Function
'*********************************************************




'*********************************************************
' 目的：GetSettingFormName with Default
'*********************************************************
Function GetSettingFormNameWithDefault(s,d)
	On Error Resume Next
	Err.Clear
	Dim x
	Call Execute("x=" & s)
	GetSettingFormNameWithDefault=x
	If Err.Number<>0 Then
		GetSettingFormNameWithDefault=d
	End If
	Err.Clear
End Function
'*********************************************************



'*********************************************************
' 目的：GetNameFormTheme
'*********************************************************
Function GetNameFormTheme(s)
	On Error Resume Next

	GetNameFormTheme=s
	Dim objXmlFile
	Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
	objXmlFile.async = False
	objXmlFile.ValidateOnParse=False
	objXmlFile.load(BlogPath & "zb_users\" & "theme" & "\" & s & "\" & "theme.xml")
	If objXmlFile.readyState=4 Then
		If objXmlFile.parseError.errorCode <> 0 Then
		Else
			GetNameFormTheme=objXmlFile.documentElement.selectSingleNode("name").text
		End If
	End If

	Err.Clear
End Function
'*********************************************************




'*********************************************************
' 目的：    Blog ReBuild 核心
'*********************************************************
Function MakeBlogReBuild_Core()

	'On Error Resume Next

	'plugin node
	bAction_Plugin_MakeBlogReBuild_Core_Begin=False
	For Each sAction_Plugin_MakeBlogReBuild_Core_Begin in Action_Plugin_MakeBlogReBuild_Core_Begin
		If Not IsEmpty(sAction_Plugin_MakeBlogReBuild_Core_Begin) Then Call Execute(sAction_Plugin_MakeBlogReBuild_Core_Begin)
		If bAction_Plugin_MakeBlogReBuild_Core_Begin=True Then Exit Function
	Next

	BlogReBuild_Statistics

	BlogReBuild_Archives

	BlogReBuild_Previous

	BlogReBuild_Comments

	BlogReBuild_TrackBacks

	BlogReBuild_Catalogs

	BlogReBuild_Calendar

	BlogReBuild_Authors

	BlogReBuild_Tags

	BlogReBuild_Categorys

	BlogReBuild_Functions

	BuildAllCache

	ExportRSS

	Call ClearGlobeCache()
	Call LoadGlobeCache()

	Dim bolOperateSuccess

	Application.Lock
	bolOperateSuccess=Application(ZC_BLOG_CLSID & "SIGNAL_OPERATESUCCESS")
	Application.UnLock

	Call SetBlogHint(bolOperateSuccess,False,Empty)

	MakeBlogReBuild_Core=True

	'plugin node
	bAction_Plugin_MakeBlogReBuild_Core_End=False
	For Each sAction_Plugin_MakeBlogReBuild_Core_End in Action_Plugin_MakeBlogReBuild_Core_End
		If Not IsEmpty(sAction_Plugin_MakeBlogReBuild_Core_End) Then Call Execute(sAction_Plugin_MakeBlogReBuild_Core_End)
		If bAction_Plugin_MakeBlogReBuild_Core_End=True Then Exit Function
	Next

	'Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    全新的部份索引程序
'*********************************************************
Function BuildAllCache()

	'plugin node
	bAction_Plugin_BuildAllCache_Begin=False
	For Each sAction_Plugin_BuildAllCache_Begin in Action_Plugin_BuildAllCache_Begin
		If Not IsEmpty(sAction_Plugin_BuildAllCache_Begin) Then Call Execute(sAction_Plugin_BuildAllCache_Begin)
		If bAction_Plugin_BuildAllCache_Begin=True Then Exit Function
	Next

	Call GetFunction()

	Dim strList

	Dim ArticleList
	Dim AuthList
	Dim CateList
	Dim TagsList

	Dim aryAllList()

	Dim objRS
	Dim i
	Dim j
	Dim n
	Dim l
	Dim k

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	objRS.Open("SELECT [log_ID] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_Level]>1) AND ([log_Istop]=0) ORDER BY [log_PostTime] DESC")

	If (Not objRS.bof) And (Not objRS.eof) Then

		objRS.PageSize = ZC_DISPLAY_COUNT
		ReDim aryAllList(objRS.PageCount+1)

		For i=1 to objRS.PageCount
			objRS.AbsolutePage=i
			For j = 1 To objRS.PageSize
				If j=1 Then aryAllList(i)="AllPage" & i & "["

				If i=1 Then
					aryAllList(i)=aryAllList(i) & objRS("log_ID") & ";"
				End If

				If j=objRS.PageSize Then aryAllList(i)=aryAllList(i) & "]"
				objRS.MoveNext
				If objRS.EOF Then aryAllList(i)=aryAllList(i) & "]":Exit For
			Next
		Next

	End If
	objRS.Close
	strList=strList & Join(aryAllList)
	Erase aryAllList



	objRS.Open("SELECT [log_ID] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_Level]>1) AND ([log_Istop]<>0) ORDER BY [log_PostTime] DESC")

	If (Not objRS.bof) And (Not objRS.eof) Then

		objRS.PageSize = ZC_DISPLAY_COUNT
		ReDim aryAllList(objRS.PageCount+1)

		For i=1 to objRS.PageCount
			objRS.AbsolutePage=i
			For j = 1 To objRS.PageSize
				If j=1 Then aryAllList(i)="IstopPage" & i & "["
				aryAllList(i)=aryAllList(i) & objRS("log_ID") & ";"
				If j=objRS.PageSize Then aryAllList(i)=aryAllList(i) & "]"
				objRS.MoveNext
				If objRS.EOF Then aryAllList(i)=aryAllList(i) & "]":Exit For
			Next
		Next

	End If
	objRS.Close
	strList=strList & Join(aryAllList)
	Erase aryAllList

	Call SaveToFile(BlogPath & "zb_users/CACHE/cache_list_"&ZC_BLOG_CLSID&".html",strList,"utf-8",False)

	BuildAllCache=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_Calendar()

	'plugin node
	bAction_Plugin_BlogReBuild_Calendar_Begin=False
	For Each sAction_Plugin_BlogReBuild_Calendar_Begin in Action_Plugin_BlogReBuild_Calendar_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Calendar_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Calendar_Begin)
		If bAction_Plugin_BlogReBuild_Calendar_Begin=True Then Exit Function
	Next

	Dim objStream
	Dim strCalendar
	Dim i,j
	Dim objRS
	Dim k,l,m,n

	'Calendar
	strCalendar=MakeCalendar("")

	strCalendar=TransferHTML(strCalendar,"[no-asp]")


	Call GetFunction()
	Functions(FunctionMetas.GetValue("calendar")).Content=strCalendar

	'Call SaveToFile(BlogPath & "zb_users/include/calendar.asp",strCalendar,"utf-8",True)

	BlogReBuild_Calendar=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_Archives()

	'plugin node
	bAction_Plugin_BlogReBuild_Archives_Begin=False
	For Each sAction_Plugin_BlogReBuild_Archives_Begin in Action_Plugin_BlogReBuild_Archives_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Archives_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Archives_Begin)
		If bAction_Plugin_BlogReBuild_Archives_Begin=True Then Exit Function
	Next

	Dim i
	Dim j
	Dim l
	Dim n
	Dim objRS
	Dim objStream

	Dim ArtList

	Call GetFunction()

	'Archives
	Dim strArchives
	Set objRS=objConn.Execute("SELECT * FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_Level]>1) ORDER BY [log_PostTime] DESC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		Dim dtmYM()
		i=0
		j=0
		ReDim Preserve dtmYM(0)
		Do While Not objRS.eof
			j=UBound(dtmYM)
			i=Year(objRS("log_PostTime")) & "-" & Month(objRS("log_PostTime"))
			If i<>dtmYM(j) Then
				ReDim Preserve dtmYM(j+1)
				dtmYM(j+1)=i
			End If
			objRS.MoveNext
		Loop
	End If
	objRS.Close
	Set objRS=Nothing

	j=Functions(FunctionMetas.GetValue("archives")).MaxLi

	If Not IsEmpty(dtmYM) Then
		For i=1 to UBound(dtmYM)

			l=Year(dtmYM(i))
			n=Month(dtmYM(i))+1
			IF n>12 Then l=l+1:n=1

			Set objRS=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_Level]>1) AND [log_PostTime] BETWEEN "& ZC_SQL_POUND_KEY & Year(dtmYM(i)) &"-"& Month(dtmYM(i)) &"-1"& ZC_SQL_POUND_KEY &" AND "& ZC_SQL_POUND_KEY & l &"-"& n &"-1" & ZC_SQL_POUND_KEY)

			If (Not objRS.bof) And (Not objRS.eof) Then
				'If CheckPluginState("STACentre") Then
				'	Dim objPostTime
				'	Set objPostTime=New STACentre_Archives
				'	If objPostTime.LoadInfoByID(Year(dtmYM(i)) & "-" & Month(dtmYM(i))) Then
				'		strArchives=strArchives & "<li><a href="""& objPostTime.Url & """>" & Year(dtmYM(i)) & " " & ZVA_Month(Month(dtmYM(i))) & "<span class=""article-nums""> (" & objRS(0) & ")</span>" +"</a></li>"
				'	End If
				'	Set objPostTime=Nothing
				'Else
					strArchives=strArchives & "<li><a href=""<#ZC_BLOG_HOST#>catalog.asp?date=" & Year(dtmYM(i)) & "-" & Month(dtmYM(i)) & """>" & Year(dtmYM(i)) & " " & ZVA_Month(Month(dtmYM(i))) & "<span class=""article-nums""> (" & objRS(0) & ")</span>" +"</a></li>"
				'End If
				If j>0 Then
					If i=j Then Exit For
				End If
			End If

			objRS.Close
			Set objRS=Nothing
		Next
	End If

	strArchives=TransferHTML(strArchives,"[no-asp]")


	Functions(FunctionMetas.GetValue("archives")).Content=strArchives

	'Call SaveToFile(BlogPath & "zb_users/include/archives.asp",strArchives,"utf-8",True)

	BlogReBuild_Archives=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_Catalogs()

	'plugin node
	bAction_Plugin_BlogReBuild_Catalogs_Begin=False
	For Each sAction_Plugin_BlogReBuild_Catalogs_Begin in Action_Plugin_BlogReBuild_Catalogs_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Catalogs_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Catalogs_Begin)
		If bAction_Plugin_BlogReBuild_Catalogs_Begin=True Then Exit Function
	Next

	IsRunGetCategory=False
	GetCategory()


	Dim objRS
	Dim objStream

	Dim ArtList

	'Catalogs
	Dim strCatalog,bolHasSubCate

	Dim aryCateInOrder 
	aryCateInOrder=GetCategoryOrder()


	Dim i,j
	For i=Lbound(aryCateInOrder)+1 To Ubound(aryCateInOrder)

			If Categorys(aryCateInOrder(i)).ParentID=0 Then
				strCatalog=strCatalog & "<li class=""li-cate""><a href="""& Categorys(aryCateInOrder(i)).Url & """>"+Categorys(aryCateInOrder(i)).Name + "<span class=""article-nums""> (" & Categorys(aryCateInOrder(i)).Count & ")</span>" +"</a></li>"

				bolHasSubCate=False
				For j=Lbound(aryCateInOrder) To UBound(aryCateInOrder)-1
					If Categorys(aryCateInOrder(j)).ParentID=Categorys(aryCateInOrder(i)).ID Then bolHasSubCate=True
				Next
				'If bolHasSubCate Then strCatalog=strCatalog & "<li class=""li-subcates""><ul class=""ul-subcates"">"
				For j=Lbound(aryCateInOrder) To UBound(aryCateInOrder)-1
					If Categorys(aryCateInOrder(j)).ParentID=Categorys(aryCateInOrder(i)).ID Then
						strCatalog=strCatalog & "<li class=""li-subcate""><a href="""& Categorys(aryCateInOrder(j)).Url & """>"+Categorys(aryCateInOrder(j)).Name + "<span class=""article-nums""> (" & Categorys(aryCateInOrder(j)).Count & ")</span>" +"</a></li>"
					End If
				Next
				'If bolHasSubCate Then strCatalog=strCatalog & "</ul></li>"
			End If

	Next


	strCatalog=TransferHTML(strCatalog,"[no-asp]")


	Call GetFunction()
	Functions(FunctionMetas.GetValue("catalog")).Content=strCatalog

	'Call SaveToFile(BlogPath & "zb_users/include/catalog.asp",strCatalog,"utf-8",True)

	BlogReBuild_Catalogs=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_Categorys()

	'plugin node
	bAction_Plugin_BlogReBuild_Categorys_Begin=False
	For Each sAction_Plugin_BlogReBuild_Categorys_Begin in Action_Plugin_BlogReBuild_Categorys_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Categorys_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Categorys_Begin)
		If bAction_Plugin_BlogReBuild_Categorys_Begin=True Then Exit Function
	Next

	GetCategory()

	Dim objRS
	Dim objStream
	Dim objArticle
	Dim i,j

	j=Functions(FunctionMetas.GetValue("previous")).MaxLi
	If j=0 Then j=10


	'Categorys
	Dim strCategory

	Dim Category
	For Each Category in Categorys

		If IsObject(Category) Then

			Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_IsAnonymous],[log_Meta] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_ID]>0) AND ([log_Level]>1) AND ([log_CateID]="&Category.ID&") ORDER BY [log_PostTime] DESC")

			If (Not objRS.bof) And (Not objRS.eof) Then
				For i=1 to j
					Set objArticle=New TArticle
					If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then
						strCategory=strCategory & "<li><a href="""& objArticle.Url & """>" & objArticle.Title & "</a></li>"
					End If
					Set objArticle=Nothing
					objRS.MoveNext
					If objRS.eof Then Exit For
				Next
			End If
			objRS.close

			strCategory=TransferHTML(strCategory,"[no-asp]")

			Call SaveToFile(BlogPath & "zb_users/include/category_"&Category.ID&".asp",strCategory,"utf-8",True)

			strCategory=""

		End If
	Next

	BlogReBuild_Categorys=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_Authors()

	'plugin node
	bAction_Plugin_BlogReBuild_Authors_Begin=False
	For Each sAction_Plugin_BlogReBuild_Authors_Begin in Action_Plugin_BlogReBuild_Authors_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Authors_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Authors_Begin)
		If bAction_Plugin_BlogReBuild_Authors_Begin=True Then Exit Function
	Next

	GetUser

	Dim objRS
	Dim objStream

	'Authors
	Dim strAuthor
	Dim User
	For Each User in Users
		If IsObject(User) Then''''''
			If User.ID>0 Then
				strAuthor=strAuthor & "<li><a href="""& User.Url & """>"+User.Name + " (" & User.Count & ")" +"</a></li>"
			End If
		End If
	Next

	strAuthor=TransferHTML(strAuthor,"[no-asp]")

	Call GetFunction()
	Functions(FunctionMetas.GetValue("authors")).Content=strAuthor

	'Call SaveToFile(BlogPath & "zb_users/include/authors.asp",strAuthor,"utf-8",True)

	BlogReBuild_Authors=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_Tags()

	'plugin node
	bAction_Plugin_BlogReBuild_Tags_Begin=False
	For Each sAction_Plugin_BlogReBuild_Tags_Begin in Action_Plugin_BlogReBuild_Tags_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Tags_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Tags_Begin)
		If bAction_Plugin_BlogReBuild_Tags_Begin=True Then Exit Function
	Next

	Call GetFunction()

	Dim objRS
	Dim objStream

	Dim i,j
	i=Functions(FunctionMetas.GetValue("tags")).MaxLi
	If i=0 Then i=20
	j=0
	'Authors
	Dim strTag

	Set objRS=objConn.Execute("SELECT * FROM [blog_Tag] ORDER BY [tag_Count] DESC,[tag_ID] ASC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		Do While Not objRS.eof
			strTag=strTag & "<li><a href=""<#ZC_BLOG_HOST#>catalog.asp?"& "tags=" & Server.URLEncode(objRS("tag_Name")) & """>"+objRS("tag_Name") + " <span class=""tag-count"">(" & objRS("tag_Count") & ")</span>" +"</a></li>"
			objRS.MoveNext
			j=j+1
			If j>i Then Exit Do
		Loop
	End If
	objRS.Close
	Set objRS=Nothing

	strTag=TransferHTML(strTag,"[no-asp]")


	Functions(FunctionMetas.GetValue("tags")).Content=strTag

	'Call SaveToFile(BlogPath & "zb_users/include/tags.asp",strTag,"utf-8",True)

	BlogReBuild_Tags=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_Previous()

	'plugin node
	bAction_Plugin_BlogReBuild_Previous_Begin=False
	For Each sAction_Plugin_BlogReBuild_Previous_Begin in Action_Plugin_BlogReBuild_Previous_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Previous_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Previous_Begin)
		If bAction_Plugin_BlogReBuild_Previous_Begin=True Then Exit Function
	Next

	Call GetFunction()

	Dim i,j
	Dim objRS
	Dim objStream
	Dim objArticle

	j=Functions(FunctionMetas.GetValue("previous")).MaxLi

	If j=0 Then j=10

	'Previous
	Dim strPrevious
	Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_IsAnonymous],[log_Meta] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_ID]>0) AND ([log_Level]>1) ORDER BY [log_PostTime] DESC")

	If (Not objRS.bof) And (Not objRS.eof) Then
		For i=1 to j
			Set objArticle=New TArticle
			If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then
				strPrevious=strPrevious & "<li><a href="""& objArticle.Url & """ title="""& objArticle.HtmlTitle &"""><span class=""article-date"">["& Right("0" & Month(objArticle.PostTime),2) & "/" & Right("0" & Day(objArticle.PostTime),2) &"]</span>" & objArticle.Title & "</a></li>"
			End If
			Set objArticle=Nothing
			objRS.MoveNext
			If objRS.eof Then Exit For
		Next
	End If
	objRS.close

	strPrevious=TransferHTML(strPrevious,"[no-asp]")

	Functions(FunctionMetas.GetValue("previous")).Content=strPrevious

	Functions(FunctionMetas.GetValue("previous")).SaveFile

	'Call SaveToFile(BlogPath & "zb_users/include/previous.asp",strPrevious,"utf-8",True)

	BlogReBuild_Previous=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_Comments()

	'plugin node
	bAction_Plugin_BlogReBuild_Comments_Begin=False
	For Each sAction_Plugin_BlogReBuild_Comments_Begin in Action_Plugin_BlogReBuild_Comments_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Comments_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Comments_Begin)
		If bAction_Plugin_BlogReBuild_Comments_Begin=True Then Exit Function
	Next

	Call GetFunction()

	Dim objRS
	Dim objStream
	Dim objArticle

	'Comments
	Dim strComments

	Dim s
	Dim i,j

	j=Functions(FunctionMetas.GetValue("comments")).MaxLi
	If j=0 Then j=10

	Set objRS=objConn.Execute("SELECT [log_ID],[comm_ID],[comm_Content],[comm_PostTime],[comm_Author] FROM [blog_Comment] WHERE [log_ID]>0 ORDER BY [comm_PostTime] DESC,[comm_ID] DESC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		For i=1 to j
			s=objRS("comm_Content")
			s=Replace(s,vbCrlf,"")
			'Set objArticle=New TArticle
			'If objArticle.LoadInfoByID(objRS("log_ID")) Then
				strComments=strComments & "<li style=""text-overflow:ellipsis;""><a href="""& GetCurrentHost & "zb_system/view.asp?nav=" & objRS("log_ID") & "#cmt" & objRS("comm_ID") & """ title=""" & objRS("comm_PostTime") & " post by " & objRS("comm_Author") & """>"+s+"</a></li>"
			'End If
			Set objArticle=Nothing
			objRS.MoveNext
			If objRS.eof Then Exit For
		Next
	End If
	objRS.close
	Set objRS=Nothing

	strComments=TransferHTML(strComments,"[no-asp]")

	Functions(FunctionMetas.GetValue("comments")).Content=strComments

	Functions(FunctionMetas.GetValue("comments")).SaveFile

	'Call SaveToFile(BlogPath & "zb_users/include/comments.asp",strComments,"utf-8",True)

	BlogReBuild_Comments=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_GuestComments()

	'plugin node
	bAction_Plugin_BlogReBuild_GuestComments_Begin=False
	For Each sAction_Plugin_BlogReBuild_GuestComments_Begin in Action_Plugin_BlogReBuild_GuestComments_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_GuestComments_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_GuestComments_Begin)
		If bAction_Plugin_BlogReBuild_GuestComments_Begin=True Then Exit Function
	Next

	BlogReBuild_GuestComments=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_TrackBacks()

	'plugin node
	bAction_Plugin_BlogReBuild_TrackBacks_Begin=False
	For Each sAction_Plugin_BlogReBuild_TrackBacks_Begin in Action_Plugin_BlogReBuild_TrackBacks_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_TrackBacks_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_TrackBacks_Begin)
		If bAction_Plugin_BlogReBuild_TrackBacks_Begin=True Then Exit Function
	Next

	BlogReBuild_TrackBacks=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function BlogReBuild_Statistics()

	'plugin node
	bAction_Plugin_BlogReBuild_Statistics_Begin=False
	For Each sAction_Plugin_BlogReBuild_Statistics_Begin in Action_Plugin_BlogReBuild_Statistics_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Statistics_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Statistics_Begin)
		If bAction_Plugin_BlogReBuild_Statistics_Begin=True Then Exit Function
	Next

	Dim i
	Dim objRS
	Dim objStream

	'重新统计分类及用户的文章数、评论数
	Dim Category,strSubCateID
	For Each Category in Categorys
		If IsObject(Category) Then
			strSubCateID=Join(GetSubCateID(Category.ID,True),",")
			Set objRS=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE [log_Level]>1 AND [log_CateID]IN(" & strSubCateID &")" )
			i=objRS(0)
			objConn.Execute("UPDATE [blog_Category] SET [cate_Count]="&i&" WHERE [cate_ID] =" & Category.ID)
			Set objRS=Nothing
		End If
	Next

	'Statistics
	Dim strStatistics
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""


	objRS.Open("SELECT COUNT([log_ID])AS allArticle,SUM([log_CommNums]) AS allCommNums,SUM([log_ViewNums]) AS allViewNums,SUM([log_TrackBackNums]) AS allTrackBackNums FROM [blog_Article]")
	If (Not objRS.bof) And (Not objRS.eof) Then
		strStatistics=strStatistics & "<li>"& ZC_MSG082 &":" & objRS("allArticle") & "</li>"
		strStatistics=strStatistics & "<li>"& ZC_MSG124 &":" & objRS("allCommNums") & "</li>"
		strStatistics=strStatistics & "<li>"& ZC_MSG125 &":" & objRS("allTrackBackNums") & "</li>"
		strStatistics=strStatistics & "<li>"& ZC_MSG129 &":" & objRS("allViewNums") & "</li>"
	End If
	objRS.Close


	strStatistics=strStatistics & "<li>"& ZC_MSG204 &":" & GetNameFormTheme(ZC_BLOG_THEME) & "</li>"
	strStatistics=strStatistics & "<li>"& ZC_MSG083 &":" & ZC_BLOG_CSS & "</li>"

	Set objRS=Nothing

	strStatistics=TransferHTML(strStatistics,"[no-asp]")

	Call GetFunction()
	Functions(FunctionMetas.GetValue("statistics")).Content=strStatistics

	'Call SaveToFile(BlogPath & "zb_users/include/statistics.asp",strStatistics,"utf-8",False)

	BlogReBuild_Statistics=True

End Function
'*********************************************************




'*********************************************************
' 目的：    BlogReBuild Functions
'*********************************************************
Function BlogReBuild_Functions

	Call GetFunction()

	Call SaveFunctionType()

	Dim i,j,s,t,f

	For Each f In Functions
		If IsObject(f)=True Then
			If f.id>0 Then
				f.SaveFile
			End If
		End If 
	Next


	Dim aryFunctionInOrder
	aryFunctionInOrder=GetFunctionOrder()

	Application.Lock
	t=GetTemplate("TEMPLATE_B_FUNCTION")
	Application.UnLock


	For i=1 To 5
		If IsArray(aryFunctionInOrder) Then
			s=""
			For j=LBound(aryFunctionInOrder)+1 To Ubound(aryFunctionInOrder)
				If Functions(aryFunctionInOrder(j)).InSidebars(i)=True Then
				s=s & Functions(aryFunctionInOrder(j)).MakeTemplate(t)
				End If
			Next
			Call SaveToFile(BlogPath & "zb_users/cache/sidebar"& IIF(i>1,i,"") &".html",s,"utf-8",False)
		End If
	Next

	BlogReBuild_Functions=True

End Function
'*********************************************************




'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    Export RSS 2.0
'*********************************************************
Function ExportRSS()

	'plugin node
	bAction_Plugin_ExportRSS_Begin=False
	For Each sAction_Plugin_ExportRSS_Begin in Action_Plugin_ExportRSS_Begin
		If Not IsEmpty(sAction_Plugin_ExportRSS_Begin) Then Call Execute(sAction_Plugin_ExportRSS_Begin)
		If bAction_Plugin_ExportRSS_Begin=True Then Exit Function
	Next

	Dim Rss2Export
	Dim objArticle

	Set Rss2Export = New TNewRss2Export

	With Rss2Export

		.TimeZone=ZC_TIME_ZONE

		.AddChannelAttribute "title",TransferHTML(ZC_BLOG_TITLE,"[html-format]")
		.AddChannelAttribute "link",TransferHTML(ZC_BLOG_HOST,"[html-format]")
		.AddChannelAttribute "description",TransferHTML(ZC_BLOG_SUBTITLE,"[html-format]")
		.AddChannelAttribute "generator","RainbowSoft Studio Z-Blog " & ZC_BLOG_VERSION
		.AddChannelAttribute "language",ZC_BLOG_LANGUAGE
		'.AddChannelAttribute "copyright",TransferHTML(ZC_BLOG_COPYRIGHT,"[nohtml][html-format]")
		.AddChannelAttribute "pubDate",GetTime(Now())

			Dim i
			Dim objRS
			Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_IsAnonymous],[log_Meta] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_ID]>0) AND ([log_Level]>2) ORDER BY [log_PostTime] DESC")

			If (Not objRS.bof) And (Not objRS.eof) Then
				For i=1 to ZC_RSS2_COUNT
					Set objArticle=New TArticle
					If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then

					If ZC_RSS_EXPORT_WHOLE Then
					.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).Name & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlContent,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
					Else
					.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).Name & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlIntro,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
					End If

					End If
					objRS.MoveNext
					If objRS.eof Then Exit For
					Set objArticle=Nothing
				Next
			End If

	End With

	'Rss2Export.Execute

	Rss2Export.SaveToFile(BlogPath & "/rss.xml")

	Set Rss2Export = Nothing

	objRS.close
	Set objRS=Nothing
	ExportRSS=True

	'Response.ContentType = "text/html"
	'Response.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    Export ATOM 1.0
'*********************************************************
Function ExportATOM()


End Function
'*********************************************************




'*********************************************************
' 目的：    Build Category
'*********************************************************
Function BuildCategory(intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType,strDirectory,strFileName)

End Function
'*********************************************************




'*********************************************************
' 目的：    Build Article
'*********************************************************
Function BuildArticle(intID,bolBuildNavigate,bolBuildCategory)

	Dim objArticle
	Set objArticle=New TArticle

	Call GetCategory()
	Call GetUser()

	If objArticle.LoadInfoByID(intID) Then
		Call GetTagsbyTagIDList(objArticle.Tag)
		objArticle.Statistic
		If objArticle.Export(ZC_DISPLAY_MODE_ALL) Then
			objArticle.SaveCache
			objArticle.Build
			objArticle.Save

		End If

		If (bolBuildNavigate=True) And (ZC_USE_NAVIGATE_ARTICLE=True) Then

			Dim objRS
			Set objRS=objConn.Execute("SELECT TOP 1 [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND ([log_CateID]<>0) AND ([log_PostTime]<" & ZC_SQL_POUND_KEY & objArticle.PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] DESC")
			If (Not objRS.bof) And (Not objRS.eof) Then
				Call BuildArticle(objRS("log_ID"),False,False)
			End If
			Set objRS=Nothing
			Set objRS=objConn.Execute("SELECT TOP 1 [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND ([log_CateID]<>0) AND ([log_PostTime]>" & ZC_SQL_POUND_KEY & objArticle.PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] ASC")
			If (Not objRS.bof) And (Not objRS.eof) Then
				Call BuildArticle(objRS("log_ID"),False,False)
			End If
			Set objRS=Nothing

		End If

		BuildArticle=True

	End If

	Set objArticle=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    GetTagsbyTagIDList
'*********************************************************
Function GetTagsbyTagIDList(strTags)
'strTags={1}{2}{3}{4}

strTags=Trim(FilterSQL(strTags))

If strTags="" Then Exit Function
If strTags="{}" Then Exit Function


Dim s,t,i
strTags=Replace(strTags,"}","")
t=Split(strTags,"{")

For i=LBound(t) To UBound(t)
	If Trim(t(i))<>"" Then
		s=s & "([tag_ID]="&t(i)&") Or"
	End If
Next

s=Left(s,Len(s)-2)


Dim objRS
Dim objTag

Set objRS=objConn.Execute("SELECT [tag_ID],[tag_Name],[tag_Intro],[tag_Order],[tag_Count],[tag_ParentID],[tag_URL],[tag_Template],[tag_FullUrl],[tag_Meta] FROM [blog_Tag] WHERE (" & s & ")")

If (Not objRS.bof) And (Not objRS.eof) Then


	Do While Not objRS.eof

		Set objTag=New TTag
		Call objTag.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9)))

		If UBound(Tags)<objTag.ID Then
			ReDim Preserve Tags(objTag.ID)
		End If

		Set Tags(objTag.ID)=objTag

		Set objTag=Nothing

		objRS.MoveNext
	Loop

End If

objRS.Close
Set objRS=Nothing

GetTagsbyTagIDList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    GetTagsbyTagNameList
'*********************************************************
Function GetTagsbyTagNameList(strTags)
'strTags=a,b,c,d,e,f,g

Set Tags(0)=New TTag

strTags=Trim(FilterSQL(strTags))

If strTags="" Then Exit Function

Dim s,t,i
t=Split(strTags,",")

For i=LBound(t) To UBound(t)
	If Trim(t(i))<>"" Then
		s=s & "([tag_Name]='"&t(i)&"') Or"
	End If
Next

s=Left(s,Len(s)-2)

Dim objRS
Dim objTag

Set objRS=objConn.Execute("SELECT [tag_ID],[tag_Name],[tag_Intro],[tag_Order],[tag_Count],[tag_ParentID],[tag_URL],[tag_Template],[tag_FullUrl],[tag_Meta] FROM [blog_Tag] WHERE (" & s & ")")

If (Not objRS.bof) And (Not objRS.eof) Then


	Do While Not objRS.eof

		Set objTag=New TTag
		Call objTag.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9)))

		If UBound(Tags)<objTag.ID Then
			ReDim Preserve Tags(objTag.ID)
		End If

		Set Tags(objTag.ID)=objTag

		Set objTag=Nothing

		objRS.MoveNext
	Loop

End If

objRS.Close
Set objRS=Nothing

GetTagsbyTagNameList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function GetCommentFloor(ID)

	Dim i,j
	i=ID
	j=0

	Dim objRS

	Do While i>0

		j=j+1

		Set objRS=objConn.Execute("SELECT [comm_ParentID] FROM [blog_Comment] WHERE [comm_ID] =" & i)

		If (Not objRS.bof) And (Not objRS.eof) Then
			i=objRS(0)
		Else
			i=0
		End If

	Loop

	GetCommentFloor=j


End Function
'*********************************************************




'*********************************************************
' 目的：    Get Function Order 输出数组.
'*********************************************************
Function GetFunctionOrder()

	Dim i
	Dim objRS

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	Dim aryCateInOrder()
	i=0

	objRS.Open("SELECT * FROM [blog_Function] ORDER BY [fn_Order] ASC,[fn_ID] ASC")
	Do While Not objRS.eof
		i=i+1
		ReDim Preserve aryCateInOrder(i)
		aryCateInOrder(i)=objRS("fn_ID")
		objRS.MoveNext
	Loop
	objRS.Close
	Set objRS=Nothing

	If i>0 Then GetFunctionOrder=aryCateInOrder
	If i=0 Then GetFunctionOrder=Array()

	Erase aryCateInOrder

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function SaveFunctionType()

	Call GetFunction()

	Dim t 

	Set t=New TMeta

	Dim i

	For i=1 To UBound(Functions)
		If IsObject(Functions(i))=True Then

			Call t.SetValue(Functions(i).FileName,Functions(i).FType)

		End If
	Next

	Call SaveToFile(BlogPath & "zb_users/CACHE/functionstype.html",t.SaveString,"utf-8",False)

	SaveFunctionType=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Create Admin Menu
'*********************************************************
Function CreateAdminLeftMenu()

'强制清空Menu,防止某些插件提前插入造成排在系统菜单之前,插件插入菜单要在系统初始化完成后
Response_Plugin_Admin_Left=""

Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("ArticleEdt"),ZC_MSG168,GetCurrentHost&"zb_system/cmd.asp?act=ArticleEdt&amp;webedit="&ZC_BLOG_WEBEDIT,"nav_new","aArticleEdt",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("FileReBuild"),ZC_MSG073,GetCurrentHost&"zb_system/cmd.asp?act=AskFileReBuild","nav_build","aAskFileReBuild",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("ArticleMng"),ZC_MSG067,GetCurrentHost&"zb_system/cmd.asp?act=ArticleMng","nav_article","aArticleMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("ArticleMng"),ZC_MSG111,GetCurrentHost&"zb_system/cmd.asp?act=ArticleMng&amp;type=Page","nav_page","aPageMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("CategoryMng"),ZC_MSG066,GetCurrentHost&"zb_system/cmd.asp?act=CategoryMng","nav_category","aCategoryMng",""))

Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("TagMng"),ZC_MSG141,GetCurrentHost&"zb_system/cmd.asp?act=TagMng","nav_tags","aTagMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("CommentMng"),ZC_MSG068,GetCurrentHost&"zb_system/cmd.asp?act=CommentMng","nav_comments","aCommentMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("FileMng"),ZC_MSG071,GetCurrentHost&"zb_system/cmd.asp?act=FileMng","nav_accessories","aFileMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("ThemeMng"),ZC_MSG223,GetCurrentHost&"zb_system/cmd.asp?act=ThemeMng","nav_themes","aThemeMng",""))


Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("FunctionMng"),ZC_MSG007,GetCurrentHost&"zb_system/cmd.asp?act=FunctionMng","nav_function","aFunctionMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("PlugInMng"),ZC_MSG107,GetCurrentHost&"zb_system/cmd.asp?act=PlugInMng","nav_plugin","aPlugInMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("UserMng"),ZC_MSG070,GetCurrentHost&"zb_system/cmd.asp?act=UserMng","nav_user","aUserMng",""))


End Function
'*********************************************************




'*********************************************************
' 目的：    Create Top Menu
'*********************************************************
Function CreateAdminTopMenu()

Response_Plugin_Admin_Top=""

Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(ZC_MSG245,GetCurrentHost&"zb_system/cmd.asp?act=admin",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(ZC_MSG247,GetCurrentHost&"zb_system/cmd.asp?act=SettingMng",""))

End Function
'*********************************************************




'*********************************************************
' 目的：    Search Child Comments'递归
'*********************************************************
Function SearchChildComments(ByVal id,ByRef allcomm)

	If IsObject(allcomm)=False Then
		Set allcomm=CreateObject("Scripting.Dictionary")
	End If

	Dim objRS
	Set objRS=objConn.Execute("SELECT [comm_ID] FROM [blog_Comment] WHERE [comm_ParentID]="&id)
	If (Not objRS.bof) And (Not objRS.eof) Then
		Do While Not objRS.eof
			If allcomm.Exists(objRS("comm_ID"))=False Then
				allcomm.add CInt(objRS("comm_ID")),""
			End If
			Call SearchChildComments(objRS("comm_ID"),allcomm)
			objRS.MoveNext
		Loop
		SearchChildComments=True
	Else
		SearchChildComments=False
	End If

End Function
'*********************************************************




'*********************************************************
' 目的：    Add Batch 
'*********************************************************
Function AddBatch(name,actioncode)


	Dim i
	i=CInt(Session("batch_order"))+1
	Session("batch_order")=i
	Call Session("batch").add("<b>" & i & "</b> : <u>" & name & "</u>",actioncode)

End Function
'*********************************************************




'*********************************************************
' 目的：    Save Config to option.asp
'*********************************************************
Function SaveConfig2Option()

	Dim strContent
	strContent=LoadFromFile(BlogPath & "zb_system\defend\c_option.asp.html","utf-8")


	Dim i

	For i=1 To BlogConfig.Count
		strContent=Replace(strContent,"<#"&BlogConfig.Meta.Names(i)&"#>",Replace(BlogConfig.Meta.GetValue(BlogConfig.Meta.Names(i)),"""",""""""))
	Next

	Call BlogConfig.Save()

	Call SaveToFile(BlogPath & "zb_users\c_option.asp",strContent,"utf-8",False)

End Function
'*********************************************************
%>
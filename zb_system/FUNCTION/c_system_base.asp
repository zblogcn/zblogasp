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


Dim BlogHost
BlogHost=GetCurrentHost()
ZC_BLOG_HOST=BlogHost

Dim BlogPath
BlogPath=GetReallyDirectory()

'补上c_option.asp未更新的参数
Call CheckUndefined()


Dim BlogVersion
BlogVersion=GetBlogVersion()

Dim ZC_BLOG_CLSID_ORIGINAL
ZC_BLOG_CLSID_ORIGINAL=ZC_BLOG_CLSID
ZC_BLOG_CLSID=MD5(BlogPath & ZC_BLOG_CLSID_ORIGINAL)

Dim StarTime
Dim EndTime
StarTime = Timer()

Dim BlogUser
Set BlogUser =New TUser

Dim Categorys()
Dim Users()
Dim Tags()
Dim Functions()

ReDim Categorys(0)
ReDim Users(0)
ReDim Tags(0)
ReDim Functions(0)

Set Categorys(0)=New TCategory
Set Users(0)=New TUser
Set Tags(0)=New TTag

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

Dim ZC_BLOG_WEBEDIT
ZC_BLOG_WEBEDIT="ueditor"

Dim ZC_TB_EXCERPT_MAX
ZC_TB_EXCERPT_MAX=250

Dim ZC_TRACKBACK_TURNOFF
ZC_TRACKBACK_TURNOFF=True

Const ZC_DISPLAY_MODE_ALL=1
Const ZC_DISPLAY_MODE_INTRO=2
Const ZC_DISPLAY_MODE_ONTOP=3
Const ZC_DISPLAY_MODE_SEARCH=4
Const ZC_DISPLAY_MODE_SYSTEMPAGE=5
Const ZC_DISPLAY_MODE_COMMENTS=6


Const ZC_POST_TYPE_ARTICLE=0
Const ZC_POST_TYPE_PAGE=1


Const ZC_DEFAULT_SIDEBAR="calendar:controlpanel:catalog:searchpanel:comments:archives:favorite:link:misc"

Dim ZC_NAVBAR_MENU_ITEM
ZC_NAVBAR_MENU_ITEM="<li id=""menu-%type-%id""><a href=""%url"">%name</a></li>"


Dim ZC_PAGE_AND_ARTICLE_PRIVATE_REGEX
ZC_PAGE_AND_ARTICLE_PRIVATE_REGEX="{%host%}/view.asp?id={%id%}"


Dim ZC_PAGE_AND_ARTICLE_DRAFT_REGEX
ZC_PAGE_AND_ARTICLE_DRAFT_REGEX="{%host%}/view.asp?id={%id%}"

'如果连接数据库为MSSQL，则应为'，默认连接Access数据库则为#
Dim ZC_SQL_POUND_KEY
ZC_SQL_POUND_KEY="#"

Dim ZC_COMMENT_VERIFY_ENABLE_INTERNAL
ZC_COMMENT_VERIFY_ENABLE_INTERNAL=True

Dim ZC_BLOG_PRODUCT
ZC_BLOG_PRODUCT="Z-Blog"

Dim ZC_BLOG_PRODUCT_FULL
ZC_BLOG_PRODUCT_FULL=""

Dim ZC_BLOG_PRODUCT_FULLHTML
ZC_BLOG_PRODUCT_FULLHTML=""

Dim BlogVersions
Set BlogVersions = New TMeta
BlogVersions.SetValue "180518","Z-Blog 2.3 Avengers Build 180518"
BlogVersions.SetValue "140101","Z-Blog 2.2 Prism Build 140101"
BlogVersions.SetValue "130801","Z-Blog 2.2 Prism Build 130801"
BlogVersions.SetValue "130722","Z-Blog 2.2 Prism Build 130722"
BlogVersions.SetValue "130128","Z-Blog 2.1 Phoenix Build 130128"
BlogVersions.SetValue "121221","Z-Blog 2.0 Doomsday Build 121221"
BlogVersions.SetValue "121028","Z-Blog 2.0 Beta2 Build 121028"
BlogVersions.SetValue "121001","Z-Blog 2.0 Beta Build 121001"

ZC_BLOG_VERSION=Replace(BlogVersions.GetValue(BlogVersions.Names(1)),"Z-Blog ","")

Response.AddHeader "Product","Z-Blog " & ZC_BLOG_VERSION


'定义了一个布尔信号，为了解决Server 2016下的Execute问题，真是日了狗了
Dim Boolean_Delay_Plugin_Signal
Boolean_Delay_Plugin_Signal = False

'*********************************************************
' 目的：    System 初始化
'*********************************************************
Sub System_Initialize()

	'Call ActivePlugin()
	If OpenConnect()=False Then
		Call ShowError(4)
	End If

	Call GetConfigs()
	BlogConfig.Load("Blog")

	BlogUser.Verify()

	Call GetCategory()
	'Call GetUser()
	'Call GetTags()
	'Call GetKeyWords()
	'Call GetFunction()

	ZC_BLOG_PRODUCT_FULL=ZC_BLOG_PRODUCT & " " & ZC_BLOG_VERSION

	ZC_BLOG_PRODUCT_FULLHTML="<a href=""http://www.zblogcn.com/"" title=""RainbowSoft Z-Blog"">" & ZC_BLOG_PRODUCT_FULL & "</a>"

	TemplateTagsDic.Item("ZC_BLOG_HOST")=BlogHost

	Boolean_Delay_Plugin_Signal = True

	Call LoadGlobeCache()
	Call CreateAdminLeftMenu()
	Call CreateAdminTopMenu()
	Call ActivePlugin()

	Boolean_Delay_Plugin_Signal = False

	Call Execute_Action_Plugin()
	Call Execute_Filter_Plugin()
	Call Execute_Response_Plugin()

	'plugin node
	bAction_Plugin_System_Initialize=False
	For Each sAction_Plugin_System_Initialize in Action_Plugin_System_Initialize
		If Not IsEmpty(sAction_Plugin_System_Initialize) Then Call Execute(sAction_Plugin_System_Initialize)
		If bAction_Plugin_System_Initialize=True Then Exit Sub
	Next

	If ZC_POST_STATIC_MODE<>"STATIC" Then
		Dim bolRebuildFiles
		Application.Lock
		bolRebuildFiles=Application(ZC_BLOG_CLSID & "SIGNAL_REBUILDFILES")
		Application.UnLock
		If IsEmpty(bolRebuildFiles)=False Then
			If bolRebuildFiles=True Then
				Call SetBlogHint(True,True,False)
			End If
		End If
	End If



	Dim bolRebuildIndex
	Application.Lock
	bolRebuildIndex=Application(ZC_BLOG_CLSID & "SIGNAL_REBUILDINDEX")
	Application.UnLock
	If IsEmpty(bolRebuildIndex)=False Then
		If bolRebuildIndex=True Then
			Call MakeBlogReBuild_Core()
		End If
	End If


	'将激活插件后移
	If sFilter_Plugin_ValidCode_Check="" Then Call Add_Filter_Plugin("Filter_Plugin_ValidCode_Check","CheckVerifyNumber")


	'plugin node
	bAction_Plugin_System_Initialize_Succeed=False
	For Each sAction_Plugin_System_Initialize_Succeed in Action_Plugin_System_Initialize_Succeed
		If Not IsEmpty(sAction_Plugin_System_Initialize_Succeed) Then Call Execute(sAction_Plugin_System_Initialize_Succeed)
		If bAction_Plugin_System_Initialize_Succeed=True Then Exit Sub
	Next

	'If Err.Number<>0 Then Call ShowError(10)

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

	Set PublicObjAdo=Nothing
	Set PublicObjFSO=Nothing
	Call CloseConnect()

End Sub
'*********************************************************




'*********************************************************
' 目的：    数据库连接
'*********************************************************
Dim IsDBConnect '数据库是否已连接
IsDBConnect=False
Function OpenConnect()

	On Error Resume Next

	If IsDBConnect=True Then
		OpenConnect=True
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

	If Err.Number=0 Then

		IsDBConnect=True
		OpenConnect=True

	Else

		Err.Clear
		Err.Raise 1

	End If

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

	Set objRS=objConn.Execute("SELECT [cate_ID],[cate_Name],[cate_Intro],[cate_Order],[cate_Count],[cate_ParentID],[cate_URL],[cate_Template],[cate_LogTemplate],[cate_FullUrl],[cate_Meta] FROM [blog_Category] ORDER BY [cate_ID] ASC")
	If (Not objRS.bof) And (Not objRS.eof) Then

		aryAllData=objRS.GetRows(objRS.RecordCount)
		objRS.Close
		Set objRS=Nothing

		k=UBound(aryAllData,1)
		l=UBound(aryAllData,2)
		For i=0 To l
			Set Categorys(aryAllData(0,i))=New TCategory
			Categorys(aryAllData(0,i)).LoadInfoByArray(Array(aryAllData(0,i),aryAllData(1,i),aryAllData(2,i),aryAllData(3,i),aryAllData(4,i),aryAllData(5,i),aryAllData(6,i),aryAllData(7,i),aryAllData(8,i),aryAllData(9,i),aryAllData(10,i)))
		Next
	End If

	Set Categorys(0)=New TCategory
	Call Categorys(0).LoadInfoByID(0)

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


	Set objRS=objConn.Execute("SELECT [mem_ID],[mem_Name],[mem_Level],[mem_Password],[mem_Email],[mem_HomePage],[mem_PostLogs],[mem_Url],[mem_Template],[mem_FullUrl],[mem_Intro],[mem_Meta] FROM [blog_Member] ORDER BY [mem_ID] ASC")
	If (Not objRS.bof) And (Not objRS.eof) Then

		aryAllData=objRS.GetRows(objRS.RecordCount)
		objRS.Close
		Set objRS=Nothing

		k=UBound(aryAllData,1)
		l=UBound(aryAllData,2)
		For i=0 To l
			Set Users(aryAllData(0,i))=New TUser
			Users(aryAllData(0,i)).LoadInfoByArray(Array(aryAllData(0,i),aryAllData(1,i),aryAllData(2,i),aryAllData(3,i),aryAllData(4,i),aryAllData(5,i),aryAllData(6,i),aryAllData(7,i),aryAllData(8,i),aryAllData(9,i),aryAllData(10,i),aryAllData(11,i)))
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
' 目的：    Configs读取
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

	Erase Functions
	ReDim Functions(0)

	Set FunctionMetas=New TMeta

	Dim objRS

	Set objRS=objConn.Execute("SELECT TOP 1 [fn_ID] FROM [blog_Function] ORDER BY [fn_ID] DESC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		i=objRS("fn_ID")
		ReDim Functions(i)
	End If

	Set objRS=objConn.Execute("SELECT [fn_ID],[fn_Name],[fn_FileName],[fn_Order],[fn_Content],[fn_IsHidden],[fn_SidebarID],[fn_HtmlID],[fn_Ftype],[fn_MaxLi],[fn_Source],[fn_ViewType],[fn_IsHideTitle],[fn_Meta] FROM [blog_Function] ORDER BY [fn_ID] ASC")
	If (Not objRS.bof) And (Not objRS.eof) Then


		aryAllData=objRS.GetRows(objRS.RecordCount)
		objRS.Close
		Set objRS=Nothing

		k=UBound(aryAllData,1)
		l=UBound(aryAllData,2)
		For i=0 To l
			Set Functions(aryAllData(0,i))=New TFunction
			Functions(aryAllData(0,i)).LoadInfoByArray(Array(aryAllData(0,i),aryAllData(1,i),aryAllData(2,i),aryAllData(3,i),aryAllData(4,i),aryAllData(5,i),aryAllData(6,i),aryAllData(7,i),aryAllData(8,i),aryAllData(9,i),aryAllData(10,i),aryAllData(11,i),aryAllData(12,i),aryAllData(13,i)))
			Call FunctionMetas.SetValue(aryAllData(2,i),aryAllData(0,i))
		Next

	End If


	Dim aryFileList

	aryFileList=LoadIncludeFilesOnlyType("zb_users\INCLUDE\")

	If IsArray(aryFileList) Then
		j=UBound(aryFileList)

		If j>0 Then
			For i=1 to j
				If Right(aryFileList(i),4)=".asp" Then
				l=Replace(aryFileList(i),".asp","")
				If FunctionMetas.Exists(l) =False Then

					k=UBound(Functions)+1

					ReDim Preserve Functions(k)

					Set Functions(k)=New TFunction

					Functions(k).ID=k
					Functions(k).FileName=l
					Functions(k).Name=l
					Functions(k).Source="other"
					Functions(k).HtmlID="fn" & l
					Functions(k).Content=LoadFromFile(BlogPath &"zb_users\INCLUDE\"&aryFileList(i),"utf-8")
					If Instr(Functions(k).Content,"</li>")>0 Then
						Functions(k).FType="ul"
					Else
						Functions(k).FType="div"
					End If
					Call FunctionMetas.SetValue(l,k)
				End If
				End If
			Next
		End If
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
		Case "ArticleAll"
			GetRights=2
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
		Case "CommentAll"
			GetRights=2
		Case "CommentMng"
			GetRights=4
		Case "CommentDel"
			GetRights=4
		Case "CommentEdt"
			GetRights=4
		Case "CommentSav"
			GetRights=4
		Case "CommentGet"
			GetRights=5
		Case "CommentAudit"
			GetRights=1
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
		Case "UserMod"
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
		Case "FileAll"
			GetRights=2
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

	CheckAuthorByID=Not objConn.Execute("SELECT [mem_ID] FROM [blog_Member] WHERE [mem_ID]=" & CLng(intAuthorId) ).BOF

End Function

Function GetAuthorByName(strName)
	Dim objRS
	Set objRS=objConn.Execute("SELECT [mem_ID] FROM [blog_Member] WHERE [mem_Name]='"&FilterSQL(strName)&"'" )
	If (Not objRS.bof) And (Not objRS.eof) Then
		GetAuthorByName=objRS(0)
	Else
		GetAuthorByName=0
	End If
End Function

Function GetAuthorByAlias(strAlias)
	Dim objRS
	Set objRS=objConn.Execute("SELECT [mem_ID] FROM [blog_Member] WHERE [mem_Url]='"&FilterSQL(strAlias)&"'" )
	If (Not objRS.bof) And (Not objRS.eof) Then
		GetAuthorByAlias=objRS(0)
	Else
		GetAuthorByAlias=0
	End If
End Function
'*********************************************************




'*********************************************************
' 目的：    检查分类是否存在
'*********************************************************
Function CheckCateByID(intCateId)

	If CLng(intCateId)=0 Then CheckCateByID=True:Exit Function

	CheckCateByID=Not objConn.Execute("SELECT [cate_ID] FROM [blog_Category] WHERE [cate_ID]=" & CLng(intCateId) ).BOF

End Function
'*********************************************************
' 目的：    根据分类名得到分类ID
'*********************************************************
Function GetCateByName(strName)

	If strName=BlogConfig.Read("ZC_UNCATEGORIZED_NAME") Then
		GetCateByName=0
		Exit Function
	End If
	Dim objRS
	Set objRS=objConn.Execute("SELECT [cate_ID] FROM [blog_Category] WHERE [cate_Name]='"&FilterSQL(strName)&"'" )
	If (Not objRS.bof) And (Not objRS.eof) Then
		GetCateByName=objRS(0)
	Else
		GetCateByName=-1
	End If
End Function

'*********************************************************
' 目的：    根据分类别名得到分类ID
'*********************************************************
Function GetCateByAlias(strAlias)

	If strAlias=BlogConfig.Read("ZC_UNCATEGORIZED_ALIAS") Then
		GetCateByAlias=0
		Exit Function
	End If
	Dim objRS
	Set objRS=objConn.Execute("SELECT [cate_ID] FROM [blog_Category] WHERE [cate_Url]='"&FilterSQL(strAlias)&"'" )
	If (Not objRS.bof) And (Not objRS.eof) Then
		GetCateByAlias=objRS(0)
	Else
		GetCateByAlias=-1
	End If
End Function
'*********************************************************




'*********************************************************
' 目的：    检查TAG是否存在
'*********************************************************
Function CheckTagByID(intTagID)

	If intTagID="" Then intTagID=0
	CheckTagByID=Not objConn.Execute("SELECT [tag_ID] FROM [blog_Tag] WHERE [tag_ID]=" & intTagID ).BOF

End Function

'*********************************************************
' 目的：    检查TAG是否存在
'*********************************************************
Function CheckTagByName(strName)

	CheckTagByName=Not objConn.Execute("SELECT [tag_ID] FROM [blog_Tag] WHERE [tag_Name]='" & FilterSQL(strName) &"'" ).BOF

End Function

'*********************************************************
' 目的：    检查TAG是否存在
'*********************************************************
Function CheckTagByIntro(strName)

	Dim strSQL
	If ZC_MSSQL_ENABLE Then
		strSQL="SELECT [tag_ID] FROM [blog_Tag] WHERE CONVERT(NVARCHAR(255),[tag_Intro])='" & FilterSQL(strName) &"'"
	Else
		strSQL="SELECT [tag_ID] FROM [blog_Tag] WHERE [tag_Intro]='" & FilterSQL(strName) &"'"
	End If
	CheckTagByIntro=Not objConn.Execute(strSQL).BOF

End Function

'*********************************************************
' 目的：   根据TAG别名得到TAG ID
'*********************************************************
Function GetTagByIntro(strName)
	Dim objRS,strSQL
	If ZC_MSSQL_ENABLE Then
		strSQL="SELECT [tag_ID] FROM [blog_Tag] WHERE CONVERT(NVARCHAR(255),[tag_Intro])='" & FilterSQL(strName) &"'"
	Else
		strSQL="SELECT [tag_ID] FROM [blog_Tag] WHERE [tag_Intro]='" & FilterSQL(strName) &"'"
	End If
	Set objRS=objConn.Execute(strSQL)
	If (Not objRS.bof) And (Not objRS.eof) Then
		GetTagByIntro=objRS(0)
	Else
		GetTagByIntro=0
	End If
End Function
'*********************************************************

'*********************************************************
' 目的：   根据TAG名得到TAG ID
'*********************************************************
Function GetTagByName(strName)
	Dim objRS
	Set objRS=objConn.Execute("SELECT [tag_ID] FROM [blog_Tag] WHERE [tag_Name]='"&FilterSQL(strName)&"'" )
	If (Not objRS.bof) And (Not objRS.eof) Then
		GetTagByName=objRS(0)
	Else
		GetTagByName=0
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
	ReDim Preserve aryCateInOrder(0)
	aryCateInOrder(0)=0

	objRS.Open("SELECT [cate_id] FROM [blog_Category] ORDER BY [cate_Order] ASC,[cate_ID] ASC")
	Do While Not objRS.eof
		i=i+1
		ReDim Preserve aryCateInOrder(i)
		aryCateInOrder(i)=objRS("cate_ID")
		objRS.MoveNext
	Loop
	objRS.Close
	Set objRS=Nothing

	GetCategoryOrder=aryCateInOrder

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
Function MakeCalendar(dtmYearMonth)

	Dim strCalendar

	Dim y
	Dim m
	Dim d
	Dim firw
	Dim lasw
	Dim ny
	Dim nm
	Dim py
	Dim pm

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
	py=y
	pm=m-1
	if m=1 then py=py-1:pm=12

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
	If ZC_MSSQL_ENABLE Then
		objRS.Open("select [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] from [blog_Article] where ([log_Type]=0) And ([log_Level]>2) And ([log_PostTime] BETWEEN "& ZC_SQL_POUND_KEY &y&"-"&m&"-1"& ZC_SQL_POUND_KEY &" AND "& ZC_SQL_POUND_KEY &ny&"-"&nm&"-1"& ZC_SQL_POUND_KEY &") AND log_ID in(SELECT max(log_ID) from [blog_Article] group by Month(log_posttime),Day(log_posttime))")
	Else
		objRS.Open("select [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] from [blog_Article] where ([log_Type]=0) And ([log_Level]>2) And ([log_PostTime] BETWEEN "& ZC_SQL_POUND_KEY &y&"-"&m&"-1"& ZC_SQL_POUND_KEY &" AND "& ZC_SQL_POUND_KEY &ny&"-"&nm&"-1"& ZC_SQL_POUND_KEY &") ")
	End If

		If (Not objRS.bof) And (Not objRS.eof) Then
			For i=1 To objRS.RecordCount
				j=CLng(Day(CDate(objRS("log_PostTime"))))
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

	s=UrlbyDateAuto(y,m-1,"")
	t=IIf(m=month(Date()),UrlbyDateAuto(y,m,""),UrlbyDateAuto(y,m+1,""))
	If m=1 Then s=UrlbyDateAuto(y-1,12,"")
	If m=12 Then t=UrlbyDateAuto(y+1,1,"")

		strCalendar="<table id=""tbCalendar""><caption><a href="""&s&""">&#00171;</a>  <a href="""&UrlbyDateAuto(y,m,"")&""">"&Replace(ZC_MSG233,"%y",y)& " " & ZVA_Month_Abbr(m)&"</a>  <a href="""&t&""">&#00187;</a></caption>"

	'thead
		strCalendar=strCalendar & "	<thead>	<tr> <th title="""&ZVA_Week(1)&""" scope=""col"" abbr="""&ZVA_Week(1)&"""><small>"&ZVA_Week_Abbr(1)&"</small></th> <th title="""&ZVA_Week(2)&""" scope=""col"" abbr="""&ZVA_Week(2)&"""><small>"&ZVA_Week_Abbr(2)&"</small></th> <th title="""&ZVA_Week(3)&""" scope=""col"" abbr="""&ZVA_Week(3)&"""><small>"&ZVA_Week_Abbr(3)&"</small></th>	<th title="""&ZVA_Week(4)&""" scope=""col"" abbr="""&ZVA_Week(4)&"""><small>"&ZVA_Week_Abbr(4)&"</small></th> <th title="""&ZVA_Week(5)&""" scope=""col"" abbr="""&ZVA_Week(5)&"""><small>"&ZVA_Week_Abbr(5)&"</small></th>	<th title="""&ZVA_Week(6)&""" scope=""col"" abbr="""&ZVA_Week(6)&"""><small>"&ZVA_Week_Abbr(6)&"</small></th> <th title="""&ZVA_Week(7)&""" scope=""col"" abbr="""&ZVA_Week(7)&"""><small>"&ZVA_Week_Abbr(7)&"</small></th>	</tr>	</thead>"

	'tfoot

	'tbody
	strCalendar=strCalendar & "	<tbody>"

	j=0
	Dim b1,b2
	Dim strDay
	For i=1 to b

		If (j Mod 7)=0 Then strCalendar=strCalendar & "<tr>"
		If (j/7)<=0 and firw<>1 then strCalendar=strCalendar & "<td class=""pad"" colspan="""& (firw-1) &"""> </td>"

		If (j=>firw-1) and (k=<d) Then

			strCalendar=strCalendar & "<td "

			If 	Cdate(y&"-"&m&"-"&k) = Date() Then
				strCalendar=strCalendar & " id =""today"" "
				b1="<b>"
				b2="</b>"
			Else
				b1=""
				b2=""
			End If

			If aryDateLink(k) Then
				strDay = Year(aryDateArticle(k).PostTime)&"-"&Month(aryDateArticle(k).PostTime)&"-"&Day(aryDateArticle(k).PostTime)
				If ZC_POST_STATIC_MODE="REWRITE" Then
					strCalendar=strCalendar & "><a title=""" & strDay & """ href=""" & Replace(ZC_DATE_REGEX,"{%date%}",strDay) & """>"&b1&(k)&b2&"</a></td>"
				Else
					strCalendar=strCalendar & "><a title=""" & strDay & """ href="""& ZC_BLOG_HOST &"catalog.asp?date=" & strDay & """>"&b1&(k)&b2&"</a></td>"
				End If
			Else
				strCalendar=strCalendar &">"&b1&(k)&b2&"</td>"
			End If

			k=k+1
		End If

		if j=b-1 AND lasw<>7 then strCalendar=strCalendar & "<td class=""pad"" colspan="""& (7-lasw) &"""> </td>"

		If (j Mod 7)=6 Then strCalendar=strCalendar & "</tr>"

		j=j+1
	Next

	strCalendar=strCalendar & "	</tbody></table>"
	strCalendar = Replace(strCalendar,"{%host%}/",ZC_BLOG_HOST)
	strCalendar = Replace(strCalendar,"/default.html","/")
	MakeCalendar=strCalendar

End Function
'*********************************************************




'*********************************************************
' 目的：    加载默认的主题模板
'*********************************************************
Function LoadDefaultTemplates()

If TemplateDic.Exists("TEMPLATE_B_ARTICLE-ISTOP")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE-ISTOP",LoadFromFile(BlogPath &"zb_system\defend\default\b_article-istop.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE-MULTI")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE-MULTI",LoadFromFile(BlogPath &"zb_system\defend\default\b_article-multi.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE-SINGLE")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE-SINGLE",LoadFromFile(BlogPath &"zb_system\defend\default\b_article-single.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_COMMENT")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_COMMENT",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_comment.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_COMMENTPOST-VERIFY")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_COMMENTPOST-VERIFY",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_commentpost-verify.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_COMMENTPOST")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_COMMENTPOST",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_commentpost.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_comment_pagebar.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_MUTUALITY")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_MUTUALITY",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_mutuality.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_NAVBAR_L")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_NAVBAR_L",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_navbar_l.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_NAVBAR_R")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_NAVBAR_R",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_navbar_r.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_TAG")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_TAG",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_tag.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_FUNCTION")=False Then Call TemplateDic.add("TEMPLATE_B_FUNCTION",LoadFromFile(BlogPath &"zb_system\defend\default\b_function.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_PAGEBAR")=False Then Call TemplateDic.add("TEMPLATE_B_PAGEBAR",LoadFromFile(BlogPath &"zb_system\defend\default\b_pagebar.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_CATALOG")=False Then Call TemplateDic.add("TEMPLATE_CATALOG",LoadFromFile(BlogPath &"zb_system\defend\default\catalog.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_DEFAULT")=False Then Call TemplateDic.add("TEMPLATE_DEFAULT",LoadFromFile(BlogPath &"zb_system\defend\default\default.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_FOOTER")=False Then Call TemplateDic.add("TEMPLATE_FOOTER",LoadFromFile(BlogPath &"zb_system\defend\default\footer.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_HEADER")=False Then Call TemplateDic.add("TEMPLATE_HEADER",LoadFromFile(BlogPath &"zb_system\defend\default\header.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_SIDEBAR")=False Then Call TemplateDic.add("TEMPLATE_SIDEBAR",LoadFromFile(BlogPath &"zb_system\defend\default\.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_SINGLE")=False Then Call TemplateDic.add("TEMPLATE_SINGLE",LoadFromFile(BlogPath &"zb_system\defend\default\single.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE-SEARCH-CONTENT")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE-SEARCH-CONTENT",LoadFromFile(BlogPath &"zb_system\defend\default\b_article-search-content.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE-PAGE")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE-PAGE",LoadFromFile(BlogPath &"zb_system\defend\default\b_article-page.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_PAGE")=False Then Call TemplateDic.add("TEMPLATE_PAGE",LoadFromFile(BlogPath &"zb_system\defend\default\page.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR_L")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR_L",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_comment_pagebar_l.html","utf-8"))
If TemplateDic.Exists("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR_R")=False Then Call TemplateDic.add("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR_R",LoadFromFile(BlogPath &"zb_system\defend\default\b_article_comment_pagebar_r.html","utf-8"))

'为2.2改了模板的默认表单
Dim s
s=TemplateDic.Item("TEMPLATE_B_ARTICLE_COMMENTPOST")

If InStr(s,"inpRevID")=0 Then
	If InStr(s,"inpLocation")>0 Then
		s=Replace(s,"inpLocation","inpRevID")
	Else
		s=Replace(s,"<input","<input type=""hidden"" name=""inpRevID"" id=""inpRevID"" value="""" /><input",1,1)
	End If
End If
TemplateDic.Item("TEMPLATE_B_ARTICLE_COMMENTPOST")=s

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

	Dim f, f1, fc, s, i

	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")

	If PublicObjFSO.FolderExists(BlogPath & strDir)=False Then
		LoadIncludeFiles=Array()
		Exit Function
	End If

	Set f = PublicObjFSO.GetFolder(BlogPath & strDir)
	Set fc = f.Files

	i=0

	For Each f1 in fc
		i=i+1
		ReDim Preserve aryFileList(i)
		aryFileList(i)=f1.name
	Next

	LoadIncludeFiles=aryFileList

End Function
'*********************************************************



'*********************************************************
' 目的：    加载指定目录的文件列表
'*********************************************************
Function LoadIncludeFilesOnlyType(strDir)

	Dim aryFileList()
	ReDim aryFileList(-1)

	Dim f, f1, fc, s, i

	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")

	If PublicObjFSO.FolderExists(BlogPath & strDir)=False Then
		LoadIncludeFilesOnlyType=Array()
		Exit Function
	End If

	Set f = PublicObjFSO.GetFolder(BlogPath & strDir)
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
' 目的：    得到模板标签
'*********************************************************
Function GetTemplateTags(Name)
	GetTemplateTags=TemplateTagsDic.Item(Name)
End Function
'*********************************************************




'*********************************************************
' 目的：    设置模板标签
'*********************************************************
Function SetTemplateTags(Name,Value)
	TemplateTagsDic.Item(Name)=Value
End Function
'*********************************************************




'*********************************************************
' 目的：    Check Template Modified Date
'*********************************************************
Function CheckTemplateModified()

	Dim f, f1, fc, s
	Dim d,nd

	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")

	If PublicObjFSO.FolderExists(BlogPath & "zb_users\" & "theme" & "\" & ZC_BLOG_THEME & "\" & ZC_TEMPLATE_DIRECTORY)=False Then Exit Function
	Set f = PublicObjFSO.GetFolder(BlogPath & "zb_users\" & "theme" & "\" & ZC_BLOG_THEME & "\" & ZC_TEMPLATE_DIRECTORY)
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

	On Error Resume Next

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
		For ii=0 to jj
			If TemplateDic.Exists(TemplatesName(ii))=False Then TemplateDic.Add TemplatesName(ii), TemplatesContent(ii)
		Next

		jj=UBound(TemplateTagsName)
		For ii=0 to jj
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

			ReDim Preserve aryTemplatesName(j)
			ReDim Preserve aryTemplatesContent(j)

			For i=1 to j

				aryFileNameTemplate(i)="theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY & "/" & aryFileList(i)
				aryFileNameTemplate_Variable(i)="TEMPLATE_" & UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
				If InStr(aryFileList(i),".")=0 Then
					aryFileNameTemplate_Variable(i)="TEMPLATE_" & UCase(aryFileList(i))
				End If
				aryTemplatesName(i)=aryFileNameTemplate_Variable(i)

				strContent=""
				strContent=LoadFromFile(BlogPath & "zb_users\" & aryFileNameTemplate(i),"utf-8")

				aryTemplatesContent(i)=strContent
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

	ReDim Preserve aryTemplatesName(j+5)
	ReDim Preserve aryTemplatesContent(j+5)


	aryTemplatesName(j+1)="CACHE_SIDEBAR"
	aryTemplatesName(j+2)="CACHE_SIDEBAR2"
	aryTemplatesName(j+3)="CACHE_SIDEBAR3"
	aryTemplatesName(j+4)="CACHE_SIDEBAR4"
	aryTemplatesName(j+5)="CACHE_SIDEBAR5"

	aryTemplatesContent(j+1)=LoadFromFile(BlogPath & "zb_users\cache" & "\sidebar.asp","utf-8" )
	aryTemplatesContent(j+2)=LoadFromFile(BlogPath & "zb_users\cache" & "\sidebar2.asp","utf-8")
	aryTemplatesContent(j+3)=LoadFromFile(BlogPath & "zb_users\cache" & "\sidebar3.asp","utf-8")
	aryTemplatesContent(j+4)=LoadFromFile(BlogPath & "zb_users\cache" & "\sidebar4.asp","utf-8")
	aryTemplatesContent(j+5)=LoadFromFile(BlogPath & "zb_users\cache" & "\sidebar5.asp","utf-8")


	'加载标签
	Dim a,b,c,d,e,a2,a3,f
	Dim t()
	Dim s()

	a=0
	b=24
	c=1
	d=350
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
				functionstype.LoadString=LoadFromFile(BlogPath & "zb_users\cache\functionstype.asp","utf-8")


				If functionstype.GetValue(modname)="div" Then
					' aryTemplateTagsValue(e+i+a)="<div id=""mod_"+modname+""" style=""display:none;""><script type=""text/javascript"">LoadFunction('"&modname&"');</script></div>"
					aryTemplateTagsValue(e+i+a)="<div class=""LoadMod"" data-mod=""" & modname & """ id=""mod_" & modname & """ style=""display:none;""></div>"
				Else
					' aryTemplateTagsValue(e+i+a)="<li id=""mod_"+modname+""" style=""display:none;""><script type=""text/javascript"">LoadFunction('"&modname&"');</script></li>"
					aryTemplateTagsValue(e+i+a)="<li class=""LoadMod"" data-mod=""" & modname & """ id=""mod_" & modname & """ style=""display:none;""></li>"
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
	t(3)="ZC_BLOG_TITLE"
	t(4)="ZC_BLOG_SUBTITLE"
	t(5)="ZC_BLOG_NAME"
	t(6)="ZC_BLOG_SUB_NAME"
	t(7)="ZC_BLOG_THEME"
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
	t(18)="ZC_HOST_TIME_ZONE"
	t(19)="ZC_VERIFYCODE_WIDTH"
	t(20)="ZC_VERIFYCODE_HEIGHT"
	t(21)="ZC_BLOG_HOST"
	t(22)="ZC_BLOG_PRODUCT"
	t(23)="ZC_BLOG_PRODUCT_FULL"
	t(24)="ZC_BLOG_PRODUCT_FULLHTML"


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
	strTag=Replace(strTag,"，",",")
	strTag=TransferHTML(strTag,"[normaltag]")

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
Function GetReallyDirectory()

	If CurrentReallyDirectory<>"" Then
		GetReallyDirectory=CurrentReallyDirectory
		Exit Function
	End If

	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")

	Dim p
	p=Server.MapPath(".") & "\"

	If PublicObjFSO.FolderExists(p & "ZB_SYSTEM\") Then
		p=p
	ElseIf PublicObjFSO.FolderExists(p & "..\ZB_SYSTEM\") Then
		p=p & "..\"
	ElseIf PublicObjFSO.FolderExists(p & "..\..\ZB_SYSTEM\") Then
		p=p & "..\..\"
	ElseIf PublicObjFSO.FolderExists(p & "..\..\..\ZB_SYSTEM\") Then
		p=p & "..\..\..\"
	ElseIf PublicObjFSO.FolderExists(p & "..\..\..\..\ZB_SYSTEM\") Then
		p=p & "..\..\..\..\"
	ElseIf PublicObjFSO.FolderExists(p & "..\..\..\..\..\ZB_SYSTEM\") Then
		p=p & "..\..\..\..\..\"
	ElseIf PublicObjFSO.FolderExists(p & "..\..\..\..\..\..\ZB_SYSTEM\") Then
		p=p & "..\..\..\..\..\..\"
	ElseIf PublicObjFSO.FolderExists(p & "..\..\..\..\..\..\..\ZB_SYSTEM\") Then
		p=p & "..\..\..\..\..\..\..\"
	End If
	Set fso=Nothing

	GetReallyDirectory=PublicObjFSO.GetFolder(p).Path & "\"

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

	Application(ZC_BLOG_CLSID & BlogUser.ID & "SIGNAL_OPERATEINFO")=Application(ZC_BLOG_CLSID & BlogUser.ID & "SIGNAL_OPERATEINFO") & vbCrlf &  strInfo

	Application.UnLock

End Function
'*********************************************************




'*********************************************************
' 目的：    设置提示标志withCLSID
'*********************************************************
Function SetBlogHintWithCLSID(bolOperateSuccess,bolRebuildIndex,bolRebuildFiles,newCLSID)

	Application.Lock

	Application(newCLSID  & BlogUser.ID & "SIGNAL_OPERATESUCCESS")=bolOperateSuccess

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
	bolOperateSuccess=Application(ZC_BLOG_CLSID & BlogUser.ID & "SIGNAL_OPERATESUCCESS")
	bolRebuildIndex=Application(ZC_BLOG_CLSID & "SIGNAL_REBUILDINDEX")
	bolRebuildFiles=Application(ZC_BLOG_CLSID & "SIGNAL_REBUILDFILES")
	strOperateInfo=Application(ZC_BLOG_CLSID & BlogUser.ID & "SIGNAL_OPERATEINFO")
	Application(ZC_BLOG_CLSID & BlogUser.ID & "SIGNAL_OPERATEINFO")=Empty
	Application.UnLock


	If IsEmpty(bolOperateSuccess)=False Then

		If bolOperateSuccess=True Then
			Response.Write "<div class='hint'><p class='hint hint_green'><font color='green'>" & ZC_MSG266 & "</font></p></div>"
		End If

		If bolOperateSuccess=False Then
			Response.Write "<div class='hint'><p class='hint hint_red'><font color='red'>" & ZC_MSG267 & "</font></p></div>"
		End If

		Application.Lock
		Application(ZC_BLOG_CLSID & BlogUser.ID & "SIGNAL_OPERATESUCCESS")=Empty
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
			Response.Write "<div class='hint'><p class='hint hint_blue'><font color='blue'>" & Replace(ZC_MSG269,"%u",BlogHost&"zb_system/cmd.asp?act=AskFileReBuild") & "</font></p></div>"
		End If

	End If


End Function
'*********************************************************




'*********************************************************
' 目的：    解析 REGEX For Path
'*********************************************************
Function ParseCustomDirectoryForPath(strRegex,strPost,strCategory,strUser,strYear,strMonth,strDay,strID,strName,strAlias)
	Dim s
	s=ParseCustomDirectory(strRegex,strPost,strCategory,strUser,strYear,strMonth,strDay,strID,strName,strAlias)
	s=Replace(s,"{%host%}",Left(BlogPath,Len(BlogPath)-1))
	ParseCustomDirectoryForPath=Replace(s,"/","\")
End Function
'*********************************************************




'*********************************************************
' 目的：    解析 REGEX For Url
'*********************************************************
Function ParseCustomDirectoryForUrl(strRegex,strPost,strCategory,strUser,strYear,strMonth,strDay,strID,strName,strAlias)
	Dim s
	s=ParseCustomDirectory(strRegex,strPost,strCategory,strUser,strYear,strMonth,strDay,strID,strName,strAlias)
	s=Replace(s,"{%host%}",Left(BlogHost,Len(BlogHost)-1))
	ParseCustomDirectoryForUrl=Replace(s,"\","/")
End Function
'*********************************************************




'*********************************************************
' 目的：    解析ZC_CUSTOM_DIRECTORY_REGEX
'*********************************************************
Function ParseCustomDirectory(strRegex,strPost,strCategory,strUser,strYear,strMonth,strDay,strID,strName,strAlias)

	On Error Resume Next

	Dim s,d
	s=strRegex

	d=strYear
	If strMonth<>"" Then
		d=d & "-" & Right("0" & strMonth,2)
	End If
	If strDay<>"" Then
		d=d & "-" & Right("0" & strDay,2)
	End If

	s=Replace(s,"{%post%}",strPost)
	s=Replace(s,"{%category%}",strCategory)
	s=Replace(s,"{%author%}",strUser)
	s=Replace(s,"{%year%}",strYear)
	s=Replace(s,"{%month%}",Right("0" & strMonth,2))
	s=Replace(s,"{%day%}",Right("0" & strDay,2))
	s=Replace(s,"{%id%}",strID)
	s=Replace(s,"{%alias%}",strAlias)
	s=Replace(s,"{%name%}",strName)
	s=Replace(s,"{%date%}",d)
	' s=Replace(s,"{%page%}",Request.QueryString("page"))
	' s=Replace(s,"//default.html","/default.html")

	ParseCustomDirectory=s

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    按照CustomDirectory指示创建相应的目录
'*********************************************************
Sub CreatDirectoryByCustomDirectoryWithFullBlogPath(ByVal strCustomDirectory)

	On Error Resume Next

	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")

	Dim s
	Dim t
	Dim i
	Dim j

	s=BlogPath

	strCustomDirectory=Replace(strCustomDirectory,"/","\")

	strCustomDirectory=Right(strCustomDirectory,Len(strCustomDirectory)-Len(BlogPath))

	t=Split(strCustomDirectory,"\")

	For i=LBound(t) To UBound(t)-1
		If (IsEmpty(t(i))=False) And (t(i)<>"") Then
			s=s & t(i) & "\"
			If (PublicObjFSO.FolderExists(s)=False) Then
				Call PublicObjFSO.CreateFolder(s)
			End If
		End If
	Next


	Err.Clear

End Sub
'*********************************************************






'*********************************************************
' 目的：    按照CustomDirectory指示创建相应的目录
'*********************************************************
Sub CreatDirectoryByCustomDirectory(ByVal strCustomDirectory)

	On Error Resume Next

	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")

	Dim s
	Dim t
	Dim i
	Dim j

	s=BlogPath

	strCustomDirectory=Replace(strCustomDirectory,"/","\")

	t=Split(strCustomDirectory,"\")

	j=0
	For i=LBound(t) To UBound(t)
		If (IsEmpty(t(i))=False) And (t(i)<>"") Then
			s=s & t(i) & "\"
			If (PublicObjFSO.FolderExists(s)=False) Then
				Call PublicObjFSO.CreateFolder(s)
			End If
			j=j+1
		End If
	Next



	Err.Clear

End Sub
'*********************************************************




'*********************************************************
' 目的：  生成左侧导航栏
'*********************************************************
Dim AdminLeftMenuCount
AdminLeftMenuCount=0
Function MakeLeftMenu(requireLevel,strName,strUrl,strLiId,strAId,strImgUrl)

	If BlogUser.Level>requireLevel Then Exit Function

	AdminLeftMenuCount=AdminLeftMenuCount+1
	dim tmp
	If Trim(strImgUrl)<>"" Then
		tmp="<li id="""&strLiId&"""><a id="""&strAId&""" href="""&strUrl&"""><span style=""background-image:url('"&strImgUrl&"')"">"&strName&"</span></a></li>"
	Else
		tmp="<li id="""&strLiId&"""><a id="""&strAId&""" href="""&strUrl&"""><span>"&strName&"</span></a></li>"
	End If
	MakeLeftMenu=tmp

End Function
'*********************************************************




'*********************************************************
' 目的：  生成头部菜单
'*********************************************************
Dim AdminTopMenuCount
AdminTopMenuCount=0
Function MakeTopMenu(requireLevel,strName,strUrl,strLiId,strTarget)

	If BlogUser.Level>requireLevel Then Exit Function

	Dim tmp
	If strTarget="" Then strTarget="_self"
	AdminTopMenuCount=AdminTopMenuCount+1
	If strLiId="" Then strLiId="topmenu"&AdminTopMenuCount
	tmp="<li id="""&strLiId&"""><a href="""&strUrl&""" target="""&strTarget&""">"&strName&"</a></li>"
	MakeTopMenu=tmp
End Function
'*********************************************************




'*********************************************************
' 目的： 加入二级菜单项
'*********************************************************
Function MakeSubMenu(strName,strUrl,strType,isNewWindows)

	Dim strSource

	strSource=strSource & "<a " & "href=""" & strUrl  & """"

	If isNewWindows=True Then strSource=strSource & " target=""_blank"""

	strSource=strSource & ">" & "<span class=""" & strType & """>"


	strSource=strSource & strName

	strSource=strSource & "</span></a>"

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
Dim IsRunActivePlugin
Function ActivePlugin()

	If IsRunActivePlugin=True Then Exit Function

	On Error Resume Next

	Dim i
	For i=0 To UBound(PluginActiveFunction)-1

		Call Execute(PluginActiveFunction(i))

	Next

	Err.Clear

	IsRunActivePlugin=True

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
	If Boolean_Delay_Plugin_Signal = True Then
		Call DelayAdd_Action_Plugin(plugname,actioncode)
		Exit Function
	End If
	On Error Resume Next
	actioncode=Replace(actioncode,"Exit Function","b" & plugname & "=True")
	actioncode=Replace(actioncode,"Exit Sub","b" & plugname & "=True")
	actioncode=Replace(actioncode,"Exit Property","b" & plugname & "=True")
	Call Execute("ReDim Preserve " & plugname & "(UBound("& plugname &")+1)")
	Call Execute(plugname & "(UBound("& plugname &"))=" & plugname & "(UBound("& plugname &"))&""" & Replace(actioncode,"""","""""") & """" & ":")
	Err.Clear
End Function
'*********************************************************




'*********************************************************
Dim String_Action_Plugin_ByOnce
String_Action_Plugin_ByOnce = ""
Function DelayAdd_Action_Plugin(plugname,actioncode)
	On Error Resume Next
	String_Action_Plugin_ByOnce = String_Action_Plugin_ByOnce & vbCrLf & "ReDim Preserve " & plugname & "(UBound("& plugname &")+1)"
	actioncode=Replace(actioncode,"Exit Function","b" & plugname & "=True")
	actioncode=Replace(actioncode,"Exit Sub","b" & plugname & "=True")
	actioncode=Replace(actioncode,"Exit Property","b" & plugname & "=True")
	String_Action_Plugin_ByOnce = String_Action_Plugin_ByOnce & vbCrLf & plugname & "(UBound("& plugname &"))=" & plugname & "(UBound("& plugname &"))&""" & Replace(actioncode,"""","""""") & """" & ":"
	Err.Clear
End Function
Function Execute_Action_Plugin()
	On Error Resume Next
	If String_Action_Plugin_ByOnce ="" Then Exit Function
	Call Execute(String_Action_Plugin_ByOnce)
	String_Action_Plugin_ByOnce = ""
	Err.Clear
End Function
'*********************************************************




'*********************************************************
' 目的：挂上Filter接口
' 参数：'plugname:接口名称
		'functionname:要挂接的函数名
'*********************************************************
Function Add_Filter_Plugin(plugname,functionname)
	If Boolean_Delay_Plugin_Signal = True Then
		Call DelayAdd_Filter_Plugin(plugname,functionname)
		Exit Function
	End If
	On Error Resume Next
	Call Execute("s" & plugname & "=" & "s" & plugname & "&""" & functionname & """" & "& ""|""")
	Err.Clear
End Function
'*********************************************************




'*********************************************************
Dim String_Filter_Plugin_ByOnce
String_Filter_Plugin_ByOnce = ""
Function DelayAdd_Filter_Plugin(plugname,functionname)
	On Error Resume Next
	String_Filter_Plugin_ByOnce = String_Filter_Plugin_ByOnce & vbCrLf & "s" & plugname & "=" & "s" & plugname & "&""" & functionname & """" & "& ""|"""
	Err.Clear
End Function
Function Execute_Filter_Plugin()
	On Error Resume Next
	If String_Filter_Plugin_ByOnce ="" Then Exit Function
	Call Execute(String_Filter_Plugin_ByOnce)
	String_Filter_Plugin_ByOnce = ""
	Err.Clear
End Function
'*********************************************************




'*********************************************************
' 目的：挂上Response接口
' 参数：'plugname:接口名称
		'parameter:要写入的内容
'*********************************************************
Function Add_Response_Plugin(plugname,parameter)
	If Boolean_Delay_Plugin_Signal = True Then
		Call DelayAdd_Response_Plugin(plugname,parameter)
		Exit Function
	End If
	On Error Resume Next
	Call Execute(plugname & "=" & plugname & "&""" & Replace(Replace(Replace(Replace(parameter,"""",""""""),vbCrlf,"""&vbCrlf&"""),vbLf,"""&vbLf&"""),vbCr,"""&vbCr&""") & """")
	Err.Clear
End Function
'*********************************************************



'*********************************************************
Dim String_Response_Plugin_ByOnce
String_Response_Plugin_ByOnce = ""
Function DelayAdd_Response_Plugin(plugname,parameter)
	On Error Resume Next
	String_Response_Plugin_ByOnce = String_Response_Plugin_ByOnce & vbCrLf & plugname & "=" & plugname & "&""" & Replace(Replace(Replace(Replace(parameter,"""",""""""),vbCrlf,"""&vbCrlf&"""),vbLf,"""&vbLf&"""),vbCr,"""&vbCr&""") & """"
	Err.Clear
End Function
Function Execute_Response_Plugin()
	On Error Resume Next
	If String_Response_Plugin_ByOnce ="" Then Exit Function
	Call Execute(String_Response_Plugin_ByOnce)
	String_Response_Plugin_ByOnce = ""
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

	BlogReBuild_Authors

	BlogReBuild_Archives

	BlogReBuild_Previous

	BlogReBuild_Comments

	BlogReBuild_Catalogs

	BlogReBuild_Categorys

	BlogReBuild_Calendar

	BlogReBuild_Tags

	BlogReBuild_Functions

	BuildAllCache

	ExportRSS

	BlogReBuild_Default

	Dim bolOperateSuccess

	Application.Lock
	bolOperateSuccess=Application(ZC_BLOG_CLSID  & BlogUser.ID & "SIGNAL_OPERATESUCCESS")
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
	Call GetFunction()
	If Functions(FunctionMetas.GetValue("calendar")).IsHidden=True Then

		Exit Function
	End If
	'Calendar

	strCalendar=MakeCalendar("")

	strCalendar=TransferHTML(strCalendar,"[no-asp]")



	Functions(FunctionMetas.GetValue("calendar")).Content=strCalendar
	Functions(FunctionMetas.GetValue("calendar")).Post()
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
	Dim s

	Dim ArtList

	Call GetFunction()
	If Functions(FunctionMetas.GetValue("archives")).IsHidden=True Then

		Exit Function
	End If
	'Archives
	Dim strArchives
	Set objRS=objConn.Execute("SELECT [log_PostTime] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_Level]>1) ORDER BY [log_PostTime] DESC")
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


If BlogConfig.Read("ZC_ARCHIVES_OLD_LISTTYPE")="True" Then
	If Not IsEmpty(dtmYM) Then
		For i=1 to UBound(dtmYM)

			l=Year(dtmYM(i))
			n=Month(dtmYM(i))+1
			IF n>12 Then l=l+1:n=1

			Set objRS=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE ([log_Type]=0) And ([log_Level]>1) AND [log_PostTime] BETWEEN "& ZC_SQL_POUND_KEY & Year(dtmYM(i)) &"-"& Month(dtmYM(i)) &"-1"& ZC_SQL_POUND_KEY &" AND "& ZC_SQL_POUND_KEY & l &"-"& n &"-1" & ZC_SQL_POUND_KEY)

			If (Not objRS.bof) And (Not objRS.eof) Then
				strArchives=strArchives & "<li><a href="""& UrlbyDateAuto(Year(dtmYM(i)),Month(dtmYM(i)),"") &""">" & Year(dtmYM(i)) & " " & ZVA_Month(Month(dtmYM(i))) & "<span class=""article-nums""> (" & objRS(0) & ")</span>" +"</a></li>"
				If j>0 Then
					If i=j Then Exit For
				End If
			End If

			objRS.Close
			Set objRS=Nothing
		Next
	End If
Else
	s="<li>"
	If Not IsEmpty(dtmYM) Then
		For i=1 to UBound(dtmYM)

			l=Year(dtmYM(i))
			n=Month(dtmYM(i))+1
			IF n>12 Then l=l+1:n=1

			Set objRS=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE ([log_Type]=0) And ([log_Level]>1) AND [log_PostTime] BETWEEN "& ZC_SQL_POUND_KEY & Year(dtmYM(i)) &"-"& Month(dtmYM(i)) &"-1"& ZC_SQL_POUND_KEY &" AND "& ZC_SQL_POUND_KEY & l &"-"& n &"-1" & ZC_SQL_POUND_KEY)

			If (Not objRS.bof) And (Not objRS.eof) Then
				If InStr(s,"<!-- year -->"&Replace(ZC_MSG233,"%y",Year(dtmYM(i)))&"<!-- year -->")=0 Then s=s & "</li><li><!-- year -->"&Replace(ZC_MSG233,"%y",Year(dtmYM(i)))&"<!-- year --><br/>"

				s=s & "<a href="""& UrlbyDateAuto(Year(dtmYM(i)),Month(dtmYM(i)),"") &""">" &  ZVA_Month_Abbr(Month(dtmYM(i)))  & "<span class=""article-nums"">(" & objRS(0) & ")</span>" +"</a>&nbsp; "
				If j>0 Then
					If i=j Then Exit For
				End If
			End If

			objRS.Close
			Set objRS=Nothing
		Next
	End If
	s=s & "</li>"
	s=Replace(s,"<li></li>","")
	strArchives=Replace(s,"<!-- year -->","")
End If

	strArchives=TransferHTML(strArchives,"[no-asp]")
	Functions(FunctionMetas.GetValue("archives")).Content=strArchives
	Functions(FunctionMetas.GetValue("archives")).Post()

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
	Call GetFunction
	If Functions(FunctionMetas.GetValue("catalog")).IsHidden=True Then

		Exit Function
	End If
	Call GetCategory()

	Dim objRS
	Dim objStream

	Dim ArtList

	'Catalogs
	Dim strCatalog,bolHasSubCate

	Dim aryCateInOrder
	aryCateInOrder=GetCategoryOrder()


	Categorys(0).Count=CLng(objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE [log_Level]>1 AND [log_Type]=0 AND [log_CateID]=0")(0))
	If Categorys(0).Count>0 Then
		strCatalog=strCatalog & "<li class=""li-cate cate-"& Categorys(0).id &"""><a href="""& Categorys(0).Url & """>"+Categorys(0).Name + "<span class=""article-nums""> (" & Categorys(0).Count & ")</span>" +"</a></li>"
	End If

	Dim i,j,c
	Dim strPrecata,strSubcate
	For i=Lbound(aryCateInOrder)+1 To Ubound(aryCateInOrder)
		strPrecata="":strSubcate=""
		If Categorys(aryCateInOrder(i)).ParentID=0 Then
			c=Categorys(aryCateInOrder(i)).Count

			bolHasSubCate=False
			For j=Lbound(aryCateInOrder)+1 To UBound(aryCateInOrder)
				If Categorys(aryCateInOrder(j)).ParentID=Categorys(aryCateInOrder(i)).ID Then bolHasSubCate=True
			Next
			'If bolHasSubCate Then strSubcate = "<ul class=""ul-subcates"">"
			For j=Lbound(aryCateInOrder)+1 To UBound(aryCateInOrder)
				If Categorys(aryCateInOrder(j)).ParentID=Categorys(aryCateInOrder(i)).ID And Categorys(aryCateInOrder(j)).Count>0 Then
					strSubcate=strSubcate & "<li class=""li-subcate cate-"& Categorys(aryCateInOrder(j)).ID &"""><a href="""& Categorys(aryCateInOrder(j)).Url & """>"+Categorys(aryCateInOrder(j)).Name + "<span class=""article-nums""> (" & Categorys(aryCateInOrder(j)).Count & ")</span>" +"</a></li>"
					c=c+Categorys(aryCateInOrder(j)).Count
				End If
			Next
			If bolHasSubCate And strSubcate<>"" Then strSubcate="<ul class=""ul-subcates"">" & strSubcate & "</ul>"

			If c>0 Then strPrecata="<li class=""li-cate cate-"& Categorys(aryCateInOrder(i)).ID &"""><a href="""& Categorys(aryCateInOrder(i)).Url & """>"+Categorys(aryCateInOrder(i)).Name + "<span class=""article-nums""> (" & c & ")</span>" +"</a>"& strSubcate &"</li>"

			strCatalog=strCatalog & strPrecata
		End If

	Next

	strCatalog=TransferHTML(strCatalog,"[no-asp]")


	Functions(FunctionMetas.GetValue("catalog")).Content=strCatalog
	Functions(FunctionMetas.GetValue("catalog")).Post()

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

	Call GetUser()
	Call GetFunction()
	If Functions(FunctionMetas.GetValue("authors")).IsHidden=True Then

		Exit Function
	End If

	Dim objRS
	Dim objStream

	'Authors
	Dim strAuthor
	Dim User
	For Each User in Users
		If IsObject(User) Then''''''
			If User.ID>0 And User.Level<4 And User.Count>0 Then
				strAuthor=strAuthor & "<li><a href="""& User.Url & """>"+User.FirstName + "<span class=""article-nums""> (" & User.Count & ")" +"</span></a></li>"
			End If
		End If
	Next

	strAuthor=TransferHTML(strAuthor,"[no-asp]")


	Functions(FunctionMetas.GetValue("authors")).Content=strAuthor
	Functions(FunctionMetas.GetValue("authors")).Post()

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

	If Functions(FunctionMetas.GetValue("tags")).IsHidden=True Then

		Exit Function
	End If

	Dim objRS
	Dim objStream

	Dim i,j,s,t,h

	i=Functions(FunctionMetas.GetValue("tags")).MaxLi
	If i=0 Then i=25
	j=0
	'Authors
	Dim strTag

	Set objRS=objConn.Execute("SELECT [tag_id] FROM [blog_Tag] ORDER BY [tag_Count] DESC,[tag_ID] ASC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		Do While Not objRS.eof
			s=s & "{" & objRS("tag_ID") & "}"
			t=t & objRS("tag_ID") & ","
			objRS.MoveNext
			j=j+1
			If j>i Then Exit Do
		Loop
	End If
	objRS.Close
	Set objRS=Nothing

	Call GetTagsbyTagIDList(s)

	Set h=CreateObject("Scripting.Dictionary")

	s=Split(t,",")
	For i=0 To UBound(s)-1
		If s(i)<>"" And Tags(s(i)).Count<>0 Then
			h.add s(i),Tags(s(i))
		End If
	Next

	For Each s In Tags
		If IsObject(s)=True Then
			If h.Exists(CStr(s.ID)) Then
				strTag=strTag & "<li class=""tag-name tag-name-size-"&TagCloud(s.Count)&"""><a href="""&s.Url&""">"+s.Name + "<span class=""tag-count""> (" & s.Count & ")</span>" +"</a></li>"
			End If
		End If
	Next


	strTag=TransferHTML(strTag,"[no-asp]")


	Functions(FunctionMetas.GetValue("tags")).Content=strTag
	Functions(FunctionMetas.GetValue("tags")).Post()

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


	Dim i,j
	Dim objRS
	Dim objStream
	Dim objArticle
	Call GetFunction()
	j=Functions(FunctionMetas.GetValue("previous")).MaxLi

	If Functions(FunctionMetas.GetValue("previous")).IsHidden=True Then

		Exit Function
	End If

	If j=0 Then j=10

	'Previous
	Dim strPrevious
	Set objRS=objConn.Execute("SELECT TOP "&j&" [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_ID]>0) AND ([log_Level]>1) ORDER BY [log_PostTime] DESC")

	If (Not objRS.bof) And (Not objRS.eof) Then
		For i=1 to j
			Set objArticle=New TArticle
			If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then
				strPrevious=strPrevious & "<li><a href="""& objArticle.HtmlUrl & """ title="""& objArticle.HtmlTitle &"""><span class=""article-date"">["& Right("0" & Month(objArticle.PostTime),2) & "/" & Right("0" & Day(objArticle.PostTime),2) &"]</span>" & objArticle.Title & "</a></li>"
			End If
			Set objArticle=Nothing
			objRS.MoveNext
			If objRS.eof Then Exit For
		Next
	End If
	objRS.close

	strPrevious=TransferHTML(strPrevious,"[no-asp]")

	Functions(FunctionMetas.GetValue("previous")).Content=strPrevious
	Functions(FunctionMetas.GetValue("previous")).Post()
	Functions(FunctionMetas.GetValue("previous")).SaveFile

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
	If Functions(FunctionMetas.GetValue("comments")).IsHidden=True Then

		Exit Function
	End If

	Dim objRS
	Dim objStream
	Dim objArticle

	'Comments
	Dim strComments

	Dim s,t
	Dim i,j

	j=Functions(FunctionMetas.GetValue("comments")).MaxLi
	If j=0 Then j=10

	Set objRS=objConn.Execute("SELECT TOP "&j&" [log_ID],[comm_ID],[comm_Content],[comm_PostTime],[comm_AuthorID],[comm_Author] FROM [blog_Comment] WHERE [log_ID]>0 AND [comm_IsCheck]=0 ORDER BY [comm_PostTime] DESC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		For i=1 to j
			Call GetUsersbyUserIDList(objRS("comm_AuthorID"))

			Set objArticle=New TArticle

			If objArticle.LoadInfoByID(objRS("log_ID")) Then
				t=objArticle.Url
			End If

			s=objRS("comm_Content")
			s=Replace(s,vbCrlf,"")
			s=Left(s,ZC_COMMENT_EXCERPT_MAX)
			s=TransferHTML(s,"[nohtml]")

			strComments=strComments & "<li style=""text-overflow:ellipsis;""><a href="""& t & "#cmt" & objRS("comm_ID") & """ title=""" & objRS("comm_PostTime") & " post by " & IIf(Users(objRS("comm_AuthorID")).Level=5,objRS("comm_Author"),Users(objRS("comm_AuthorID")).FirstName) & """>"+s+"</a></li>"
			Set objArticle=Nothing
			objRS.MoveNext
			If objRS.eof Then Exit For
		Next
	End If
	objRS.close
	Set objRS=Nothing

	strComments=TransferHTML(strComments,"[no-asp]")

	Functions(FunctionMetas.GetValue("comments")).Content=strComments
	Functions(FunctionMetas.GetValue("comments")).Post()
	Functions(FunctionMetas.GetValue("comments")).SaveFile

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
	Call GetFunction()
	If Functions(FunctionMetas.GetValue("statistics")).IsHidden=True Then

		Exit Function
	End If

	'Statistics
	Dim strStatistics
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	objRS.Open("SELECT COUNT([log_ID])AS allArticle,SUM([log_ViewNums]) AS allViewNums FROM [blog_Article] WHERE [log_Type]=0")
	If (Not objRS.bof) And (Not objRS.eof) Then
		strStatistics=strStatistics & "<li>"& ZC_MSG082 &":" & objRS("allArticle") & "</li>"
		strStatistics=strStatistics & "<li>"& ZC_MSG124 &":" & objConn.Execute("SELECT SUM([log_CommNums]) FROM [blog_Article]")(0) & "</li>"
		strStatistics=strStatistics & "<li>"& ZC_MSG129 &":" & objRS("allViewNums") & "</li>"
	End If
	objRS.Close

	strStatistics=strStatistics & "<li>"& ZC_MSG204 &":" & GetNameFormTheme(ZC_BLOG_THEME) & "</li>"
	'strStatistics=strStatistics & "<li>"& ZC_MSG083 &":" & ZC_BLOG_CSS & "</li>"

	Set objRS=Nothing

	strStatistics=TransferHTML(strStatistics,"[no-asp]")


	Functions(FunctionMetas.GetValue("statistics")).Content=strStatistics
	Functions(FunctionMetas.GetValue("statistics")).Post()

	BlogReBuild_Statistics=True

End Function
'*********************************************************




'*********************************************************
' 目的：    BlogReBuild Functions
'*********************************************************
Function BlogReBuild_Functions

	'plugin node
	bAction_Plugin_BlogReBuild_Functions_Begin=False
	For Each sAction_Plugin_BlogReBuild_Functions_Begin in Action_Plugin_BlogReBuild_Functions_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Functions_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Functions_Begin)
		If bAction_Plugin_BlogReBuild_Functions_Begin=True Then Exit Function
	Next

	IsRunFunctions=False
	Call GetFunction()

	Call SaveFunctionType()

	Dim i,j,s,t,f

	For Each f In Functions
		If IsObject(f)=True Then
			If f.id>0 And f.SourceType<>"other" Then
				f.SaveFile
			End If
		End If
	Next


	Dim aryFunctionInOrder
	'aryFunctionInOrder=GetFunctionOrder()

	t=GetTemplate("TEMPLATE_B_FUNCTION")

	aryFunctionInOrder=Split(ZC_SIDEBAR_ORDER,":")
	s=""
	For Each f In aryFunctionInOrder
		If FunctionMetas.Exists(f)=True Then
			If Functions(FunctionMetas.GetValue(f)).IsHidden=False Then
				s=s & Functions(FunctionMetas.GetValue(f)).MakeTemplate(t)
			End If
		End If
	Next
	Call SaveToFile(BlogPath & "zb_users/cache/sidebar.asp",s,"utf-8",False)

	aryFunctionInOrder=Split(ZC_SIDEBAR_ORDER2,":")
	s=""
	For Each f In aryFunctionInOrder
		If FunctionMetas.Exists(f)=True Then
			If Functions(FunctionMetas.GetValue(f)).IsHidden=False Then
				s=s & Functions(FunctionMetas.GetValue(f)).MakeTemplate(t)
			End If
		End If
	Next
	Call SaveToFile(BlogPath & "zb_users/cache/sidebar2.asp",s,"utf-8",False)

	aryFunctionInOrder=Split(ZC_SIDEBAR_ORDER3,":")
	s=""
	For Each f In aryFunctionInOrder
		If FunctionMetas.Exists(f)=True Then
			If Functions(FunctionMetas.GetValue(f)).IsHidden=False Then
				s=s & Functions(FunctionMetas.GetValue(f)).MakeTemplate(t)
			End If
		End If
	Next
	Call SaveToFile(BlogPath & "zb_users/cache/sidebar3.asp",s,"utf-8",False)

	aryFunctionInOrder=Split(ZC_SIDEBAR_ORDER4,":")
	s=""
	For Each f In aryFunctionInOrder
		If FunctionMetas.Exists(f)=True Then
			If Functions(FunctionMetas.GetValue(f)).IsHidden=False Then
				s=s & Functions(FunctionMetas.GetValue(f)).MakeTemplate(t)
			End If
		End If
	Next
	Call SaveToFile(BlogPath & "zb_users/cache/sidebar4.asp",s,"utf-8",False)

	aryFunctionInOrder=Split(ZC_SIDEBAR_ORDER5,":")
	s=""
	For Each f In aryFunctionInOrder
		If FunctionMetas.Exists(f)=True Then
			If Functions(FunctionMetas.GetValue(f)).IsHidden=False Then
				s=s & Functions(FunctionMetas.GetValue(f)).MakeTemplate(t)
			End If
		End If
	Next
	Call SaveToFile(BlogPath & "zb_users/cache/sidebar5.asp",s,"utf-8",False)
'	Next

	BlogReBuild_Functions=True

End Function
'*********************************************************




'*********************************************************
' 目的：    BlogReBuild Default
'*********************************************************
Function BlogReBuild_Default

	'plugin node
	bAction_Plugin_BlogReBuild_Default_Begin=False
	For Each sAction_Plugin_BlogReBuild_Default_Begin in Action_Plugin_BlogReBuild_Default_Begin
		If Not IsEmpty(sAction_Plugin_BlogReBuild_Default_Begin) Then Call Execute(sAction_Plugin_BlogReBuild_Default_Begin)
		If bAction_Plugin_BlogReBuild_Default_Begin=True Then Exit Function
	Next

	Call ClearGlobeCache()
	Call LoadGlobeCache()

	TemplateTagsDic.Item("ZC_BLOG_HOST")="<#ZC_BLOG_HOST#>"

	Dim ArtList
	Set ArtList=New TArticleList

	ArtList.LoadCache

	ArtList.template="DEFAULT"

	If ArtList.Export("","","","","",ZC_DISPLAY_MODE_INTRO) Then

		ArtList.Build

		Call SaveToFile(BlogPath & "zb_users/CACHE/default.asp",ArtList.html,"utf-8",False)

	End If


	BlogReBuild_Default=True

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
		.AddChannelAttribute "link",TransferHTML(BlogHost,"[html-format]")
		.AddChannelAttribute "description",TransferHTML(ZC_BLOG_SUBTITLE,"[html-format]")
		.AddChannelAttribute "generator","RainbowSoft Studio Z-Blog " & ZC_BLOG_VERSION
		.AddChannelAttribute "language",ZC_BLOG_LANGUAGE
		.AddChannelAttribute "pubDate",GetTime(Now())

			Dim i
			Dim objRS
			Set objRS=objConn.Execute("SELECT TOP "&ZC_RSS2_COUNT&" [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) AND ([log_Level]>2) ORDER BY [log_PostTime] DESC")

			If (Not objRS.bof) And (Not objRS.eof) Then
				For i=1 to ZC_RSS2_COUNT
					Set objArticle=New TArticle
					If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then

					If ZC_RSS_EXPORT_WHOLE Then
					.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).FirstName & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlContent,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
					Else
					.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).FirstName & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlIntro,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
					End If

					End If
					objRS.MoveNext
					If objRS.eof Then Exit For
					Set objArticle=Nothing
				Next
			End If

	End With

	Rss2Export.SaveToFile(BlogPath & "zb_users/cache/rss.xml")

	Set Rss2Export = Nothing

	objRS.close
	Set objRS=Nothing
	ExportRSS=True

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

	If objArticle.LoadInfoByID(intID) Then
		If objArticle.Export(ZC_DISPLAY_MODE_ALL) Then
			objArticle.SaveCache
			objArticle.Statistic
			objArticle.Build
			objArticle.Save
		End If

		If ZC_POST_STATIC_MODE="STATIC" Then

			If (bolBuildNavigate=True) And (ZC_USE_NAVIGATE_ARTICLE=True) Then

				Dim objRS
				Set objRS=objConn.Execute("SELECT TOP 1 [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND ([log_Type]=0) AND ([log_PostTime]<" & ZC_SQL_POUND_KEY & objArticle.PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] DESC")
				If (Not objRS.bof) And (Not objRS.eof) Then
					Call BuildArticle(objRS("log_ID"),False,False)
				End If
				Set objRS=Nothing
				Set objRS=objConn.Execute("SELECT TOP 1 [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND ([log_Type]=0) AND ([log_PostTime]>" & ZC_SQL_POUND_KEY & objArticle.PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] ASC")
				If (Not objRS.bof) And (Not objRS.eof) Then
					Call BuildArticle(objRS("log_ID"),False,False)
				End If
				Set objRS=Nothing

			End If

		End If

		BuildArticle=True

	End If

	Set objArticle=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    GetUsersbyUserIDList
'*********************************************************
Function GetUsersbyUserIDList(strUsers)
'strTags=1,2,3,4,5,6,7,8

If strUsers="" Then Exit Function

Dim s,t,i,j,d

t=Split(strUsers,",")

Set d = CreateObject("Scripting.Dictionary")

For i=LBound(t) To UBound(t)
	If Trim(t(i))<>"" Then
		If UBound(Users)>=t(i) Then
			If IsObject(Users(t(i)))=False Then
				If d.Exists(t(i))=False Then Call d.add(t(i),t(i))
			End If
		Else
			If d.Exists(t(i))=False Then Call d.add(t(i),t(i))
		End If
	End If
Next

t = d.Keys
For i = 0 To d.Count -1

	If UBound(Users)>=CLng(t(i)) Then
		If IsObject(Users(t(i)))=False Then
			j=j+1
			s=s & "([mem_ID]="&CLng(t(i))&") Or"
		End If
	Else
		j=j+1
		s=s & "([mem_ID]="&CLng(t(i))&") Or"
	End If
	If j=100 Then
		Call GetUsers_Sub(s)
		j=0
		s=""
	End If
Next

Call GetUsers_Sub(s)

GetUsersbyUserIDList=True



End Function
'*********************************************************




'*********************************************************
' 目的：    GetUsersbyUser子函数
'*********************************************************
Sub GetUsers_Sub(s)

If Len(s)>2 Then
	s=Left(s,Len(s)-2)
Else
	Exit Sub
End If

Dim objRS
Dim objUser

Set objRS=objConn.Execute("SELECT [mem_ID],[mem_Name],[mem_Level],[mem_Password],[mem_Email],[mem_HomePage],[mem_PostLogs],[mem_Url],[mem_Template],[mem_FullUrl],[mem_Intro],[mem_Meta] FROM [blog_Member] WHERE (" & s & ")")

If (Not objRS.bof) And (Not objRS.eof) Then

	Do While Not objRS.eof

		Set objUser=New TUser
		Call objUser.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11)))

		If UBound(Users)<objUser.ID Then
			ReDim Preserve Users(objUser.ID)
		End If

		Set Users(objUser.ID)=objUser

		Set objUser=Nothing

		objRS.MoveNext
	Loop

End If

objRS.Close
Set objRS=Nothing

End Sub
'*********************************************************




'*********************************************************
' 目的：    GetTagsbyTagIDList
'*********************************************************
Function GetTagsbyTagIDList(strTags)
'strTags={1}{2}{3}{4}

If strTags="" Then Exit Function
If strTags="{}" Then Exit Function
If IsNull(strTags) Then Exit Function

Dim s,t,i,j,d
strTags=Replace(strTags,"}","")
t=Split(strTags,"{")

Set d = CreateObject("Scripting.Dictionary")

For i=LBound(t) To UBound(t)
	If Trim(t(i))<>"" Then
		If UBound(Tags)>=t(i) Then
			If IsObject(Tags(t(i)))=False Then
				If d.Exists(t(i))=False Then Call d.add(t(i),t(i))
			End If
		Else
			If d.Exists(t(i))=False Then Call d.add(t(i),t(i))
		End If
	End If
Next

t = d.Keys
For i = 0 To d.Count -1

	If UBound(Tags)>=CLng(t(i)) Then
		If IsObject(Tags(t(i)))=False Then
			j=j+1
			s=s & "([tag_ID]="&CLng(t(i))&") Or"
		End If
	Else
		j=j+1
		s=s & "([tag_ID]="&CLng(t(i))&") Or"
	End If
	If j=100 Then
		Call GetTags_Sub(s)
		j=0
		s=""
	End If
Next

Call GetTags_Sub(s)

GetTagsbyTagIDList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    GetTagsbyTag子函数
'*********************************************************
Sub GetTags_Sub(s)

If Len(s)>2 Then
	s=Left(s,Len(s)-2)
Else
	Exit Sub
End If

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

End Sub
'*********************************************************




'*********************************************************
' 目的：    GetTagsbyTagNameList
'*********************************************************
Function GetTagsbyTagNameList(strTags)
'strTags=a,b,c,d,e,f,g

Set Tags(0)=New TTag

If strTags="" Then Exit Function

Dim s,t,i
t=Split(strTags,",")

For i=LBound(t) To UBound(t)
	If Trim(t(i))<>"" Then
		s=s & "([tag_Name]='"&FilterSQL(t(i))&"') Or"
	End If
Next

Call GetTags_Sub(s)

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
' 目的：    废弃
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

	objRS.Open("SELECT [fn_id] FROM [blog_Function] ORDER BY [fn_Order] ASC,[fn_ID] ASC")
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

	Call SaveToFile(BlogPath & "zb_users/CACHE/functionstype.asp",t.SaveString,"utf-8",False)

	SaveFunctionType=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Create Admin Menu
'*********************************************************
Function CreateAdminLeftMenu()

'强制清空Menu,防止某些插件提前插入造成排在系统菜单之前,插件插入菜单要在系统初始化完成后
Response_Plugin_Admin_Left=""

Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("ArticleEdt"),ZC_MSG168,BlogHost&"zb_system/cmd.asp?act=ArticleEdt&amp;webedit="&ZC_BLOG_WEBEDIT,"nav_new","aArticleEdt",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("ArticleMng"),ZC_MSG067,BlogHost&"zb_system/cmd.asp?act=ArticleMng","nav_article","aArticleMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("ArticleAll"),ZC_MSG111,BlogHost&"zb_system/cmd.asp?act=ArticleMng&amp;type=Page","nav_page","aPageMng",""))

Call Add_Response_Plugin("Response_Plugin_Admin_Left","<li class='split'><hr/></li>")


Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("CategoryMng"),ZC_MSG066,BlogHost&"zb_system/cmd.asp?act=CategoryMng","nav_category","aCategoryMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("TagMng"),ZC_MSG141,BlogHost&"zb_system/cmd.asp?act=TagMng","nav_tags","aTagMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("CommentMng"),ZC_MSG068,BlogHost&"zb_system/cmd.asp?act=CommentMng","nav_comments","aCommentMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("FileMng"),ZC_MSG071,BlogHost&"zb_system/cmd.asp?act=FileMng","nav_accessories","aFileMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("UserMng"),ZC_MSG070,BlogHost&"zb_system/cmd.asp?act=UserMng","nav_user","aUserMng",""))

Call Add_Response_Plugin("Response_Plugin_Admin_Left","<li class='split'><hr/></li>")

Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("ThemeMng"),ZC_MSG223,BlogHost&"zb_system/cmd.asp?act=ThemeMng","nav_themes","aThemeMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("FunctionMng"),ZC_MSG007,BlogHost&"zb_system/cmd.asp?act=FunctionMng","nav_function","aFunctionMng",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(GetRights("PlugInMng"),ZC_MSG107,BlogHost&"zb_system/cmd.asp?act=PlugInMng","nav_plugin","aPlugInMng",""))

End Function
'*********************************************************




'*********************************************************
' 目的：    Create Top Menu
'*********************************************************
Function CreateAdminTopMenu()

Response_Plugin_Admin_Top=""

Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(GetRights("admin"),ZC_MSG245,BlogHost&"zb_system/cmd.asp?act=admin","",""))
Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(GetRights("SettingMng"),ZC_MSG247,BlogHost&"zb_system/cmd.asp?act=SettingMng","",""))
If Not ZC_POST_STATIC_MODE<>"STATIC" Then
	Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(GetRights("AskFileReBuild"),ZC_MSG073,BlogHost&"zb_system/cmd.asp?act=AskFileReBuild","",""))
End If

End Function
'*********************************************************




'*********************************************************
' 目的：    html=Response_Plugin_Admin_Left
'*********************************************************
Function ResponseAdminLeftMenu(html)
	Call Filter_Plugin_ResponseAdminLeftMenu(html)
	ResponseAdminLeftMenu=html
End Function
'*********************************************************




'*********************************************************
' 目的：    html=Response_Plugin_Admin_Top
'*********************************************************
Function ResponseAdminTopMenu(html)
	Call Filter_Plugin_ResponseAdminTopMenu(html)
	ResponseAdminTopMenu=html
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
				allcomm.add CLng(objRS("comm_ID")),""
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

	If IsObject(Session("batch"))=False THen
		Set Session("batch")=CreateObject("Scripting.Dictionary")
		Session("batchorder")=0
		Session("batchtime")=0
	End If

	'检则是否未完成批操作,未完成的话不新增加批操作.
	If Session("batch").Count <> CLng(Session("batchorder")) Then Exit Function

	Dim i
	i=Session("batchorder")+1

	Session("batchorder")=i
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

		If Trim(BlogConfig.Meta.GetValue(BlogConfig.Meta.Names(i)))="" And InStr(strContent,""""& "<#"&BlogConfig.Meta.Names(i)&"#>" &"""")=0 Then
			strContent=Replace(strContent,"<#"&BlogConfig.Meta.Names(i)&"#>","Empty")
		Else
			strContent=Replace(strContent,"<#"&BlogConfig.Meta.Names(i)&"#>",Replace(BlogConfig.Meta.GetValue(BlogConfig.Meta.Names(i)),"""",""""""))
		End If

	Next

	If Instr(strContent,"<#ZC_BLOG_LANGUAGEPACK#>")>0 Then
		strContent=Replace(strContent,"<#ZC_BLOG_LANGUAGEPACK#>","SimpChinese")
		Call BlogConfig.Write("ZC_BLOG_LANGUAGEPACK","SimpChinese")
	End If

	If Instr(strContent,"<#ZC_COMMENT_EXCERPT_MAX#>")>0 Then
		strContent=Replace(strContent,"<#ZC_COMMENT_EXCERPT_MAX#>","20")
		Call BlogConfig.Write("ZC_COMMENT_EXCERPT_MAX","20")
	End If

	Call BlogConfig.Save()

	Call SaveToFile(BlogPath & "zb_users\c_option.asp",strContent,"utf-8",False)

End Function
'*********************************************************




'*********************************************************
' 目的：    日期类的简化函数 Regex
'*********************************************************
Function RegexbyDate(y,m,d)

	RegexbyDate=ZC_DATE_REGEX

End Function
'*********************************************************
'*********************************************************
' 目的：    日期类的简化函数 FullPath
'*********************************************************
Function FullPath(y,m,d)
	FullPath=ParseCustomDirectoryForPath(FullRegex,ZC_STATIC_DIRECTORY,"","",y,m,d,"","","")
End Function
'*********************************************************
'*********************************************************
' 目的：    日期类的简化函数 Url
'*********************************************************
Function UrlbyDate(y,m,d)

	UrlbyDate=ParseCustomDirectoryForUrl(RegexbyDate(y,m,d),ZC_STATIC_DIRECTORY,"","",y,m,d,"","","")

End Function
'*********************************************************
'*********************************************************
' 目的：    日期类的简化函数 Url auto
'*********************************************************
Function UrlbyDateAuto(y,m,d)
	Dim s
	If ZC_STATIC_MODE="MIX" Then
		s=ParseCustomDirectoryForUrl("{%host%}/catalog.asp?date={%year%}-{%month%}",ZC_STATIC_DIRECTORY,"","",y,m,d,"","","")
	Else
		s=UrlbyDate(y,m,d)
	End If
	If Right(s,12)="default.html" Then s=Left(s,Len(s)-12)
	UrlbyDateAuto=s
End Function
'*********************************************************




'*********************************************************
' 目的：  刷新c_option.asp至数据库
'*********************************************************
Function RefreshOptionFormFileToDB()
	On Error Resume Next

	If Not CheckUpdateDB("[fn_Source]","[blog_Function]") Then
		IF ZC_MSSQL_ENABLE=True Then
			objConn.execute("ALTER TABLE [blog_Function] ADD fn_Source nvarchar(50) default ''")
			objConn.execute("ALTER TABLE [blog_Function] ADD fn_ViewType nvarchar(50) default ''")
			objConn.execute("ALTER TABLE [blog_Function] ADD fn_IsHidden bit default 0")
		Else
			objConn.execute("ALTER TABLE [blog_Function] ADD COLUMN fn_Source VARCHAR(50) default """"")
			objConn.execute("ALTER TABLE [blog_Function] ADD COLUMN fn_ViewType VARCHAR(50) default """"")
			objConn.execute("ALTER TABLE [blog_Function] ADD COLUMN fn_IsHidden YESNO DEFAULT 0")
		End If

		objConn.execute("UPDATE [blog_Function] SET [fn_Source]='system' WHERE [fn_IsSystem]<>0")
		objConn.execute("UPDATE [blog_Function] SET [fn_Source]='users' WHERE [fn_IsSystem]=0")
		objConn.execute("UPDATE [blog_Function] SET [fn_IsHidden]=0")
		objConn.execute("UPDATE [blog_Function] SET [fn_ViewType]=''")
		objConn.execute("UPDATE [blog_Function] SET [fn_Meta]=''")

		IF ZC_MSSQL_ENABLE=True Then
			objConn.execute("ALTER TABLE [blog_Function] DROP COLUMN fn_IsSystem")
		Else
			objConn.execute("ALTER TABLE [blog_Function] DROP COLUMN fn_IsSystem")
		End If

		IF ZC_MSSQL_ENABLE=True Then
			objConn.execute("ALTER TABLE [blog_Function] ADD fn_IsHideTitle bit default 0")
		Else
			objConn.execute("ALTER TABLE [blog_Function] ADD COLUMN fn_IsHideTitle YESNO DEFAULT 0")
		End If
		objConn.execute("UPDATE [blog_Function] SET [fn_IsHideTitle]=0")
	End If

	Dim a,b
	b=LoadFromFile(BlogPath &"zb_users\c_option.asp","utf-8")
	For Each a In BlogConfig.Meta.Names
		If InStr(b,"Dim "& a)>0 Then
			Call Execute("Call BlogConfig.Write("""&a&""","&a&")")
		End If
	Next
	Call BlogConfig.Write("ZC_BLOG_VERSION",Replace(BlogVersions.GetValue(BlogVersions.Names(1)),"Z-Blog ",""))
	Call BlogConfig.Write("ZC_BLOG_CLSID",ZC_BLOG_CLSID_ORIGINAL)

	Call BlogConfig.Save()


	Call GetFunction()
	If Functions(FunctionMetas.GetValue("controlpanel")).Content="<span class=""cp-login""><a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=login"">[<#ZC_MSG009#>]</a></span>&nbsp;&nbsp;<span class=""cp-vrs""><a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=vrs"">[<#ZC_MSG021#>]</a></span>" Then
		Functions(FunctionMetas.GetValue("controlpanel")).Content="<span class=""cp-hello"">您好,欢迎到访网站!</span><br/><span class=""cp-login""><a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=login"">[<#ZC_MSG009#>]</a></span>&nbsp;&nbsp;<span class=""cp-vrs""><a href=""<#ZC_BLOG_HOST#>zb_system/cmd.asp?act=vrs"">[<#ZC_MSG021#>]</a></span>"
		Functions(FunctionMetas.GetValue("controlpanel")).Post()
	End If

	Err.Clear
End Function
'*********************************************************




'*********************************************************
Function GetVersionByBuild(b)

	Dim s
	b=CStr(b)

	If BlogVersions.Exists(b)=True Then
		s=BlogVersions.GetValue(b)
	Else
		s="Z-Blog 2.X Other Build " & b
	End If

	GetVersionByBuild=s

End Function
'*********************************************************




'*********************************************************
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
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function GetMetaValueWithForm(obj)

	Dim a,b

	For Each a In Request.Form
		If Left(a,5)="meta_" Then
			b=Mid(a,6,Len(a))
			Call obj.Meta.SetValue( b,Request.Form(a) )
		End If
	Next

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function GetFunctionByFileName(fn)

	Call GetFunction()

	Dim objRSsub
	Set objRSsub=objConn.Execute("SELECT * FROM [blog_Function] WHERE [fn_FileName]='"& fn &"'" )
	If (Not objRSsub.bof) And (Not objRSsub.eof) Then
		Set GetFunctionByFileName=Functions(objRSsub("fn_ID"))
	Else
		Set GetFunctionByFileName=New TFunction
	End If
	Set objRSsub=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：
' mode={[add][modif][del]}
' itemtype={page,cate,tags,等等}
'*********************************************************
Function AddNavBar(itemtype,id,name,url,mode)

	Call GetFunction()

	Dim s,s2,t,i,j,b

	b=False

	Dim re
	Set re = New RegExp
	re.Global = True
	re.Pattern = "href=""[^\""]+?"""
	re.IgnoreCase = True

	s=Functions(FunctionMetas.GetValue("navbar")).Content
	s2=s

	t=Split(s,"</li>")

	If UBound(t)>0 Then
		For i=0 To UBound(t)-1
			t(i)=t(i) & "</li>"
			j=t(i)
			If InStr(j,"menu-"&itemtype&"-"&id)>0 Then
				If InStr(mode,"[modif]")>0 Then
					j=re.Replace(j,"href="""&url&"""")
				End If
				b=True
			End If

			t(i)=j
		Next
	End If

	'ZC_NAVBAR_MENU_ITEM="<li id=""menu-%type-%id""><a href=""%url"">%name</a></li>"
	If b=False And InStr(mode,"[add]")>0 Then
		i=UBound(t)
		ReDim Preserve t(i+1)
		j=ZC_NAVBAR_MENU_ITEM
		j=Replace(j,"%type",itemtype)
		j=Replace(j,"%id",id)
		j=Replace(j,"%url",url)
		j=Replace(j,"%name",name)
		t(i)=j
	End If

	If InStr(mode,"[del]")>0 Then
		For i=0 To UBound(t)-1
			j=t(i)
			If InStr(j,"menu-"&itemtype&"-"&id)>0 Then
				j=""
			End If
			t(i)=j
		Next
	End If

	s=Join(t,"")

	Functions(FunctionMetas.GetValue("navbar")).Content=s
	Functions(FunctionMetas.GetValue("navbar")).Post()

	If s<>s2 Then
		Call SetBlogHint(Empty,Empty,True)
	End If

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function GetCateIDByNameAndAlias(Name,Alias)
	Dim aryCateInOrder : aryCateInOrder=GetCategoryOrder()
	Dim m,n
	For m=LBound(aryCateInOrder) To Ubound(aryCateInOrder)
		If Alias=Categorys(aryCateInOrder(m)).Alias Then
			GetCateIDByNameAndAlias=Categorys(aryCateInOrder(m)).ID
			Exit Function
		End If
		If Name=Categorys(aryCateInOrder(m)).Name Then
			GetCateIDByNameAndAlias=Categorys(aryCateInOrder(m)).ID
			Exit Function
		End If
	Next
	GetCateIDByNameAndAlias=Empty
End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function CookiesPath()

	Dim s
	s=BlogHost
	s=Right(BlogHost,Len(BlogHost)-InStr(BLogHost,"//")-1)
	s=Replace(s,Request.ServerVariables("HTTP_HOST"),"")
	If s="" Then s="/"

	CookiesPath=s

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function CheckUndefined()

	On Error Resume Next
	Dim a
	a=UCase(LoadFromFile(BlogPath &"zb_users\c_option.asp","utf-8"))

	If Trim(a)="" Then Exit Function

	If InStr(a,"DIM ZC_SYNTAXHIGHLIGHTER_ENABLE")=0 Then
		Call ExecuteGlobal("ZC_SYNTAXHIGHLIGHTER_ENABLE=True")
	End If

	If InStr(a,"DIM ZC_CODEMIRROR_ENABLE")=0 Then
		Call ExecuteGlobal("ZC_CODEMIRROR_ENABLE=True")
	End If

	If InStr(a,"DIM ZC_ARTICLE_EXCERPT_MAX")=0 Then
		Call ExecuteGlobal("ZC_ARTICLE_EXCERPT_MAX=250")
	End If

	If InStr(a,"DIM ZC_POST_STATIC_MODE")=0 Then
		Call ExecuteGlobal("ZC_POST_STATIC_MODE=""STATIC""")
	End If

	If InStr(a,"DIM ZC_HTTP_LASTMODIFIED")=0 Then
		Call ExecuteGlobal("ZC_HTTP_LASTMODIFIED=False")
	End If

	If InStr(a,"DIM ZC_PERMANENT_DOMAIN_ENABLE")=0 Then
		Call ExecuteGlobal("ZC_PERMANENT_DOMAIN_ENABLE=False")
	End If

	If InStr(a,"DIM ZC_DEFAULT_PAGES_TEMPLATE")=0 Then
		Call ExecuteGlobal("ZC_DEFAULT_PAGES_TEMPLATE=""""")
	End If

	If InStr(a,"DIM ZC_SIDEBAR_ORDER")=0 Then
		Call ExecuteGlobal("ZC_SIDEBAR_ORDER="""&ZC_DEFAULT_SIDEBAR&"""")
	End If

	If InStr(a,"DIM ZC_SIDEBAR_ORDER2")=0 Then
		Call ExecuteGlobal("ZC_SIDEBAR_ORDER2=""""")
	End If

	If InStr(a,"DIM ZC_SIDEBAR_ORDER3")=0 Then
		Call ExecuteGlobal("ZC_SIDEBAR_ORDER3=""""")
	End If

	If InStr(a,"DIM ZC_SIDEBAR_ORDER4")=0 Then
		Call ExecuteGlobal("ZC_SIDEBAR_ORDER4=""""")
	End If

	If InStr(a,"DIM ZC_SIDEBAR_ORDER5")=0 Then
		Call ExecuteGlobal("ZC_SIDEBAR_ORDER5=""""")
	End If

	If InStr(a,"DIM ZC_BLOG_LANGUAGEPACK")=0 Then
		Call ExecuteGlobal("ZC_BLOG_LANGUAGEPACK=""SimpChinese""")
	End If


	If InStr(a,"DIM ZC_COMMENT_EXCERPT_MAX")=0 Then
		Call ExecuteGlobal("ZC_COMMENT_EXCERPT_MAX=20")
	End If


End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function GetBlogVersion()
	Dim s
	s=ZC_BLOG_VERSION
	s=Right(s,6)
	s=Trim(s)
	s=CLng(s)
	GetBlogVersion=s
End Function
'*********************************************************




'*********************************************************
' 目的：为主题提供的便捷函数,可以生成自己的模块
' 参数:主题ID,模块名,模块ID(文件名),模块HtmlID,模块类型(div/ul),模块Maxli(默认0),模块内容
'*********************************************************
Function AddThemeFunction(ThemeID,FunctionName,FunctionFileName,FunctionHtmlID,FunctionType,FunctionMaxLi,FunctionHideTitle,FunctionContent)

	Dim objFunction
	Set objFunction=GetFunctionByFileName(FunctionFileName)

	objFunction.Name=FunctionName
	objFunction.FileName=FunctionFileName
	objFunction.HtmlID=FunctionHtmlID
	objFunction.Ftype=FunctionType
	objFunction.MaxLi=FunctionMaxLi
	objFunction.Content=FunctionContent
	objFunction.IsHideTitle=CBool(FunctionHideTitle)
	objFunction.Source="theme_"& ThemeID

	If objFunction.Post Then
		AddThemeFunction=True
	End If
	Set objFunction=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：为插件...
' 参数:插件ID,模块名,模块ID(文件名),模块HtmlID,模块类型(div/ul),模块Maxli(默认0),模块内容
'*********************************************************
Function AddPluginFunction(PluginID,FunctionName,FunctionFileName,FunctionHtmlID,FunctionType,FunctionMaxLi,FunctionHideTitle,FunctionContent)

	Dim objFunction
	Set objFunction=GetFunctionByFileName(FunctionFileName)

	objFunction.Name=FunctionName
	objFunction.FileName=FunctionFileName
	objFunction.HtmlID=FunctionHtmlID
	objFunction.Ftype=FunctionType
	objFunction.MaxLi=FunctionMaxLi
	objFunction.Content=FunctionContent
	objFunction.IsHideTitle=CBool(FunctionHideTitle)
	objFunction.Source="plugin_"& PluginID

	If objFunction.Post Then
		AddPluginFunction=True
	End If
	Set objFunction=Nothing

End Function
'*********************************************************




Function CreateValidCode(strVerifyNumber)
	Call Filter_Plugin_ValidCode_Create(strVerifyNumber)
End Function


Function CheckValidCode(strInputNumber)
	CheckValidCode=Filter_Plugin_ValidCode_Check(strInputNumber)
End Function


%>

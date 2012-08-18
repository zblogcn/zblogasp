<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8及以上的版本
'// 插件制作:  zblog管理员之家(www.zbadmin.com)
'// 备    注:   Mini缩略图插件代码
'// 最后修改：   2012/2/20
'// 最后版本:    0.1
'///////////////////////////////////////////////////////////////////////////////
%>
<!-- #include file="Function.asp" -->
<%
Dim objConfig
Dim MiniTu_MiniImgWidth
Dim MiniTu_MiniImgHeight
Dim MiniTu_NoHtmlIntro
Dim MiniTu_SearchInContent

Sub MiniTu_Initialize()
	Set objConfig=new TConfig
	objConfig.Load "MiniTu"
	MiniTu_MiniImgWidth=CInt(objConfig.Read("MiniImgWidth"))
	MiniTu_MiniImgHeight=CInt(objConfig.Read("MiniImgHeight"))
	MiniTu_NoHtmlIntro=CBool(objConfig.Read("NoHtmlIntro"))
	MiniTu_SearchInContent=CBool(objConfig.Read("SearchInContent"))
End Sub
'=======================================================
'注册插件并挂接口
'=======================================================
Call RegisterPlugin("MiniTu","ActivePlugin_MiniTu")

Function ActivePlugin_MiniTu()

	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Export_TemplateTags","MiniTu_Core")
	Call Add_Filter_Plugin("Filter_Plugin_TUpLoadFile_Del","MiniTu_Del")
	'Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Core","MiniTu_Filter")

	Call Add_Action_Plugin("Action_Plugin_Searching_Begin","Call MiniTu_Search:Response.End")

End Function


'=======================================================
'为文章增加标签
'=======================================================
Function MiniTu_Core(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)

		If MiniTu_NoHtmlIntro Then
			aryTemplateTagsValue(4)=TransferHTML(aryTemplateTagsValue(4),"[nohtml]")
		End If

		Dim c:c=UBOUND(aryTemplateTagsName)+1

		ReDim Preserve aryTemplateTagsName(c)
		ReDim Preserve aryTemplateTagsValue(c)

		Dim strImgTag
		strImgTag=MiniTu_Build(aryTemplateTagsValue(3),aryTemplateTagsValue(5),aryTemplateTagsValue(11))

		aryTemplateTagsName(c)="article/intro/minitu"
		aryTemplateTagsValue(c)=strImgTag

End Function


'=======================================================
'函数: 从正文中提取图片路径.
'输入: 文章全文.
'返回: 有图则返回图片路径, 无图返回空.
'=======================================================
Function MiniTu_OriginalURL(ByVal strContent)
	'On Error Resume Next

	Dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase=True
	objRegExp.Global=False

	objRegExp.Pattern="(<img[^>]+src[^""]+"")([^""]+)([^>]+>)"

	Dim Match, Matches, Value
	Set Matches=objRegExp.Execute(strContent)
		For Each Match in Matches
			Value=objRegExp.Replace(Match.value,"$2")
		Next
	Set Matches=Nothing

	Set objRegExp=Nothing

	MiniTu_OriginalURL=Value

	'Err.Clear
End Function


'=======================================================
'函数: 生成缩略图.
'输入: 图片路径.
'返回: 有图则返回图片路径, 无图返回空.
'=======================================================
Function MiniTu_MiniURL(ByVal strUrl)

	Dim strOriginalPath,strFileName,strMiniPath

	If strUrl="" Then
		MiniTu_MiniURL=""
	Else
		If InStr(LCase(strUrl),LCase(GetCurrentHost & ZC_UPLOAD_DIRECTORY))>0 Then
			strOriginalPath=Replace(LCase(strUrl),LCase(GetCurrentHost),BlogPath)
			strOriginalPath=MiniTu_URLDecode(strOriginalPath)
			strFileName=LCase(Mid(strOriginalPath,InStrRev(Replace(strOriginalPath,"/","\"),"\")+1))
			strMiniPath=Left(strOriginalPath,InStrRev(strOriginalPath,".")-1) & "_mini.jpg"
			MiniTu_MiniURL=Replace(strMiniPath,BlogPath,GetCurrentHost)
			Call MiniTu_CreatMini(strOriginalPath,strMiniPath)
		Else
			MiniTu_MiniURL=strUrl
		End If
	End If

End Function


'=======================================================
'函数: 生成链接, 含无图校验.
'输入: 图片路径.
'返回: 返回含有图片的链接.
'=======================================================
Function MiniTu_Build(ByVal strTitle,ByVal strContent,ByVal strHref)
	MiniTu_Initialize
	Dim strUrl

	strUrl=MiniTu_OriginalURL(strContent)
	strUrl=MiniTu_MiniURL(strUrl)

	If strUrl="" Then
		strUrl=GetCurrentHost & "zb_users/plugin/MiniTu/noimg.jpg"
	End If

	MiniTu_Build="<a href="""& strHref &""" target=""_blank"" title="""& strTitle &"""><img src="""& strUrl &""" alt="""& strTitle &""" /></a>"

End Function


'=======================================================
'删除附件时删除缩略图
'=======================================================
Function MiniTu_Del(byval ID,byval AuthorID,byval FileSize,byval FileName,byval PostTime,byval DirByTime)
	'On Error Resume Next
	MiniTu_Initialize
	Call CheckParameter(ID,"int",0)

	Dim objRS,strFilePath

	Set objRS=objConn.Execute("SELECT * FROM [blog_UpLoad] WHERE [ul_ID] = " & ID)

	If (Not objRS.bof) And (Not objRS.eof) Then

		If objRS("ul_DownNum")=-1 Then
			objConn.Execute("UPDATE [blog_Upload] SET [ul_DownNum]=(0) WHERE [ul_DownNum]=("& ID &")")
		End If

		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")

		strFilePath = BlogPath & "/"& ZC_UPLOAD_DIRECTORY &"/" & objRS("ul_FileName")
		strFilePath=Left(strFilePath,InStrRev(strFilePath,".")-1) & "_mini.jpg"
		If fso.FileExists( strFilePath ) Then
			fso.DeleteFile( strFilePath )
		End If

		strFilePath = BlogPath & "/"& ZC_UPLOAD_DIRECTORY & "/" & Year(objRS("ul_PostTime")) & "/" & Month(objRS("ul_PostTime")) &"/" & objRS("ul_FileName")
		strFilePath=Left(strFilePath,InStrRev(strFilePath,".")-1) & "_mini.jpg"
		If fso.FileExists( strFilePath ) Then
			fso.DeleteFile( strFilePath )
		End If

		Set fso = Nothing

	Else

		Exit Function

	End If

	objRS.Close
	Set objRS=Nothing

	'Err.Clear
End Function

'=======================================================
'过滤FCK相对路径, 换成绝对路径.
'=======================================================
'Function MiniTu_Filter(ByRef objArticle)
'	objArticle.Content=Replace(objArticle.Content,"../../../",GetCurrentHost)
'	objArticle.Intro=Replace(objArticle.Intro,"../../../",GetCurrentHost)
'End Function


'=======================================================
'安装插件
'=======================================================
Function MiniTu_Search()
	MiniTu_Initialize
	'检查权限
	If Not CheckRights("Search") Then Call ShowError(6)

	TemplateTagsDic.Item("ZC_BLOG_HOST")=GetCurrentHost()

'检查权限
	If Not CheckRights("Search") Then Call ShowError(6)

	Dim strQuestion
	strQuestion=TransferHTML(Request.QueryString("q"),"[nohtml]")

	Dim objArticle
	Set objArticle=New TArticle
	objArticle.LoadCache
	strQuestion=Trim(strQuestion)
	strQuestion=FilterSQL(strQuestion)
	Dim i
	Dim j
	Dim s
	Dim aryArticleList()

	Dim objRS
	Dim intPageCount
	Dim objSubArticle

	Dim cate
	If IsEmpty(Request.QueryString("cate"))=False Then
	cate=CInt(Request.QueryString("cate"))
	End If

	strQuestion=FilterSQL(strQuestion)
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn


	Dim sql
	If Not Len(strQuestion)=0 Then

		sql="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_ID]>0) AND ([log_Level]>2)"

		If ZC_MSSQL_ENABLE=False Then
			sql=sql& "AND( (InStr(1,LCase([log_Title]),LCase('"&strQuestion&"'),0)<>0)  OR (InStr(1,LCase([log_Content]),LCase('"&strQuestion&"'),0)<>0) )"
		Else
			sql=sql& "AND( (CHARINDEX('"&strQuestion&"',[log_Title])<>0) OR (CHARINDEX('"&strQuestion&"',[log_Content])<>0) )"
		End If
		If MiniTu_SearchInContent Then
			If ZC_MSSQL_ENABLE=False Then
				sql=sql & "OR (InStr(1,LCase([log_Intro]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('"&strQuestion&"'),0)<>0) "
			Else
				sql=sql & " OR (CHARINDEX('"&strQuestion&"',[log_Intro])<>0) "

			End If
		End If
		sql=sql & " ORDER BY [log_PostTime] DESC,[log_ID] DESC"
		objRs.Source=sql
		objRS.Open()
		s=Replace(Replace(ZC_MSG086,"%s","<strong>" & TransferHTML(Replace(strQuestion,Chr(39)&Chr(39),Chr(39),1,-1,0),"[html-format]") & "</strong>",vbTextCompare,1),"%s","<strong>" & objRS.RecordCount & "</strong>",1,-1,0)

		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = ZC_SEARCH_COUNT
			intPageCount=objRS.PageCount
			objRS.AbsolutePage = 1

			For i = 1 To objRS.PageSize

				ReDim Preserve aryArticleList(i)

				Set objSubArticle=New TArticle
				If objSubArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then
					objSubArticle.SearchText=Request.QueryString("q")
					If objSubArticle.Export(ZC_DISPLAY_MODE_SEARCH)= True Then
						aryArticleList(i)=objSubArticle.subhtml
					End If
				End If
				Set objSubArticle=Nothing
		
				objRS.MoveNext
				If objRS.EOF Then Exit For
		
			Next

		Else
			ReDim Preserve aryArticleList(0)
		End If

		objRS.Close()
		Set objRS=Nothing

		objArticle.FType=ZC_POST_TYPE_PAGE
		objArticle.Content=Join(aryArticleList)
		
		objArticle.Title=ZC_MSG085 + ":" + TransferHTML(strQuestion,"[html-format]")
		
		If objArticle.Export(ZC_DISPLAY_MODE_SYSTEMPAGE) Then
			objArticle.Build
			Response.Write objArticle.html
		End If
	End If
End Function

'=======================================================
'安装插件
'=======================================================
Function InstallPlugin_MiniTu

	
	MiniTu_Initialize
	If objConfig.Exists("a")=False Then
		objConfig.Write "a","0.2"
		objConfig.Write "MiniImgWidth",300
		objConfig.Write "MiniImgHeight","0"
		objConfig.Write "NoHtmlIntro",True
		objConfig.Write "SearchInContent",True
		objConfig.Save
	End If
	Call SetBlogHint(True,Empty,True)
	
End Function

'=======================================================
'卸载插件
'=======================================================
Function UnInstallPlugin_MiniTu

	Call SetBlogHint(True,Empty,True)

End Function

%>
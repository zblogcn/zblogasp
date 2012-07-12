<!-- #include file="config.asp" -->
<!-- #include file="function.asp" -->
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Arwen 其它版本的Z-blog未知
'// 插件制作:    haphic(http://www.esloy.com/)
'// 备    注:    STACentre - 挂口函数页
'// 最后修改：   2011-5-1
'// 最后版本:    1.x
'///////////////////////////////////////////////////////////////////////////////

'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************

'注册插件
Call RegisterPlugin("STACentre","ActivePlugin_STACentre")

'挂口部分
Function ActivePlugin_STACentre()

	'插入菜单
	Call Add_Response_Plugin("Response_Plugin_AskFileReBuild_SubMenu",MakeSubMenu("静态化中心",ZC_BLOG_HOST&"ZB_USERS/PLUGIN/STACentre/main.asp","m-left",False))

	'后台管理部分
	'Filter_Plugin_TArticle_xxx
	Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Succeed","STACentre_BuildPageByArticleEdt")
	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Del","STACentre_BuildPageByArticleDel")
	'Filter_Plugin_PostCategory_Succeed
	Call Add_Filter_Plugin("Filter_Plugin_PostCategory_Succeed","STACentre_BuildPageByCateEdt")
	Call Add_Filter_Plugin("Filter_Plugin_TCategory_Del","STACentre_BuildPageByCateDel")
	'Filter_Plugin_TTag_xxx
	Call Add_Filter_Plugin("Filter_Plugin_PostTag_Succeed","STACentre_BuildPageByTagEdt")
	Call Add_Filter_Plugin("Filter_Plugin_TTag_Del","STACentre_BuildPageByTagDel")
	'Filter_Plugin_TUser_xxx
	Call Add_Filter_Plugin("Filter_Plugin_EditUser_Succeed","STACentre_BuildPageByUserEdt")
	Call Add_Filter_Plugin("Filter_Plugin_TUser_Del","STACentre_BuildPageByUserDel")

	'前台跳转
	'Action_Plugin_Catalog_Begin
	'Call Add_Action_Plugin("Action_Plugin_Catalog_Begin","STACentre_DynamicToStatic")


End Function
'*********************************************************




'*********************************************************
' 挂口: 直接挂入的函数
'*********************************************************

'操作文章时重建
Function STACentre_BuildPageByArticleEdt(ByRef objArticle)

	Call STACentre_BuildPageByCateID(objArticle.cateID,True)
	Call STACentre_BuildPageByTagCode(objArticle.Tag,True)
	Call STACentre_BuildPageByAuthorID(objArticle.AuthorID,True)
	Call STACentre_BuildPageByPostTime(objArticle.PostTime,True)

End Function

Function STACentre_BuildPageByArticleDel(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef IsAnonymous,ByRef MetaString)

	Call STACentre_BuildPageByCateID(cateID,True)
	Call STACentre_BuildPageByTagCode(Tag,True)
	Call STACentre_BuildPageByAuthorID(AuthorID,True)
	Call STACentre_BuildPageByPostTime(PostTime,True)

End Function


'操作分类时重建
Function STACentre_BuildPageByCateEdt(ByRef objCategory)

	Call STACentre_BuildPageByCateID(objCategory.ID,False)
	Call GetCategory()
	Call STACentre_BuildPageByCateID(objCategory.ID,True)

End Function

Function STACentre_BuildPageByCateDel(ByRef ID,ByRef Name,ByRef Intro,ByRef Order,ByRef Count,ByRef ParentID,ByRef Alias,ByRef TemplateName,ByRef MetaString)

	Call STACentre_BuildPageByCateID(ID,False)

End Function

'操作标签时重建
Function STACentre_BuildPageByTagEdt(ByRef objTag)

	Call STACentre_BuildPageByTagID(objTag.ID,False)
	Call GetTags()
	Call STACentre_BuildPageByTagID(objTag.ID,True)

End Function

Function STACentre_BuildPageByTagDel(ByRef ID,ByRef Name,ByRef Intro,ByRef Order,ByRef Count,ByRef ParentID,ByRef Alias,ByRef TemplateName,ByRef MetaString)

	Call STACentre_BuildPageByTagID(ID,False)

End Function

'操作用户时重建
Function STACentre_BuildPageByUserEdt(ByRef objUser)

	Call STACentre_BuildPageByAuthorID(objUser.ID,False)
	Call GetUser()
	Call STACentre_BuildPageByAuthorID(objUser.ID,True)

End Function

Function STACentre_BuildPageByUserDel(ByRef ID,ByRef Name,ByRef Level,ByRef Password,ByRef Email,ByRef HomePage,ByRef Count,ByRef Alias,ByRef MetaString,ByRef currentUser)

	Call STACentre_BuildPageByAuthorID(ID,False)

End Function

'*********************************************************




'*********************************************************
' 外延部分: 和输出直接相关
'*********************************************************

'根据文章 Catelog 字段建立页面
Function STACentre_BuildPageByCateID(ByVal strID,ByVal Build)

	Dim objCate
	Set objCate=New STACentre_Categorys
	If objCate.LoadInfoByID(strID) Then
		If Build Then
			objCate.Build
		Else
			objCate.Del
		End If
	End If
	Set objCate=Nothing

End Function

'根据 Tag 字段建立页面
Function STACentre_BuildPageByTagID(ByVal strID,ByVal Build)

	Dim objTag
	Set objTag=New STACentre_Tags
	If objTag.LoadInfoById(strID) Then
		If Build Then
			objTag.Build
		Else
			objTag.Del
		End If
	End If
	Set objTag=Nothing

End Function

'根据文章 Tag 字段建立页面
Function STACentre_BuildPageByTagCode(ByVal strTagCode,ByVal Build)
	'On Error Resume Next

	Dim t,s,i
	Dim objTag

	If strTagCode<>"" Then
		s=strTagCode
		s=Replace(s,"}","")
		t=Split(s,"{")

		For i=LBound(t) To UBound(t)
			If t(i)<>"" Then
				Set objTag=New STACentre_Tags
				If objTag.LoadInfoById(t(i)) Then
					If Build Then
						objTag.Build
					Else
						objTag.Del
					End If
				End If
				Set objTag=Nothing
			End If
		Next
	End If

	Err.Clear
End Function

'根据文章 Author 字段建立页面
Function STACentre_BuildPageByAuthorID(ByVal strID,ByVal Build)

	Dim objAuth
	Set objAuth=New STACentre_Authors
	If objAuth.LoadInfoByID(strID) Then
		If Build Then
			objAuth.Build
		Else
			objAuth.Del
		End If
	End If
	Set objAuth=Nothing

End Function

'根据文章 PostTime 字段建立页面
Function STACentre_BuildPageByPostTime(ByVal strPostTime,ByVal Build)

	Dim objPostTime
	Set objPostTime=New STACentre_Archives
	If objPostTime.LoadInfoByID(Year(strPostTime)&"-"&Month(strPostTime)) Then
		If Build Then
			objPostTime.Build
		Else
			objPostTime.Del
		End If
	End If
	Set objPostTime=Nothing

End Function

'*********************************************************
'动态 Tag 地址 301 跳转至静态
Function STACentre_DynamicToStatic()

	If Not IsEmpty(Request.QueryString("tags")) Then
		Dim strTagName
		Dim strTagID
		Dim strTagURL
		strTagName=Request.QueryString("tags")
		strTagID=STACentre_TransToTagID(strTagName)
		strTagURL=STACentre_TransToTagURL(strTagID)
		RedirectBy301(strTagURL)
		Response.End
	End If

End Function

'*********************************************************

'*********************************************************




'*********************************************************
' 挂口: 激活和停用
'*********************************************************

'安装插件
Function InstallPlugin_STACentre()
	On Error Resume Next

	Dim fso_chkright_1 : fso_chkright_1=False
	Dim fso_chkright_2 : fso_chkright_2=False

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
		fso.CreateFolder(BlogPath & "esloy")
		Call SaveToFile(BlogPath & "esloy/haphic.html","test by haphic","utf-8",False)
		If LoadFromFile(BlogPath & "esloy/haphic.html","utf-8")="test by haphic" Then fso_chkright_1=True
		fso.DeleteFolder(BlogPath & "esloy")
		If fso.FolderExists(BlogPath & "esloy")=False Then fso_chkright_2=True

		If fso_chkright_1 And fso_chkright_2 Then
			If (fso.FolderExists(BlogPath & STACentre_Dir_Primary)=False) Then fso.CreateFolder(BlogPath & STACentre_Dir_Primary)
			Call SetBlogHint_Custom("&raquo; 文本读写和目录操作权限正常, 插件可以使用.")
			Call SetBlogHint(True,Empty,True)
		Else
			Call SetBlogHint_Custom("&raquo; 文本读写和目录操作权限可能不正常, 插件可能存在使用问题, 请联系主机客服调整配置或更换主机.")
			Call SetBlogHint(False,Empty,Empty)
		End If
	Set fso = Nothing

	Err.Clear
End Function

'卸载插件
Function UnInstallPlugin_STACentre()
	On Error Resume Next

	Call STACentre_ClearAllDirsByHistory()

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(Server.MapPath("progress.txt")) Then
		fso.DeleteFile(Server.MapPath("progress.txt"))
	End If
	Set fso = Nothing

	Call SetBlogHint(True,Empty,True)

	Err.Clear
End Function
'*********************************************************
%>
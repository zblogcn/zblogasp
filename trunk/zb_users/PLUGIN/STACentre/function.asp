<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Arwen 其它版本的Z-blog未知
'// 插件制作:    haphic(http://www.esloy.com/)
'// 备    注:    STACentre - 函数库
'// 最后修改：   2011-5-1
'// 最后版本:    1.x
'///////////////////////////////////////////////////////////////////////////////


'*********************************************************
' 基本部分: 路径解析与文件生成等
'*********************************************************

'将设置解析为路径
Function STACentre_ParseCustomDirectory(strRegex,strPost,strType,strID,strName,strAlias)

	Dim s
	s=strRegex

	s=Replace(s,"{%post%}",strPost)
	s=Replace(s,"{%type%}",strType)
	s=Replace(s,"{%id%}",strID)
	s=Replace(s,"{%name%}",strName)
	s=Replace(s,"{%alias%}",strAlias)

	STACentre_ParseCustomDirectory=s

End Function


'得到文件路径(不含文件名)
Function STACentre_GetStaticDirectory(strRegex,strType,strID,strName,strAlias)

	Dim strDirectory
	strDirectory=STACentre_ParseCustomDirectory(strRegex,ZC_STATIC_DIRECTORY,strType,strID,strName,strAlias)
	strDirectory=Replace(strDirectory,"/","\")
	If Right(BlogPath & strDirectory,1)<>"\" Then
		strDirectory=strDirectory & "\"
	End If
	STACentre_GetStaticDirectory=strDirectory

End Function
'*********************************************************




'*********************************************************
' 基本部分: 得到日期数组
'*********************************************************
Function STACentre_GetArchivesList()

	Dim i
	Dim j
	Dim l
	Dim n
	Dim objRS

	'Archives
	Dim aryArchives()
	Set objRS=objConn.Execute("SELECT * FROM [blog_Article] WHERE ([log_Level]>1) ORDER BY [log_PostTime] DESC")
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

	If Not IsEmpty(dtmYM) Then
		For i=1 to UBound(dtmYM)

			l=Year(dtmYM(i))
			n=Month(dtmYM(i))+1
			IF n>12 Then l=l+1:n=1

			Set objRS=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE ([log_Level]>1) AND [log_PostTime] BETWEEN "& ZC_SQL_POUND_KEY & Year(dtmYM(i)) &"-"& Month(dtmYM(i)) &"-1"& ZC_SQL_POUND_KEY &" AND "& ZC_SQL_POUND_KEY & l &"-"& n &"-1" & ZC_SQL_POUND_KEY)

			If (Not objRS.bof) And (Not objRS.eof) Then

				ReDim Preserve aryArchives(i)

				aryArchives(i)=Year(dtmYM(i)) & "-" & Month(dtmYM(i))

			End If

			objRS.Close
			Set objRS=Nothing
		Next
	End If


	STACentre_GetArchivesList=aryArchives

	Erase aryArchives

End Function
'*********************************************************




'*********************************************************
' 核心部分: 核心类
'*********************************************************

'Categorys
Class STACentre_Categorys

	Public ID
	Public Name
	Public Alias

	Public FID
	Public FName
	Public FAlias

	Public PageContent

	Private Property Get sType
		sType="categorys"
	End Property

	Private Property Get Enable
		Enable=var_STACentre_Dir_Categorys_Enable
		If IsEmpty(Enable) Or Enable="" Then
			Enable=STACentre_Dir_Categorys_Enable
		End If
	End Property

	Private Property Get Regex
		Regex=var_STACentre_Dir_Categorys_Regex
		If IsEmpty(Regex) Or Regex="" Then
			Regex=STACentre_Dir_Categorys_Regex
		End If
	End Property

	Private Property Get Anonymous
		Anonymous=var_STACentre_Dir_Categorys_Anonymous
		If IsEmpty(Anonymous) Or Anonymous="" Then
			Anonymous=STACentre_Dir_Categorys_Anonymous
		End If
	End Property

	Private Property Get FCate
		FCate=var_STACentre_Dir_Categorys_FCate
		If IsEmpty(FCate) Or FCate="" Then
			FCate=STACentre_Dir_Categorys_FCate
		End If
	End Property

	Public Property Get Filename
		Filename = Alias
		If Anonymous Then Filename="default"
		FileName = FileName & "." & ZC_STATIC_TYPE
	End Property

	Public Property Get Directory
		If FCate Then
			Directory=STACentre_GetStaticDirectory(Regex,sType,FID,FName,FAlias)
		Else
			Directory=STACentre_GetStaticDirectory(Regex,sType,ID,Name,Alias)
		End If
	End Property

	Public Property Get FullPath
		FullPath=BlogPath & Directory & Filename
	End Property

	Public Property Get Url
		If Anonymous Then
			Url=ZC_BLOG_HOST & Directory
		Else
			Url=ZC_BLOG_HOST & Directory & Filename
		End If
		If Not Enable Then Url=ZC_BLOG_HOST & "catalog.asp?cate=" & ID
		Url=Replace(Url,"\","/")
	End Property

	Public Function LoadInfoByID(ByVal intID)
		Call CheckParameter(intID,"int",0)
		If intID<=UBound(Categorys) Then
			If IsObject(Categorys(intID)) Then
				ID=Categorys(intID).ID
				Name=Categorys(intID).Name
				Alias=Categorys(intID).Alias

				If IsNull(Alias) Or IsEmpty(Alias) Or Alias="" Then Alias=ID

				FID=ID
				FName=Name
				FAlias=Alias

				If STACentre_Dir_Categorys_FCate Then
					If Categorys(intID).ParentID>0 Then
						Dim intFID : intFID=Categorys(intID).ParentID
						If intFID<=UBound(Categorys) Then
							If IsObject(Categorys(intFID)) Then
								FID=Categorys(intFID).ID
								FName=Categorys(intFID).Name
								FAlias=Categorys(intFID).Alias
								If IsNull(FAlias) Or IsEmpty(FAlias) Or FAlias="" Then FAlias=FID

								FID=FID & "/" & ID
								FName=FName & "/" & Name
								FAlias=FAlias & "/" & Alias
							End If
						End If
					End If
				End If

				Dim ArtList
				Set ArtList=New TArticleList
				ArtList.LoadCache
				ArtList.template="CATALOG"
				If ArtList.ExportByMixed(Empty,ID,Empty,Empty,Empty,ZC_DISPLAY_MODE_ALL) Then
					ArtList.Build
					PageContent=ArtList.html
					LoadInfoByID=True
				End If
				Set ArtList=Nothing
			End If
		End If
	End Function

	Public Function LoadInfoByName(ByVal strName)
		LoadInfoByName=False
	End Function

	Public Function Build
		If Enable Then
			Call STACentre_AddFolderRecord(Directory)
			Call CreatDirectoryByCustomDirectory(Directory)
			Call STACentre_AddDirectoryRecord("file",FullPath)
			Call SaveToFile(FullPath,PageContent,"utf-8",True)
		End If
	End Function

	Public Function Del
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
			If fso.FileExists(FullPath) Then
				fso.DeleteFile(FullPath)
				Call STACentre_DelDirectoryRecord("file",FullPath)
			End If
		Set fso = Nothing
	End Function

	Private Sub class_initialize()
	End Sub

	Private sub class_terminate()
	End Sub

End Class


'Tags
Class STACentre_Tags

	Public ID
	Public Name
	Public Alias
	Public PageContent

	Private Property Get sType
		sType="tags"
	End Property

	Private Property Get Enable
		Enable=var_STACentre_Dir_Tags_Enable
		If IsEmpty(Enable) Or Enable="" Then
			Enable=STACentre_Dir_Tags_Enable
		End If
	End Property

	Private Property Get Regex
		Regex=var_STACentre_Dir_Tags_Regex
		If IsEmpty(Regex) Or Regex="" Then
			Regex=STACentre_Dir_Tags_Regex
		End If
	End Property

	Private Property Get Anonymous
		Anonymous=var_STACentre_Dir_Tags_Anonymous
		If IsEmpty(Anonymous) Or Anonymous="" Then
			Anonymous=STACentre_Dir_Tags_Anonymous
		End If
	End Property

	Public Property Get Filename
		Filename = Alias
		Filename = Replace(Filename,"　"," ")
		Filename = Replace(Filename," ","-")
		If Anonymous Then Filename="default"
		FileName = FileName & "." & ZC_STATIC_TYPE
	End Property

	Public Property Get Directory
		Directory=STACentre_GetStaticDirectory(Regex,sType,ID,Name,Alias)
	End Property

	Public Property Get FullPath
		FullPath=BlogPath & Directory & Filename
	End Property

	Public Property Get Url
		If Anonymous Then
			Url=ZC_BLOG_HOST & Directory
		Else
			Url=ZC_BLOG_HOST & Directory & Filename
		End If
		If Not Enable Then Url=ZC_BLOG_HOST & "catalog.asp?"& "tags=" & Server.URLEncode(Name)
		Url=Replace(Url,"\","/")
	End Property

	Public Function LoadInfoByID(ByVal intID)
		Call CheckParameter(intID,"int",0)
		If intID<=UBound(Tags) Then
			If IsObject(Tags(intID)) Then
				ID=Tags(intID).ID
				Name=Tags(intID).Name
				Alias=Tags(intID).Alias

				If IsNull(Alias) Or IsEmpty(Alias) Or Alias="" Then Alias=Name

				Dim ArtList
				Set ArtList=New TArticleList
				ArtList.LoadCache
				ArtList.template="CATALOG"
				If ArtList.ExportByMixed(Empty,Empty,Empty,Empty,Name,ZC_DISPLAY_MODE_ALL) Then
					ArtList.Build
					PageContent=ArtList.html
					LoadInfoByID=True
				End If
				Set ArtList=Nothing
			End If
		End If
	End Function

	Public Function LoadInfoByName(ByVal strName)
		Dim s,t
		t=0
		For Each s in Tags
			If IsObject(s) Then
				If UCase(s.Name)=UCase(strName) Then
					t=s.ID
					Exit For
				End If
			End If
		Next
		If Not t=0 Then
			If LoadInfoByID(t) Then
				LoadInfoByName=True
			End If
		End If
	End Function

	Public Function Build
		If Enable Then
			Call STACentre_AddFolderRecord(Directory)
			Call CreatDirectoryByCustomDirectory(Directory)
			Call STACentre_AddDirectoryRecord("file",FullPath)
			Call SaveToFile(FullPath,PageContent,"utf-8",True)
		End If
	End Function

	Public Function Del
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
			If fso.FileExists(FullPath) Then
				fso.DeleteFile(FullPath)
				Call STACentre_DelDirectoryRecord("file",FullPath)
			End If
		Set fso = Nothing
	End Function

	Private Sub class_initialize()
	End Sub

	Private sub class_terminate()
	End Sub

End Class

'Authors
Class STACentre_Authors

	Public ID
	Public Name
	Public Alias
	Public PageContent

	Private Property Get sType
		sType="Authors"
	End Property

	Private Property Get Enable
		Enable=var_STACentre_Dir_Authors_Enable
		If IsEmpty(Enable) Or Enable="" Then
			Enable=STACentre_Dir_Authors_Enable
		End If
	End Property

	Private Property Get Regex
		Regex=var_STACentre_Dir_Authors_Regex
		If IsEmpty(Regex) Or Regex="" Then
			Regex=STACentre_Dir_Authors_Regex
		End If
	End Property

	Private Property Get Anonymous
		Anonymous=var_STACentre_Dir_Authors_Anonymous
		If IsEmpty(Anonymous) Or Anonymous="" Then
			Anonymous=STACentre_Dir_Authors_Anonymous
		End If
	End Property

	Public Property Get Filename
		Filename = Alias
		If Anonymous Then Filename="default"
		FileName = FileName & "." & ZC_STATIC_TYPE
	End Property

	Public Property Get Directory
		Directory=STACentre_GetStaticDirectory(Regex,sType,ID,Name,Alias)
	End Property

	Public Property Get FullPath
		FullPath=BlogPath & Directory & Filename
	End Property

	Public Property Get Url
		If Anonymous Then
			Url=ZC_BLOG_HOST & Directory
		Else
			Url=ZC_BLOG_HOST & Directory & Filename
		End If
		If Not Enable Then Url=ZC_BLOG_HOST & "catalog.asp?"& "auth=" & ID
		Url=Replace(Url,"\","/")
	End Property

	Public Function LoadInfoByID(ByVal intID)
		Call CheckParameter(intID,"int",0)
		If intID<=UBound(Users) Then
			If IsObject(Users(intID)) Then
				ID=Users(intID).ID
				Name=Users(intID).Name
				Alias=Users(intID).Alias

				If IsNull(Alias) Or IsEmpty(Alias) Or Alias="" Then Alias=Name

				Dim ArtList
				Set ArtList=New TArticleList
				ArtList.LoadCache
				ArtList.template="CATALOG"
				If ArtList.ExportByMixed(Empty,Empty,ID,Empty,Empty,ZC_DISPLAY_MODE_ALL) Then
					ArtList.Build
					PageContent=ArtList.html
					LoadInfoByID=True
				End If
				Set ArtList=Nothing
			End If
		End If
	End Function

	Public Function LoadInfoByName(ByVal strName)
		LoadInfoByName=False
	End Function

	Public Function Build
		If Enable Then
			Call STACentre_AddFolderRecord(Directory)
			Call CreatDirectoryByCustomDirectory(Directory)
			Call STACentre_AddDirectoryRecord("file",FullPath)
			Call SaveToFile(FullPath,PageContent,"utf-8",True)
		End If
	End Function

	Public Function Del
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
			If fso.FileExists(FullPath) Then
				fso.DeleteFile(FullPath)
				Call STACentre_DelDirectoryRecord("file",FullPath)
			End If
		Set fso = Nothing
	End Function

	Private Sub class_initialize()
	End Sub

	Private sub class_terminate()
	End Sub

End Class

'Archives
Class STACentre_Archives

	Public ID '标准日期 2011-11 std
	Public Name '日期全称 Novermber-2011 full
	Public Alias '日期缩写 Nov-2011 abbr
	Public PageContent

	Private Property Get sType
		sType="Archives"
	End Property

	Private Property Get Enable
		Enable=var_STACentre_Dir_Archives_Enable
		If IsEmpty(Enable) Or Enable="" Then
			Enable=STACentre_Dir_Archives_Enable
		End If
	End Property

	Private Property Get Regex
		Regex=var_STACentre_Dir_Archives_Regex
		If IsEmpty(Regex) Or Regex="" Then
			Regex=STACentre_Dir_Archives_Regex
		End If
	End Property

	Private Property Get Anonymous
		Anonymous=var_STACentre_Dir_Archives_Anonymous
		If IsEmpty(Anonymous) Or Anonymous="" Then
			Anonymous=STACentre_Dir_Archives_Anonymous
		End If
	End Property

	Private Property Get Format
		Format=var_STACentre_Dir_Archives_Format
		If IsEmpty(Format) Or Format="" Then
			Format=STACentre_Dir_Archives_Format
		End If
	End Property

	Public Property Get Filename
		Select Case Format
			Case "std"
				FileName=ID
			Case "full"
				Filename=Name
			Case "abbr"
				Filename=Alias
			Case Else
				Filename=ID
		End Select
		If Anonymous Then Filename="default"
		FileName = FileName & "." & ZC_STATIC_TYPE
	End Property

	Public Property Get Directory
		Directory=STACentre_GetStaticDirectory(Regex,sType,ID,Name,Alias)
	End Property

	Public Property Get FullPath
		FullPath=BlogPath & Directory & Filename
	End Property

	Public Property Get Url
		If Anonymous Then
			Url=ZC_BLOG_HOST & Directory
		Else
			Url=ZC_BLOG_HOST & Directory & Filename
		End If
		If Not Enable Then Url=ZC_BLOG_HOST & "catalog.asp?"& "date=" & Year(ID) & "-" & Month(ID)
		Url=Replace(Url,"\","/")
	End Property

	Public Function LoadInfoByID(ByVal strID)
		Call CheckParameter(strID,"sql","2012-12")
		Dim ArtList
		Set ArtList=New TArticleList
		ArtList.LoadCache
		ArtList.template="CATALOG"
		If ArtList.ExportByMixed(Empty,Empty,Empty,strID,Empty,ZC_DISPLAY_MODE_ALL) Then
			ArtList.Build
			PageContent=ArtList.html

			ID=Year(strID) & "-" & Right("0"&Month(strID),2)
			Name=ZVA_Month(Month(strID)) & "-" & Year(strID)
			Alias=ZVA_Month_Abbr(Month(strID)) & "-" & Year(strID)

			LoadInfoByID=True
		End If
		Set ArtList=Nothing
	End Function

	Public Function LoadInfoByName(ByVal strName)
		LoadInfoByName=False
	End Function

	Public Function Build
		If Enable Then
			Call STACentre_AddFolderRecord(Directory)
			Call CreatDirectoryByCustomDirectory(Directory)
			Call STACentre_AddDirectoryRecord("file",FullPath)
			Call SaveToFile(FullPath,PageContent,"utf-8",True)
		End If
	End Function

	Public Function Del
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
			If fso.FileExists(FullPath) Then
				fso.DeleteFile(FullPath)
				Call STACentre_DelDirectoryRecord("file",FullPath)
			End If
		Set fso = Nothing
	End Function

	Private Sub class_initialize()
	End Sub

	Private sub class_terminate()
	End Sub

End Class
'*********************************************************




'*********************************************************
' 扩展部分: 记录与清理
'*********************************************************

'载入数据与存储数据
Function STACentre_LoadLogData()
	'Application.Contents.RemoveAll
	Application.Lock
	STACentre_LoadLogData=Application(ZC_BLOG_CLSID & "STACENTRE_LOG")
	Application.UnLock
	If IsEmpty(STACentre_LoadLogData) Or STACentre_LoadLogData="" Then
		STACentre_LoadLogData=LoadFromFile(BlogPath & "ZB_USERS/PLUGIN/STACentre/log.txt","utf-8")
	End If
End Function

Function STACentre_SaveLogData(strLog)
	Application.Lock
	Application(ZC_BLOG_CLSID & "STACENTRE_LOG")=strLog
	Application.UnLock
	Call SaveToFile(BlogPath & "ZB_USERS/PLUGIN/STACentre/log.txt",strLog,"utf-8",True)
End Function

'得到建立的最上级目录地址, 仅供完全清除时使用.
Function STACentre_AddFolderRecord(strDirectory)

	Dim s
	Dim t
	Dim i
	Dim j

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	s=BlogPath
	strDirectory=Replace(strDirectory,"/","\")
	t=Split(strDirectory,"\")
	j=0
	For i=LBound(t) To UBound(t)
		If (IsEmpty(t(i))=False) And (t(i)<>"") Then
			If j=0 And LCase(Left(t(i),3))="zb_" Then Exit For
			s=s & t(i)
			If (fso.FolderExists(s)=False) Then
				Call STACentre_AddDirectoryRecord("fldr",s)
				Exit For
			End If
			s=s & "\"
			j=j+1
		End If
	Next
	Set fso = Nothing

End Function


'创建单条生成记录
Function STACentre_AddDirectoryRecord(strType,strDirectory)

	Dim strLog
	Dim strList
	
	strDirectory = Replace(strDirectory,blogPath,"{%blogpath%}")


	strLog =  STACentre_LoadLogData()
	strList = strType & ":" & strDirectory & vbCrlf
	strList = Lcase(strList)

	If IsEmpty(strLog)=False Or strLog<>"" Then
		If InStr(strLog,strList)>0 Then Exit Function
	End If

	strLog= strList & strLog

	Call STACentre_SaveLogData(strLog)

End Function

'删除单条生成记录
Function STACentre_DelDirectoryRecord(strType,strDirectory)

	Dim strLog
	Dim strList

	strLog =  STACentre_LoadLogData()
	strList = strType & ":" & strDirectory & vbCrlf
	strList = Lcase(strList)

	If IsEmpty(strLog)=False Or strLog<>"" Then
		If InStr(strLog,strList)>0 Then
			strLog=Replace(strLog,strList,"")
		Else
			Exit Function
		End If
	End If

	Call STACentre_SaveLogData(strLog)

End Function

'删除所有生成的记录
Function STACentre_ClearAllDirsByHistory()
	'On Error Resume Next

	Dim logPath
	Dim strLog
	Dim aryLog
	Dim fso
	Dim s

	logPath=BlogPath & "ZB_USERS/PLUGIN/STACentre/log.txt"
	strLog = LoadFromFile(logPath,"utf-8")
	aryLog = Split(strLog,vbCrlf)

	Set fso = CreateObject("Scripting.FileSystemObject")

	For Each s In aryLog
		If s<>"" Then
			s=LCase(s)
			s=Replace(s,"{%blogpath%}",BlogPath)
			If Left(s,5)="fldr:" Then
				s=Replace(s,"fldr:","")
				If fso.FolderExists(s) Then
					fso.DeleteFolder(s)
				End If
			End If
			If Left(s,5)="file:" Then
				s=Replace(s,"file:","")
				If fso.FileExists(s) Then
					fso.DeleteFile(s)
				End If
			End If
		End If
	Next

	If fso.FileExists(logPath) Then
		fso.DeleteFile(logPath)
	End If

	Set fso = Nothing

	Application.Lock
	Application(ZC_BLOG_CLSID & "STACENTRE_LOG")=Empty
	Application.UnLock

	'Err.Clear
End Function
'*********************************************************





'*********************************************************
' 外延部分: 和输出直接相关, 见 Include.asp 文件
'*********************************************************
%>






























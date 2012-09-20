<%
'///////////////////////////////////////////////////////////////////////////////
'//				Z-Blog
'// 作	 者:    	瑜廷
'// 技术支持:    33195#qq.com
'// 程序名称:    	YT.Build
'// 开始时间:    	2010.12.21
'// 最后修改:    2012.08.24
'// 备	 注:    	only for zblog1.8
'///////////////////////////////////////////////////////////////////////////////
Class YTStatic
	Private jArticle
	Private jArticleList
	Private Fhtml
	Private Ftemplate
	
	Public Property Get html
		html = Fhtml
	End Property
	
	Public Property Let html(strFileName)
		Fhtml = strFileName
	End Property

	Private Sub Class_Initialize()

	End Sub
	
	Private Sub Class_Terminate()
		Set jArticle = Nothing
		Set jArticleList = Nothing
	End Sub
	
	Public Function Catalog(intPage,intCate,intAuth,strDate,strTags)
		Set jArticleList = new TArticleList
			If jArticleList.Export(intPage,intCate,intAuth,strDate,strTags,ZC_DISPLAY_MODE_ALL) Then
				jArticleList.Build
				Fhtml = jArticleList.html
				Catalog = True
			End If
	End Function
	
	Public Function View(ID)
		Set jArticle = new TArticle
		If jArticle.LoadInfoByID(ID) Then
			If jArticle.Level = 1 Then
				Fhtml = ZVA_ErrorMsg(9)
				Exit Function
			End If
			If jArticle.Level = 2 Then
				If Not CheckRights("Root") Then
					If (jArticle.AuthorID <> BlogUser.ID) Then
						Fhtml = ZVA_ErrorMsg(6)
						Exit Function
					End If
				End If
			End If
			If jArticle.Export(ZC_DISPLAY_MODE_ALL)= True Then
				jArticle.Build
				Fhtml = jArticle.html
				View = True
			End If
		End If
	End Function
	
	Public Sub GetData(intCateId,intAuthorId,dtmYearMonth,TagId,Byref intPageCount,Byref Row)
		Dim objRS
		Set objRS=Server.CreateObject("ADODB.Recordset")
			objRS.CursorType = adOpenKeyset
			objRS.LockType = adLockReadOnly
			objRS.ActiveConnection=objConn
			objRS.Source="SELECT [log_ID] FROM [blog_Article] WHERE ([log_ID]>0) AND([log_Type]=0) AND ([log_Level]>1)"
		If Not IsEmpty(intCateId) Then
			objRS.Source=objRS.Source & "AND([log_CateID]="&intCateId&")"
		End if
		If Not IsEmpty(intAuthorId) Then
			objRS.Source=objRS.Source & "AND([log_AuthorID]="&intAuthorId&")"
		End If
		If IsDate(dtmYearMonth) Then
			Dim y
			Dim m
			Dim ny
			Dim nm
			If IsDate(dtmYearMonth) Then
				y=Year(dtmYearMonth)
				m=Month(dtmYearMonth)
				objRS.Source=objRS.Source & "AND(Year([log_PostTime])="&y&") AND(Month([log_PostTime])="&m&")"
			End If
		End If
		If Not isEmpty(TagId) And isNumeric(TagId) Then
			objRS.Source=objRS.Source & "AND([log_Tag] LIKE '%{"&TagId&"}%')"
		End If
		
		objRS.Source=objRS.Source & "ORDER BY [log_PostTime] DESC,[log_ID] DESC"
		objRS.Open()
		
		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = ZC_DISPLAY_COUNT
			intPageCount = objRS.PageCount
			Dim j
			ReDim Row(-1)
			Do Until objRS.eof
				j = UBound(Row) + 1
				ReDim Preserve Row(j)
				Row(j) = objRS(0)
				objRS.MoveNext
			Loop
		Else
			ReDim Row(0)
		End If
		objRS.Close()
		Set objRS=Nothing
	End Sub
	
	Function GetdtmYM()
		Dim i,j,objRS,SQL,a,d
		If ZC_MSSQL_ENABLE Then
			SQL="SELECT distinct CONVERT(varchar(100),[log_PostTime],23) FROM [blog_Article] WHERE [log_Type]=0 AND ([log_Level]>1)"
		Else
			SQL="SELECT distinct format([log_PostTime],'YYYY-MM-DD') FROM [blog_Article] WHERE [log_Type]=0 AND ([log_Level]>1)"
		End If
		Set objRS=objConn.Execute(SQL)
		If (Not objRS.bof) And (Not objRS.eof) Then
			Dim dtmYM()
			i=0
			j=0
			Set d=CreateObject("Scripting.Dictionary")
				Do While Not objRS.eof
					i=Year(objRS(0)) & "-" & Month(objRS(0))
					If Not d.Exists(i) Then Call d.Add(i,i)
					objRS.MoveNext
				Loop
				ReDim dtmYM(0)
				For Each a In d.Keys
					j=UBound(dtmYM)
					ReDim Preserve dtmYM(j+1)
					dtmYM(j+1)=d.Item(a)
				Next
			Set d=Nothing
		End If
		objRS.Close
		Set objRS=Nothing
		GetdtmYM=dtmYM
	End Function

	Public Sub Save(Path)
		Call SaveToFile(Path,Fhtml,"utf-8",True)
	End Sub
End Class
%>
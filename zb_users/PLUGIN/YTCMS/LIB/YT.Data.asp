<%
Class YT_Table
	Function List()
		Dim aryFileList
		aryFileList=LoadIncludeFiles("ZB_USERS/PLUGIN/YTCMS/DATA/")
		If IsArray(aryFileList) Then
			Dim t(),j,i,s,e
			Redim t(-1)
			For i=1 to UBound(aryFileList)
				s=aryFileList(i)
				e=Mid(s,(InStrRev(s,".")+1),Len(s)-InStrRev(s,"."))
				If e="xml" Then
					j=UBound(t)+1
					ReDim Preserve t(j)
					t(j)=aryFileList(i)
				End If
			Next
		End If
		List = t
	End Function
	Sub Import(t)
		Dim x,d
			d=Server.MapPath("/ZB_USERS/PLUGIN/YTCMS/DATA/"&t)
		Set x=CreateObject("Microsoft.XMLDOM")
			x.async=False
			x.ValidateOnParse=False
			x.load(d)
			If x.readyState=4 Then
				If x.parseError.errorCode=0 Then
					Dim n,s,r,j,k,w,l,Field(),Value()
						t=Left(t,InStrRev(t,".")-1)
						If Exist(t) Then
							Dim sql
							For Each n In x.selectNodes("//"&t)
								Redim T2(-1)
								sql="INSERT INTO [@TABLE](@FIELDS) VALUES (@VALUE)"
								Set r=objConn.Execute("SELECT TOP 1 * FROM "&t)
									For Each k In r.Fields
										w=UBound(T2)+1
										ReDim Preserve T2(w)
										T2(w)=k.properties("ISAUTOINCREMENT")
									Next
								Set r = Nothing
								Redim Field(-1)
								Redim Value(-1)
								For s=0 To n.childNodes.length-1
									If T2(s)=False Then
										j=UBound(Field)+1
										ReDim Preserve Field(j)
										ReDim Preserve Value(j)
										Field(j)="["&n.childNodes(s).nodeName&"]"
										Value(j)="'"&n.childNodes(s).Text&"'"
									End If
								Next
								sql=Replace(sql,"@TABLE",t)
								sql=Replace(sql,"@FIELDS",Join(Field,","))
								sql=Replace(sql,"@VALUE",Join(Value,","))
								objConn.Execute(Sql)
							Next
						End If
					Dim fso,XmlFile
					Set fso = CreateObject("Scripting.FileSystemObject")
						Set XmlFile = fso.GetFile(d)
							XmlFile.Delete
						Set XmlFile = Nothing
					Set fso=Nothing
				End If
			End If
		Set x=Nothing
	End Sub 
	Function Exist(TableName)
		On Error Resume Next
		Dim Rs
		Set Rs=objConn.Execute("SELECT TOP 1 * FROM "&TableName)
		Set Rs=Nothing
			If Err.Number=0 Then
			Exist=True
		Else
			Err.Clear
			Exist=False
		End If	
	End Function
	Sub Delete(Node)
		Dim Sql
		Sql = "DROP TABLE ["&Node.selectSingleNode("Table/Name").Text&"]"
		objConn.Execute(Sql)
	End Sub
	Sub Create(Node)
		Dim Field,Sql
		Sql = "CREATE TABLE ["&Node.selectSingleNode("Table/Name").Text&"] ("
		For Each Field In Node.selectNodes("Field")
			Sql = Sql & "["&Field.selectSingleNode("Name").Text&"] "
			Sql = Sql & Field.selectSingleNode("Property").Text
			If ZC_MSSQL_ENABLE Then
				Sql=Replace(Sql,"COUNTER(1,1)","INT IDENTITY(1,1) NOT NULL")
			End If
			Sql = Sql & ","
		Next
		Sql = Sql & "[log_ID] INT)"
		objConn.Execute(Sql)
	End Sub
	Function GetFields(TableName)
		Dim Rs,fs(),n,i
		Set Rs = objConn.Execute("SELECT TOP 1 * FROM "&TableName)
			ReDim fs(Rs.Fields.Count-1)
			i = 0
			For Each n In Rs.Fields
				fs(i) = n.Name
				i = i + 1
			Next
		Set Rs = Nothing
		GetFields=fs
	End Function
	Function FieldExist(Fields,Field)
		Dim n
		For Each n In Fields
			If n=Field Then
				FieldExist=True
				Exit Function
			End If
		Next
		FieldExist=False
	End Function
End Class
Class YT_Article
	Private YTTemplate
	Private Sub Class_Initialize()
		Set YTTemplate = new YT_Template
	End Sub
	Private Sub Class_Terminate()
		Set YTTemplate = Nothing
	End Sub
	'单篇文章
	Function GetArticleModel(ID)
		ID = Split(ID,",")
		If isArray(ID) Then
			Dim Rs
			Set Rs = objConn.Execute("select * from blog_Article WHERE [log_ID] IN ("&Join(ID,",")&")")
				If Not (Rs.EOF and Rs.BOF) Then GetArticleModel = Rs.GetRows
			Set Rs = Nothing
		End If
	End Function
	'最新文章
	Function GetArticleRandomSortNew(Rows)
		If IsNumeric(Rows) Then
			Dim Rs
			Set Rs = objConn.Execute("select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_ID]>0) AND ([log_Level]>2) order by log_ID desc")
				If Not (Rs.EOF and Rs.BOF) Then GetArticleRandomSortNew = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'随机文章
	Function GetArticleRandomSortRand(Rows)
		If IsNumeric(Rows) Then
			Dim Rs,sql
			If ZC_MSSQL_ENABLE Then
				sql="select top "& CStr(Rows) &" [log_ID] from blog_Article WHERE ([log_Level]>2) order by newid()"
			Else
				Randomize
				sql="select top "& CStr(Rows) &" [log_ID] from blog_Article WHERE ([log_Level]>2) order by rnd("& (-1 * (Int(1000 * Rnd) + 1)) &" * log_ID)"
			End If
			Set Rs = objConn.Execute(sql)
				If Not (Rs.EOF and Rs.BOF) Then GetArticleRandomSortRand = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	

	'本月评论排行
	Function GetArticleRandomSortComMonth(Rows)
		If IsNumeric(Rows) Then
			Dim Rs,sql
			If ZC_MSSQL_ENABLE Then
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND DATEDIFF(MONTH,GETDATE(),log_PostTime)=0 ORDER BY log_CommNums DESC"
			Else
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND (log_PostTime>Now()-90) ORDER BY log_CommNums DESC"
			End If
			Set Rs = objConn.Execute(sql)
				If Not (Rs.EOF and Rs.BOF) Then GetArticleRandomSortComMonth = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'本年评论排行
	Function GetArticleRandomSortComYear(Rows)
		If IsNumeric(Rows) Then
			Dim Rs,sql
			If ZC_MSSQL_ENABLE Then
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND DATEDIFF(YEAR,GETDATE(),log_PostTime)=0 ORDER BY log_CommNums DESC"
			Else
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND  (log_PostTime>Now()-365) ORDER BY log_CommNums DESC "
			End If
			Set Rs = objConn.Execute(sql)
				If Not (Rs.EOF and Rs.BOF) Then GetArticleRandomSortComYear = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'本月排行
	Function GetArticleRandomSortTopMonth(Rows)
		If IsNumeric(Rows) Then
			Dim Rs,sql
			If ZC_MSSQL_ENABLE Then
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND  DATEDIFF(MONTH,GETDATE(),log_PostTime)=0 ORDER BY log_ViewNums DESC "
			Else
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND  (log_PostTime>Now()-30) ORDER BY log_ViewNums DESC "
			End If
			Set Rs = objConn.Execute(sql)
				If Not (Rs.EOF and Rs.BOF) Then GetArticleRandomSortTopMonth = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'本年排行
	Function GetArticleRandomSortTopYear(Rows)
		If IsNumeric(Rows) Then
			Dim Rs,sql
			If ZC_MSSQL_ENABLE Then
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND  DATEDIFF(YEAR,GETDATE(),log_PostTime)=0 ORDER BY log_ViewNums DESC "
			Else
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND  (log_PostTime>Now()-365) ORDER BY log_ViewNums DESC "
			End If
			Set Rs = objConn.Execute(sql)
				If Not (Rs.EOF and Rs.BOF) Then GetArticleRandomSortTopYear = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'热文排行
	Function GetArticleRandomSortTopHot(Rows)	
		If IsNumeric(Rows) Then
			Dim Rs,sql
			If ZC_MSSQL_ENABLE Then
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) ORDER BY log_CommNums*100 + log_TrackBackNums*200 + SQRT(log_ViewNums)*10 - DATEDIFF(DAY,GETDATE(),Log_PostTime)*DATEDIFF(DAY,GETDATE(),Log_PostTime) DESC"
			Else
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) ORDER BY log_CommNums*100 + log_TrackBackNums*200 + sqr(log_ViewNums)*10 - (date()-Log_PostTime)*(date()-Log_PostTime) DESC "
			End If
			Set Rs = objConn.Execute(sql)
				If Not (Rs.EOF and Rs.BOF) Then GetArticleRandomSortTopHot = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
		
	'分类文章列表
	Function GetArticleCategorys(Rows,CategoryID)
		If IsNumeric(Rows) Then
			Dim Rs
			Set Rs = objConn.Execute("SELECT top "& CStr(Rows) &" [log_ID] FROM [blog_Article] WHERE ([log_ID]>0) AND ([log_Level]>1) AND ([log_CateID] IN ("& CStr(CategoryID) &")) ORDER BY [log_PostTime] DESC")
				If Not (Rs.EOF and Rs.BOF) Then GetArticleCategorys = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'全站Limit(by:流年)
	
	Function GetArticleLimit(Rows,Index)
		If IsNumeric(Rows) And IsNumeric(Index) Then
			Dim Rs
			Set Rs = objConn.Execute("SELECT top "& CStr(Rows) &" [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND (log_ID>0) AND ([log_Istop]=0) AND [log_ID] NOT IN (SELECT top "& CStr(Index) &" [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND (log_ID>0) AND ([log_Istop]=0) ORDER BY [log_PostTime] DESC) ORDER BY [log_PostTime] DESC")
				If Not (Rs.EOF and Rs.BOF) Then GetArticleLimit = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'分类Limit(by:流年)
	Function GetArticleCategorysLimit(Rows,Index,CategoryID)
		If IsNumeric(Rows) And IsNumeric(Index) Then
			Dim Rs
			Set Rs = objConn.Execute("SELECT top "& CStr(Rows) &" [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND (log_ID>0) AND ([log_Istop]=0) AND [log_CateID] IN ("& CStr(CategoryID) &") AND [log_ID] NOT IN (SELECT top "& CStr(Index) &" [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND (log_ID>0) AND ([log_Istop]=0) AND [log_CateID] IN ("& CStr(CategoryID) &") ORDER BY [log_PostTime] DESC) ORDER BY [log_PostTime] DESC")
				If Not (Rs.EOF and Rs.BOF) Then GetArticleCategorysLimit = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	
	'分类热门文章列表
	Function GetArticleCategorysTophot(Rows,CategoryID)
		If IsNumeric(Rows) Then
			Dim Rs,sql
			If ZC_MSSQL_ENABLE Then
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND ([log_CateID] IN ("&CStr(CategoryID)&")) ORDER BY log_CommNums*100 + log_TrackBackNums*200 + SQRT(log_ViewNums)*10 - DATEDIFF(DAY,GETDATE(),Log_PostTime)*DATEDIFF(DAY,GETDATE(),Log_PostTime) DESC"
			Else
				sql="select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND ([log_CateID] IN ("&CStr(CategoryID)&")) ORDER BY log_CommNums*100 + log_TrackBackNums*200 + sqr(log_ViewNums)*10 - (date()-Log_PostTime)*(date()-Log_PostTime) DESC"
			End If
			Set Rs = objConn.Execute(sql)
				If Not (Rs.EOF and Rs.BOF) Then GetArticleCategorysTophot = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'Tag文章列表
	Function GetArticleTag(Rows,TagID)
		If IsNumeric(Rows) And IsNumeric(TagID) Then
			Dim Rs
			Set Rs = objConn.Execute("select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) And blog_Article.[log_Tag] LIKE '%{"&CStr(TagID)&"}%' ORDER BY Log_PostTime DESC")
				If Not (Rs.EOF and Rs.BOF) Then GetArticleTag = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'分类Tag文章列表
	Function GetArticleCategoryTag(Rows,CategoryID,TagID)
		If IsNumeric(Rows) And IsNumeric(TagID) Then
			Dim Rs
			Set Rs = objConn.Execute("select top " & CStr(Rows) & " [log_ID] from blog_Article WHERE ([log_Level]>2) AND (log_ID>0) AND ([log_CateID] IN ("&CStr(CategoryID)&")) And blog_Article.[log_Tag] LIKE '%{"&CStr(TagID)&"}%' ORDER BY Log_PostTime DESC")
				If Not (Rs.EOF and Rs.BOF) Then GetArticleCategoryTag = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'置顶文章列表
	Function GetArticleTop(Rows)
		If IsNumeric(Rows) Then
			Dim Rs
			Set Rs = objConn.Execute("SELECT top " & CStr(Rows) & " [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND (log_ID>0) AND ([log_Istop]=1) ORDER BY [log_PostTime] DESC")
				If Not (Rs.EOF and Rs.BOF) Then GetArticleTop = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'分类置顶文章列表
	Function GetArticleCategoryTop(Rows,CategoryID)
		If IsNumeric(Rows) Then
			Dim Rs
			Set Rs = objConn.Execute("SELECT top " & CStr(Rows) & " [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND (log_ID>0) AND ([log_Istop]=1) AND ([log_CateID] IN ("&CStr(CategoryID)&")) ORDER BY [log_PostTime] DESC")
				If Not (Rs.EOF and Rs.BOF) Then GetArticleCategoryTop = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
End Class

Class YT_Comment
	'最新回复列表
	Function GetCommentComments(Rows)
		If IsNumeric(Rows) Then
			Dim Rs
			Set Rs = objConn.Execute("SELECT top "& CStr(Rows) &" [comm_ID] FROM [blog_Article],[blog_Comment] WHERE blog_Article.log_ID>0 and blog_Comment.log_ID=blog_Article.log_ID ORDER BY [comm_PostTime] DESC,[comm_ID] DESC")
				If Not (Rs.EOF and Rs.BOF) Then GetCommentComments = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	
	'分类最新回复列表
	Function GetCommentCategorysComments(Rows,CategoryID)
		If IsNumeric(Rows) Then
			Dim Rs
			Set Rs = objConn.Execute("SELECT top "& CStr(Rows) &" [comm_ID] FROM [blog_Article],[blog_Comment] WHERE blog_Article.log_ID>0 AND ([log_CateID] IN ("& CStr(CategoryID) &")) and blog_Comment.log_ID=blog_Article.log_ID ORDER BY [comm_PostTime] DESC,[comm_ID] DESC")
				If Not (Rs.EOF and Rs.BOF) Then GetCommentCategorysComments = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
	'文章评论列表
	Function GetCommentArticleComments(Rows,ID)
		If IsNumeric(Rows) Then
			Dim Rs
			Set Rs = objConn.Execute("SELECT top "& CStr(Rows) &" [comm_ID] FROM [blog_Comment] WHERE (blog_Comment.log_ID IN ("& CStr(ID) &")) ORDER BY [comm_PostTime] DESC,[comm_ID] DESC")
				If Not (Rs.EOF and Rs.BOF) Then GetCommentArticleComments = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
End Class

Class YT_Tag
	' 标签列表
	Function GetTagLists(Rows)
		If IsNumeric(Rows) Then
			Dim Rs
			Set Rs = objConn.Execute("SELECT top " & CStr(Rows) & " [tag_ID] FROM [blog_Tag] ORDER BY [tag_Order] DESC,[tag_Count] DESC,[tag_ID] ASC")
				If Not (Rs.EOF and Rs.BOF) Then GetTagLists = Rs.GetRows(Rows)
			Set Rs = Nothing
		End If
	End Function
End Class
%>
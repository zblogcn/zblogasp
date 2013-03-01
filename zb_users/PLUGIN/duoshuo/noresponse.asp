<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("duoshuo")=False Then Call ShowError(48)
Call DuoShuo_Initialize

Select Case Request.QueryString("act")
	Case "callback":Call CallBack
	Case "export":Call Export
End Select


Sub CallBack()
	If Not IsEmpty(duoshuo.get("short_name")) Then
		duoshuo.config.Write "short_name",duoshuo.get("short_name")
		duoshuo.config.Write "secret",duoshuo.get("secret")
		duoshuo.config.Save
	End If
	Call SetBlogHint(True,Empty,Empty)
	Response.Write "<script>top.location.reload()</script>"
End Sub

Sub Export
	
	Response.ContentType="application/json"
	Response.AddHeader "Content-Disposition", "attachment; filename=duoshuo_export.json"

	Response.Write "{""threads"":"&ArticleData().jsString&",""posts"":"
	Response.Write QueryToJson(objConn,"SELECT comm_AuthorID As author_key,comm_ID As post_key,log_id As thread_key,comm_ParentID As parent_key"&_
					",comm_Author As author_name,comm_Email As author_email,comm_HomePage As author_url,comm_PostTime As created_at"&_
					",comm_ip As ip,comm_agent As agent,comm_Content As message FROM blog_Comment WHERE comm_IsCheck=0").jsString
	Response.Write "}"
	'Dim aryData(),rs,i
	'i=0
	'Set rs=objConn.Execute("SELECT * FROM blog_Comment")
	'Redim aryData(rs.PageSize)
	'Do Until rs.Eof
'		aryData(i)=rs("comm_id")
'		rs.MoveNext
'	Loop
'	Dim s
'	s=(new duoshuo_Duoshuo_aspjson).toJSON(aryData)
'	Response.Write s
End Sub

Function ArticleData()
        Dim rs, jsa, col , o
        Set rs = objConn.Execute("SELECT [log_ID] As thread_key,[log_CateID],[log_Title] as title,[log_Intro] as excerpt,[log_Level],[log_AuthorID] as author_key,[log_PostTime],[log_ViewNums] as views,[log_Url] as url,[log_Type] FROM [blog_Article]")
        Set jsa = jsArray()
		jsa.Kind=1
        While Not (rs.EOF Or rs.BOF)
				Set o=New TArticle
				If o.LoadInfoByArray(Array(rs(0),"",rs(1),rs(2),rs(3),"",rs(4),rs(5),rs(6),0,rs(7),0,rs(8),False,"","",rs(9),"")) Then
	                Set jsa(Null) = jsObject()
					For Each col In rs.Fields
						If col.Name<>"url" And Left(col.Name,4)<>"log_" Then
	    	            	jsa(Null)(col.Name) = col.Value
						ElseIf col.Name = "create_at" Then
							jsa(Null)(col.Name) = Year(col.Value) & "-" & Month(col.Value) & "-" & Day(col.Value) & " " & Hour(col.Value) & ":" & Minute(col.Value) & ":" & Second(col.Value)
						ElseIf col.Name="url" Then
							jsa(Null)(col.Name) = TransferHTML(o.FullUrl,"[zc_blog_host]")
						End If
					Next
        		End If
				Set o=Nothing
		rs.MoveNext
        Wend
        Set ArticleData = jsa
End Function

Function QueryToJSON(dbc, sql)
        Dim rs, jsa, col, k
        Set rs = dbc.Execute(sql)
        Set jsa = jsArray()
		jsa.Kind=1
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
						If col.Name = "created_at" Then
							k=CStr(col.Value)
							jsa(Null)(col.Name) = Year(k) & "-" & Right("0"&Month(k),2) & "-" & Right("0"&Day(k),2) & " " & Right("0"&Hour(k),2) & ":" & Right("0"&Minute(k),2) & ":" & Right("0"&Second(k),2)
						Else
	                        jsa(Null)(col.Name) = col.Value
                		End If
				Next
        rs.MoveNext
        Wend
        Set QueryToJSON = jsa
End Function


Function jsObject
	Set jsObject = new duoshuo_aspjson
	jsObject.Kind = 0
End Function

Function jsArray
	Set jsArray = new duoshuo_aspjson
	jsArray.Kind = 1
End Function

%>
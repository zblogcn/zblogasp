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
	'文章导出，未完工
	'Response.Write QueryToJson(objConn,"SELECT log_ID As thread_key,log_Title As title,log_id As thread_key,comm_ParentID As parent_key"&_
					",comm_Author As author_name,comm_Email As author_email,comm_HomePage As author_url,comm_PostTime As created_at"&_
					",comm_ip As ip,comm_agent As agent,comm_Content As message FROM blog_Comment").jsString
	Response.Write QueryToJson(objConn,"SELECT comm_AuthorID As author_key,comm_ID As post_key,log_id As thread_key,comm_ParentID As parent_key"&_
					",comm_Author As author_name,comm_Email As author_email,comm_HomePage As author_url,comm_PostTime As created_at"&_
					",comm_ip As ip,comm_agent As agent,comm_Content As message FROM blog_Comment").jsString
	'Dim aryData(),objRs,i
	'i=0
	'Set objRs=objConn.Execute("SELECT * FROM blog_Comment")
	'Redim aryData(objRs.PageSize)
	'Do Until objRs.Eof
'		aryData(i)=objRs("comm_id")
'		objRs.MoveNext
'	Loop
'	Dim s
'	s=(new duoshuo_Duoshuo_aspjson).toJSON(aryData)
'	Response.Write s
End Sub

Function QueryToJSON(dbc, sql)
        Dim rs, jsa, col
        Set rs = dbc.Execute(sql)
        Set jsa = jsArray()
		jsa.Kind=1
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null)(col.Name) = col.Value
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
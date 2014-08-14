<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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
<!-- #include file="ASPJSON.asp" -->
<%
Call System_Initialize()
If BlogUser.Level > 1 Then Call ShowError(6)
If Not CheckPluginState("A2P") Then Call ShowError(48)

Dim strType
Dim intMin, intMax, intMaxPackID, intPackID
Dim strFieldName, strTableName, strSQL
Dim resObject
Set resObject = A2P_jsObject()

intMin = Request.Form("min")
intMax = Request.Form("max")
intPackID = Request.Form("count")
intMaxPackID = Request.Form("sum")
strType = Request.Form("type")


Call CheckParameter(intMin, "int", 1)
Call CheckParameter(intMax, "int", 1)
Call CheckParameter(intPackID, "int", 1)
Call CheckParameter(intMaxPackID, "int", 1)

strTableName = strType
Select Case strType
	Case "article":  strFieldName = "log_ID"
	Case "comment":  strFieldName = "comm_ID"
	Case "member":   strFieldName = "mem_ID"
	Case "category": strFieldName = "cate_ID"
	Case "upload":   strFieldName = "ul_ID"
End Select
strSQL = "SELECT * FROM [blog_" & strTableName & "] WHERE (" & strFieldName & " BETWEEN " & intMin & " AND " & intMax & ") ORDER BY " & strFieldName & " DESC"
	
resObject("time") = 0
resObject("message") = ""
resObject("error") = 0

If Not PublicObjFSO.FolderExists(BlogPath & "zb_users/PLUGIN/A2P/output/") Then
	Call PublicObjFSO.CreateFolder(BlogPath & "zb_users/PLUGIN/A2P/output/")
End If 

Dim objReturn
Set objReturn = A2P_jsObject()
Set objReturn("data") = QueryToJson(objConn, strSQL)
objReturn("table") = strType
Call SaveToFile(BlogPath & "zb_users/PLUGIN/A2P/output/data_pack_" & intPackID & ".json", A2P_toJSON(objReturn), "UTF-8", True)

resObject("time") = RunTime()
resObject("message") = strType & "数据包分块" & intPackID & "/" & intMaxPackID & "导出完成，耗时" & resObject("time") & "ms。"
resObject.Flush

	

Function BuildDoubleTopSql(intPage, PageSize, intMax, strFieldName, strTableName, strWhere, strOrder, strIDField)
	Dim PageSize2, aryResult(1)
	PageSize2 = PageSize
	If intPage < 1 Then intPage = 1
	
	If PageSize * intPage > intMax Then
		PageSize2 = CLng(intMax Mod PageSize)
		If PageSize * (intPage - 1) + PageSize2 >intMax Then
			aryResult(0) = False
			aryResult(1) = False
			A2P_GetDoubleTop = aryResult
			Exit Function
		End If
	End If
	
	aryResult(0) = PageSize2
	aryResult(1) = "SELECT * FROM (SELECT TOP " & PageSize2 & " *  FROM (SELECT TOP " & (PageSize * intPage) & " " & strFieldName
	aryResult(1) = aryResult(1) & " FROM [" & strTableName & "]  " & IIf(strWhere = "" , "", "WHERE (" & strWhere & ")")
	aryResult(1) = aryResult(1) & " ORDER BY " & strIDField & " ASC) AS [TEST] ORDER BY " & strIDField & " DESC) " 
	aryResult(1) = aryResult(1) & IIf(strOrder = "", "", " AS [TEST] " & strOrder) 
	
	BuildDoubleTopSql = aryResult
	
End Function

Function QueryToJSON(dbc, sql)
	Dim rs, jsa, col
	Set rs = dbc.Execute(sql)
	Set jsa = A2P_jsArray()
	While Not (rs.EOF Or rs.BOF)
		Set jsa(Null) = A2P_jsObject()
			For Each col In rs.Fields
				jsa(Null)(col.Name) = col.Value
		Next
	rs.MoveNext
	Wend
	Set QueryToJSON = jsa
End Function
%>
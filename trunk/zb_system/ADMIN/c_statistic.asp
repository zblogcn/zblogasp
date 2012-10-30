<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)&(sipo)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_updateinfo.asp
'// 开始时间:    2007-1-26
'// 最后修改:    
'// 备    注:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_event.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查权限
If Not CheckRights("SiteInfo") Then Call ShowError(6)

Response.ExpiresAbsolute   =   Now()   -   1           
Response.Expires   =   0   
Response.CacheControl   =   "no-cache"   

Dim strContent

Dim b
b=False
Dim fso,f
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(BlogPath & "zb_users\CACHE\statistic.txt")=True Then
	Set f = fso.GetFile(BlogPath & "zb_users\CACHE\statistic.txt")
	strContent=LoadFromFile(BlogPath & "zb_users\CACHE\statistic.txt","utf-8")
	If DateDiff("h",f.DateLastModified,Now)>24 Then
		b=True
	End If
Else
	b=True
End If

If IsEmpty(Request.QueryString("reload"))=False Then
	b=True
End If

If b=True Then
	Dim objRS
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	Dim allArticle,allCommNums,allViewNums,allUserNums,allCateNums,allTagsNums,allPage

	Call getUser()
	Dim User,i
	For Each User in Users
		If IsObject(User) Then
			Set objRS=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE [log_Level]>1 AND [log_AuthorID]=" & User.ID )
			i=objRS(0)
			objConn.Execute("UPDATE [blog_Member] SET [mem_PostLogs]="&i&" WHERE [mem_ID] =" & User.ID)
			Set objRS=Nothing

			Set objRS=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [comm_AuthorID]=" & User.ID )
			i=objRS(0)
			objConn.Execute("UPDATE [blog_Member] SET [mem_PostComms]="&i&" WHERE [mem_ID] =" & User.ID)
			Set objRS=Nothing
		End If
	Next

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""
	objRS.Open("SELECT COUNT([log_ID])AS allArticle,SUM([log_ViewNums]) AS allViewNums FROM [blog_Article] WHERE [log_Type]=0")
	If (Not objRS.bof) And (Not objRS.eof) Then
		allArticle=objRS("allArticle")
		allPage=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE [log_Type]=1")(0)
		allCommNums=objConn.Execute("SELECT SUM([log_CommNums]) FROM [blog_Article]")(0)
		allViewNums=objRS("allViewNums")
	End If
	objRS.Close

	objRS.Open("SELECT COUNT([tag_ID])AS allTagsNums FROM [blog_Tag]")
	If (Not objRS.bof) And (Not objRS.eof) Then
		allTagsNums=objRS("allTagsNums")
	End If
	objRS.Close

	objRS.Open("SELECT COUNT([mem_ID])AS allUserNums FROM [blog_Member]")
	If (Not objRS.bof) And (Not objRS.eof) Then
		allUserNums=objRS("allUserNums")
	End If
	objRS.Close

	objRS.Open("SELECT COUNT([cate_ID])AS allCateNums FROM [blog_Category]")
	If (Not objRS.bof) And (Not objRS.eof) Then
		allCateNums=objRS("allCateNums")
	End If
	objRS.Close

	Call CheckParameter(allArticle,"int",0)
	Call CheckParameter(allCommNums,"int",0)
	Call CheckParameter(allViewNums,"int",0)
	Call CheckParameter(allUserNums,"int",0)
	Call CheckParameter(allCateNums,"int",0)
	Call CheckParameter(allTagsNums,"int",0)

	'strContent = "<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"" width=""100%"" class=""tableBorder"">"
	strContent = "<tr class=""color1""><th height=""32"" colspan=""4""  align=""center"">&nbsp;" & ZC_MSG167& "&nbsp;<a href=""javascript:statistic('?reload');"">["& ZC_MSG225 &"]</a> <img id=""statloading"" style=""display:none"" src=""../image/admin/loading.gif""></th></tr>"
	strContent = strContent & "<tr>"
	strContent = strContent & "<td width=""20%"">" & ZC_MSG005& " </td>"
	strContent = strContent & "<td width=""30%"">" & BlogUser.Name& "  (" & ZVA_User_Level_Name(BlogUser.Level)& ")</td>"
	strContent = strContent & "<td width=""20%"">" & ZC_MSG150& " </td>"
	strContent = strContent & "<td width=""30%"">" & ZC_BLOG_VERSION& " </td>"
	strContent = strContent & "</tr>"
	strContent = strContent & "<tr>"
	strContent = strContent & "<td width=""20%"">" & ZC_MSG082& " </td>"
	strContent = strContent & "<td width=""30%"">" & allArticle& " </td>"
	strContent = strContent & "<td width=""20%"">" & ZC_MSG124& " </td>"
	strContent = strContent & "<td width=""30%"">" & allCommNums& " </td>"
	strContent = strContent & "</tr>"
	strContent = strContent & "<tr>"
	strContent = strContent & "<td width=""20%"">" & ZC_MSG125& " </td>"
	strContent = strContent & "<td width=""30%"">" & allPage& " </td>"
	strContent = strContent & "<td width=""20%"">" & ZC_MSG129& " </td>"
	strContent = strContent & "<td width=""30%"">" & allViewNums& " </td>"
	strContent = strContent & "</tr>"
	strContent = strContent & "<tr>"
	strContent = strContent & "<td width=""20%"">" & ZC_MSG163& " </td>"
	strContent = strContent & "<td width=""30%"">" & allTagsNums& " </td>"
	strContent = strContent & "<td width=""20%"">" & ZC_MSG162& " </td>"
	strContent = strContent & "<td width=""30%"">" & allCateNums& " </td>"
	strContent = strContent & "</tr>"
	strContent = strContent & "<tr>"
	strContent = strContent & "<td width=""20%"">" & ZC_MSG204& " /" & ZC_MSG083& " </td>"
	strContent = strContent & "<td width=""30%"">" & GetNameFormTheme(ZC_BLOG_THEME)& "  / " & ZC_BLOG_CSS& " .css</td>"
	strContent = strContent & "<td width=""20%"">" & ZC_MSG166& " </td>"
	strContent = strContent & "<td width=""30%"">" & allUserNums& " </td>"
	strContent = strContent & "</tr>"
	strContent = strContent & "<tr>"
	strContent = strContent & "<td width=""20%"">MetaWeblog API</td>"
	strContent = strContent & "<td colspan=""3"" width=""80%"">" & BlogHost& "zb_system/xml-rpc/index.asp</td>"
	strContent = strContent & "</tr>"
	'strContent = strContent & "</table>"

	Call SaveToFile(BlogPath & "zb_users\CACHE\statistic.txt",strContent,"utf-8",False)

End If

Response.Write strContent

Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
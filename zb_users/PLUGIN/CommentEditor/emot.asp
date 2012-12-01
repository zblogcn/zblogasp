<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<%' Response.ContentType="text/json" %>
<!-- #include file="../../../zb_users/c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<% Response.Clear %>{<%	
	Dim fso,f(),f1,fb,fc
	Dim aryFileList,a,i,j,e,x,y,p

	'f=Split(ZC_EMOTICONS_FILENAME,"|")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set fb = fso.GetFolder(BlogPath & "zb_users/emotion" & "/")
	Set fc = fb.SubFolders
		i=0
	For Each f1 in fc	
		ReDim Preserve f(i)
		f(i)=f1.name
		i=i+1
	Next
	'f=LoadIncludeFiles("zb_users\emotion\")
	y=UBound(f)
	For x=0 To y
		aryFileList=LoadIncludeFiles("zb_users\emotion\"&f(x)) 
		If IsArray(aryFileList) Then
			j=UBound(aryFileList)
			For i=1 to j
				If InStr(ZC_EMOTICONS_FILETYPE,Right(aryFileList(i),3))>0 Then 
					e="'"&Replace(Server.URLEncode(aryFileList(i)),"+","%20")&"':'"&aryFileList(i)&"',"& e 
					p=i
				End If 
			Next
			e=Left(e,Len(e)-1)
		End If 
	%>'<%=f(x)%>':{name:'<%=f(x)%>',list:{<%=e%>},width:50,height:50,line:10}<% If x<y Then Response.Write "," 
		e=""
	Next
	%>}
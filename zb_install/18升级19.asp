<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:   
'// 开始时间:   
'// 最后修改:    
'// 备    注:   
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<%Response.Buffer=True %>
<!-- #include file="c_option.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn" lang="zh-cn">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="zh-cn" />
	<meta http-equiv="pragma" content="no-cache">
	<meta http-equiv="cache-control" content="no-cache,must-revalidate">
	<meta http-equiv="expires" content="0">
	<title>Z-Blog 1.8升级1.9工具</title>
<style type="text/css">
<!--
*{
	font-size:14px;
}
body{
	margin:0;
	padding:0;
	color: #000000;
	font-size:12px;
	background:#EDF5FB;
	font-family:"宋体","黑体";
}
h1,h2,h3,h4,h5,h6{
	font-size:18px;
	padding:0;
	margin:0;
}
a{
	text-decoration: none;
}
a:link {
	color:#0066CC;
	text-decoration: none;
}
a:visited {
	color:#0066CC;
	text-decoration: none;
}
a:hover {
	color:#FF7F50;
	text-decoration: underline;
}
a:active {
	color:#FF7F50;
	text-decoration: underline;
}
p{
	margin:0;
	padding:5px;
}
table {
	border-collapse: collapse;
	border:1px solid #527b9a;
	background:#ffffff;
	margin-top:10px;
}
td{
	border:1px solid #527b9a;
	margin:0;
	padding:3px;
}
img{
	border:0;
}
hr{
	border:0px;
	border-top:1px solid #666666;
	background:#666666;
	margin:2px 0 4px 0;
	padding:0;
	height:0px;
}
img{
	margin:0;
	padding:0;
}
form{
	margin:0;
	padding:0;
}
input{
	background:#eeeeee;
}
select{
	background:#eeeeee;
}
textarea{
	background:#eeeeee;
}
input.button{
	background:#eeeeee url("../image/edit/fade-butt.png");
	border: 3px double #909090;
	border-left-color: #c0c0c0;
	border-top-color: #c0c0c0;
	color: #333;
	padding: 0.05em 0.25em 0.05em 0.25em;
}

#frmLogin{
	position:absolute;
	left: 50%;
	top: 40%;
	margin: -150px 0px 0px -300px;
	padding:0;
	overflow:hidden;
	width:600px;
	height:400px;
	background-color:white;
	border:1px solid #527b9a;
}

#frmLogin h3{
	padding:10px 0 10px 0;
	margin:1px 1px 0 1px;
	text-align:center;
	color:black;
	background:#c5dbec;
	font-size:24px;
	height:30px;
}

#divHeader{
	background:#e4f0f9;
	margin:0 1px 0 1px;
	padding:8px;
}
#divMain{
	height:290px;
}
#divFooter{
	border-top:1px solid #A1B0B9;
	margin:0 1px 0 1px;
	text-align:center;
	padding:2px;
}

#divMain_Top{
	padding:8px;
	padding-bottom:0;	
}
#divMain_Center{
	padding:5px;
}
#divMain_Bottom{
	text-align:right;
	padding:5px;
}
#txaContent{
	border:1px solid #A1B0B9;
	background:#FFFFFF;
}
-->
</style>
</head>
<body>


<form id="frmLogin" method="post">
<h3>Z-Blog 1.8升级1.9工具</h3>
<div id="divHeader"><a href="http://www.rainbowsoft.org/" target="_blank">Z-Blog主页</a> | <a href="http://bbs.rainbowsoft.org" class="here" target="_blank">Zblogger社区</a> | <a href="http://wiki.rainbowsoft.org/" target="_blank">Z-Wiki</a> | <a href="http://blog.rainbowsoft.org/" target="_blank">菠萝阁</a> | <a href="http://show.rainbowsoft.org/" target="_blank">菠萝秀</a> | <a href="http://download.rainbowsoft.org/" target="_blank">菠萝的海</a> | <a href="http://www.dbshost.cn/" target="_blank">DBS主机</a></div>
<div id="divMain">
<input type="hidden" name="ok" id="ok" value="ok" />
<%

Dim BlogPath
BlogPath=Server.MapPath("cmd.asp")
BlogPath=Left(BlogPath,Len(BlogPath)-Len("cmd.asp"))

If Request.Form("ok")="ok" Then


'*********************************************************
' 目的：    Load Value For Setting
'*********************************************************
Function LoadValueForSetting(strContent,bolConst,strTypeVar,strItem,ByRef strValue)

	Dim i,j,s,t
	Dim strConst
	Dim objRegExp
	Dim Matches,Match

	If bolConst=True Then strConst="Const"

	Set objRegExp=New RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True


	If strTypeVar="String" Then

		objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))(.+?)(\r\n|\n|$)"
		Set Matches = objRegExp.Execute(strContent)
		If Matches.Count=1 Then

			t=Matches(0).Value
			t=Replace(t,VbCrlf,"")
			t=Replace(t,Vblf,"")
			objRegExp.Pattern="( *)""(.*)""( *)($)"
			Set Matches = objRegExp.Execute(t)

			If Matches.Count>0 Then

				s=Trim(Matches(0).Value)
				s=Mid(s,2,Len(s)-2)
				s=Replace(s,"""""","""")
				strValue=s

				LoadValueForSetting=True
				Exit Function

			End If
		End If

	End If

	If strTypeVar="Boolean" Then

		objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([a-z]+)( *)(\r\n|\n|$)"
		Set Matches = objRegExp.Execute(strContent)
		If Matches.Count=1 Then
			t=Matches(0).Value
			t=Replace(t,VbCrlf,"")
			t=Replace(t,Vblf,"")
			objRegExp.Pattern="( *)((True)|(False))( *)($)"
			Set Matches = objRegExp.Execute(t)

			If Matches.Count>0 Then

				s=Trim(Matches(0).Value)
				s=LCase(Matches(0).Value)
				If InStr(s,"true")>0 Then
					strValue=True
				ElseIf InStr(s,"false")>0 Then
					strValue=False
				End If

				LoadValueForSetting=True
				Exit Function

			End If
		End If

	End If

	If strTypeVar="Numeric" Then

		objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([0-9.]+)( *)(\r\n|\n|$)"
		Set Matches = objRegExp.Execute(strContent)
		If Matches.Count=1 Then
			t=Matches(0).Value
			t=Replace(t,VbCrlf,"")
			t=Replace(t,Vblf,"")
			objRegExp.Pattern="( *)([0-9.]+)( *)($)"
			Set Matches = objRegExp.Execute(t)

			If Matches.Count>0 Then

				s=Trim(Matches(0).Value)
				If IsNumeric(s)=True Then

					strValue=s

					LoadValueForSetting=True
					Exit Function

				End If

			End If
		End If

	End If

	LoadValueForSetting=False

End Function
'*********************************************************


'*********************************************************
' 目的：    Save Value For Setting
'*********************************************************
Function SaveValueForSetting(ByRef strContent,bolConst,strTypeVar,strItem,strValue)

	Dim i,j,s,t
	Dim strConst
	Dim objRegExp

	If bolConst=True Then strConst="Const"

	Set objRegExp=New RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True

	If strTypeVar="String" Then

		strValue=Replace(strValue,"""","""""")
		strValue=""""& strValue &""""

		objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))(.+?)(\r\n|\n|$)"
		If objRegExp.Test(strContent)=True Then
			strContent=objRegExp.Replace(strContent,"$1$2"& strValue &"$8")
			SaveValueForSetting=True
			Exit Function
		End If

	End If

	If strTypeVar="Boolean" Then

		strValue=Trim(strValue)
		If LCase(strValue)="true" Then
			strValue="True"
		Else
			strValue="False"
		End If

		If objRegExp.Test(strContent)=True Then
			objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([a-z]+)( *)(\r\n|\n|$)"
			strContent=objRegExp.Replace(strContent,"$1$2"& strValue &"$9")
			SaveValueForSetting=True
			Exit Function
		End If


	End If

	If strTypeVar="Numeric" Then

		strValue=Trim(strValue)
		If IsNumeric(strValue)=False Then
			strValue=0
		End If

		If objRegExp.Test(strContent)=True Then
			objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([0-9.]+)( *)(\r\n|\n|$)"
			strContent=objRegExp.Replace(strContent,"$1$2"& strValue &"$9")
			SaveValueForSetting=True
			Exit Function
		End If

	End If

	SaveValueForSetting=False

End Function
'*********************************************************



'*********************************************************
' 目的：    Load Text form File
' 输入：    
' 输入：    
' 返回：    
'*********************************************************
Function LoadFromFile(strFullName,strCharset)

	On Error Resume Next

	Dim objStream

	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = adTypeText
	.Mode = adModeReadWrite
	.Open
	.Charset = strCharset
	.Position = objStream.Size
	.LoadFromFile strFullName
	LoadFromFile=.ReadText
	.Close
	End With
	Set objStream = Nothing

	Err.Clear

End Function
'*********************************************************


'*********************************************************
' 目的：    Save Text to File
' 输入：    
' 输入：    
' 返回：    
'*********************************************************
Function SaveToFile(strFullName,strContent,strCharset,bolRemoveBOM)

	On Error Resume Next

	Dim objStream

	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = adTypeText
	.Mode = adModeReadWrite
	.Open
	.Charset = strCharset
	.Position = objStream.Size
	.WriteText = strContent
	.SaveToFile strFullName,adSaveCreateOverWrite
	.Close
	End With
	Set objStream = Nothing

	Err.Clear

End Function
'*********************************************************


Function CheckUpdateDB(a,b)
	Err.Clear
	On Error Resume Next
	Dim Rs
	Set Rs=objConn.execute("SELECT "&a&" FROM "&b)
	Set Rs=Nothing
	If Err.Number=0 Then
	CheckUpdateDB=True
	Else
	Err.Clear
	CheckUpdateDB=False
	End If	
End Function


Function UpdateDateBase()
	
	If Not CheckUpdateDB("[log_IsTop]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_IsTop] YESNO DEFAULT FALSE")
		objConn.execute("UPDATE [blog_Article] SET [log_IsTop]=0")
	End If

	If Not CheckUpdateDB("[log_Tag]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Tag] VARCHAR(255)")
	End If

	If Not CheckUpdateDB("[tag_ID]","[blog_Tag]") Then
		objConn.execute("CREATE TABLE [blog_Tag] (tag_ID AutoIncrement primary key,tag_Name VARCHAR(255),tag_Intro text,tag_ParentID int,tag_URL VARCHAR(255),tag_Order int,tag_Count int)")
	End If

	If Not CheckUpdateDB("[coun_ID]","[blog_Counter]") Then
		objConn.execute("CREATE TABLE [blog_Counter] (coun_ID AutoIncrement primary key,coun_IP VARCHAR(20),coun_Agent text,coun_Refer VARCHAR(255),coun_PostTime TIME DEFAULT Now())")
	End If

	If Not CheckUpdateDB("[key_ID]","[blog_Keyword]") Then
		objConn.execute("CREATE TABLE [blog_Keyword] (key_ID AutoIncrement primary key,key_Name VARCHAR(255),key_Intro text,key_URL VARCHAR(255))")
	End If

	If Not CheckUpdateDB("[ul_Quote]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_Quote] VARCHAR(255)")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_Quote]=''")
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_DownNum] int DEFAULT 0")
	End If

	If Not CheckUpdateDB("[ul_FileIntro]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_FileIntro] VARCHAR(255)")
	End If

	If Not CheckUpdateDB("[ul_DirByTime]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_DirByTime] YESNO DEFAULT FALSE")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_DirByTime]=[ul_Quote]")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_Quote]=''")
	End If

	If Not CheckUpdateDB("[tag_Template]","[blog_Tag]") Then
		objConn.execute("ALTER TABLE [blog_Tag] ADD COLUMN [tag_Template] VARCHAR(255)")
	End If

	If Not CheckUpdateDB("[cate_Template]","[blog_Category]") Then
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Template] VARCHAR(255)")
	End If

	If Not CheckUpdateDB("[log_Template]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Template] VARCHAR(255)")
	End If

	If Not CheckUpdateDB("[log_Meta]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Yea] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Nay] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Delate] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Meta] text")
	End If

	If Not CheckUpdateDB("[tag_Meta]","[blog_Tag]") Then
		objConn.execute("ALTER TABLE [blog_Tag] ADD COLUMN [tag_Meta] text")
	End If

	If Not CheckUpdateDB("[ul_Meta]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_Meta] text")
	End If

	If Not CheckUpdateDB("[cate_Meta]","[blog_Category]") Then
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Temp] VARCHAR(255)")
		objConn.execute("UPDATE [blog_Category] SET [cate_Temp]=[cate_Url]")
		objConn.execute("UPDATE [blog_Category] SET [cate_URL]=[cate_Intro]")
		objConn.execute("UPDATE [blog_Category] SET [cate_Intro]=[cate_Temp]")
		objConn.execute("ALTER TABLE [blog_Category] DROP COLUMN [cate_Temp]")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_ParentID] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Meta] text")
	End If

	If Not CheckUpdateDB("[comm_Meta]","[blog_Comment]") Then
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Reply] text")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_LastReplyIP] VARCHAR(15)")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_LastReplyTime] datetime default now()")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Yea] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Nay] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Delate] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Meta] text")
	End If
	
	If Not CheckUpdateDB("[mem_Meta]","[blog_Member]") Then
		objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_Meta] text")
	End If

	If Not CheckUpdateDB("[tb_Meta]","[blog_TrackBack]") Then
		objConn.execute("ALTER TABLE [blog_TrackBack] ADD COLUMN [tb_Meta] text")
	End If


End Function





	Dim tmpSng
	Dim objFSO
	Set objFSO=Server.CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(BlogPath & "c_option.asp") Then
		tmpSng=LoadFromFile(BlogPath & "/c_option.asp","utf-8")
		Call SaveValueForSetting(tmpSng,True,"String","ZC_BLOG_VERSION","1.9 Beta1 Build 110401")
		Call SaveToFile(BlogPath & "/c_option.asp",tmpSng,"utf-8",false)
	End If 

	Dim objConn	
	Dim objCat
	'这段是升级数据库的
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(ZC_DATABASE_PATH)

	Call UpdateDateBase()

%>
<p><center>升级成功!</center></p>
<%
Else
%>
<p></p>
<p>请先将原数据库和c_custom.asp与c_option.asp文件分别放入新版程序的对应位置,再将THEME目录,UPLOAD目录和PLUGIN目录覆盖新版的相应目录.</p>
<p><center><input type='submit' value='升级' onclick=''></input></center></p>
<%
End If



%>


</div>
<div id="divFooter"><b><font color="blue">[使用必看]:升级完成后，请通过ftp删除此asp程序文件．</font></b></div>
</form>
</body>
</html>
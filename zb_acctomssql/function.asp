<%

Const TempString="\nSET IDENTITY_INSERT [dbo].[blog_Article] ON\nINSERT INTO [dbo].[blog_Article] (log_ID,log_CateID,log_AuthorID,log_Level,log_Url,log_Title,log_Intro,log_Content,log_IP,log_PostTime,log_CommNums,log_ViewNums,log_TrackBackNums,log_Tag,log_IsTop,log_Yea,log_Nay,log_Ratting,log_Template,log_FullUrl,log_Type,log_Meta) SELECT log_ID,log_CateID,log_AuthorID,log_Level,log_Url,log_Title,log_Intro,log_Content,log_IP,log_PostTime,log_CommNums,log_ViewNums,log_TrackBackNums,log_Tag,log_IsTop,log_Yea,log_Nay,log_Ratting,log_Template,log_FullUrl,log_Type,log_Meta FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_Article]\nSET IDENTITY_INSERT [dbo].[blog_Article] Off\n\nSET IDENTITY_INSERT [dbo].[blog_Category] ON\nINSERT INTO [dbo].[blog_Category] (cate_ID,cate_Name,cate_Order,cate_Intro,cate_Count,cate_URL,cate_ParentID,cate_Template,cate_LogTemplate,cate_FullUrl,cate_Meta) SELECT cate_ID,cate_Name,cate_Order,cate_Intro,cate_Count,cate_URL,cate_ParentID,cate_Template,cate_LogTemplate,cate_FullUrl,cate_Meta FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_Category]\nSET IDENTITY_INSERT [dbo].[blog_Category] Off\n\nSET IDENTITY_INSERT [dbo].[blog_Comment] ON\nINSERT INTO [dbo].[blog_Comment] (comm_ID,log_ID,comm_AuthorID,comm_Author,comm_Content,comm_Email,comm_HomePage,comm_PostTime,comm_IP,comm_Agent,comm_Reply,comm_LastReplyIP,comm_LastReplyTime,comm_Yea,comm_Nay,comm_Ratting,comm_ParentID,comm_IsCheck,comm_Meta) SELECT comm_ID,log_ID,comm_AuthorID,comm_Author,comm_Content,comm_Email,comm_HomePage,comm_PostTime,comm_IP,comm_Agent,comm_Reply,comm_LastReplyIP,comm_LastReplyTime,comm_Yea,comm_Nay,comm_Ratting,comm_ParentID,comm_IsCheck,comm_Meta FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_Comment]\nSET IDENTITY_INSERT [dbo].[blog_Comment] Off\n\nINSERT INTO [dbo].[blog_Config] (conf_Name,conf_Value) SELECT conf_Name,conf_Value FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_Config]\n	\nSET IDENTITY_INSERT [dbo].[blog_Counter] ON\nINSERT INTO [dbo].[blog_Counter](coun_ID,coun_IP,coun_Agent,coun_Refer,coun_PostTime,coun_Content,coun_UserID,coun_PostData,coun_URL,coun_AllRequestHeader,coun_LogName) SELECT coun_ID,coun_IP,coun_Agent,coun_Refer,coun_PostTime,coun_Content,coun_UserID,coun_PostData,coun_URL,coun_AllRequestHeader,coun_LogName FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_Counter]\nSET IDENTITY_INSERT [dbo].[blog_Counter] Off\n\nSET IDENTITY_INSERT [dbo].[blog_Function] ON\nINSERT INTO [dbo].[blog_Function] (fn_ID,fn_Name,fn_FileName,fn_Order,fn_Content,fn_IsSystem,fn_SidebarID,fn_HtmlID,fn_Ftype,fn_MaxLi,fn_Meta) 	SELECT fn_ID,fn_Name,fn_FileName,fn_Order,fn_Content,fn_IsSystem,fn_SidebarID,fn_HtmlID,fn_Ftype,fn_MaxLi,fn_Meta FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_Function]\nSET IDENTITY_INSERT [dbo].[blog_Function] Off\n\n\nSET IDENTITY_INSERT [dbo].[blog_Keyword] ON\nINSERT INTO [dbo].[blog_Keyword] (key_ID,key_Name,key_Intro,key_URL) SELECT key_ID,key_Name,key_Intro,key_URL FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_Keyword]\nSET IDENTITY_INSERT [dbo].[blog_Keyword] Off\n\n\nSET IDENTITY_INSERT [dbo].[blog_Member] ON\nINSERT INTO [dbo].[blog_Member] (mem_ID,mem_Level,mem_Name,mem_Password,mem_Sex,mem_Email,mem_MSN,mem_QQ,mem_HomePage,mem_LastVisit,mem_Status,mem_PostLogs,mem_PostComms,mem_Intro,mem_IP,mem_Count,mem_Template,mem_FullUrl,mem_Guid,mem_Meta) SELECT mem_ID,mem_Level,mem_Name,mem_Password,mem_Sex,mem_Email,mem_MSN,mem_QQ,mem_HomePage,mem_LastVisit,mem_Status,mem_PostLogs,mem_PostComms,mem_Intro,mem_IP,mem_Count,mem_Template,mem_FullUrl,mem_Guid,mem_Meta FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_Member]\nSET IDENTITY_INSERT [dbo].[blog_Member] Off\n\n\nSET IDENTITY_INSERT [dbo].[blog_Tag] ON\nINSERT INTO [dbo].[blog_Tag] (tag_ID,tag_Name,tag_Intro,tag_ParentID,tag_URL,tag_Order,tag_Count,tag_Template,tag_FullUrl,tag_Meta) SELECT tag_ID,tag_Name,tag_Intro,tag_ParentID,tag_URL,tag_Order,tag_Count,tag_Template,tag_FullUrl,tag_Meta FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_Tag]\nSET IDENTITY_INSERT [dbo].[blog_Tag] Off\n\n\nSET IDENTITY_INSERT [dbo].[blog_TrackBack] ON\nINSERT INTO [dbo].[blog_TrackBack] (tb_ID,log_ID,tb_URL,tb_Title,tb_Blog,tb_Excerpt,tb_PostTime,tb_IP,tb_Agent,tb_Meta) SELECT tb_ID,log_ID,tb_URL,tb_Title,tb_Blog,tb_Excerpt,tb_PostTime,tb_IP,tb_Agent,tb_Meta FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_TrackBack]\nSET IDENTITY_INSERT [dbo].[blog_TrackBack] Off\n\n\nSET IDENTITY_INSERT [dbo].[blog_UpLoad] ON\nINSERT INTO [dbo].[blog_UpLoad] (ul_ID,ul_AuthorID,ul_FileSize,ul_FileName,ul_PostTime,ul_Quote,ul_DownNum,ul_FileIntro,ul_DirByTime,ul_Meta) SELECT ul_ID,ul_AuthorID,ul_FileSize,ul_FileName,ul_PostTime,ul_Quote,ul_DownNum,ul_FileIntro,ul_DirByTime,ul_Meta FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""<#MDBPath#>""')...[blog_UpLoad]\nSET IDENTITY_INSERT [dbo].[blog_UpLoad] Off\n"
Dim username,password,userguid
Dim dbtype,dbpath,dbserver,dbname,dbusername,dbpassword
dbtype=Request("dbtype")
dbpath=Request("dbpath")
dbserver=Request("dbserver")
dbname=Request("dbname")
dbusername=Request("dbusername")
dbpassword=Request("dbpassword")


Dim zblogstep
zblogstep=Request.QueryString("step")


If zblogstep="" Then zblogstep=1

Function OpenConnect2(t)
	On Error Resume Next
	Set objConn = Server.CreateObject("ADODB.Connection")
	If t=0 Then
		objConn.Open "Provider=SqlOLEDB;Data Source="&ZC_MSSQL_SERVER&";Initial Catalog="&ZC_MSSQL_DATABASE&";Persist Security Info=True;User ID="&ZC_MSSQL_USERNAME&";Password="&ZC_MSSQL_PASSWORD&";"
	ElseIf t=2 Then
		objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BlogPath & ZC_DATABASE_PATH
	Else
		objConn.Open "Provider=SqlOLEDB;Data Source="&ZC_MSSQL_SERVER&";Persist Security Info=True;User ID="&ZC_MSSQL_USERNAME&";Password="&ZC_MSSQL_PASSWORD&";"
	End If
	If Err.Number=0 Then
		OpenConnect2=True
	Else
		Err.Clear
		Err.Raise 1
	End If
End Function

Function FilterSQL2(ByVal strSQL,f)
	If IsNull(strSQL) Then strSQL=""
	
	If VarType(strSQL)=vbBoolean Then
		If strSQL=True Then strSQL=1 Else strSQL=0
	End If

	
	strSQL=CStr(Replace(strSQL,chr(39),chr(39)&chr(39)))
	If f<>"num" Then 
		strSQL="'"&strSQL&"'"
	Else
		If IsNumeric(strSQL) Then
			strSQL=strSQL
		Else
			strSQL=0
		End If
	End If
	
	FilterSQL2=strSQL
	
End Function


function cleanlog(byref obj)
	on error resume next
	obj.execute "DBCC SHRINKFILE("&dbpath&"_log,0)"&vbcrlf&"DUMP TRANSACTION "&dbpath&" WITH NO_LOG"
	if err.number=0 then ExportLog "事务日志初始化成功" else ExportLog "事务日志初始化失败，可能您没有权限初始化，这并不影响使用，程序将继续运行"
	err.clear
end function


Function ExportLog(str)
	Response.Write "[" & Now & "]" & str & "<br/>"
End Function

Function CheckBoolean(str)

End Function

Sub makeloading(text)
	Response.Write "<script>$(""#loading"").html('" & text & "')</script>"
	Response.Flush
End Sub

Sub ExportError(text)
	Response.Write "<script>$(""#loading"").css(""background"",""red"")</script>"
	ExportLog text
end sub
%>
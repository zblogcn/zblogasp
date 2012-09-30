<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Devo Or Newer
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    批量管理文章插件 - 跳转页
'// 最后修改：   2008-11-18
'// 最后版本:    1.4.1
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!-- #include file="config.asp" -->
<%

Dim strAct
strAct=Request.QueryString("act")

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>4 Then Call ShowError(6) 

Dim strRedirect
strRedirect = "articleList.asp?page="&Request.QueryString("page")&"&cate="&ReQuest("cate")&"&level="&ReQuest("level")&"&istop="&ReQuest("istop")&"&user="&ReQuest("user")&"&tag="&ReQuest("tag")&"&title="&Escape(ReQuest("title"))&""

Select Case strAct

	Case "SkipWarning"
		Call DisableShowWarning

	Case "ShowWarning"
		Call EnableShowWarning

	Case "EnableTagMng"
		Call EnableTagMng

	Case "DisableTagMng"
		Call DisableTagMng

	Case "EnableTagCloud"
		Call EnableTagCloud

	Case "DisableTagCloud"
		Call DisableTagCloud

	Case "EnableTagHint"
		Call EnableTagHint

	Case "DisableTagHint"
		Call DisableTagHint

	Case "BasicEdit"
		Call BasicEdit()

End Select

Function DisableShowWarning
	objConfig.Write "ShowWarning",False
	objConfig.Save
	Response.Redirect "articleList.asp"
End Function

Function EnableShowWarning
	objConfig.Write "ShowWarning",True
	objConfig.Save
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "help.asp"
End Function


Function EnableTagMng
	objConfig.Write "UseTagMng",True
	objConfig.Save
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "help.asp"
End Function

Function DisableTagMng
	objConfig.Write "UseTagMng",False
	objConfig.Save
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "help.asp"
End Function


Function EnableTagCloud
	objConfig.Write "UseTagCloud",True
	objConfig.Save
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "help.asp"
End Function

Function DisableTagCloud
	objConfig.Write "UseTagCloud",False
	objConfig.Save
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "help.asp"
End Function


Function EnableTagHint
	objConfig.Write "UseTagHint",True
	objConfig.Save
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "help.asp"
End Function

Function DisableTagHint
	objConfig.Write "UseTagHint",False
	objConfig.Save
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "help.asp"
End Function


Function BasicEdit()

	Dim cata_ID
	Dim log_Level
	Dim log_Istop
	Dim Log_Author
	Dim Log_AddTag
	Dim Log_RmvTag
	Dim batch_Del
	Dim batch_Del_Success
	Dim suc_Log_ID,err_Log_ID

	cata_ID=Request.Form("MoveCata")
	log_Level=Request.Form("EdtLevel")
	log_Istop=Request.Form("EdtIstop")
	Log_Author=Request.Form("EdtUser")
	Log_AddTag=Request.Form("AddTag")
	Log_RmvTag=Request.Form("RmvTag")
	batch_Del=Request.Form("BatchDel")

	Call CheckParameter(cata_ID,"int",-1)
	Call CheckParameter(log_Level,"int",-1)
	Call CheckParameter(log_Istop,"int",-1)
	Call CheckParameter(log_Author,"int",-1)
	Call CheckParameter(Log_AddTag,"int",-1)
	Call CheckParameter(Log_RmvTag,"int",-1)

	batch_Del=CBool(batch_Del)
	batch_Del_Success=True
	
	Dim Log_ID,aryLog_ID,strLog_ID
	If Request.Form("edtBatch")<>"" Then
		Log_ID=Request.Form("edtBatch")
	Else
		Call SetBlogHint(False,Empty,Empty)
		Call SetBlogHint_Custom("? 提示:没有选择文章.")
		Response.Redirect strRedirect
	End If

	aryLog_ID=split(Log_ID,",")

	Dim i,n
	n=UBound(aryLog_ID,1)
	For i=0 to n
		Call CheckParameter(aryLog_ID(i),"int",-1)
		If aryLog_ID(i)>0 then
			If i<n-1 Then
				strLog_ID=strLog_ID & aryLog_ID(i) & ","
			Else
				strLog_ID=strLog_ID & aryLog_ID(i)
			End If
		End If
	Next
	aryLog_ID = split(strLog_ID,",")
	n=UBound(aryLog_ID)

	If cata_ID=-1 And log_Level=-1 And log_Istop=-1 And log_Author=-1 And Log_AddTag=-1 And Log_RmvTag=-1 And batch_Del=False Then
		Call SetBlogHint(False,Empty,Empty)
		Call SetBlogHint_Custom("&raquo; 提示:没有选择执行方式.")
	ElseIf batch_Del Then

		'plugin node
		

		'If n > Int(ZC_REBUILD_FILE_COUNT/2) Then
		'	Call SetBlogHint(False,Empty,Empty)
		'	Call SetBlogHint_Custom("? 提示:同时执行的文章数量超过 "& Int(ZC_REBUILD_FILE_COUNT/2) &", 无法执行操作.")
		'Else

			For i=0 To n

				If DelArticle(aryLog_ID(i)) Then
					suc_Log_ID=suc_Log_ID & aryLog_ID(i) & ","
				Else
					batch_Del_Success=False
					err_Log_ID=err_Log_ID & aryLog_ID(i) & ","
				End If

			Next

			If batch_Del_Success Then
				'plugin node
				
				Call SetBlogHint(True,True,Empty)
				Call SetBlogHint_Custom("&raquo; 提示:已成功删除ID为 "& suc_Log_ID &" 的文章.")
			Else
				Call SetBlogHint(False,True,Empty)
				Call SetBlogHint_Custom("&raquo; 提示:在删除ID为 "& err_Log_ID &" 的文章时发生错误.")
			End If

		'End If

	Else
		If cata_ID <> -1  Then
			objConn.Execute("UPDATE [blog_Article] SET [log_CateID]="& cata_ID &" WHERE [log_ID] in("& strLog_ID &")")
		End If

		If log_Level <> -1 Then
			objConn.Execute("UPDATE [blog_Article] SET [log_Level]="& log_Level &" WHERE [log_ID] in("& strLog_ID &")")
		End if

		If log_Istop <> -1 Then
			If log_Istop = 1 Then
			objConn.Execute("UPDATE [blog_Article] SET [log_IsTop]=1 WHERE [log_ID] in("& strLog_ID &")")
			End If
			if log_Istop = 0 then
			objConn.Execute("UPDATE [blog_Article] SET [log_IsTop]=0 WHERE [log_ID] in("& strLog_ID &")")
			End If
		End If

		If log_Author <> -1 Then
			objConn.Execute("UPDATE [blog_Article] SET [log_AuthorID]="& log_Author &" WHERE [log_ID] in("& strLog_ID &")")
		End if

		If Log_AddTag <> -1 Then
			Dim AddTag_ObjRS,tmpAddTags
			Set AddTag_ObjRS=objConn.Execute("SELECT [log_ID],[log_Tag] FROM [blog_Article] WHERE [log_ID] in("& strLog_ID &")")
				If (Not AddTag_ObjRS.bof) And (Not AddTag_ObjRS.eof) Then
					Do While Not AddTag_ObjRS.eof
						tmpAddTags=AddTag_ObjRS("log_Tag")
						If Not InStr(tmpAddTags,"{"& Log_AddTag &"}")>0 Then tmpAddTags=tmpAddTags & "{"& Log_AddTag &"}"
						objConn.Execute("UPDATE [blog_Article] SET [log_Tag]='"& tmpAddTags &"' WHERE [log_ID]="& AddTag_ObjRS("log_ID"))
						AddTag_ObjRS.MoveNext
					Loop
				End If
			Set AddTag_ObjRS=Nothing
			Call ScanTagCount("{"&Log_AddTag&"}")
			'If CheckPluginState("EC_HTMLTAGS")=True Then Call EC_HTMLTAGS_BuildPageByTagID(Log_AddTag)
		End if

		If Log_RmvTag <> -1 Then
			Dim RmvTag_ObjRS,tmpRmvTags
			Set RmvTag_ObjRS=objConn.Execute("SELECT [log_ID],[log_Tag] FROM [blog_Article] WHERE [log_ID] in("& strLog_ID &")")
				If (Not RmvTag_ObjRS.bof) And (Not RmvTag_ObjRS.eof) Then
					Do While Not RmvTag_ObjRS.eof
						tmpRmvTags=RmvTag_ObjRS("log_Tag")
						tmpRmvTags=Replace(tmpRmvTags,"{"& Log_RmvTag &"}","")
						objConn.Execute("UPDATE [blog_Article] SET [log_Tag]='"& tmpRmvTags &"' WHERE [log_ID]="& RmvTag_ObjRS("log_ID"))
						RmvTag_ObjRS.MoveNext
					Loop
				End If
			Set RmvTag_ObjRS=Nothing
			Call ScanTagCount("{"&Log_RmvTag&"}")
			'If CheckPluginState("EC_HTMLTAGS")=True Then Call EC_HTMLTAGS_BuildPageByTagID(Log_RmvTag)
		End if

		'If n > Int(ZC_REBUILD_FILE_COUNT/3) Then
		'	Call SetBlogHint(True,True,True)
		'	Call SetBlogHint_Custom("? 提示:单次数量大于 "&Int(ZC_REBUILD_FILE_COUNT/3)&" 的批量管理请自行 ""文件重建"".")
		'Else

			'plugin node
			For Each sAction_Plugin_FileReBuild_Begin in Action_Plugin_FileReBuild_Begin
				If Not IsEmpty(sAction_Plugin_FileReBuild_Begin) Then Call Execute(sAction_Plugin_FileReBuild_Begin)
				If bAction_Plugin_FileReBuild_Begin=True Then Exit Function
			Next

			Dim objArticle
			For i=0 To n
				Set objArticle=New TArticle
				If objArticle.LoadInfoByID(aryLog_ID(i)) Then
					objArticle.DelFile()
				End If
				Set objArticle=Nothing

				If BuildArticle(aryLog_ID(i),True,True) Then
					suc_Log_ID=suc_Log_ID & aryLog_ID(i) & ","
				Else
					batch_Del_Success=False
					err_Log_ID=err_Log_ID & aryLog_ID(i) & ","
				End If
			Next

			If batch_Del_Success Then
				Call SetBlogHint(True,True,Empty)
				Call SetBlogHint_Custom("&raquo; 提示:已成功对ID为 "& suc_Log_ID &" 的文章进行重建.")

				'plugin node

			Else
				Call SetBlogHint(False,True,Empty)
				Call SetBlogHint_Custom("&raquo; 提示:在重建ID为 "& err_Log_ID &" 的文章时发生错误.")
			End If

		'End If

	End If

	Response.Redirect strRedirect


End Function

Call System_Terminate()

'If Err.Number<>0 then
'	Call ShowError(0)
'End If
%>
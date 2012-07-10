<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 Devo
'// 插件制作:    haphic(http://haphic.com)
'// 备    注:    Deep09 参数设定
'// 最后修改：   2008-2-9
'// 最后版本:    0.4
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%

'----------------------------------------------------------
	Response.Clear
	Response.ExpiresAbsolute   =   Now()   -   1
	Response.Expires   =   0
	Response.CacheControl   =   "no-cache"
'----------------------------------------------------------

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("STACentre")=False Then Call ShowError(48)

Dim strAct
strAct=Request.QueryString("act")

Select Case strAct

	Case "view"
		Call ExportUrlsPreview()

	Case "save"
		Call ExportSaveSetting()

	Case "build"
		Call ExportPageRebuild()

End Select

'=====================================================================================
' 函数调用部分
'=====================================================================================

'预览静态页路径
Function ExportUrlsPreview()

	Dim i,j,k,l,m

	Dim strType
	strType=Request.QueryString("type")

	Dim listCount
	listCount=8

	Dim Color(7)
	Color(0)="#CCC"
	Color(1)="#999"
	Color(2)="#666"
	Color(3)="#333"
	Color(4)="#000"
	Color(5)="#000"
	Color(6)="#000"
	Color(7)="#000"

	Select Case strType

		Case "Categorys"
			var_STACentre_Dir_Categorys_Enable=Request.Form("STACentre_Dir_Categorys_Enable")
			var_STACentre_Dir_Categorys_Regex=Request.Form("STACentre_Dir_Categorys_Regex")
			var_STACentre_Dir_Categorys_Anonymous=Request.Form("STACentre_Dir_Categorys_Anonymous")
			var_STACentre_Dir_Categorys_FCate=Request.Form("STACentre_Dir_Categorys_FCate")
			j=1
			For i=1 To UBound(Categorys)
				If j>listCount Then Exit For
				If IsObject(Categorys(i)) Then
					Set k = New STACentre_Categorys
						If k.LoadInfoByID(Categorys(i).ID) Then
							m=m & "<tr style=""color:"& Color(listCount-j) &";""><td width=""100"">" & Categorys(i).HTMLName & "</td><td>" & k.Url & "</td></tr>"
							j=j+1
						End If
					Set k = Nothing
				End If
			Next

		Case "Tags"
			var_STACentre_Dir_Tags_Enable=Request.Form("STACentre_Dir_Tags_Enable")
			var_STACentre_Dir_Tags_Regex=Request.Form("STACentre_Dir_Tags_Regex")
			var_STACentre_Dir_Tags_Anonymous=Request.Form("STACentre_Dir_Tags_Anonymous")
			j=1
			For i=1 To UBound(Tags)
				If j>listCount Then Exit For
				If IsObject(Tags(i)) Then
					Set k = New STACentre_Tags
						If k.LoadInfoByID(Tags(i).ID) Then
							m=m & "<tr style=""color:"& Color(listCount-j) &";""><td width=""100"">" & Tags(i).HTMLName & "</td><td>" & k.Url & "</td></tr>"
							j=j+1
						End If
					Set k = Nothing
				End If
			Next

		Case "Authors"
			var_STACentre_Dir_Authors_Enable=Request.Form("STACentre_Dir_Authors_Enable")
			var_STACentre_Dir_Authors_Regex=Request.Form("STACentre_Dir_Authors_Regex")
			var_STACentre_Dir_Authors_Anonymous=Request.Form("STACentre_Dir_Authors_Anonymous")
			j=1
			For i=1 To UBound(Users)
				If j>listCount Then Exit For
				If IsObject(Users(i)) Then
					Set k = New STACentre_Authors
						If k.LoadInfoByID(Users(i).ID) Then
							m=m & "<tr style=""color:"& Color(listCount-j) &";""><td width=""100"">" & Users(i).Name & "</td><td>" & k.Url & "</td></tr>"
							j=j+1
						End If
					Set k = Nothing
				End If
			Next

		Case "Archives"
			var_STACentre_Dir_Archives_Enable=Request.Form("STACentre_Dir_Archives_Enable")
			var_STACentre_Dir_Archives_Regex=Request.Form("STACentre_Dir_Archives_Regex")
			var_STACentre_Dir_Archives_Anonymous=Request.Form("STACentre_Dir_Archives_Anonymous")
			var_STACentre_Dir_Archives_Format=Request.Form("STACentre_Dir_Archives_Format")
			l=STACentre_GetArchivesList()
			j=1
			For i=1 To UBound(l)
				If j>listCount Then Exit For
				If IsDate(l(i)) Then
					Set k = New STACentre_Archives
						If k.LoadInfoByID(l(i)) Then
							m=m & "<tr style=""color:"& Color(listCount-j) &";""><td width=""100"">" & l(i) & "</td><td>" & k.Url & "</td></tr>"
							j=j+1
						End If
					Set k = Nothing
				End If
			Next
	End Select

	If Err.Number=0 Then Response.Write m

End Function


'保存设置
Function ExportSaveSetting()

	Dim strContent,tmpContent

	strContent=LoadFromFile(Server.MapPath("config.asp"),"utf-8")
	tmpContent=strContent

	Dim strZC_STACentre_Dir_Categorys_Enable
	strZC_STACentre_Dir_Categorys_Enable=Request.Form("STACentre_Dir_Categorys_Enable")
	Call SaveValueForSetting(strContent,True,"Boolean","STACentre_Dir_Categorys_Enable",strZC_STACentre_Dir_Categorys_Enable)

	Dim strZC_STACentre_Dir_Categorys_Regex
	strZC_STACentre_Dir_Categorys_Regex=Request.Form("STACentre_Dir_Categorys_Regex")
	Call SaveValueForSetting(strContent,True,"String","STACentre_Dir_Categorys_Regex",strZC_STACentre_Dir_Categorys_Regex)

	Dim strZC_STACentre_Dir_Categorys_Anonymous
	strZC_STACentre_Dir_Categorys_Anonymous=Request.Form("STACentre_Dir_Categorys_Anonymous")
	Call SaveValueForSetting(strContent,True,"Boolean","STACentre_Dir_Categorys_Anonymous",strZC_STACentre_Dir_Categorys_Anonymous)

	Dim strZC_STACentre_Dir_Categorys_FCate
	strZC_STACentre_Dir_Categorys_FCate=Request.Form("STACentre_Dir_Categorys_FCate")
	Call SaveValueForSetting(strContent,True,"Boolean","STACentre_Dir_Categorys_FCate",strZC_STACentre_Dir_Categorys_FCate)


	Dim strZC_STACentre_Dir_Tags_Enable
	strZC_STACentre_Dir_Tags_Enable=Request.Form("STACentre_Dir_Tags_Enable")
	Call SaveValueForSetting(strContent,True,"Boolean","STACentre_Dir_Tags_Enable",strZC_STACentre_Dir_Tags_Enable)

	Dim strZC_STACentre_Dir_Tags_Regex
	strZC_STACentre_Dir_Tags_Regex=Request.Form("STACentre_Dir_Tags_Regex")
	Call SaveValueForSetting(strContent,True,"String","STACentre_Dir_Tags_Regex",strZC_STACentre_Dir_Tags_Regex)

	Dim strZC_STACentre_Dir_Tags_Anonymous
	strZC_STACentre_Dir_Tags_Anonymous=Request.Form("STACentre_Dir_Tags_Anonymous")
	Call SaveValueForSetting(strContent,True,"Boolean","STACentre_Dir_Tags_Anonymous",strZC_STACentre_Dir_Tags_Anonymous)


	Dim strZC_STACentre_Dir_Authors_Enable
	strZC_STACentre_Dir_Authors_Enable=Request.Form("STACentre_Dir_Authors_Enable")
	Call SaveValueForSetting(strContent,True,"Boolean","STACentre_Dir_Authors_Enable",strZC_STACentre_Dir_Authors_Enable)

	Dim strZC_STACentre_Dir_Authors_Regex
	strZC_STACentre_Dir_Authors_Regex=Request.Form("STACentre_Dir_Authors_Regex")
	Call SaveValueForSetting(strContent,True,"String","STACentre_Dir_Authors_Regex",strZC_STACentre_Dir_Authors_Regex)

	Dim strZC_STACentre_Dir_Authors_Anonymous
	strZC_STACentre_Dir_Authors_Anonymous=Request.Form("STACentre_Dir_Authors_Anonymous")
	Call SaveValueForSetting(strContent,True,"Boolean","STACentre_Dir_Authors_Anonymous",strZC_STACentre_Dir_Authors_Anonymous)


	Dim strZC_STACentre_Dir_Archives_Enable
	strZC_STACentre_Dir_Archives_Enable=Request.Form("STACentre_Dir_Archives_Enable")
	Call SaveValueForSetting(strContent,True,"Boolean","STACentre_Dir_Archives_Enable",strZC_STACentre_Dir_Archives_Enable)

	Dim strZC_STACentre_Dir_Archives_Regex
	strZC_STACentre_Dir_Archives_Regex=Request.Form("STACentre_Dir_Archives_Regex")
	Call SaveValueForSetting(strContent,True,"String","STACentre_Dir_Archives_Regex",strZC_STACentre_Dir_Archives_Regex)

	Dim strZC_STACentre_Dir_Archives_Anonymous
	strZC_STACentre_Dir_Archives_Anonymous=Request.Form("STACentre_Dir_Archives_Anonymous")
	Call SaveValueForSetting(strContent,True,"Boolean","STACentre_Dir_Archives_Anonymous",strZC_STACentre_Dir_Archives_Anonymous)

	Dim strZC_STACentre_Dir_Archives_Format
	strZC_STACentre_Dir_Archives_Format=Request.Form("STACentre_Dir_Archives_Format")
	Call SaveValueForSetting(strContent,True,"String","STACentre_Dir_Archives_Format",strZC_STACentre_Dir_Archives_Format)

	If Err.Number=0 Then
		If strContent<>tmpContent Then
			Call SaveToFile(Server.MapPath("config.asp"),strContent,"utf-8",False)
			Call SaveToFile(Server.MapPath("progress.txt"),0,"utf-8",True)
			Response.Write "true"
		Else
			Response.Write "false"
		End If
	End If

End Function


'重建静态页
Function ExportPageRebuild()

	Dim i,j,k,l,a,b
	Dim aryCommands()

			ReDim Preserve aryCommands(0)
			aryCommands(0)="Call STACentre_ClearAllDirsByHistory()"

			ReDim Preserve aryCommands(1)
			aryCommands(1)="Call ClearGlobeCache()"

			ReDim Preserve aryCommands(2)
			aryCommands(2)="Call LoadGlobeCache()"

			ReDim Preserve aryCommands(3)
			aryCommands(3)="Call MakeBlogReBuild_Core()"

			j=4
			For i=1 To UBound(Categorys)
				If IsObject(Categorys(i)) Then
					ReDim Preserve aryCommands(j)
					aryCommands(j)="Call STACentre_BuildPageByCateID("& Categorys(i).ID &",True)"
					j=j+1
				End If
			Next

			For i=1 To UBound(Tags)
				If IsObject(Tags(i)) Then
					ReDim Preserve aryCommands(j)
					aryCommands(j)="Call STACentre_BuildPageByTagID("& Tags(i).ID &",True)"
					j=j+1
				End If
			Next

			For i=1 To UBound(Users)
				If IsObject(Users(i)) Then
					ReDim Preserve aryCommands(j)
					aryCommands(j)="Call STACentre_BuildPageByAuthorID("& Users(i).ID &",True)"
					j=j+1
				End If
			Next

			l=STACentre_GetArchivesList()
			For i=1 To UBound(l)
				If IsDate(l(i)) Then
					ReDim Preserve aryCommands(j)
					aryCommands(j)="Call STACentre_BuildPageByPostTime("""& l(i) &""",True)"
					j=j+1
				End If
			Next

			a=aryCommands
			b=UBound(a)

	Erase aryCommands

	k=Request.QueryString("tasknum")
	Call CheckParameter(k,"int",0)

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

	If k=0 Then
		If fso.FileExists(Server.MapPath("progress.txt")) Then
			k=LoadFromFile(Server.MapPath("progress.txt"),"utf-8")
			Call CheckParameter(k,"int",0)
		End If
	End If

	If k>b Then
		If fso.FileExists(Server.MapPath("progress.txt")) Then
			fso.DeleteFile(Server.MapPath("progress.txt"))
		End If
		Response.Write b+1 &"/"& b+1
		Set fso = Nothing
		Exit Function
	End If

	Set fso = Nothing

	Call Execute(a(k))

	If Err.Number=0 Then
		Call SaveToFile(Server.MapPath("progress.txt"),k,"utf-8",True)
		Response.Write k &"/"& b+1 &"/"& a(k)
	End If


End Function

'=====================================================================================
Call System_Terminate()
%>


<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.7
'// 插件制作:    
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 

If CheckPluginState("Totoro")=False Then Call ShowError(48)
%>
<%

Dim act,delid
act=Request.Form("act")
delid=Request.Form("id")
Dim strContent
Dim strZC_TOTORO_BADWORD_LIST,StrTMP,NEW_BADWORD,bolTOTORO_DEL_DIRECTLY
strContent=LoadFromFile(BlogPath & "/PLUGIN/totoro/include.asp","utf-8")
Call LoadValueForSetting(strContent,True,"String","TOTORO_BADWORD_LIST",strZC_TOTORO_BADWORD_LIST)
Call LoadValueForSetting(strContent,True,"Boolean","TOTORO_DEL_DIRECTLY",bolTOTORO_DEL_DIRECTLY)
If act="delcm" then

	Dim objComment
	Set objComment=New TComment
	If objComment.LoadInfobyID(delid) Then
	
		StrTMP=TOTORO_checkStr(objComment.HomePage & "|" & objComment.Content,strZC_TOTORO_BADWORD_LIST)
		strZC_TOTORO_BADWORD_LIST=strZC_TOTORO_BADWORD_LIST & StrTMP
		NEW_BADWORD=StrTMP
		Response.Write Totoro_dealIt(objComment,bolTOTORO_DEL_DIRECTLY)

	End If		
		
Elseif act="deltb" then

	Dim objTrackBack
	Set objTrackBack=New TTrackBack
	If objTrackBack.LoadInfobyID(delid) Then
	
		StrTMP=TOTORO_checkStr(objTrackBack.URL & "|" & objTrackBack.Excerpt,strZC_TOTORO_BADWORD_LIST)
		strZC_TOTORO_BADWORD_LIST=strZC_TOTORO_BADWORD_LIST & StrTMP
		NEW_BADWORD=StrTMP
		Response.Write Totoro_dealIt(objTrackBack,bolTOTORO_DEL_DIRECTLY)
	
	End If
	
End If

If left(strZC_TOTORO_BADWORD_LIST,1)="|" then strZC_TOTORO_BADWORD_LIST=Right(strZC_TOTORO_BADWORD_LIST, Len(strZC_TOTORO_BADWORD_LIST) - 1)
Call SaveValueForSetting(strContent,True,"String","TOTORO_BADWORD_LIST",strZC_TOTORO_BADWORD_LIST)
Call SaveToFile(BlogPath & "/PLUGIN/totoro/include.asp",strContent,"utf-8",False)
'If NEW_BADWORD<>"" Then Response.write ",TotoroⅡ新增下列黑词： " & Right(NEW_BADWORD, Len(NEW_BADWORD) - 1)

%>
<%
Function TOTORO_checkStr(strToCheck,BADWORD_LIST)
		Dim objReg,objMatches,Match
		Set objReg = New RegExp
		objReg.IgnoreCase = True
		objReg.Global = True
		objReg.Pattern = "http://([\w-]+\.)+[\w-]+"
		Set objMatches = objReg.Execute(strToCheck)
		For Each Match In objMatches
			If Totoro_checkNewBadWord(Match.Value,BADWORD_LIST & TOTORO_checkStr) then
				TOTORO_checkStr=TOTORO_checkStr & "|" & Right(Match.Value, Len(Match.Value) - 7)
			End if
		Next
		Set objReg = Nothing
		Set objMatches = Nothing
		Set Match = Nothing
End Function

Function Totoro_checkNewBadWord(content,BADWORD_LIST)

	Totoro_checkNewBadWord=True
	Dim i,j
	j=0
    Dim strFilter
    strFilter = Split(BADWORD_LIST, "|")
	For i = 0 To UBound(strFilter)
		If strFilter(i)<>"" Then
			If InStr (LCase(content), LCase(strFilter(i))) > 0 Then
				Totoro_checkNewBadWord=False
				Exit For
			End If
		End If
    Next

End Function


Function Totoro_dealIt(objToDeal,bolDel)

	Dim logId
	logId=objToDeal.log_ID

	If bolDel Then
		If objToDeal.Del() Then Totoro_dealIt = "删除成功"
	Else
		objToDeal.log_ID=-1-objToDeal.log_ID
		If objToDeal.Post Then Totoro_dealIt = "已加入审核"
	End If
	
	Call BuildArticle(logId,False,False)
	Call SetBlogHint(Null,True,Null)
	Set objToDeal = Nothing	
	
End Function
%>
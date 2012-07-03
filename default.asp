<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    default.asp
'// 开始时间:    2004.07.25
'// 最后修改:    
'// 备    注:    主页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_function_md5.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%



Dim cca,ccb,cci

ReDim cca(10000)
ReDim ccb(10000)

Application.Lock
cca=Application(ZC_BLOG_CLSID & "STOPCCA")
ccb=Application(ZC_BLOG_CLSID & "STOPCCB")
Application.UnLock

If IsArray(cca)=False Then
ReDim cca(10000)
ReDim ccb(10000)
Application.Lock
Application(ZC_BLOG_CLSID & "STOPCCT")=Now()
Application.UnLock
End If


For cci=0 To 9999 
	If cca(cci)="" Then
		cca(cci)=Request.ServerVariables("Remote_Addr")
		Exit for
	End If 
Next

For cci=0 To 9999
	If cca(cci)=Request.ServerVariables("Remote_Addr") Then
		ccb(cci)=ccb(cci)+1
		'同一IP超过1000次就被屏蔽
		If ccb(cci)>1000 Then
			response.end
		End If 
	End If
Next

Application.Lock
If DateDiff("d", Now(), Application(ZC_BLOG_CLSID & "STOPCCT"))<0 Then
	ReDim cca(10000)
	ReDim ccb(10000)
End If 
Application(ZC_BLOG_CLSID & "STOPCCA")=cca
Application(ZC_BLOG_CLSID & "STOPCCB")=ccb
Application.UnLock




If (InStr(LCase(Request.ServerVariables("HTTP_ACCEPT")),"text/vnd.wap.wml") > 0) And (InStr(LCase(Request.ServerVariables("HTTP_ACCEPT")),"text/html") = 0)  Then Response.Redirect "wap.asp"

'向导部分wizard
If ZC_DATABASE_PATH="data/zblog.mdb" Then Response.Redirect "wizard.asp?verify=" & MD5(ZC_DATABASE_PATH & Replace(LCase(Request.ServerVariables("PATH_TRANSLATED")),"default.asp",""))

Call System_Initialize_WithOutDB()

'plugin node
For Each sAction_Plugin_Default_Begin in Action_Plugin_Default_Begin
	If Not IsEmpty(sAction_Plugin_Default_Begin) Then Call Execute(sAction_Plugin_Default_Begin)
Next

Dim ArtList
Set ArtList=New TArticleList

ArtList.LoadCache

ArtList.template="DEFAULT"

If ArtList.ExportByCache("","","","","","") Then

	ArtList.Build

	Response.Write ArtList.html

End If

'plugin node
For Each sAction_Plugin_Default_End in Action_Plugin_Default_End
	If Not IsEmpty(sAction_Plugin_Default_End) Then Call Execute(sAction_Plugin_Default_End)
Next

Call System_Terminate_WithOutDB()

%><!-- <%=RunTime()%>ms --><%
If Err.Number<>0 then
	Call ShowError(0)
End If
%>
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

Response.ExpiresAbsolute = FormatDateTime(Now()) - 1           
Response.Expires = 0
Response.CacheControl = "no-cache"   

Dim strContent

Dim b
b=False
Dim fso,f
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(BlogPath & "zb_users\CACHE\statistic.asp")=True Then
	If DateDiff("h",fso.GetFile(BlogPath & "zb_users\CACHE\statistic.asp").DateLastModified,Now)>24 Then
		b=True
	Else
		strContent=LoadFromFile(BlogPath & "zb_users\CACHE\statistic.asp","utf-8")
	End If
Else
	b=True
End If

If IsEmpty(Request.QueryString("reload"))=False Then
	b=True
End If

If b=True Then strContent=RefreshStatistic()

strContent=Replace(strContent,"<"&"%=BlogUser",BlogUser.Name& "  (" & ZVA_User_Level_Name(BlogUser.Level)& ")")
strContent=Replace(strContent,"<"&"%=Theme",GetNameFormTheme(ZC_BLOG_THEME)& "  / " & ZC_BLOG_CSS& ".css")
strContent=Replace(strContent,"<"&"%=Version",ZC_BLOG_VERSION)
strContent=Replace(strContent,"<"&"%=BlogHost",BlogHost)
Response.Write strContent
Set Fso=Nothing
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
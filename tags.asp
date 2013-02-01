<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
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
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%
Dim html

Call System_Initialize()

'plugin node
For Each sAction_Plugin_Tags_Begin in Action_Plugin_Tags_Begin
	If Not IsEmpty(sAction_Plugin_Tags_Begin) Then Call Execute(sAction_Plugin_Tags_Begin)
Next

Call GetTags()

Dim objArticle
Set objArticle=New TArticle

objArticle.LoadCache

objArticle.Title="TagCloud"

Dim Tag
Dim strTagCloud()
Dim i,j

For Each Tag in Tags
	If IsObject(Tag) Then 
		If Tag.Count<>0 Then
			i=TagCloud(Tag.Count)
			ReDim Preserve strTagCloud(j+1)
			strTagCloud(j) = "<a href=""" & Tag.Url &""" title=""" & Tag.Count & """  class=""tag-name tag-name-size-"&i&""">" &Tag.name & "</a>"
		End If 
	End If
	j=j+1
Next

objArticle.FType=ZC_POST_TYPE_PAGE
objArticle.Content="<div class=""tags-cloud"">"&Join(strTagCloud)&"</div>"
objArticle.Title="TagCloud"
objArticle.FullRegex="{%host%}/{%alias%}.html"

If GetTemplate("TEMPLATE_TAGS")<>empty Then
	objArticle.template="TAGS"
End If

If objArticle.Export(ZC_DISPLAY_MODE_SYSTEMPAGE) Then
	objArticle.Build
	html=objArticle.html
	Response.Write html
End If

'plugin node
For Each sAction_Plugin_Tags_End in Action_Plugin_Tags_End
	If Not IsEmpty(sAction_Plugin_Tags_End) Then Call Execute(sAction_Plugin_Tags_End)
Next

Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
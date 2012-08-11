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
<% 'On Error Resume Next %>
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
Call System_Initialize()

'plugin node
For Each sAction_Plugin_Tags_Begin in Action_Plugin_Tags_Begin
	If Not IsEmpty(sAction_Plugin_Tags_Begin) Then Call Execute(sAction_Plugin_Tags_Begin)
Next

TemplateTagsDic.Item("ZC_BLOG_HOST")=GetCurrentHost()

Call GetTags()

Dim ArtList
Set ArtList=New TArticleList

ArtList.LoadCache


If GetTemplate("TEMPLATE_TAGS")<>empty Then
ArtList.template="TAGS"
End If

ArtList.Title="TagCloud"

Dim Tag
Dim strTagCloud()
Dim i,j

Dim objRS
Set objRS=objConn.Execute("SELECT [tag_ID] FROM [blog_Tag] ORDER BY [tag_Name] ASC")
If (Not objRS.bof) And (Not objRS.eof) Then
	Do While Not objRS.eof

		If Tags(objRS("tag_ID")).Count<=50 Then
			i=Tags(objRS("tag_ID")).Count*4
		ElseIf Tags(objRS("tag_ID")).Count>50 And Tags(objRS("tag_ID")).Count<=100 Then
			i=Tags(objRS("tag_ID")).Count*2
		ElseIf Tags(objRS("tag_ID")).Count>100 And Tags(objRS("tag_ID")).Count<=200 Then
			i=Tags(objRS("tag_ID")).Count*1.5
		ElseIf Tags(objRS("tag_ID")).Count>200 Then
			i=Tags(objRS("tag_ID")).Count*1
		End If

		ReDim Preserve strTagCloud(j+1)
		strTagCloud(j) = "<span style='font-family:verdana,sans-serif;line-height:150%;font-size:"& (100 + (i)) &"%;margin:10px;'><a title='" & Tags(objRS("tag_ID")).Count & "' href='" & Tags(objRS("tag_ID")).Url &"'>" & Tags(objRS("tag_ID")).name & "</a></span> "
		j=j+1
		objRS.MoveNext
	Loop
End If
objRS.Close
Set objRS=Nothing

ArtList.SetVar "CUSTOM_TAGS",Join(strTagCloud)

ArtList.SetVar "CUSTOM_TAGS_TITLE","TagCloud"

ArtList.Build

ArtList.SetVar "CUSTOM_TAGS_DESC","TagCloud"

Response.Write ArtList.html

'plugin node
For Each sAction_Plugin_Tags_End in Action_Plugin_Tags_End
	If Not IsEmpty(sAction_Plugin_Tags_End) Then Call Execute(sAction_Plugin_Tags_End)
Next

%><!-- <%=RunTime()%>ms --><%
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
<%
'注册插件
Call RegisterPlugin("KindEditor","ActivePlugin_KindEditor")
'挂口部分
Function ActivePlugin_KindEditor()
	Call Add_Action_Plugin("Action_Plugin_ArticleEdt_Begin","RedirectURL()")
	'Add_Action_Plugin("Action_Plugin_Edit_ueditor_Begin","Server.Transfer(""..\..\zb_users\PLUGIN\KindEditor\edit_kindeditor.asp""):Response.End")
End Function

Function UseKindEditor()
	'Server.Transfer ("../../zb_users/plugin/KindEditor/edit_kindeditor.asp")
	Server.Transfer ("..\..\zb_users\PLUGIN\KindEditor\edit_kindeditor.asp")
	Response.End
End Function

Function RedirectURL()
	If Request.QueryString("type")="Page" Then
		If IsEmpty(Request.QueryString("id")) Then
			Response.Redirect "../zb_users/plugin/KindEditor/edit_kindeditor.asp?type=Page"
		Else
			Response.Redirect "../zb_users/plugin/KindEditor/edit_kindeditor.asp?type=Page&id="&Request.QueryString("id")
		End If
	Else
		If IsEmpty(Request.QueryString("id")) Then
			Response.Redirect "../zb_users/plugin/KindEditor/edit_kindeditor.asp"
		Else
			Response.Redirect "../zb_users/plugin/KindEditor/edit_kindeditor.asp?id="&Request.QueryString("id")
		End If
	End If
End Function
%>
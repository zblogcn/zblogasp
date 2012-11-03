<%
'************************************
' Powered by ThemePluginEditor<%=版本号%>
' zsx http://www.zsxsoft.com
'************************************
Dim <%=主题名%>_theme(1)
<%=主题名%>_theme(0)=Array(<%=文件注释数组%>)
<%=主题名%>_theme(1)=Array(<%=文件名数组%>)

Call RegisterPlugin("<%=主题名%>","ActivePlugin_<%=主题名%>")

Function ActivePlugin_<%=主题名%>()
	'如果插件需要include代码，则直接在这里加。
	'这里加文件管理
	If CheckPluginState("FileManage") Then
		Call Add_Action_Plugin("Action_Plugin_FileManage_ExportInformation_NotFound","<%=主题名%>_exportdetail(""{path}"",""{f}"")")
	End If
    '这里是给后台加管理按钮
	Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu","<script type='text/javascript'>$(document).ready(function(){$(""#theme-<%=主题名%> .theme-name"").append('<input class=""button"" style=""float:right;margin:0;padding-left:10px;padding-right:10px;"" type=""button"" value=""配置"" onclick=""location.href=\'"&BlogHost&"/zb_users/theme/<%=主题名%>/plugin/editor.asp\'"">')})</script>")
End Function

Function <%=主题名%>_exportdetail(p,f)
	On Error Resume Next
	dim z,k,l,i
	z=LCase(f)
	k=LCase(p)
	l=lcase(blogpath)
	k=IIf(Right(k,1)="\",Left(k,Len(k)-1),k)
	l=IIf(Right(l,1)="\",Left(l,Len(l)-1),l)
	if k=l & "\zb_users\theme\<%=主题名%>\include" Then
		For i=0 To Ubound(<%=主题名%>_theme(1))
			If <%=主题名%>_theme(1)(i)=z Then <%=主题名%>_exportdetail=<%=主题名%>_theme(0)(i)
		Next
	End If
End Function
%>
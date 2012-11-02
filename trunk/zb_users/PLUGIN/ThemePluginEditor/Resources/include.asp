<%
'************************************
' Powered by ThemePluginEditor
' zsx http://www.zsxsoft.com
'************************************
Dim default_theme(1)
default_theme(0)=Array(<%=文件注释数组%>)
default_theme(1)=Array(<%=文件名数组%>)

Call RegisterPlugin("default","ActivePlugin_default")

Function ActivePlugin_default()
	'如果插件需要include代码，则直接在这里加。
	'这里加文件管理
	If CheckPluginState("FileManage") Then
		Call Add_Action_Plugin("Action_Plugin_FileManage_ExportInformation_NotFound","default_exportdetail(""{path}"",""{f}"")")
	End If
End Function

Function default_exportdetail(p,f)
	On Error Resume Next
	dim z,k,l,i
	z=LCase(f)
	k=LCase(p)
	l=lcase(blogpath)
	k=IIf(Right(k,1)="\",Left(k,Len(k)-1),k)
	l=IIf(Right(l,1)="\",Left(l,Len(l)-1),l)
	if k=l & "\zb_users\theme\<%=主题名%>\include" Then
		For i=0 To Ubound(default_theme(1))
			If default_theme(1)(i)=z Then default_exportdetail=default_theme(0)(i)
		Next
	End If
End Function
%>
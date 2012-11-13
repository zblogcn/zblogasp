<%
'************************************
' Powered by ThemePluginEditor 1.1
' zsx http://www.zsxsoft.com
'************************************
Dim default_theme(1)
default_theme(0)=Array("头图")
default_theme(1)=Array("bg-nav.jpg")

Call RegisterPlugin("default","ActivePlugin_default")

Function ActivePlugin_default()
	'如果插件需要include代码，则直接在这里加。
    Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(1,"主题配置",BlogHost & "/zb_users/theme/default/plugin/editor.asp","adefaultManage",""))
	'这里加文件管理
	If CheckPluginState("FileManage") Then
		Call Add_Action_Plugin("Action_Plugin_FileManage_ExportInformation_NotFound","default_exportdetail(""{path}"",""{f}"")")
	End If
    '这里是给后台加管理按钮
    'If BlogVersion<=121028 Then Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu","<script type='text/javascript'>$(document).ready(function(){$(""#theme-default .theme-name"").append('<input class=""button"" style=""float:right;margin:0;padding-left:10px;padding-right:10px;"" type=""button"" value=""配置"" onclick=""location.href=\'"&BlogHost&"/zb_users/theme/default/plugin/editor.asp\'"">')})</sc"&"ript>")
End Function

Function default_exportdetail(p,f)
	On Error Resume Next
	dim z,k,l,i
	z=LCase(f)
	k=LCase(p)
	l=lcase(blogpath)
	k=IIf(Right(k,1)="\",Left(k,Len(k)-1),k)
	l=IIf(Right(l,1)="\",Left(l,Len(l)-1),l)
	if k=l & "\zb_users\theme\default\include" Then
		For i=0 To Ubound(default_theme(1))
			If default_theme(1)(i)=z Then default_exportdetail=default_theme(0)(i)
		Next
	End If
End Function
%>
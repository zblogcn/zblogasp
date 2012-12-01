<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Devo Or Newer
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    批量管理文章插件 - 跳转页
'// 最后修改：   2008-10-24
'// 最后版本:    1.4
'///////////////////////////////////////////////////////////////////////////////
Call System_Initialize()

Dim ShowWarning

Dim UseTagMng
Dim UseTagCloud
Dim UseTagHint
Dim objConfig
Set objConfig=New TConfig
objConfig.Load "BatchArticles"
ShowWarning=CBool(objConfig.Read("ShowWarning"))
UseTagMng=CBool(objConfig.Read("UseTagMng"))
UseTagCloud=CBool(objConfig.Read("UseTagCloud"))
UseTagHint=CBool(objConfig.Read("UseTagHint"))

Function ExportSearch(table,value)
	If ZC_MSSQL_ENABLE Then
		ExportSearch="( (CHARINDEX('" & value &"',["&table&"]))<>0)"
	Else
		ExportSearch="( (InStr(1,LCase(["&table&"]),LCase('" & value &"'),0)<>0) )"
	End If
End Function

%>
<%
'****************************************
' MenuManage 子菜单
'****************************************
Function MenuManage_SubMenu(id)
	Dim aryName,aryPath,aryFloat,aryInNewWindow,i
	aryName=Array("首页")
	aryPath=Array("main.asp")
	aryFloat=Array("m-left")
	aryInNewWindow=Array(False)
	For i=0 To Ubound(aryName)
		MenuManage_SubMenu=MenuManage_SubMenu & MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i))
	Next
End Function
%>
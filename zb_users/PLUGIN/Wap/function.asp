<%
'****************************************
' Wap 子菜单
'****************************************
Function Wap_SubMenu(id)
	Dim aryName,aryPath,aryFloat,aryInNewWindow,i
	aryName=Array("首页")
	aryPath=Array("main.asp")
	aryFloat=Array("m-left")
	aryInNewWindow=Array(False)
	For i=0 To Ubound(aryName)
		Wap_SubMenu=Wap_SubMenu & MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i))
	Next
End Function
%>
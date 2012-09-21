<%




Const DownLoad_URL="http://download.rainbowsoft.org/Plugins/ps.asp"
Const Resource_URL="http://download.rainbowsoft.org/Plugins/"    '注意. Include 文件里还有一同名变量要修改
Const Update_URL="http://download.rainbowsoft.org/Plugin/dlcs/download2.asp?plugin="

Const XML_Pack_Ver="1.0"
Const XML_Pack_Type="Plugin"
Const XML_Pack_Version="Z-Blog 2.0"




Sub SubMenu(id)
	Dim aryName,aryValue,aryPos
	aryName=Array("在线安装插件","在线安装主题","新建插件","新建主题")
	aryValue=Array("plugin_list.asp","theme_list.asp","plugin_edit.asp","theme_edit.asp")
	aryPos=Array("m-left","m-left","m-left","m-left")
	Dim i 
	For i=0 To Ubound(aryName)
		Response.Write MakeSubMenu(aryName(i),aryValue(i),aryPos(i) & IIf(id=i," m-now",""),False)
	Next
End Sub

%>
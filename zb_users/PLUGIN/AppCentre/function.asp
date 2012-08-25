<%


Dim App_ID,App_Name,App_URL,App_Note,App_PubDate
Dim App_Adapted,App_Version,App_Modified
Dim App_Type,App_Path,App_Include,App_Level
Dim App_Author_Name,App_Author_Url,App_Author_Email

Dim Action,SelectedPlugin,SelectedPluginName
Dim objXmlVerChk,NewVersionExists

Const DownLoad_URL="http://download.rainbowsoft.org/Plugins/ps.asp"
Const Resource_URL="http://download.rainbowsoft.org/Plugins/"    '注意. Include 文件里还有一同名变量要修改
Const Update_URL="http://download.rainbowsoft.org/Plugin/dlcs/download2.asp?plugin="

Const XML_Pack_Ver="1.0"
Const XML_Pack_Type="Plugin"
Const XML_Pack_Version="Z-Blog_1_8"




Sub SubMenu(id)
	Dim aryName,aryValue,aryPos
	aryName=Array("应用下载","主题管理","插件管理","升级应用")
	aryValue=Array("download.asp","theme.asp","plugin.asp","update.asp")
	aryPos=Array("m-left","m-left","m-left","m-left")
	Dim i 
	For i=0 To Ubound(aryName)
		Response.Write MakeSubMenu(aryName(i),aryValue(i),aryPos(i) & IIf(id=i," m-now",""),False)
	Next
End Sub


%>
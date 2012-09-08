<%
'注册插件
Call RegisterPlugin("CustomMeta","ActivePlugin_CustomMeta")

Function ActivePlugin_CustomMeta() 

	Call Add_Action_Plugin("Action_Plugin_Edit_Form","Call CustomMeta_AddLogEdit(EditArticle)")
	Call Add_Action_Plugin("Action_Plugin_EditCatalog_Form","Call CustomMeta_AddCateEdit(EditCategory)")
	Call Add_Action_Plugin("Action_Plugin_EditUser_Form","Call CustomMeta_AddUserEdit(EditUser)")

End Function



Function CustomMeta_AddLogEdit(obj)

	Dim c
	Set c=New TConfig
	c.Load "CustomMeta"

	Dim m,i,s
	Set m=New TMeta
	m.LoadString=c.Read("LogMeta")
	For i=LBound(m.Names)+1 To UBound(m.Names)
		s=s & "<div style='clear:both;width:100%'><p style='width:20%;float:left;'><span class='title'>"&m.GetValue(m.Names(i))&"字段:</span></p><p style='width:80%;float:left;'><input style='width:100%;' type='text' name='meta_"&m.Names(i)&"' value='"&obj.Meta.GetValue(m.Names(i))&"' /></p></div>"
	Next

	Call Add_Response_Plugin("Response_Plugin_Edit_Form2",s)

End Function

Function CustomMeta_AddCateEdit(obj)

	Dim c
	Set c=New TConfig
	c.Load "CustomMeta"

	Dim m,i,s
	Set m=New TMeta
	m.LoadString=c.Read("CateMeta")
	For i=LBound(m.Names)+1 To UBound(m.Names)
		s=s & "<p style=''><span class='title'>"&m.GetValue(m.Names(i))&"字段:</span><br/><input style='width:600px;' type='text' name='meta_"&m.Names(i)&"' value='"&obj.Meta.GetValue(m.Names(i))&"' /></p>"
	Next

	Call Add_Response_Plugin("Response_Plugin_EditCatalog_Form",s)

End Function


Function CustomMeta_AddUserEdit(obj)

	Dim c
	Set c=New TConfig
	c.Load "CustomMeta"

	Dim m,i,s
	Set m=New TMeta
	m.LoadString=c.Read("UserMeta")
	For i=LBound(m.Names)+1 To UBound(m.Names)
		s=s & "<p style=''><span class='title'>"&m.GetValue(m.Names(i))&"字段:</span><br/><input style='width:600px;' type='text' name='meta_"&m.Names(i)&"' value='"&obj.Meta.GetValue(m.Names(i))&"' /></p>"
	Next

	Call Add_Response_Plugin("Response_Plugin_EditUser_Form",s)

End Function

Function InstallPlugin_CustomMeta()


End Function


Function UnInstallPlugin_CustomMeta()


End Function

%>